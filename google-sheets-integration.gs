/**
 * REINVENTA by Mary Méndez — Apps Script
 *
 * HOJAS DEL SPREADSHEET:
 *
 * [Registros]
 *   A  Fecha registro       B  Nombre            C  Correo
 *   D  WhatsApp             E  Contacto preferido F  ¿Qué busca?
 *   G  Fase de compra       H  Monto pagado (MXN) I  Fecha de pago
 *   J  Pagó ✓              K  Stripe Payment ID  L  Origen
 *   M  Acepta marketing     N  Autoriza imagen    O  Canal UTM
 *   P  ¿Cómo se enteró?    Q  ¿A qué se dedica?
 *
 * [Comunicaciones]
 *   A  Correo               B  Nombre             C  WhatsApp
 *   D  Contacto preferido
 *   — Confirmación —
 *   E  WA Confirmación      F  Estado Conf. WA    G  Correo Conf. Enviado
 *   — Recordatorio —
 *   H  WA Recordatorio      I  Estado Record. WA  J  Correo Record. Enviado
 *   — QR de entrada —
 *   K  WA QR                L  Estado QR WA       M  Correo QR Enviado
 *   — Agradecimiento —
 *   N  WA Agradecimiento    O  Estado Agradec. WA P  Correo Agradec. Enviado
 *
 * [Asistencia]
 *   A  ID Único (RNV-001)   B  Nombre             C  Correo
 *   D  Fase                 E  Asistió ✓          F  Fecha entrada
 *
 * [Dashboard] — solo fórmulas, no la toca el script
 */

var SHEET_REGISTROS      = 'Registros';
var SHEET_COMUNICACIONES = 'Comunicaciones';
var SHEET_ASISTENCIA     = 'Asistencia';

var EMAILS_NOTIFICACION = ['mejoracontinua@caceca.org', 'alopez@alumbrastudios.com'];
var LIMITE_TOTAL        = 40;

/* ── Getters de hojas ────────────────────────────────────────── */
function getSheet()               { return getSheetByName(SHEET_REGISTROS); }
function getComunicacionesSheet() { return getSheetByName(SHEET_COMUNICACIONES); }
function getAsistenciaSheet()     { return getSheetByName(SHEET_ASISTENCIA); }

function getSheetByName(nombre) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(nombre);
  if (!sheet) throw new Error('Hoja no encontrada: ' + nombre);
  return sheet;
}

/* ── doGet ───────────────────────────────────────────────────── */
function doGet(e) {
  var action = e.parameter.action || '';
  if (action === 'entrada') return handleEntrada(e.parameter.id || '');
  if (action === 'hub')     return handleHub(e.parameter.id || '');
  if (action === 'admin')   return handleAdmin();
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ── doPost ──────────────────────────────────────────────────── */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if (data.type && data.type === 'checkout.session.completed') {
      return handleStripeWebhook(data);
    }
    if (data.action === 'encuesta')        return handleEncuesta(data);
    if (data.action === 'encuesta_previa') return handleEncuestaPrevia(data);
    return handleFormSubmit(data);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/* ── Popup de la landing → Registros ────────────────────────── */
function handleFormSubmit(data) {
  var sheet  = getSheet();
  var correo = (data.correo || '').toLowerCase().trim();
  var existingRow = findRowByEmail(sheet, correo);

  if (existingRow) {
    sheet.getRange(existingRow, 2).setValue(data.nombre      || sheet.getRange(existingRow, 2).getValue());
    sheet.getRange(existingRow, 4).setValue(data.whatsapp    || sheet.getRange(existingRow, 4).getValue());
    sheet.getRange(existingRow, 5).setValue(data.contacto    || sheet.getRange(existingRow, 5).getValue());
    sheet.getRange(existingRow, 6).setValue(data.transformar || sheet.getRange(existingRow, 6).getValue());
    sheet.getRange(existingRow, 12).setValue('landing + stripe');
    sheet.getRange(existingRow, 13).setValue(data.marketing       || '');
    sheet.getRange(existingRow, 14).setValue(data.autoriza_imagen || '');
    if (data.canal_utm)      sheet.getRange(existingRow, 15).setValue(data.canal_utm);
    if (data.como_se_entero) sheet.getRange(existingRow, 16).setValue(data.como_se_entero);
    if (data.ocupacion)      sheet.getRange(existingRow, 17).setValue(data.ocupacion);
  } else {
    sheet.appendRow([
      new Date(), data.nombre || '', correo,
      data.whatsapp    || '', data.contacto   || '', data.transformar || '',
      data.servicio    || '', '', '', '',
      '', data.origen  || 'landing',
      data.marketing       || '', data.autoriza_imagen || '',
      data.canal_utm       || '', data.como_se_entero  || '', data.ocupacion || ''
    ]);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ result: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ── Webhook de Stripe ───────────────────────────────────────── */
function handleStripeWebhook(event) {
  var session  = event.data.object;
  var correo   = ((session.customer_details && session.customer_details.email)
                   ? session.customer_details.email
                   : (session.customer_email || '')).toLowerCase().trim();
  var nombre   = session.customer_details && session.customer_details.name
                   ? session.customer_details.name : '';
  var monto    = session.amount_total ? (session.amount_total / 100).toFixed(2) : '';
  var fecha    = session.created ? new Date(session.created * 1000) : new Date();
  var stripeId = session.id || '';

  var fase = 'Taller';
  if      (monto == '1300.00') fase = 'Early Bird';
  else if (monto == '1500.00') fase = 'Preventa';
  else if (monto == '1700.00') fase = 'Últimos lugares';

  var sheet = getSheet();

  if (stripeId && findRowByStripeId(sheet, stripeId)) {
    return ContentService.createTextOutput(JSON.stringify({ result: 'duplicate' })).setMimeType(ContentService.MimeType.JSON);
  }

  var existingRow = findRowByEmail(sheet, correo);
  var tel         = '';

  if (existingRow) {
    var yaPago = sheet.getRange(existingRow, 10).getValue() === '✓';
    tel = sheet.getRange(existingRow, 4).getValue();
    var contacto = sheet.getRange(existingRow, 5).getValue();
    if (!yaPago) {
      var nombreFinal = nombre || sheet.getRange(existingRow, 2).getValue();
      if (nombre) sheet.getRange(existingRow, 2).setValue(nombre);
      sheet.getRange(existingRow, 7).setValue(fase);
      sheet.getRange(existingRow, 8).setValue(monto);
      sheet.getRange(existingRow, 9).setValue(fecha);
      sheet.getRange(existingRow, 10).setValue('✓');
      sheet.getRange(existingRow, 11).setValue(stripeId);
      sheet.getRange(existingRow, 12).setValue('landing + stripe');
      actualizarAsistencia(correo, nombreFinal, fase);
      var idNuevo = obtenerIdAsistente(correo);
      sincronizarComunicaciones(correo, nombreFinal, tel, contacto, idNuevo);
    }
  } else {
    sheet.appendRow([
      new Date(), nombre, correo, '', '', '', fase, monto, fecha, '✓', stripeId, 'stripe directo',
      '', '', '', '', ''
    ]);
    actualizarAsistencia(correo, nombre, fase);
    var idNuevo = obtenerIdAsistente(correo);
    sincronizarComunicaciones(correo, nombre, '', '', idNuevo);
  }

  var totalPagos = contarPagosSheet(sheet);
  if (totalPagos === LIMITE_TOTAL) {
    var props = PropertiesService.getScriptProperties();
    if (!props.getProperty('sold_out_enviado')) {
      notificarCupoAgotado(sheet);
      props.setProperty('sold_out_enviado', 'true');
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ result: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ── Sincronizar fila en Comunicaciones ──────────────────────── */
function sincronizarComunicaciones(correo, nombre, telefono, contacto, idUnico) {
  var sheet = getComunicacionesSheet();
  var fila  = findRowByEmailInSheet(sheet, correo);

  if (!fila) {
    sheet.appendRow([correo, nombre, telefono, contacto, '', '', '', '', '', '', '', '', '', '', '', '']);
    fila = sheet.getLastRow();
  } else {
    if (!sheet.getRange(fila, 3).getValue() && telefono) sheet.getRange(fila, 3).setValue(telefono);
    if (!sheet.getRange(fila, 4).getValue() && contacto) sheet.getRange(fila, 4).setValue(contacto);
  }

  // Generar botón WA confirmación si no existe
  var yaLink = sheet.getRange(fila, 5).getValue();
  if (!yaLink) {
    var id = idUnico || '';
    generarBotonWA(sheet, fila, nombre, telefono,
      function(n, t) { return generateWhatsAppLinkConfirmacion(n, t, id); },
      5, 6, 'Enviar WhatsApp');
  }
}

/* ── Asistencia ──────────────────────────────────────────────── */
function actualizarAsistencia(correo, nombre, fase) {
  var sheet = getAsistenciaSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][2] || '').toLowerCase().trim() === correo.toLowerCase().trim()) return;
  }
  var id = generarSiguienteId(sheet);
  sheet.appendRow([id, nombre, correo, fase, '', '', '', '', '', '']);
}

function generarSiguienteId(sheet) {
  var data     = sheet.getDataRange().getValues();
  var contador = 1;
  for (var i = 1; i < data.length; i++) {
    var id = (data[i][0] || '').toString();
    if (id.startsWith('RNV-')) {
      var num = parseInt(id.replace('RNV-', ''), 10);
      if (num >= contador) contador = num + 1;
    }
  }
  return 'RNV-' + String(contador).padStart(3, '0');
}

/* ── Hub: perfil del asistente ───────────────────────────────── */
function handleHub(id) {
  if (!id) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'ID requerido' })).setMimeType(ContentService.MimeType.JSON);
  }
  var sheet = getAsistenciaSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString().trim() === id) {
      // Verificar encuesta previa en hoja separada
      var ss        = SpreadsheetApp.getActiveSpreadsheet();
      var prevSheet = ss.getSheetByName('Encuesta Previa');
      var encPrev   = false;
      if (prevSheet) {
        var prevData = prevSheet.getDataRange().getValues();
        for (var j = 1; j < prevData.length; j++) {
          if ((prevData[j][0] || '').toString().trim() === id) { encPrev = true; break; }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({
        id:             data[i][0],
        nombre:         data[i][1],
        correo:         data[i][2],
        fase:           data[i][3],
        asistio:        data[i][4] === '✓',
        fechaEntrada:   data[i][5] ? data[i][5].toString() : '',
        encuesta:       data[i][6] === '✓',
        calificacion:   data[i][8] || 0,
        encuestaPrevia: encPrev
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ error: 'Registro no encontrado' })).setMimeType(ContentService.MimeType.JSON);
}

/* ── Admin: todos los datos del panel ───────────────────────── */
function handleAdmin() {
  var asiSheet  = getAsistenciaSheet();
  var asiData   = asiSheet.getDataRange().getValues();
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var prevSheet = ss.getSheetByName('Encuesta Previa');

  var asistentes = [];
  for (var i = 1; i < asiData.length; i++) {
    var row = asiData[i];
    if (!row[0]) continue;
    asistentes.push({
      id:             row[0],
      nombre:         row[1],
      correo:         row[2],
      fase:           row[3],
      asistio:        row[4] === '✓',
      fechaEntrada:   row[5] ? row[5].toString() : '',
      encuesta:       row[6] === '✓',
      fechaEncuesta:  row[7] ? row[7].toString() : '',
      calificacion:   row[8] || 0,
      comentario:     row[9] || ''
    });
  }

  var encuestasPrevia = [];
  if (prevSheet) {
    var prevData = prevSheet.getDataRange().getValues();
    for (var j = 1; j < prevData.length; j++) {
      var p = prevData[j];
      if (!p[0]) continue;
      encuestasPrevia.push({
        id:           p[0],
        nombre:       p[1],
        correo:       p[2],
        satisfaccion: p[4],
        coherencia:   p[5],
        confianza:    p[6],
        proyeccion:   p[7],
        motivacion:   p[8],
        expectativa:  p[9],
        piel:         p[10],
        cabello:      p[11],
        ojos:         p[12]
      });
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ asistentes: asistentes, encuestasPrevia: encuestasPrevia }))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ── Hub: encuesta previa al evento (imagen y propósito) ────── */
function handleEncuestaPrevia(data) {
  if (!data.id) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'ID requerido' })).setMimeType(ContentService.MimeType.JSON);
  }

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Encuesta Previa');
  if (!sheet) {
    sheet = ss.insertSheet('Encuesta Previa');
    sheet.getRange(1, 1, 1, 14).setValues([[
      'ID','Nombre','Correo','Fecha',
      'Satisfacción (1-10)','Coherencia (1-10)','Confianza (1-10)','Proyección (1-10)',
      'Motivación principal','Expectativa',
      'Tono de piel','Color de cabello','Color de ojos','Ya respondió'
    ]]);
    sheet.getRange(1, 1, 1, 14).setFontWeight('bold').setBackground('#2A0F25').setFontColor('#C6A56A');
  }

  // Buscar si ya respondió
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if ((rows[i][0] || '').toString().trim() === data.id) {
      return ContentService.createTextOutput(JSON.stringify({ result: 'ya_enviada' })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Buscar nombre y correo en Asistencia
  var asiSheet = getAsistenciaSheet();
  var asiData  = asiSheet.getDataRange().getValues();
  var nombre = '', correo = '';
  for (var j = 1; j < asiData.length; j++) {
    if ((asiData[j][0] || '').toString().trim() === data.id) {
      nombre = asiData[j][1];
      correo = asiData[j][2];
      break;
    }
  }

  sheet.appendRow([
    data.id, nombre, correo, new Date(),
    data.satisfaccion  || '',
    data.coherencia    || '',
    data.confianza     || '',
    data.proyeccion    || '',
    data.motivacion    || '',
    data.expectativa   || '',
    data.piel          || '',
    data.cabello       || '',
    data.ojos          || '',
    '✓'
  ]);

  return ContentService.createTextOutput(JSON.stringify({ result: 'success' })).setMimeType(ContentService.MimeType.JSON);
}

/* ── Hub: guardar encuesta ───────────────────────────────────── */
function handleEncuesta(data) {
  if (!data.id) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'ID requerido' })).setMimeType(ContentService.MimeType.JSON);
  }
  var sheet = getAsistenciaSheet();
  var rows  = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if ((rows[i][0] || '').toString().trim() === data.id) {
      if (rows[i][4] !== '✓') {
        return ContentService.createTextOutput(JSON.stringify({ error: 'Check-in pendiente' })).setMimeType(ContentService.MimeType.JSON);
      }
      if (rows[i][6] === '✓') {
        return ContentService.createTextOutput(JSON.stringify({ result: 'ya_enviada' })).setMimeType(ContentService.MimeType.JSON);
      }
      sheet.getRange(i + 1, 7).setValue('✓');
      sheet.getRange(i + 1, 8).setValue(new Date());
      sheet.getRange(i + 1, 9).setValue(data.calificacion || '');
      sheet.getRange(i + 1, 10).setValue(data.comentario  || '');
      return ContentService.createTextOutput(JSON.stringify({ result: 'success' })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ error: 'ID no encontrado' })).setMimeType(ContentService.MimeType.JSON);
}

/* ── Registro de entrada (QR) ────────────────────────────────── */
function handleEntrada(id) {
  if (!id) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'ID requerido' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var sheet = getAsistenciaSheet();
  var data  = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString().trim() === id) {
      var nombre    = data[i][1];
      var fase      = data[i][3];
      var yaAsistio = data[i][4];

      if (yaAsistio === '✓') {
        return ContentService
          .createTextOutput(JSON.stringify({ yaRegistrado: true, nombre: nombre }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      sheet.getRange(i + 1, 5).setValue('✓');
      sheet.getRange(i + 1, 6).setValue(new Date());

      return ContentService
        .createTextOutput(JSON.stringify({ nombre: nombre, fase: fase }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ error: 'No se encontró este registro' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ── WhatsApp: generador genérico ────────────────────────────── */
function normalizeWhatsAppNumber(raw) {
  if (!raw) return null;
  var digits = raw.toString().replace(/\D/g, '');
  if (digits.length === 10) return '52' + digits;
  if (digits.length === 12 && digits.slice(0,2) === '52') return digits;
  if (digits.length === 13 && digits.slice(0,3) === '521') return '52' + digits.slice(3);
  if (digits.length === 11 && digits[0] === '1') return digits;
  if (digits.length >= 10 && digits.length <= 15) return digits;
  return null;
}

function generarBotonWA(sheet, fila, nombre, telefono, generadorFn, colLink, colEstado, textoBoton) {
  var link = generadorFn(nombre, telefono);
  if (link) {
    var richText = SpreadsheetApp.newRichTextValue()
      .setText(textoBoton)
      .setLinkUrl(link)
      .build();
    sheet.getRange(fila, colLink).setRichTextValue(richText);
    if (!sheet.getRange(fila, colEstado).getValue()) {
      sheet.getRange(fila, colEstado).setValue('Pendiente');
    }
  } else {
    sheet.getRange(fila, colLink).setValue('SIN TELEFONO');
    sheet.getRange(fila, colEstado).setValue('SIN TELEFONO');
  }
}

/* ── WhatsApp: mensajes por campaña ─────────────────────────── */
function generateWhatsAppLinkConfirmacion(nombre, telefono, idUnico) {
  var numero = normalizeWhatsAppNumber(telefono);
  if (!numero) return null;
  var p      = nombre ? nombre.trim().split(' ')[0] : 'participante';
  var e      = encodeURIComponent;
  var NL     = '%0A';
  var hubUrl = idUnico
    ? 'https://reinventabymarymendez.com.mx/hub?id=' + idUnico
    : 'https://reinventabymarymendez.com.mx/hub';
  var msg =
    e('*REINVENTA by Mary Méndez*') + NL + NL +
    e('Hola, ' + p + '. Tu lugar está confirmado. 🤍') + NL + NL +
    e('*Lo que tu imagen comunica*') + NL +
    e('- Sábado 15 de agosto') + NL +
    e('- 10:00 a 12:00 am') + NL +
    e('- The University Club of Mexico') + NL +
    e('- Av. Reforma 150, Juárez, CDMX') + NL + NL +
    e('*Cómo llegar:*') + NL +
    e('https://maps.app.goo.gl/Uo7tYiQz23jMCmKw7') + NL + NL +
    e('- - - - - - - - - - - - -') + NL + NL +
    e('*Tu espacio personal del evento:*') + NL +
    e(hubUrl) + NL + NL +
    e('Aquí encontrarás tu pase de entrada con código QR, la agenda del día y los recursos del taller.') + NL + NL +
    e('⭐ *Una cosa importante:*') + NL +
    e('Dentro de tu espacio hay una encuesta breve que te pedimos contestar _antes del evento_. Mary la revisa personalmente para preparar materiales y recomendaciones a la medida de cada asistente. Entre más detallada seas, más personalizada será tu experiencia ese día.') + NL + NL +
    e('No toma más de 5 minutos y hace una gran diferencia. 🙏') + NL + NL +
    e('- - - - - - - - - - - - -') + NL + NL +
    e('Nos da mucho gusto tenerte. Mary estará encantada de acompañarte.') + NL + NL +
    e('_Con cariño,_') + NL +
    e('_Reinventa by Mary Méndez_') + NL + NL +
    e('_Este es un mensaje informativo, por favor no respondas a este chat._') + NL + NL +
    e('- - - - - - - - - - - - -') + NL +
    e('_Organizado integralmente por_') + NL +
    e('*Alumbra Studios*') + NL +
    e('https://www.alumbrastudios.com');
  return 'https://wa.me/' + numero + '?text=' + msg;
}

function generateWhatsAppLinkRecordatorio(nombre, telefono) {
  var numero = normalizeWhatsAppNumber(telefono);
  if (!numero) return null;
  var p  = nombre ? nombre.trim().split(' ')[0] : 'participante';
  var e  = encodeURIComponent;
  var NL = '%0A';
  var msg =
    e('*REINVENTA by Mary Mendez*') + NL + NL +
    e('Hola, ' + p + '. Mañana es el gran día.') + NL + NL +
    e('*Lo que tu imagen comunica*') + NL +
    e('- Mañana sábado 15 de agosto') + NL +
    e('- 10:00 a 12:00 am') + NL +
    e('- The University Club of Mexico') + NL +
    e('- Av. Reforma 150, Juárez, CDMX') + NL + NL +
    e('*Cómo llegar:*') + NL +
    e('https://maps.app.goo.gl/Uo7tYiQz23jMCmKw7') + NL + NL +
    e('Te esperamos puntual. Mary tiene algo muy especial preparado para ti.') + NL + NL +
    e('_Con cariño,_') + NL +
    e('_Reinventa by Mary Mendez_') + NL + NL +
    e('_Este es un mensaje informativo, por favor no respondas a este chat._') + NL + NL +
    e('- - - - - - - - - - - - -') + NL +
    e('_Organizado integralmente por_') + NL +
    e('*Alumbra Studios*') + NL +
    e('https://www.alumbrastudios.com');
  return 'https://wa.me/' + numero + '?text=' + msg;
}

function generateWhatsAppLinkQR(nombre, telefono, idUnico) {
  var numero = normalizeWhatsAppNumber(telefono);
  if (!numero) return null;
  var p  = nombre ? nombre.trim().split(' ')[0] : 'participante';
  var e  = encodeURIComponent;
  var NL = '%0A';
  var urlEntrada = 'https://reinventabymarymendez.com.mx/entrada?id=' + idUnico;
  var msg =
    e('*REINVENTA by Mary Mendez*') + NL + NL +
    e('Hola, ' + p + '. Aquí está tu acceso para el taller.') + NL + NL +
    e('*Tu código de entrada:*') + NL +
    e(urlEntrada) + NL + NL +
    e('Muestra este link en la entrada el día del evento.') + NL + NL +
    e('*Lo que tu imagen comunica*') + NL +
    e('- Sábado 15 de agosto') + NL +
    e('- 10:00 a 12:00 am') + NL +
    e('- The University Club of Mexico') + NL +
    e('- Av. Reforma 150, Juárez, CDMX') + NL + NL +
    e('Nos vemos pronto.') + NL + NL +
    e('_Con cariño,_') + NL +
    e('_Reinventa by Mary Mendez_') + NL + NL +
    e('_Este es un mensaje informativo, por favor no respondas a este chat._') + NL + NL +
    e('- - - - - - - - - - - - -') + NL +
    e('_Organizado integralmente por_') + NL +
    e('*Alumbra Studios*') + NL +
    e('https://www.alumbrastudios.com');
  return 'https://wa.me/' + numero + '?text=' + msg;
}

function generateWhatsAppLinkAgradecimiento(nombre, telefono) {
  var numero = normalizeWhatsAppNumber(telefono);
  if (!numero) return null;
  var p  = nombre ? nombre.trim().split(' ')[0] : 'participante';
  var e  = encodeURIComponent;
  var NL = '%0A';
  var msg =
    e('*REINVENTA by Mary Mendez*') + NL + NL +
    e('Hola, ' + p + '. Gracias por acompañarnos hoy.') + NL + NL +
    e('Fue un honor compartir este espacio contigo. Esperamos que lo que viviste hoy te acompañe mucho tiempo.') + NL + NL +
    e('*Conoce más sobre Mary:*') + NL +
    e('https://reinventabymarymendez.com.mx') + NL + NL +
    e('_Con cariño,_') + NL +
    e('_Reinventa by Mary Mendez_') + NL + NL +
    e('_Este es un mensaje informativo, por favor no respondas a este chat._') + NL + NL +
    e('- - - - - - - - - - - - -') + NL +
    e('_Organizado integralmente por_') + NL +
    e('*Alumbra Studios*') + NL +
    e('https://www.alumbrastudios.com');
  return 'https://wa.me/' + numero + '?text=' + msg;
}

/* ── Generación masiva de botones WA ────────────────────────── */
function generarBotonesConfirmacionMasivo() {
  _generarBotonesMasivo(generateWhatsAppLinkConfirmacion, 5, 6, 'Enviar confirmación', false);
}

function generarBotonesRecordatorioMasivo() {
  _generarBotonesMasivo(generateWhatsAppLinkRecordatorio, 8, 9, 'Enviar recordatorio', false);
}

function generarBotonesAgradecimientoMasivo() {
  _generarBotonesMasivo(generateWhatsAppLinkAgradecimiento, 14, 15, 'Enviar agradecimiento', false);
}

function _generarBotonesMasivo(generadorFn, colLink, colEstado, textoBoton, soloSinLink) {
  var regSheet  = getSheet();
  var comSheet  = getComunicacionesSheet();
  var data      = regSheet.getDataRange().getValues();
  var generados = 0;

  for (var i = 1; i < data.length; i++) {
    if (data[i][9] !== '✓') continue;
    var correo = (data[i][2] || '').toLowerCase().trim();
    var nombre = data[i][1];
    var tel    = data[i][3];
    var fila   = findRowByEmailInSheet(comSheet, correo);
    if (!fila) continue;
    var yaLink = comSheet.getRange(fila, colLink).getValue();
    if (soloSinLink && yaLink) continue;
    generarBotonWA(comSheet, fila, nombre, tel, generadorFn, colLink, colEstado, textoBoton);
    generados++;
  }
  Logger.log('Botones generados (' + textoBoton + '): ' + generados);
}

function generarBotonesQRMasivo() {
  var regSheet  = getSheet();
  var comSheet  = getComunicacionesSheet();
  var asiSheet  = getAsistenciaSheet();
  var regData   = regSheet.getDataRange().getValues();
  var asiData   = asiSheet.getDataRange().getValues();
  var generados = 0;

  var idPorCorreo = {};
  for (var j = 1; j < asiData.length; j++) {
    var c = (asiData[j][2] || '').toLowerCase().trim();
    if (c) idPorCorreo[c] = asiData[j][0];
  }

  for (var i = 1; i < regData.length; i++) {
    if (regData[i][9] !== '✓') continue;
    var correo = (regData[i][2] || '').toLowerCase().trim();
    var nombre = regData[i][1];
    var tel    = regData[i][3];
    var id     = idPorCorreo[correo];
    if (!id) continue;
    var fila = findRowByEmailInSheet(comSheet, correo);
    if (!fila) continue;

    var link = generateWhatsAppLinkQR(nombre, tel, id);
    if (link) {
      var richText = SpreadsheetApp.newRichTextValue()
        .setText('Enviar QR')
        .setLinkUrl(link)
        .build();
      comSheet.getRange(fila, 11).setRichTextValue(richText);
      if (!comSheet.getRange(fila, 12).getValue()) comSheet.getRange(fila, 12).setValue('Pendiente');
    } else {
      comSheet.getRange(fila, 11).setValue('SIN TELEFONO');
      comSheet.getRange(fila, 12).setValue('SIN TELEFONO');
    }
    generados++;
  }
  Logger.log('Botones QR generados: ' + generados);
}

/* ── Correos: plantillas ─────────────────────────────────────── */
function _headerCorreo() {
  return '<div style="background:#E8E2DB;padding:2rem 1rem;font-family:\'Gill Sans\',Calibri,\'Segoe UI\',sans-serif;">'
    + '<div style="max-width:540px;margin:0 auto;background:#EFE9E2;box-shadow:0 4px 40px rgba(42,15,37,.13);">'
    + '<div style="background:#2A0F25;padding:2rem 2.4rem 1.6rem;text-align:center;">'
    + '<span style="font-family:Georgia,serif;font-weight:400;font-size:1rem;letter-spacing:.22em;text-transform:uppercase;color:#C6A56A;display:block;margin-bottom:.2rem;">Reinventa</span>'
    + '<span style="font-size:.62rem;letter-spacing:.16em;text-transform:uppercase;color:rgba(198,165,106,.5);">by Mary Méndez</span>'
    + '</div>'
    + '<div style="height:2px;background:#C6A56A;opacity:.45;"></div>';
}

function _footerCorreo(correo) {
  return '<div style="background:#2A0F25;padding:1.1rem 2rem;text-align:center;">'
    + '<p style="font-size:.63rem;letter-spacing:.07em;color:rgba(198,165,106,.5);line-height:1.7;margin:0 0 .5rem;">REINVENTA by Mary Méndez &middot; Ciudad de México<br>'
    + 'Este correo fue enviado a ' + correo + '.</p>'
    + '<p style="font-size:.6rem;color:rgba(198,165,106,.3);margin:0;">Evento organizado integralmente por <strong style="color:rgba(198,165,106,.5);">Alumbra Studios</strong> &middot; <a href="https://www.alumbrastudios.com" style="color:rgba(198,165,106,.4);text-decoration:none;">alumbrastudios.com</a></p>'
    + '</div></div></div>';
}

function _firmaCorreo() {
  return '<hr style="border:none;border-top:1px solid rgba(42,15,37,.1);margin:0 0 1.4rem;" />'
    + '<p style="font-family:Georgia,serif;font-size:.98rem;color:#2A0F25;margin:.8rem 0 .1rem;">Mary Méndez</p>'
    + '<p style="font-size:.72rem;color:#8F7383;margin:0;">Consultora de imagen y liderazgo</p>';
}

function _detallesEvento() {
  return '<div style="border:1px solid rgba(42,15,37,.12);padding:1.4rem 1.6rem;margin-bottom:1.6rem;">'
    + '<p style="font-size:.6rem;letter-spacing:.16em;text-transform:uppercase;color:#8F7383;margin:0 0 .8rem;">Detalles del evento</p>'
    + '<p style="font-family:Georgia,serif;font-size:1.05rem;font-weight:400;color:#2A0F25;margin:0 0 .1rem;">Taller de imagen y liderazgo</p>'
    + '<p style="font-size:.78rem;color:#8F7383;font-style:italic;margin:0 0 1rem;">Lo que tu imagen comunica</p>'
    + '<p style="font-size:.85rem;color:#4a3545;margin:0 0 .5rem;">📅 &nbsp;Sábado 15 de agosto de 2026 &middot; 10:00–12:00 pm</p>'
    + '<p style="font-size:.85rem;color:#4a3545;margin:0;">📍 &nbsp;The University Club of Mexico<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#8F7383;font-size:.8rem;">Av. Paseo de la Reforma 150, Juárez, CDMX</span></p>'
    + '</div>';
}

function enviarCorreoConfirmacion(nombre, correo, fase, idUnico) {
  var p = nombre ? nombre.split(' ')[0] : 'Hola';
  var calLink = 'https://calendar.google.com/calendar/render?action=TEMPLATE'
    + '&text=Taller+de+imagen+y+liderazgo+%E2%80%94+REINVENTA'
    + '&dates=20260815T160000Z/20260815T180000Z'
    + '&details=Taller+Lo+que+tu+imagen+comunica+%7C+REINVENTA+by+Mary+M%C3%A9ndez'
    + '&location=The+University+Club+of+Mexico%2C+Av.+Paseo+de+la+Reforma+150%2C+Ju%C3%A1rez%2C+CDMX';
  var hubUrl = idUnico
    ? 'https://reinventabymarymendez.com.mx/hub?id=' + idUnico
    : 'https://reinventabymarymendez.com.mx/hub';

  var html = _headerCorreo()
    + '<div style="padding:2.2rem 2.6rem 2rem;">'
    + '<div style="display:inline-block;background:rgba(42,15,37,.07);border-left:2px solid #C6A56A;padding:.4rem .75rem;font-size:.65rem;letter-spacing:.13em;text-transform:uppercase;color:#2A0F25;margin-bottom:1.4rem;">Pago confirmado &middot; ' + fase + '</div>'
    + '<h1 style="font-family:Georgia,serif;font-weight:400;font-size:1.5rem;line-height:1.35;color:#2A0F25;margin:0 0 1rem;">Tu lugar está reservado,<br>' + p + '.</h1>'
    + '<p style="font-size:.92rem;line-height:1.7;color:#4a3545;margin:0 0 1.8rem;">Nos da mucho gusto tenerte en el taller. Mary estará encantada de acompañarte en este proceso. Guarda la fecha en tu calendario para que no se te pase ningún detalle.</p>'
    + _detallesEvento()
    + '<a href="' + calLink + '" style="display:block;background:#2A0F25;color:#EFE9E2;text-align:center;padding:.9rem 1.2rem;text-decoration:none;font-size:.8rem;letter-spacing:.1em;text-transform:uppercase;margin-bottom:1.6rem;">📅 &nbsp;Agregar a Google Calendar</a>'
    + '<div style="border:1px solid rgba(42,15,37,.12);padding:1.4rem 1.6rem;margin-bottom:1.6rem;">'
    + '<p style="font-size:.6rem;letter-spacing:.16em;text-transform:uppercase;color:#8F7383;margin:0 0 .6rem;">Tu espacio personal del evento</p>'
    + '<p style="font-size:.87rem;color:#4a3545;line-height:1.6;margin:0 0 .8rem;">Aquí encontrarás tu pase de entrada con código QR, la agenda del día y los recursos del taller.</p>'
    + '<a href="' + hubUrl + '" style="display:block;background:#C6A56A;color:#2A0F25;text-align:center;padding:.8rem 1.2rem;text-decoration:none;font-size:.8rem;letter-spacing:.1em;text-transform:uppercase;margin-bottom:1rem;font-weight:600;">Acceder a mi espacio →</a>'
    + '<div style="border-top:1px solid rgba(42,15,37,.1);padding-top:.9rem;">'
    + '<p style="font-size:.78rem;color:#2A0F25;font-weight:600;margin:0 0 .3rem;">⭐ Una cosa importante</p>'
    + '<p style="font-size:.8rem;color:#4a3545;line-height:1.6;margin:0;">Dentro de tu espacio hay una encuesta breve que te pedimos contestar <strong>antes del evento</strong>. Mary la revisa personalmente para preparar materiales y recomendaciones a la medida de cada asistente. No toma más de 5 minutos y hace una gran diferencia.</p>'
    + '</div></div>'
    + _firmaCorreo()
    + '</div>'
    + _footerCorreo(correo);

  MailApp.sendEmail({ to: correo, bcc: 'alopez@alumbrastudios.com', name: 'Reinventa by Mary Méndez',
    subject: 'Tu lugar en el taller está confirmado ✦ REINVENTA', htmlBody: html });
}

function enviarCorreoRecordatorio(nombre, correo) {
  var p = nombre ? nombre.split(' ')[0] : 'Hola';

  var html = _headerCorreo()
    + '<div style="padding:2.2rem 2.6rem 2rem;">'
    + '<div style="display:inline-block;background:rgba(42,15,37,.07);border-left:2px solid #C6A56A;padding:.4rem .75rem;font-size:.65rem;letter-spacing:.13em;text-transform:uppercase;color:#2A0F25;margin-bottom:1.4rem;">Recordatorio &middot; Mañana es el taller</div>'
    + '<h1 style="font-family:Georgia,serif;font-weight:400;font-size:1.5rem;line-height:1.35;color:#2A0F25;margin:0 0 1rem;">' + p + ',<br>mañana es el gran día.</h1>'
    + '<p style="font-size:.92rem;line-height:1.7;color:#4a3545;margin:0 0 1.8rem;">Mary tiene algo muy especial preparado para ti. Te esperamos puntual y con muchas ganas de transformar la manera en que tu imagen comunica quién eres.</p>'
    + _detallesEvento()
    + '<p style="font-size:.85rem;color:#4a3545;margin:-1rem 0 1.6rem;padding:0 1.6rem;"><a href="https://maps.app.goo.gl/Uo7tYiQz23jMCmKw7" style="color:#C6A56A;text-decoration:none;">Ver en Google Maps &rarr;</a></p>'
    + '<p style="font-size:.87rem;line-height:1.7;color:#4a3545;margin:0 0 .5rem;">Si tienes alguna duda de último momento no dudes en contactarnos. ¡Nos vemos mañana!</p>'
    + _firmaCorreo()
    + '</div>'
    + _footerCorreo(correo);

  MailApp.sendEmail({ to: correo, bcc: 'alopez@alumbrastudios.com', name: 'Reinventa by Mary Méndez',
    subject: 'Mañana te esperamos ✦ REINVENTA', htmlBody: html });
}

function enviarCorreoQR(nombre, correo, idUnico) {
  var p          = nombre ? nombre.split(' ')[0] : 'Hola';
  var urlEntrada = 'https://reinventabymarymendez.com.mx/entrada?id=' + idUnico;
  var urlQR      = 'https://api.qrserver.com/v1/create-qr-code/?size=200x200&data=' + encodeURIComponent(urlEntrada);

  var html = _headerCorreo()
    + '<div style="padding:2.2rem 2.6rem 2rem;text-align:center;">'
    + '<div style="display:inline-block;background:rgba(42,15,37,.07);border-left:2px solid #C6A56A;padding:.4rem .75rem;font-size:.65rem;letter-spacing:.13em;text-transform:uppercase;color:#2A0F25;margin-bottom:1.4rem;">Tu acceso al evento</div>'
    + '<h1 style="font-family:Georgia,serif;font-weight:400;font-size:1.5rem;line-height:1.35;color:#2A0F25;margin:0 0 .6rem;text-align:left;">' + p + ',<br>aquí está tu entrada.</h1>'
    + '<p style="font-size:.92rem;line-height:1.7;color:#4a3545;margin:0 0 1.8rem;text-align:left;">Presenta este código QR en la entrada el día del evento. Puedes mostrarlo desde tu teléfono.</p>'
    + '<div style="background:#2A0F25;display:inline-block;padding:1rem;margin-bottom:1rem;">'
    + '<img src="' + urlQR + '" width="180" height="180" style="display:block;" alt="Código QR de entrada" /></div>'
    + '<p style="font-size:.72rem;color:#8F7383;margin:0 0 1.6rem;">Si la imagen no carga, usa este enlace:<br><a href="' + urlEntrada + '" style="color:#C6A56A;">' + urlEntrada + '</a></p>'
    + _detallesEvento()
    + '<p style="font-size:.87rem;line-height:1.7;color:#4a3545;margin:0 0 .5rem;text-align:left;">¡Nos vemos el 15 de agosto. Será un día increíble!</p>'
    + _firmaCorreo()
    + '</div>'
    + _footerCorreo(correo);

  MailApp.sendEmail({ to: correo, bcc: 'alopez@alumbrastudios.com', name: 'Reinventa by Mary Méndez',
    subject: 'Tu código de entrada ✦ REINVENTA', htmlBody: html });
}

function enviarCorreoAgradecimiento(nombre, correo) {
  var p = nombre ? nombre.split(' ')[0] : 'Hola';

  var html = _headerCorreo()
    + '<div style="padding:2.2rem 2.6rem 2rem;">'
    + '<div style="display:inline-block;background:rgba(42,15,37,.07);border-left:2px solid #C6A56A;padding:.4rem .75rem;font-size:.65rem;letter-spacing:.13em;text-transform:uppercase;color:#2A0F25;margin-bottom:1.4rem;">Gracias por estar aquí</div>'
    + '<h1 style="font-family:Georgia,serif;font-weight:400;font-size:1.5rem;line-height:1.35;color:#2A0F25;margin:0 0 1rem;">' + p + ',<br>fue un honor acompañarte.</h1>'
    + '<p style="font-size:.92rem;line-height:1.7;color:#4a3545;margin:0 0 1.2rem;">Gracias por confiar en este espacio y por abrirte a transformar la manera en que tu imagen comunica quién eres. Lo que viviste hoy es solo el comienzo.</p>'
    + '<p style="font-size:.92rem;line-height:1.7;color:#4a3545;margin:0 0 1.8rem;">Mary estará siempre disponible para seguir acompañándote en este camino.</p>'
    + '<div style="border:1px solid rgba(42,15,37,.12);padding:1.4rem 1.6rem;margin-bottom:1.6rem;">'
    + '<p style="font-size:.6rem;letter-spacing:.16em;text-transform:uppercase;color:#8F7383;margin:0 0 .8rem;">Conoce más sobre Mary</p>'
    + '<p style="font-size:.87rem;color:#4a3545;margin:0;"><a href="https://reinventabymarymendez.com.mx" style="color:#C6A56A;text-decoration:none;">reinventabymarymendez.com.mx &rarr;</a></p>'
    + '</div>'
    + _firmaCorreo()
    + '</div>'
    + _footerCorreo(correo);

  MailApp.sendEmail({ to: correo, bcc: 'alopez@alumbrastudios.com', name: 'Reinventa by Mary Méndez',
    subject: 'Gracias por acompañarnos ✦ REINVENTA', htmlBody: html });
}

/* ── Envíos masivos de correo ────────────────────────────────── */
function enviarConfirmacionExistentes() {
  _enviarCorreosMasivo('enviarCorreoConfirmacion', 7, 'Correos confirmación');
}

function enviarCorreosRecordatorioMasivo() {
  _enviarCorreosMasivo('enviarCorreoRecordatorio', 10, 'Correos recordatorio');
}

function enviarCorreosAgradecimientoMasivo() {
  _enviarCorreosMasivo('enviarCorreoAgradecimiento', 16, 'Correos agradecimiento');
}

function _enviarCorreosMasivo(fnNombre, colEnviado, label) {
  var regSheet = getSheet();
  var comSheet = getComunicacionesSheet();
  var data     = regSheet.getDataRange().getValues();
  var enviados = 0;
  var vistos   = {};

  for (var i = 1; i < data.length; i++) {
    if (data[i][9] !== '✓') continue;
    var correo = (data[i][2] || '').toLowerCase().trim();
    var nombre = data[i][1];
    var fase   = data[i][6];
    if (!correo || vistos[correo]) continue;

    var filaComm = findRowByEmailInSheet(comSheet, correo);
    var yaEnv    = filaComm ? comSheet.getRange(filaComm, colEnviado).getValue() : '';

    if (yaEnv !== 'Sí') {
      if (fnNombre === 'enviarCorreoConfirmacion') enviarCorreoConfirmacion(nombre, correo, fase);
      if (fnNombre === 'enviarCorreoRecordatorio') enviarCorreoRecordatorio(nombre, correo);
      if (fnNombre === 'enviarCorreoAgradecimiento') enviarCorreoAgradecimiento(nombre, correo);
      if (filaComm) comSheet.getRange(filaComm, colEnviado).setValue('Sí');
      vistos[correo] = true;
      enviados++;
      Utilities.sleep(1000);
    }
  }
  Logger.log(label + ' enviados: ' + enviados);
}

function enviarCorreosQRMasivo() {
  var regSheet = getSheet();
  var comSheet = getComunicacionesSheet();
  var asiSheet = getAsistenciaSheet();
  var regData  = regSheet.getDataRange().getValues();
  var asiData  = asiSheet.getDataRange().getValues();
  var enviados = 0;
  var vistos   = {};

  var idPorCorreo = {};
  for (var j = 1; j < asiData.length; j++) {
    var c = (asiData[j][2] || '').toLowerCase().trim();
    if (c) idPorCorreo[c] = asiData[j][0];
  }

  for (var i = 1; i < regData.length; i++) {
    if (regData[i][9] !== '✓') continue;
    var correo = (regData[i][2] || '').toLowerCase().trim();
    var nombre = regData[i][1];
    if (!correo || vistos[correo]) continue;

    var id       = idPorCorreo[correo];
    var filaComm = findRowByEmailInSheet(comSheet, correo);
    var yaEnv    = filaComm ? comSheet.getRange(filaComm, 13).getValue() : '';

    if (id && yaEnv !== 'Sí') {
      enviarCorreoQR(nombre, correo, id);
      if (filaComm) comSheet.getRange(filaComm, 13).setValue('Sí');
      vistos[correo] = true;
      enviados++;
      Utilities.sleep(1000);
    }
  }
  Logger.log('Correos QR enviados: ' + enviados);
}

/* ── Correos de prueba ───────────────────────────────────────── */
function enviarCorreoPrueba()              { enviarCorreoConfirmacion('Valeria García', 'mejoracontinua@caceca.org', 'Early Bird'); }
function enviarCorreoRecordatorioPrueba()  { enviarCorreoRecordatorio('Valeria García', 'mejoracontinua@caceca.org'); }
function enviarCorreoQRPrueba()            { enviarCorreoQR('Valeria García', 'mejoracontinua@caceca.org', 'RNV-001'); }
function enviarCorreoAgradecimientoPrueba(){ enviarCorreoAgradecimiento('Valeria García', 'mejoracontinua@caceca.org'); }

/* ── Prueba WhatsApp ─────────────────────────────────────────── */
function probarEnlaceWhatsApp() {
  var sheet = getComunicacionesSheet();
  var fila  = sheet.getLastRow() + 1;
  var link  = generateWhatsAppLinkConfirmacion('Estef PRUEBA', '5536599392');
  Logger.log('Link: ' + link);
  var rt = SpreadsheetApp.newRichTextValue().setText('Enviar WhatsApp').setLinkUrl(link).build();
  sheet.getRange(fila, 5).setRichTextValue(rt);
  sheet.getRange(fila, 6).setValue('PRUEBA');
}

/* ── Configurar dropdowns ────────────────────────────────────── */
function configurarDropdowns() {
  var sheet      = getComunicacionesSheet();
  var ultimaFila = Math.max(sheet.getLastRow(), 2);
  var regla      = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pendiente', 'Enviado', 'SIN TELEFONO'], true)
    .setAllowInvalid(false)
    .build();
  // F, I, L, O = cols 6, 9, 12, 15 = estados de WA
  [6, 9, 12, 15].forEach(function(col) {
    sheet.getRange(2, col, ultimaFila - 1, 1).setDataValidation(regla);
  });
  Logger.log('Dropdowns configurados');
}

/* ── Notificaciones internas ─────────────────────────────────── */
function notificarCupoAgotado(sheet) {
  var eb  = contarFaseSheet(sheet, 'Early Bird');
  var pre = contarFaseSheet(sheet, 'Preventa');
  var fin = contarFaseSheet(sheet, 'Últimos lugares');
  var recaudado = (eb * 1300) + (pre * 1500) + (fin * 1700);
  var asunto = 'REINVENTA — Cupo completo 40/40 lugares vendidos';
  var cuerpo = 'SOLD OUT!\n\n— Early Bird: ' + eb + '\n— Preventa: ' + pre + '\n— Últimos lugares: ' + fin
    + '\n\nTotal estimado: $' + recaudado.toLocaleString('es-MX') + ' MXN';
  EMAILS_NOTIFICACION.forEach(function(email) { MailApp.sendEmail(email, asunto, cuerpo); });
}

/* ── Configuración inicial ───────────────────────────────────── */
function configurarHojas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  function crearOActualizar(nombre, headers) {
    var sheet = ss.getSheetByName(nombre);
    if (!sheet) sheet = ss.insertSheet(nombre);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#2A0F25').setFontColor('#C6A56A');
    return sheet;
  }

  crearOActualizar('Registros', [
    'Fecha registro','Nombre','Correo','WhatsApp','Contacto preferido',
    '¿Qué busca?','Fase de compra','Monto pagado (MXN)','Fecha de pago','Pagó ✓',
    'Stripe Payment ID','Origen','Acepta marketing','Autoriza uso de imagen',
    'Canal UTM','¿Cómo se enteró?','¿A qué se dedica?'
  ]);

  crearOActualizar('Comunicaciones', [
    'Correo','Nombre','WhatsApp','Contacto preferido',
    'WA Confirmación','Estado Conf. WA','Correo Conf. Enviado',
    'WA Recordatorio','Estado Record. WA','Correo Record. Enviado',
    'WA QR','Estado QR WA','Correo QR Enviado',
    'WA Agradecimiento','Estado Agradec. WA','Correo Agradec. Enviado'
  ]);

  crearOActualizar('Asistencia', [
    'ID Único','Nombre','Correo','Fase','Asistió ✓','Fecha entrada',
    'Encuesta ✓','Fecha encuesta','Calificación','Comentario'
  ]);

  Logger.log('Hojas configuradas.');
}

/* ── Migración de datos existentes ───────────────────────────── */
function migrarDatosExistentes() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var origen = ss.getSheetByName('REINVENTA - Registro')
            || ss.getSheetByName('Reinventa - Registros');

  if (!origen) { Logger.log('Hoja original no encontrada.'); return; }

  var regSheet = getSheet();
  var comSheet = getComunicacionesSheet();
  var asiSheet = getAsistenciaSheet();
  var data     = origen.getDataRange().getValues();
  var migrados = 0;

  for (var i = 1; i < data.length; i++) {
    var fila   = data[i];
    var correo = (fila[2] || '').toLowerCase().trim();
    if (!correo) continue;

    if (!findRowByEmail(regSheet, correo)) {
      regSheet.appendRow([
        fila[0], fila[1], correo, fila[3], fila[4], fila[5],
        fila[6], fila[7], fila[8], fila[9], fila[10], fila[11],
        fila[12], fila[13], fila[15], fila[16], fila[17]
      ]);
    }

    if (fila[9] === '✓') {
      if (!findRowByEmailInSheet(comSheet, correo)) {
        comSheet.appendRow([correo, fila[1], fila[3], fila[4], '', '', '', '', '', '', '', '', '', '', '', '']);
      }
      if (!findRowByEmailInSheet(asiSheet, correo)) {
        asiSheet.appendRow([generarSiguienteId(asiSheet), fila[1], correo, fila[6], '', '']);
      }
    }
    migrados++;
  }
  Logger.log('Migrados: ' + migrados);
}

/* ── Helpers de confirmación ─────────────────────────────────── */
function obtenerIdAsistente(correo) {
  var sheet = getAsistenciaSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][2] || '').toLowerCase().trim() === correo.toLowerCase().trim()) {
      return data[i][0] || '';
    }
  }
  return '';
}

function marcarCorreoEnviado(correo, col) {
  var sheet = getComunicacionesSheet();
  var fila  = findRowByEmailInSheet(sheet, correo);
  if (fila) sheet.getRange(fila, col).setValue('Sí');
}

/* ── Helpers ─────────────────────────────────────────────────── */
function findRowByEmail(sheet, correo) {
  if (!correo) return null;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][2] || '').toString().toLowerCase().trim() === correo) return i + 1;
  }
  return null;
}

function findRowByEmailInSheet(sheet, correo) {
  if (!correo) return null;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString().toLowerCase().trim() === correo.toLowerCase().trim()) return i + 1;
  }
  return null;
}

function findRowByStripeId(sheet, stripeId) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][10] || '').toString().trim() === stripeId) return i + 1;
  }
  return null;
}

function contarPagosSheet(sheet) {
  var data = sheet.getDataRange().getValues();
  var n = 0;
  for (var i = 1; i < data.length; i++) { if (data[i][9] === '✓') n++; }
  return n;
}

function contarFaseSheet(sheet, fase) {
  var data = sheet.getDataRange().getValues();
  var n = 0;
  for (var i = 1; i < data.length; i++) { if (data[i][9] === '✓' && data[i][6] === fase) n++; }
  return n;
}
