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
 *   — Indicaciones (jueves 13 ago) —
 *   Q  WA Indicaciones      R  Estado Indic. WA  S  Correo Indic. Enviado
 *
 * [Asistencia]
 *   A  ID Único (RNV-001)   B  Nombre             C  Correo
 *   D  Fase                 E  Asistió ✓          F  Fecha entrada
 *   G  Encuesta ✓           H  Fecha encuesta     I  Calificación  J  Comentario
 *
 * [Dashboard] — solo fórmulas, no la toca el script
 */

var SHEET_REGISTROS      = 'Registros';
var SHEET_COMUNICACIONES = 'Comunicaciones';
var SHEET_ASISTENCIA     = 'Asistencia';
var SHEET_MEDIDAS        = 'Medidas';

var EMAILS_NOTIFICACION = ['mejoracontinua@caceca.org', 'alopez@alumbrastudios.com'];
var LIMITE_TOTAL        = 40;
var STAFF_PIN           = '1508';

/* ── Getters de hojas ────────────────────────────────────────── */
function getSheet()               { return getSheetByName(SHEET_REGISTROS); }
function getComunicacionesSheet() { return getSheetByName(SHEET_COMUNICACIONES); }
function getAsistenciaSheet()     { return getSheetByName(SHEET_ASISTENCIA); }
function getMedidasSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_MEDIDAS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_MEDIDAS);
    sh.appendRow(['ID', 'Nombre', 'Hombros (cm)', 'Busto (cm)', 'Cintura (cm)', 'Cadera (cm)', 'Forma de cuerpo', 'Fecha']);
    sh.getRange(1, 1, 1, 8).setFontWeight('bold');
  }
  return sh;
}

function getSheetByName(nombre) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(nombre);
  if (!sheet) throw new Error('Hoja no encontrada: ' + nombre);
  return sheet;
}

/* ── doGet ───────────────────────────────────────────────────── */
function doGet(e) {
  var action = e.parameter.action || '';
  if (action === 'entrada') return handleEntrada(e.parameter.id || '', e.parameter.preview === '1');
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
    if (data.action === 'guardar_medidas') return handleGuardarMedidas(data);
    if (data.action === 'admin_action')    return handleAdminAction(data);
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
  else if (monto == '2600.00') fase = 'Early Bird x2';
  else if (monto == '3000.00') fase = 'Preventa x2';
  else if (monto == '3400.00') fase = 'Últimos lugares x2';
  else if (monto == '3900.00') fase = 'Early Bird x3';
  else if (monto == '4500.00') fase = 'Preventa x3';
  else if (monto == '5100.00') fase = 'Últimos lugares x3';

  var sheet = getSheet();

  // Anti-duplicado por Stripe ID
  if (stripeId && findRowByStripeId(sheet, stripeId)) {
    return ContentService.createTextOutput(JSON.stringify({ result: 'duplicate' })).setMimeType(ContentService.MimeType.JSON);
  }

  var existingRow = findRowByEmail(sheet, correo);
  var tel         = '';
  var contacto    = '';

  if (existingRow) {
    var yaPago = sheet.getRange(existingRow, 10).getValue() === '✓';
    tel      = sheet.getRange(existingRow, 4).getValue();
    contacto = sheet.getRange(existingRow, 5).getValue();

    if (!yaPago) {
      // Primera compra de este correo — actualiza la fila existente
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
    } else {
      // Segunda (o tercera) compra del mismo correo — boleto adicional
      // Se agrega NUEVA fila en Registros con el nuevo Stripe ID
      sheet.appendRow([
        new Date(), nombre || sheet.getRange(existingRow, 2).getValue(), correo,
        tel, contacto, '', fase, monto, fecha, '✓', stripeId, 'stripe directo',
        '', '', '', '', ''
      ]);
      // Agrega acompañante en Asistencia (sin tocar Comunicaciones)
      var nombreComprador = sheet.getRange(existingRow, 2).getValue() || nombre;
      agregarAcompanante(nombreComprador, fase);
    }
  } else {
    // Correo completamente nuevo
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
  // Verificar si ya existe por correo
  for (var i = 1; i < data.length; i++) {
    if ((data[i][2] || '').toLowerCase().trim() === correo.toLowerCase().trim()) return;
  }
  var id = generarSiguienteId(sheet);
  sheet.appendRow([id, nombre, correo, fase, '', '', '', '', '', '']);
}

/* ── Agregar acompañante en Asistencia ───────────────────────── */
/* Se llama cuando el comprador ya tiene su fila pero compra un
   boleto adicional (segunda o tercera compra del mismo correo).
   También se usa en sincronizarRegistrosFaltantes para las fases
   que ya traen x2 o x3 en un solo Stripe event.                 */
function agregarAcompanante(nombreComprador, fase) {
  var sheet     = getAsistenciaSheet();
  var nombreAc  = 'Acompañante de ' + nombreComprador;
  // Verificar que no exista ya ese acompañante
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][1] || '').toLowerCase().trim() === nombreAc.toLowerCase().trim()) return;
  }
  var id = generarSiguienteId(sheet);
  sheet.appendRow([id, nombreAc, '', fase, '', '', '', '', '', '']);
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

  var idNorm = id.toString().trim().toUpperCase();
  var ss     = SpreadsheetApp.getActiveSpreadsheet();

  // ── Buscar en Asistencia: primero por ID (col A), luego por correo (col C) ──
  var asiSheet = getAsistenciaSheet();
  var asiData  = asiSheet.getDataRange().getValues();
  var fila     = -1;

  // Pasada 1: coincidencia exacta por ID
  for (var i = 1; i < asiData.length; i++) {
    if ((asiData[i][0] || '').toString().trim().toUpperCase() === idNorm) { fila = i; break; }
  }

  // Pasada 2: si no encontró por ID, busca por correo (quien entra con su correo en vez del código)
  if (fila === -1) {
    for (var i = 1; i < asiData.length; i++) {
      if ((asiData[i][2] || '').toString().trim().toLowerCase() === id.toString().trim().toLowerCase()) { fila = i; break; }
    }
  }

  // ── Si encontró en Asistencia → devolver datos ────────────────────────────
  if (fila !== -1) {
    var foundId   = (asiData[fila][0] || '').toString().trim();
    var prevSheet = ss.getSheetByName('Encuesta Previa');
    var encPrev   = false;
    if (prevSheet) {
      var prevData = prevSheet.getDataRange().getValues();
      for (var j = 1; j < prevData.length; j++) {
        if ((prevData[j][0] || '').toString().trim() === foundId) { encPrev = true; break; }
      }
    }
    return ContentService.createTextOutput(JSON.stringify({
      id:             asiData[fila][0],
      nombre:         asiData[fila][1],
      correo:         asiData[fila][2],
      fase:           asiData[fila][3],
      asistio:        asiData[fila][4] === '✓',
      fechaEntrada:   asiData[fila][5] ? asiData[fila][5].toString() : '',
      encuesta:       asiData[fila][6] === '✓',
      calificacion:   asiData[fila][8] || 0,
      encuestaPrevia: encPrev
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // ── Fallback: buscar en Registros por correo o por ID (si ya fue sincronizado)
  // Cubre el caso de personas que pagaron pero no fueron sincronizadas a Asistencia aún.
  try {
    var regSheet = getSheet();
    var regData  = regSheet.getDataRange().getValues();
    var idBaja   = id.toString().trim().toLowerCase();
    var esCodigoRNV = /^rnv-\d+$/i.test(idBaja);

    for (var r = 1; r < regData.length; r++) {
      var regCorreo = (regData[r][2] || '').toString().trim().toLowerCase();
      var regPago   = (regData[r][9] || '').toString().trim();
      if (!regCorreo || regPago !== '✓') continue;

      // Coincide si el input es el correo de la persona
      var coincide = (regCorreo === idBaja);

      if (coincide) {
        var regNombre = (regData[r][1] || '').toString().trim();
        var regFase   = (regData[r][6] || '').toString().trim();
        actualizarAsistencia(regCorreo, regNombre, regFase);
        // Recuperar la fila recién creada
        var asiData2 = asiSheet.getDataRange().getValues();
        for (var k = 1; k < asiData2.length; k++) {
          if ((asiData2[k][2] || '').toString().trim().toLowerCase() === regCorreo) {
            return ContentService.createTextOutput(JSON.stringify({
              id:             asiData2[k][0],
              nombre:         asiData2[k][1],
              correo:         asiData2[k][2],
              fase:           asiData2[k][3],
              asistio:        false,
              fechaEntrada:   '',
              encuesta:       false,
              calificacion:   0,
              encuestaPrevia: false
            })).setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
    }
  } catch(e) {
    // Si falla el fallback, caemos al error estándar
  }

  return ContentService.createTextOutput(JSON.stringify({ error: 'Registro no encontrado' })).setMimeType(ContentService.MimeType.JSON);
}

/* ── Admin: todos los datos del panel ───────────────────────── */
/* ── Guardar medidas de cuerpo desde el hub ─────────────────── */
function handleGuardarMedidas(data) {
  if (!data.id) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'ID requerido' })).setMimeType(ContentService.MimeType.JSON);
  }
  var sheet   = getMedidasSheet();
  var medData = sheet.getDataRange().getValues();
  var idNorm  = data.id.toString().trim();

  // Buscar fila existente para este ID
  var filaExistente = -1;
  for (var i = 1; i < medData.length; i++) {
    if ((medData[i][0] || '').toString().trim() === idNorm) { filaExistente = i + 1; break; }
  }

  // Obtener nombre del asistente desde Asistencia
  var nombre = data.nombre || '';
  if (!nombre) {
    try {
      var asiData = getAsistenciaSheet().getDataRange().getValues();
      for (var j = 1; j < asiData.length; j++) {
        if ((asiData[j][0] || '').toString().trim() === idNorm) { nombre = asiData[j][1]; break; }
      }
    } catch(e) {}
  }

  var fila = [
    idNorm,
    nombre,
    parseFloat(data.hombros) || '',
    parseFloat(data.busto)   || '',
    parseFloat(data.cintura) || '',
    parseFloat(data.cadera)  || '',
    data.forma || '',
    new Date().toLocaleString('es-MX')
  ];

  if (filaExistente > 0) {
    sheet.getRange(filaExistente, 1, 1, 8).setValues([fila]);
  } else {
    sheet.appendRow(fila);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

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

  // Mapa de medidas por ID
  var medidasPorId = {};
  try {
    var medSheet = getMedidasSheet();
    var medData  = medSheet.getDataRange().getValues();
    for (var m = 1; m < medData.length; m++) {
      var mid = (medData[m][0] || '').toString().trim();
      if (mid) medidasPorId[mid] = {
        hombros: medData[m][2] || '',
        busto:   medData[m][3] || '',
        cintura: medData[m][4] || '',
        cadera:  medData[m][5] || '',
        forma:   medData[m][6] || ''
      };
    }
  } catch(e) {}

  var encuestasPrevia = [];
  if (prevSheet) {
    var prevData = prevSheet.getDataRange().getValues();
    for (var j = 1; j < prevData.length; j++) {
      var p = prevData[j];
      if (!p[0]) continue;
      var pid = (p[0] || '').toString().trim();
      encuestasPrevia.push({
        id:           pid,
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
        ojos:         p[12],
        medidas:      medidasPorId[pid] || null
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
function handleEntrada(id, preview) {
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

      // Preview mode: return info without marking attendance
      if (preview) {
        return ContentService.createTextOutput(JSON.stringify({ nombre: nombre, fase: fase, preview: true })).setMimeType(ContentService.MimeType.JSON);
      }

      if (yaAsistio === '✓') {
        return ContentService
          .createTextOutput(JSON.stringify({ yaRegistrado: true, nombre: nombre }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      // Gate: requiere encuesta previa
      var ss2 = SpreadsheetApp.getActiveSpreadsheet();
      var prevSheet2 = ss2.getSheetByName('Encuesta Previa');
      var tienePrevia = false;
      if (prevSheet2) {
        var prevData2 = prevSheet2.getDataRange().getValues();
        for (var k = 1; k < prevData2.length; k++) {
          if ((prevData2[k][0]||'').toString().trim() === id) { tienePrevia = true; break; }
        }
      }
      if (!tienePrevia) {
        return ContentService.createTextOutput(JSON.stringify({ error: 'encuesta_previa_pendiente', nombre: nombre })).setMimeType(ContentService.MimeType.JSON);
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
    e('Hola, ' + p + '. Tu lugar está confirmado.') + NL + NL +
    e('*Lo que tu imagen comunica*') + NL +
    e('- Sábado 15 de agosto de 2026') + NL +
    e('- 10:00 a 12:00 pm') + NL +
    e('- The University Club of Mexico') + NL +
    e('- Av. Paseo de la Reforma 150, Juárez, CDMX') + NL + NL +
    e('*Cómo llegar:*') + NL +
    e('https://maps.app.goo.gl/Uo7tYiQz23jMCmKw7') + NL + NL +
    e('- - - - - - - - - - - - -') + NL + NL +
    e('*Tu espacio personal del evento:*') + NL +
    e(hubUrl) + NL + NL +
    e('Aquí encontrarás tu pase de entrada con código QR, la agenda del día y los recursos del taller.') + NL + NL +
    e('- - - - - - - - - - - - -') + NL + NL +
    e('*POR FAVOR LEE ESTO ANTES DE CERRAR*') + NL + NL +
    e('Dentro de tu espacio hay una *encuesta previa* que Mary necesita que contestes _antes del taller_.') + NL + NL +
    e('Mary lee personalmente cada respuesta para preparar los materiales, los ejemplos y las recomendaciones específicas para cada asistente. Si no la contestas antes del evento, Mary no podrá personalizar tu experiencia ese día.') + NL + NL +
    e('No toma más de 5 minutos. Por favor, hazlo hoy mismo.') + NL + NL +
    e('- - - - - - - - - - - - -') + NL + NL +
    e('Nos da mucho gusto tenerte. Mary estará encantada de acompañarte.') + NL + NL +
    e('_Con cariño,_') + NL +
    e('_Reinventa by Mary Méndez_') + NL + NL +
    e('_Este es un mensaje informativo, por favor no respondas a este chat._') + NL + NL +
    e('- - - - - - - - - - - - -') + NL +
    (idUnico ? e('_Tu código de acceso: *' + idUnico + '*_') + NL : '') +
    e('_Organizado integralmente por_') + NL +
    e('*Alumbra Studios*') + NL +
    e('https://www.alumbrastudios.com');
  return 'https://wa.me/' + numero + '?text=' + msg;
}

function generateWhatsAppLinkIndicaciones(nombre, telefono, idUnico) {
  var numero = normalizeWhatsAppNumber(telefono);
  if (!numero) return null;
  var p      = nombre ? nombre.trim().split(' ')[0] : 'participante';
  var e      = encodeURIComponent;
  var NL     = '%0A';
  var hubUrl = idUnico
    ? 'https://reinventabymarymendez.com.mx/hub?id=' + idUnico
    : 'https://reinventabymarymendez.com.mx/hub';
  var msg =
    e('*REINVENTA by Mary Mendez*') + NL + NL +
    e('¡Hola a todas!') + NL + NL +
    e('Estamos muy cerca de nuestro taller y, para que aproveches al máximo tu experiencia y podamos realizar tu análisis de cuerpo y rostro de forma precisa, te comparto los siguientes requerimientos e indicaciones para el día del evento:') + NL + NL +
    e('- - - - - - - - - - - - -') + NL + NL +
    e('*👕 Vestimenta (Análisis de cuerpo):*') + NL + NL +
    e('*Parte inferior:* Asiste con un pantalón ajustado (tipo leggings o jeans ajustados).') + NL + NL +
    e('*Parte superior:* Lleva una blusa o playera básica ajustada (de preferencia en color blanco o neutro).') + NL + NL +
    e('*Capa extra:* Encima de tu playera básica, puedes llevar un saco, blazer o blusón en el color que más te guste o con el que te sientas más cómoda.') + NL + NL +
    e('- - - - - - - - - - - - -') + NL + NL +
    e('*💇‍♀️ Rostro (Análisis visagismo):*') + NL + NL +
    e('Es muy importante despejar tu rostro. Te sugerimos acudir con el cabello recogido; si no te es posible, no te preocupes, aquí contaremos con pinzas para facilitártelo.') + NL + NL +
    e('- - - - - - - - - - - - -') + NL + NL +
    e('*📏 Pasos para medir tu cuerpo:*') + NL + NL +
    e('• Hombros: Mide de un extremo al otro pasando por la parte más alta de la espalda.') + NL +
    e('• Busto: Rodea la parte más voluminosa del pecho, a la altura de los pezones.') + NL +
    e('• Cintura: Ubica la zona más angosta del torso, justo 2 dedos arriba del ombligo, sin meter el abdomen.') + NL +
    e('• Cadera: Mide la parte más ancha de los glúteos y los huesos de la cadera. Toma nota de cada medida.') + NL + NL +
    e('*Identifica la forma de tu cuerpo:*') + NL + NL +
    e('• Reloj de arena: Busto y caderas similares, cintura notablemente más pequeña (diferencia de 20 cm o más).') + NL +
    e('• Rectángulo: Hombros, cintura y cadera con medidas muy parecidas.') + NL +
    e('• Triángulo o Pera: Las caderas son más anchas (5% o más) que hombros y busto.') + NL +
    e('• Triángulo invertido: Los hombros o busto son más anchos que las caderas.') + NL +
    e('• Manzana: La medida de la cintura es mayor o similar a la de hombros y caderas.') + NL + NL +
    e('- - - - - - - - - - - - -') + NL + NL +
    e('¡Asegúrate de venir cómoda y lista para descubrir tu mejor versión! Nos vemos muy pronto. 🌟') + NL + NL +
    e('- - - - - - - - - - - - -') + NL + NL +
    e('*Tu espacio personal del evento:*') + NL +
    e(hubUrl) + NL + NL +
    e('_Con cariño,_') + NL +
    e('_Reinventa by Mary Mendez_') + NL + NL +
    e('_Este es un mensaje informativo, por favor no respondas a este chat._') + NL + NL +
    e('- - - - - - - - - - - - -') + NL +
    (idUnico ? e('_Tu código de acceso: *' + idUnico + '*_') + NL : '') +
    e('_Organizado integralmente por_') + NL +
    e('*Alumbra Studios*') + NL +
    e('https://www.alumbrastudios.com');
  return 'https://wa.me/' + numero + '?text=' + msg;
}

function generateWhatsAppLinkRecordatorio(nombre, telefono, idUnico) {
  var numero = normalizeWhatsAppNumber(telefono);
  if (!numero) return null;
  var p      = nombre ? nombre.trim().split(' ')[0] : 'participante';
  var e      = encodeURIComponent;
  var NL     = '%0A';
  var hubUrl = idUnico
    ? 'https://reinventabymarymendez.com.mx/hub?id=' + idUnico
    : 'https://reinventabymarymendez.com.mx/hub';
  var msg =
    e('*REINVENTA by Mary Mendez*') + NL + NL +
    e('Hola, ' + p + '. Mañana es el gran día.') + NL + NL +
    e('*Lo que tu imagen comunica*') + NL +
    e('- Mañana sábado 15 de agosto') + NL +
    e('- 10:00 a 12:00 pm') + NL +
    e('- The University Club of Mexico') + NL +
    e('- Av. Paseo de la Reforma 150, Juárez, CDMX') + NL + NL +
    e('*Cómo llegar:*') + NL +
    e('https://maps.app.goo.gl/Uo7tYiQz23jMCmKw7') + NL + NL +
    e('- - - - - - - - - - - - -') + NL + NL +
    e('*Tu espacio personal del evento:*') + NL +
    e(hubUrl) + NL + NL +
    e('Ahí encuentras la agenda del día, tu código QR de entrada y los materiales del taller.') + NL + NL +
    e('- - - - - - - - - - - - -') + NL + NL +
    e('*👕 Recuerda el dresscode de mañana:*') + NL + NL +
    e('• Parte inferior: pantalón ajustado (leggings o jeans).') + NL +
    e('• Parte superior: blusa o playera básica ajustada (blanco o neutro de preferencia).') + NL +
    e('• Capa extra: saco, blazer o blusón en el color que prefieras.') + NL +
    e('• Rostro: de preferencia con el cabello recogido para el análisis de visagismo.') + NL + NL +
    e('- - - - - - - - - - - - -') + NL + NL +
    e('*Ultimo recordatorio: encuesta previa*') + NL + NL +
    e('Si aún no has contestado la encuesta dentro de tu espacio, este es el momento. Mary la revisa antes del taller para personalizar tu experiencia.') + NL + NL +
    e('Quienes no la contesten antes del evento no tendrán acceso al material digital posterior — guías, recursos y demás — que se desbloquea en tu perfil tras el taller.') + NL + NL +
    e('- - - - - - - - - - - - -') + NL + NL +
    e('Te esperamos puntual. ¡Nos vemos mañana!') + NL + NL +
    e('_Con cariño,_') + NL +
    e('_Reinventa by Mary Mendez_') + NL + NL +
    e('_Este es un mensaje informativo, por favor no respondas a este chat._') + NL + NL +
    e('- - - - - - - - - - - - -') + NL +
    (idUnico ? e('_Tu código de acceso: *' + idUnico + '*_') + NL : '') +
    e('_Organizado integralmente por_') + NL +
    e('*Alumbra Studios*') + NL +
    e('https://www.alumbrastudios.com');
  return 'https://wa.me/' + numero + '?text=' + msg;
}

function generateWhatsAppLinkQR(nombre, telefono, idUnico) {
  var numero = normalizeWhatsAppNumber(telefono);
  if (!numero) return null;
  var p      = nombre ? nombre.trim().split(' ')[0] : 'participante';
  var e      = encodeURIComponent;
  var NL     = '%0A';
  var hubUrl = 'https://reinventabymarymendez.com.mx/hub?id=' + idUnico;
  var msg =
    e('*REINVENTA by Mary Mendez*') + NL + NL +
    e('Hola, ' + p + '. Hoy te esperamos.') + NL + NL +
    e('*Lo que tu imagen comunica*') + NL +
    e('- Hoy sábado 15 de agosto') + NL +
    e('- 10:00 a 12:00 pm') + NL +
    e('- The University Club of Mexico') + NL +
    e('- Av. Paseo de la Reforma 150, Juárez, CDMX') + NL + NL +
    e('*Cómo llegar:*') + NL +
    e('https://maps.app.goo.gl/Uo7tYiQz23jMCmKw7') + NL + NL +
    e('- - - - - - - - - - - - -') + NL + NL +
    e('*Tu código QR de entrada:*') + NL + NL +
    e('Entra a tu espacio personal y muestra el código QR al llegar. El staff lo escaneará en la entrada.') + NL + NL +
    e(hubUrl) + NL + NL +
    e('- - - - - - - - - - - - -') + NL + NL +
    e('¡Nos vemos en un momento. Sera un día increíble!') + NL + NL +
    e('_Con cariño,_') + NL +
    e('_Reinventa by Mary Mendez_') + NL + NL +
    e('_Este es un mensaje informativo, por favor no respondas a este chat._') + NL + NL +
    e('- - - - - - - - - - - - -') + NL +
    e('_Tu código de acceso: *' + idUnico + '*_') + NL +
    e('_Organizado integralmente por_') + NL +
    e('*Alumbra Studios*') + NL +
    e('https://www.alumbrastudios.com');
  return 'https://wa.me/' + numero + '?text=' + msg;
}

function generateWhatsAppLinkAgradecimiento(nombre, telefono, idUnico) {
  var numero = normalizeWhatsAppNumber(telefono);
  if (!numero) return null;
  var p      = nombre ? nombre.trim().split(' ')[0] : 'participante';
  var e      = encodeURIComponent;
  var NL     = '%0A';
  var hubUrl = idUnico
    ? 'https://reinventabymarymendez.com.mx/hub?id=' + idUnico
    : 'https://reinventabymarymendez.com.mx/hub';
  var msg =
    e('*REINVENTA by Mary Mendez*') + NL + NL +
    e('Hola, ' + p + '. Fue un honor acompañarte hoy.') + NL + NL +
    e('Gracias por confiar en este espacio y por abrirte a transformar la manera en que tu imagen comunica quién eres. Lo que viviste hoy es solo el comienzo.') + NL + NL +
    e('Mary estará siempre disponible para seguir acompañándote en este camino.') + NL + NL +
    e('- - - - - - - - - - - - -') + NL + NL +
    e('*Recursos del taller*') + NL + NL +
    e('Las guías y materiales se desbloquean en tu espacio personal una vez que completes la encuesta de satisfacción.') + NL + NL +
    e('Para acceder:') + NL +
    e('1. Entra a tu espacio personal') + NL +
    e('2. Completa la encuesta de satisfacción') + NL + NL +
    e(hubUrl) + NL + NL +
    e('- - - - - - - - - - - - -') + NL + NL +
    e('*Conoce más sobre Mary:*') + NL +
    e('https://reinventabymarymendez.com.mx') + NL + NL +
    e('_Con cariño,_') + NL +
    e('_Reinventa by Mary Mendez_') + NL + NL +
    e('_Este es un mensaje informativo, por favor no respondas a este chat._') + NL + NL +
    e('- - - - - - - - - - - - -') + NL +
    (idUnico ? e('_Tu código de acceso: *' + idUnico + '*_') + NL : '') +
    e('_Organizado integralmente por_') + NL +
    e('*Alumbra Studios*') + NL +
    e('https://www.alumbrastudios.com');
  return 'https://wa.me/' + numero + '?text=' + msg;
}

/* ── Generación masiva de botones WA ────────────────────────── */
function generarBotonesConfirmacionMasivo() {
  _generarBotonesMasivoConId(generateWhatsAppLinkConfirmacion, 5, 6, 'Enviar confirmación', false);
}

function generarBotonesIndicacionesMasivo() {
  _generarBotonesMasivoConId(generateWhatsAppLinkIndicaciones, 17, 18, 'Enviar indicaciones (jueves)', false);
}

function generarBotonesRecordatorioMasivo() {
  _generarBotonesMasivoConId(generateWhatsAppLinkRecordatorio, 8, 9, 'Enviar recordatorio', false);
}

function generarBotonesAgradecimientoMasivo() {
  _generarBotonesMasivoConId(generateWhatsAppLinkAgradecimiento, 14, 15, 'Enviar agradecimiento', false);
}

/* Genera botones WA pasando el ID único de cada asistente al mensaje */
function _generarBotonesMasivoConId(generadorFn, colLink, colEstado, textoBoton, soloSinLink) {
  var regSheet  = getSheet();
  var comSheet  = getComunicacionesSheet();
  var asiSheet  = getAsistenciaSheet();
  var data      = regSheet.getDataRange().getValues();
  var asiData   = asiSheet.getDataRange().getValues();
  var generados = 0;
  var vistos    = {};

  // Mapa correo → ID
  var idPorCorreo = {};
  for (var j = 1; j < asiData.length; j++) {
    var c = (asiData[j][2] || '').toLowerCase().trim();
    if (c) idPorCorreo[c] = asiData[j][0];
  }

  for (var i = 1; i < data.length; i++) {
    if (data[i][9] !== '✓') continue;
    var correo = (data[i][2] || '').toLowerCase().trim();
    var nombre = data[i][1];
    var tel    = data[i][3];
    if (!correo || vistos[correo]) continue;
    vistos[correo] = true;

    var id   = idPorCorreo[correo] || '';
    var fila = findRowByEmailInSheet(comSheet, correo);
    if (!fila) continue;
    var yaLink = comSheet.getRange(fila, colLink).getValue();
    if (soloSinLink && yaLink) continue;

    generarBotonWA(comSheet, fila, nombre, tel,
      function(fn, theId) {
        return function(n, t) { return fn(n, t, theId); };
      }(generadorFn, id),
      colLink, colEstado, textoBoton);
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

    var link = generateWhatsAppLinkQR(nombre, tel, id); // id ya viene de idPorCorreo
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
    + '<p style="font-size:.85rem;color:#4a3545;margin:0 0 .5rem;">Sábado 15 de agosto de 2026 &middot; 10:00-12:00 pm</p>'
    + '<p style="font-size:.85rem;color:#4a3545;margin:0;">The University Club of Mexico<br><span style="color:#8F7383;font-size:.8rem;">Av. Paseo de la Reforma 150, Juárez, CDMX</span></p>'
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
    + '<a href="' + calLink + '" style="display:block;background:#2A0F25;color:#EFE9E2;text-align:center;padding:.9rem 1.2rem;text-decoration:none;font-size:.8rem;letter-spacing:.1em;text-transform:uppercase;margin-bottom:1.6rem;">Agregar a Google Calendar</a>'
    + '<div style="border:1px solid rgba(42,15,37,.12);padding:1.4rem 1.6rem;margin-bottom:1.6rem;">'
    + '<p style="font-size:.6rem;letter-spacing:.16em;text-transform:uppercase;color:#8F7383;margin:0 0 .6rem;">Tu espacio personal del evento</p>'
    + '<p style="font-size:.87rem;color:#4a3545;line-height:1.6;margin:0 0 .8rem;">Aquí encontrarás tu pase de entrada con código QR, la agenda del día y los recursos del taller.</p>'
    + '<a href="' + hubUrl + '" style="display:block;background:#C6A56A;color:#2A0F25;text-align:center;padding:.8rem 1.2rem;text-decoration:none;font-size:.8rem;letter-spacing:.1em;text-transform:uppercase;margin-bottom:1rem;font-weight:600;">Acceder a mi espacio &rarr;</a>'
    + '<div style="border-top:1px solid rgba(42,15,37,.1);padding-top:.9rem;">'
    + '<p style="font-size:.78rem;color:#2A0F25;font-weight:600;margin:0 0 .3rem;">Una cosa importante</p>'
    + '<p style="font-size:.8rem;color:#4a3545;line-height:1.6;margin:0;">Dentro de tu espacio hay una encuesta breve que te pedimos contestar <strong>antes del evento</strong>. Mary la revisa personalmente para preparar materiales y recomendaciones a la medida de cada asistente. No toma más de 5 minutos y hace una gran diferencia.</p>'
    + '</div></div>'
    + _firmaCorreo()
    + '</div>'
    + _footerCorreo(correo);

  MailApp.sendEmail({ to: correo, bcc: 'alopez@alumbrastudios.com', name: 'Reinventa by Mary Méndez',
    subject: 'Tu lugar en el taller está confirmado — REINVENTA', htmlBody: html });
}

/* ── Jueves 13 ago: indicaciones de vestimenta y análisis ───────── */
function enviarCorreoIndicaciones(nombre, correo, idUnico) {
  var p      = nombre ? nombre.split(' ')[0] : 'hola';
  var hubUrl = idUnico
    ? 'https://reinventabymarymendez.com.mx/hub?id=' + idUnico
    : 'https://reinventabymarymendez.com.mx/hub';

  var html = _headerCorreo()
    + '<div style="padding:2.2rem 2.6rem 2rem;">'
    + '<div style="display:inline-block;background:rgba(42,15,37,.07);border-left:2px solid #C6A56A;padding:.4rem .75rem;font-size:.65rem;letter-spacing:.13em;text-transform:uppercase;color:#2A0F25;margin-bottom:1.4rem;">Preparación &middot; Jueves 13 de agosto</div>'
    + '<h1 style="font-family:Georgia,serif;font-weight:400;font-size:1.5rem;line-height:1.35;color:#2A0F25;margin:0 0 .8rem;">' + p + ',<br>ya casi llegamos. 🌟</h1>'
    + '<p style="font-size:.92rem;line-height:1.7;color:#4a3545;margin:0 0 1.6rem;">Para que tu experiencia del sábado sea increíble — y podamos hacer tu análisis de cuerpo y rostro de forma precisa — aquí van las indicaciones para llegar lista:</p>'

    // Hub destacado primero
    + '<div style="background:#2A0F25;padding:1.4rem 1.6rem;margin-bottom:1.6rem;">'
    + '<p style="font-size:.6rem;letter-spacing:.16em;text-transform:uppercase;color:rgba(198,165,106,.6);margin:0 0 .5rem;">Tu espacio personal</p>'
    + '<p style="font-size:.87rem;color:#EFE9E2;line-height:1.6;margin:0 0 .9rem;">Entra a tu espacio para ver la agenda, tu QR de entrada y registrar tus medidas de cuerpo antes del taller.</p>'
    + '<a href="' + hubUrl + '" style="display:block;background:#C6A56A;color:#2A0F25;text-align:center;padding:.8rem 1.2rem;text-decoration:none;font-size:.8rem;letter-spacing:.1em;text-transform:uppercase;font-weight:600;">Entrar a mi espacio &rarr;</a>'
    + '</div>'

    // Vestimenta
    + '<div style="border:1px solid rgba(42,15,37,.12);padding:1.4rem 1.6rem;margin-bottom:1rem;">'
    + '<p style="font-size:.6rem;letter-spacing:.16em;text-transform:uppercase;color:#8F7383;margin:0 0 .9rem;">👕 Qué ponerte &mdash; Análisis de cuerpo</p>'
    + '<p style="font-size:.88rem;color:#4a3545;line-height:1.7;margin:0 0 .5rem;"><strong style="color:#2A0F25;">Abajo:</strong> Pantalón ajustado — leggings o jeans pegados funcionan perfecto.</p>'
    + '<p style="font-size:.88rem;color:#4a3545;line-height:1.7;margin:0 0 .5rem;"><strong style="color:#2A0F25;">Arriba:</strong> Blusa o playera básica ajustada. Si puedes, en blanco o neutro.</p>'
    + '<p style="font-size:.88rem;color:#4a3545;line-height:1.7;margin:0;"><strong style="color:#2A0F25;">Encima:</strong> Un saco, blazer o blusón en el color que más te guste — ese va a poder ser cualquiera.</p>'
    + '</div>'

    // Rostro
    + '<div style="border:1px solid rgba(42,15,37,.12);padding:1.4rem 1.6rem;margin-bottom:1rem;">'
    + '<p style="font-size:.6rem;letter-spacing:.16em;text-transform:uppercase;color:#8F7383;margin:0 0 .9rem;">💇‍♀️ Rostro &mdash; Análisis visagismo</p>'
    + '<p style="font-size:.88rem;color:#4a3545;line-height:1.7;margin:0;">De preferencia con el cabello recogido para el análisis. Si no puedes, no te preocupes &mdash; habrá pinzas disponibles. ✌️</p>'
    + '</div>'

    // Medidas
    + '<div style="border:1px solid rgba(42,15,37,.12);padding:1.4rem 1.6rem;margin-bottom:1.6rem;">'
    + '<p style="font-size:.6rem;letter-spacing:.16em;text-transform:uppercase;color:#8F7383;margin:0 0 .9rem;">📏 Tus medidas (opcional pero súper útil)</p>'
    + '<p style="font-size:.88rem;color:#4a3545;line-height:1.7;margin:0 0 1rem;">Si puedes tomarte las medidas antes de llegar, Mary podrá hacer tu análisis aún más personalizado. Son solo 4: hombros, busto, cintura y cadera. Necesitas una cinta métrica y 5 minutos.</p>'
    + '<p style="font-size:.85rem;color:#4a3545;line-height:1.7;margin:0 0 .35rem;"><strong style="color:#2A0F25;">Hombros:</strong> De un extremo al otro por la parte más alta de la espalda.</p>'
    + '<p style="font-size:.85rem;color:#4a3545;line-height:1.7;margin:0 0 .35rem;"><strong style="color:#2A0F25;">Busto:</strong> La parte más voluminosa del pecho, a la altura de los pezones.</p>'
    + '<p style="font-size:.85rem;color:#4a3545;line-height:1.7;margin:0 0 .35rem;"><strong style="color:#2A0F25;">Cintura:</strong> La zona más angosta, 2 dedos arriba del ombligo, sin meter el abdomen.</p>'
    + '<p style="font-size:.85rem;color:#4a3545;line-height:1.7;margin:0 0 1rem;"><strong style="color:#2A0F25;">Cadera:</strong> La parte más ancha de los glúteos y los huesos de la cadera.</p>'
    + '<a href="' + hubUrl + '" style="display:block;background:transparent;border:1px solid #C6A56A;color:#2A0F25;text-align:center;padding:.7rem 1.2rem;text-decoration:none;font-size:.78rem;letter-spacing:.09em;text-transform:uppercase;font-weight:600;">Registrar mis medidas en mi espacio &rarr;</a>'
    + '</div>'

    + '<p style="font-size:.92rem;line-height:1.7;color:#4a3545;margin:0 0 .5rem;">¡Nos vemos el sábado! Va a ser un día para descubrir muchas cosas lindas.</p>'
    + _firmaCorreo()
    + '</div>'
    + _footerCorreo(correo);

  MailApp.sendEmail({ to: correo, bcc: 'alopez@alumbrastudios.com', name: 'Reinventa by Mary Méndez',
    subject: p + ', esto necesitas para el taller 🌟 — REINVENTA', htmlBody: html });
}

function enviarCorreoRecordatorio(nombre, correo, idUnico) {
  var p      = nombre ? nombre.split(' ')[0] : 'Hola';
  var hubUrl = idUnico
    ? 'https://reinventabymarymendez.com.mx/hub?id=' + idUnico
    : 'https://reinventabymarymendez.com.mx/hub';

  var html = _headerCorreo()
    + '<div style="padding:2.2rem 2.6rem 2rem;">'
    + '<div style="display:inline-block;background:rgba(42,15,37,.07);border-left:2px solid #C6A56A;padding:.4rem .75rem;font-size:.65rem;letter-spacing:.13em;text-transform:uppercase;color:#2A0F25;margin-bottom:1.4rem;">Recordatorio &middot; Mañana es el taller</div>'
    + '<h1 style="font-family:Georgia,serif;font-weight:400;font-size:1.5rem;line-height:1.35;color:#2A0F25;margin:0 0 1rem;">' + p + ',<br>mañana es el gran día.</h1>'
    + '<p style="font-size:.92rem;line-height:1.7;color:#4a3545;margin:0 0 1.8rem;">Mary tiene algo muy especial preparado para ti. Te esperamos puntual y con muchas ganas de transformar la manera en que tu imagen comunica quién eres.</p>'
    + _detallesEvento()
    + '<p style="font-size:.85rem;color:#4a3545;margin:-1rem 0 1.6rem;padding:0 1.6rem;"><a href="https://maps.app.goo.gl/Uo7tYiQz23jMCmKw7" style="color:#C6A56A;text-decoration:none;">Ver en Google Maps &rarr;</a></p>'
    // Acceso al hub
    + '<div style="border:1px solid rgba(42,15,37,.12);padding:1.4rem 1.6rem;margin-bottom:1.6rem;">'
    + '<p style="font-size:.6rem;letter-spacing:.16em;text-transform:uppercase;color:#8F7383;margin:0 0 .6rem;">Tu espacio personal del evento</p>'
    + '<p style="font-size:.87rem;color:#4a3545;line-height:1.6;margin:0 0 .8rem;">Ahí encontrarás la agenda del día, tu código QR de entrada y los materiales del taller.</p>'
    + '<a href="' + hubUrl + '" style="display:block;background:#C6A56A;color:#2A0F25;text-align:center;padding:.8rem 1.2rem;text-decoration:none;font-size:.8rem;letter-spacing:.1em;text-transform:uppercase;margin-bottom:0;font-weight:600;">Acceder a mi espacio &rarr;</a>'
    + '</div>'
    // Dresscode — recordatorio compacto
    + '<div style="border:1px solid rgba(42,15,37,.15);border-left:3px solid #C6A56A;padding:1.2rem 1.4rem;margin-bottom:1.4rem;background:rgba(198,165,106,.04);">'
    + '<p style="font-size:.75rem;color:#2A0F25;font-weight:700;margin:0 0 .5rem;letter-spacing:.03em;">👕 Recuerda el dresscode del día</p>'
    + '<p style="font-size:.82rem;color:#4a3545;line-height:1.6;margin:0 0 .35rem;"><strong>Parte inferior:</strong> Pantalón ajustado (leggings o jeans).</p>'
    + '<p style="font-size:.82rem;color:#4a3545;line-height:1.6;margin:0 0 .35rem;"><strong>Parte superior:</strong> Blusa o playera básica ajustada (blanco o neutro de preferencia).</p>'
    + '<p style="font-size:.82rem;color:#4a3545;line-height:1.6;margin:0 0 .35rem;"><strong>Capa extra:</strong> Saco, blazer o blusón en el color que prefieras.</p>'
    + '<p style="font-size:.82rem;color:#4a3545;line-height:1.6;margin:0;"><strong>Rostro:</strong> De preferencia con el cabello recogido para el análisis de visagismo (si no puedes, hay pinzas disponibles).</p>'
    + '</div>'
    // Encuesta previa — último recordatorio
    + '<div style="border:1px solid #C6A56A;border-left:3px solid #C6A56A;padding:1.2rem 1.4rem;margin-bottom:1.6rem;background:rgba(198,165,106,.06);">'
    + '<p style="font-size:.75rem;color:#2A0F25;font-weight:700;margin:0 0 .4rem;letter-spacing:.03em;">Último recordatorio: encuesta previa</p>'
    + '<p style="font-size:.82rem;color:#4a3545;line-height:1.6;margin:0;">Si aún no has contestado la encuesta dentro de tu espacio, este es el momento. Mary la revisa personalmente antes del taller para personalizar tu experiencia.<br><br><strong>Importante:</strong> quienes no la contesten antes del evento no tendrán acceso al material digital posterior — guías, recursos y demás — que se desbloquea en tu perfil tras el taller.</p>'
    + '</div>'
    + '<p style="font-size:.87rem;line-height:1.7;color:#4a3545;margin:0 0 .5rem;">Si tienes alguna duda de último momento no dudes en contactarnos. ¡Nos vemos mañana!</p>'
    + _firmaCorreo()
    + '</div>'
    + _footerCorreo(correo);

  MailApp.sendEmail({ to: correo, bcc: 'alopez@alumbrastudios.com', name: 'Reinventa by Mary Méndez',
    subject: 'Mañana te esperamos — REINVENTA', htmlBody: html });
}

function enviarCorreoQR(nombre, correo, idUnico) {
  var p      = nombre ? nombre.split(' ')[0] : 'Hola';
  var hubUrl = 'https://reinventabymarymendez.com.mx/hub?id=' + idUnico;
  var urlQR  = 'https://api.qrserver.com/v1/create-qr-code/?size=220x220&data=' + encodeURIComponent(hubUrl);

  var html = _headerCorreo()
    + '<div style="padding:2.2rem 2.6rem 2rem;text-align:center;">'
    + '<div style="display:inline-block;background:rgba(42,15,37,.07);border-left:2px solid #C6A56A;padding:.4rem .75rem;font-size:.65rem;letter-spacing:.13em;text-transform:uppercase;color:#2A0F25;margin-bottom:1.4rem;">Hoy es el día &middot; Sábado 15 de agosto</div>'
    + '<h1 style="font-family:Georgia,serif;font-weight:400;font-size:1.5rem;line-height:1.35;color:#2A0F25;margin:0 0 .6rem;text-align:left;">' + p + ',<br>hoy te esperamos.</h1>'
    + '<p style="font-size:.92rem;line-height:1.7;color:#4a3545;margin:0 0 1.8rem;text-align:left;">Muestra este código QR al llegar al evento. El staff lo escaneará en la entrada.</p>'
    + '<div style="background:#2A0F25;display:inline-block;padding:1.2rem;margin-bottom:1.6rem;">'
    + '<img src="' + urlQR + '" width="200" height="200" style="display:block;" alt="Código QR de entrada" /></div>'
    + _detallesEvento()
    + '<p style="font-size:.87rem;line-height:1.7;color:#4a3545;margin:0 0 .5rem;text-align:left;">¡Nos vemos en un momento. Será un día increíble!</p>'
    + _firmaCorreo()
    + '</div>'
    + _footerCorreo(correo);

  MailApp.sendEmail({ to: correo, bcc: 'alopez@alumbrastudios.com', name: 'Reinventa by Mary Méndez',
    subject: 'Tu entrada para hoy — REINVENTA', htmlBody: html });
}

function enviarCorreoAgradecimiento(nombre, correo, idUnico) {
  var p      = nombre ? nombre.split(' ')[0] : 'Hola';
  var hubUrl = idUnico
    ? 'https://reinventabymarymendez.com.mx/hub?id=' + idUnico
    : 'https://reinventabymarymendez.com.mx/hub';

  var html = _headerCorreo()
    + '<div style="padding:2.2rem 2.6rem 2rem;">'
    + '<div style="display:inline-block;background:rgba(42,15,37,.07);border-left:2px solid #C6A56A;padding:.4rem .75rem;font-size:.65rem;letter-spacing:.13em;text-transform:uppercase;color:#2A0F25;margin-bottom:1.4rem;">Gracias por estar aquí</div>'
    + '<h1 style="font-family:Georgia,serif;font-weight:400;font-size:1.5rem;line-height:1.35;color:#2A0F25;margin:0 0 1rem;">' + p + ',<br>fue un honor acompañarte.</h1>'
    + '<p style="font-size:.92rem;line-height:1.7;color:#4a3545;margin:0 0 1.2rem;">Gracias por confiar en este espacio y por abrirte a transformar la manera en que tu imagen comunica quién eres. Lo que viviste hoy es solo el comienzo.</p>'
    + '<p style="font-size:.92rem;line-height:1.7;color:#4a3545;margin:0 0 1.8rem;">Mary estará siempre disponible para seguir acompañándote en este camino.</p>'
    // Recursos del taller
    + '<div style="border:1px solid rgba(42,15,37,.12);padding:1.4rem 1.6rem;margin-bottom:1.6rem;">'
    + '<p style="font-size:.6rem;letter-spacing:.16em;text-transform:uppercase;color:#8F7383;margin:0 0 .8rem;">Recursos del taller</p>'
    + '<p style="font-size:.87rem;color:#4a3545;line-height:1.6;margin:0 0 1rem;">Las guías y materiales del taller se desbloquean en tu espacio personal una vez que completes la encuesta de satisfacción. Toma solo un momento y nos ayuda muchísimo.</p>'
    + '<div style="border-top:1px solid rgba(42,15,37,.08);padding-top:1rem;margin-bottom:1rem;">'
    + '<p style="font-size:.78rem;color:#2A0F25;margin:0 0 .3rem;">Para acceder a tus recursos:</p>'
    + '<p style="font-size:.82rem;color:#4a3545;margin:0 0 .2rem;">1. Entra a tu espacio personal</p>'
    + '<p style="font-size:.82rem;color:#4a3545;margin:0;">2. Completa la encuesta de satisfacción</p>'
    + '</div>'
    + '<a href="' + hubUrl + '" style="display:block;background:#C6A56A;color:#2A0F25;text-align:center;padding:.8rem 1.2rem;text-decoration:none;font-size:.8rem;letter-spacing:.1em;text-transform:uppercase;font-weight:600;">Ir a mi espacio y completar encuesta &rarr;</a>'
    + '</div>'
    + '<div style="border:1px solid rgba(42,15,37,.12);padding:1.4rem 1.6rem;margin-bottom:1.6rem;">'
    + '<p style="font-size:.6rem;letter-spacing:.16em;text-transform:uppercase;color:#8F7383;margin:0 0 .8rem;">Conoce más sobre Mary</p>'
    + '<p style="font-size:.87rem;color:#4a3545;margin:0;"><a href="https://reinventabymarymendez.com.mx" style="color:#C6A56A;text-decoration:none;">reinventabymarymendez.com.mx &rarr;</a></p>'
    + '</div>'
    + _firmaCorreo()
    + '</div>'
    + _footerCorreo(correo);

  MailApp.sendEmail({ to: correo, bcc: 'alopez@alumbrastudios.com', name: 'Reinventa by Mary Méndez',
    subject: 'Gracias por acompañarnos — REINVENTA', htmlBody: html });
}

/* ── Envíos masivos de correo ────────────────────────────────── */
function enviarConfirmacionExistentes() {
  _enviarCorreosMasivo('enviarCorreoConfirmacion', 7, 'Correos confirmación');
}

function enviarCorreosIndicacionesMasivo() {
  _enviarCorreosMasivo('enviarCorreoIndicaciones', 19, 'Correos indicaciones (jueves)');
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
      var idUnicoCorreo = obtenerIdAsistente(correo);
      if (fnNombre === 'enviarCorreoConfirmacion')  enviarCorreoConfirmacion(nombre, correo, fase, idUnicoCorreo);
      if (fnNombre === 'enviarCorreoIndicaciones')  enviarCorreoIndicaciones(nombre, correo, idUnicoCorreo);
      if (fnNombre === 'enviarCorreoRecordatorio')  enviarCorreoRecordatorio(nombre, correo, idUnicoCorreo);
      if (fnNombre === 'enviarCorreoAgradecimiento') enviarCorreoAgradecimiento(nombre, correo, idUnicoCorreo);
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

/* ── Trigger automático: QR a las 8am del 15 agosto 2026 ────── */
/* Ejecuta esta función UNA SOLA VEZ desde el editor de Apps Script
   para programar el envío automático. Puedes verificarlo en
   Edición > Triggers del proyecto actual.                        */
function programarEnvioQR8AM() {
  // Eliminar triggers previos del mismo nombre para no duplicar
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'enviarCorreosQRMasivo') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // Crear trigger: sábado 15 agosto 2026 a las 8:00 am (hora México = UTC-6)
  // Apps Script usa la zona horaria del spreadsheet, asegúrate de tenerla en America/Mexico_City
  ScriptApp.newTrigger('enviarCorreosQRMasivo')
    .timeBased()
    .at(new Date('2026-08-15T08:00:00'))
    .create();
  SpreadsheetApp.getUi().alert('Listo. El correo con código QR se mandará automáticamente el sábado 15 de agosto a las 8:00 am.');
}

/* Cancela el trigger por si necesitas moverlo o ajustarlo */
function cancelarEnvioQR() {
  var triggers = ScriptApp.getProjectTriggers();
  var cancelados = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'enviarCorreosQRMasivo') {
      ScriptApp.deleteTrigger(triggers[i]);
      cancelados++;
    }
  }
  SpreadsheetApp.getUi().alert('Trigger cancelado (' + cancelados + ' eliminado(s)).');
}

/* ── Correos de prueba ───────────────────────────────────────── */
function enviarCorreoPrueba()              { enviarCorreoConfirmacion('Valeria García', 'mejoracontinua@caceca.org', 'Early Bird', 'RNV-001'); }
function enviarCorreoIndicacionesPrueba()  { enviarCorreoIndicaciones('Valeria García', 'mejoracontinua@caceca.org', 'RNV-001'); }
function enviarCorreoRecordatorioPrueba()  { enviarCorreoRecordatorio('Valeria García', 'mejoracontinua@caceca.org', 'RNV-001'); }
function enviarCorreoQRPrueba()            { enviarCorreoQR('Valeria García', 'mejoracontinua@caceca.org', 'RNV-001'); }
function enviarCorreoAgradecimientoPrueba(){ enviarCorreoAgradecimiento('Valeria García', 'mejoracontinua@caceca.org', 'RNV-001'); }

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

/* ── Menú REINVENTA ──────────────────────────────────────────── */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('REINVENTA')
    .addItem('Sincronizar registros faltantes', 'sincronizarRegistrosFaltantes')
    .addItem('Regenerar links WA sin ID', 'regenerarLinksWASinId')
    .addSeparator()
    .addItem('Correos indicaciones masivo (jue 13 ago)', 'enviarCorreosIndicacionesMasivo')
    .addItem('Generar botones WA indicaciones (jue 13 ago)', 'generarBotonesIndicacionesMasivo')
    .addItem('Programar indicaciones automático (jue 13 ago 12pm)', 'adminProgramarIndicaciones')
    .addItem('Cancelar indicaciones automático', 'adminCancelarIndicaciones')
    .addSeparator()
    .addItem('Enviar recordatorio masivo ahora', 'adminEnviarRecordatorioMasivo')
    .addItem('Programar recordatorio automático (vie 14 ago 12pm)', 'adminProgramarRecordatorio')
    .addItem('Cancelar recordatorio automático', 'adminCancelarRecordatorio')
    .addSeparator()
    .addItem('Programar envío QR automático (sáb 15 ago 8am)', 'programarEnvioQR8AM')
    .addItem('Cancelar envío QR automático', 'cancelarEnvioQR')
    .addSeparator()
    .addItem('Enviar agradecimiento masivo ahora (con check-in)', 'adminEnviarAgradecimientoMasivo')
    .addItem('Programar agradecimiento automático (lun 17 ago 10am)', 'adminProgramarAgradecimiento')
    .addItem('Cancelar agradecimiento automático', 'adminCancelarAgradecimiento')
    .addToUi();
}

/* Regenera el link de WhatsApp de confirmación para filas en
   Comunicaciones cuyo link actual NO contiene ?id=              */
function regenerarLinksWASinId() {
  var sheet = getComunicacionesSheet();
  var data  = sheet.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    var correo   = (data[i][0] || '').toString().trim();
    var nombre   = (data[i][1] || '').toString().trim();
    var telefono = (data[i][2] || '').toString().trim();
    var linkActual = (data[i][4] || '').toString();
    if (!correo || !telefono) continue;
    if (linkActual.indexOf('?id=') !== -1) continue; // ya tiene ID, saltar
    var id = obtenerIdAsistente(correo);
    if (!id) continue; // sin ID en Asistencia, saltar
    var nuevoLink = generateWhatsAppLinkConfirmacion(nombre, telefono, id);
    if (nuevoLink) {
      sheet.getRange(i + 1, 5).setValue(nuevoLink);
      sheet.getRange(i + 1, 5).setFormula('=HYPERLINK("' + nuevoLink + '","Enviar WhatsApp")');
      count++;
    }
  }
  SpreadsheetApp.getUi().alert('Listo. Se actualizaron ' + count + ' link(s) con su ID personalizado.');
}

/* ── Sincronizar registros faltantes ─────────────────────────── */
/* Recorre Registros, y para cada fila con Pagó ✓:
   - Si el correo no está en Asistencia → lo agrega
   - Si la fase es x2 → agrega 1 "Acompañante de..." en Asistencia
   - Si la fase es x3 → agrega 2 "Acompañante de..." en Asistencia
   - Si el correo no está en Comunicaciones → lo agrega con botón WA
   No envía ningún correo ni WhatsApp, todo es manual.           */
function sincronizarRegistrosFaltantes() {
  var sheet    = getSheet();
  var data     = sheet.getDataRange().getValues();
  var vistos   = {};
  var agregados = 0;

  for (var i = 1; i < data.length; i++) {
    var pago    = (data[i][9] || '').toString().trim();
    if (pago !== '✓') continue;

    var correo  = (data[i][2] || '').toString().toLowerCase().trim();
    var nombre  = (data[i][1] || '').toString().trim();
    var telefono= (data[i][3] || '').toString().trim();
    var contacto= (data[i][4] || '').toString().trim();
    var fase    = (data[i][6] || '').toString().trim();

    if (!correo) continue;

    if (!vistos[correo]) {
      vistos[correo] = { nombre: nombre, fase: fase, boletos: 1 };

      // Agregar a Asistencia si no existe
      actualizarAsistencia(correo, nombre, fase);

      // Agregar acompañantes según fase (x2 = 1 acomp., x3 = 2 acomp.)
      var cantAcompanantes = 0;
      if (fase.indexOf('x2') !== -1) cantAcompanantes = 1;
      if (fase.indexOf('x3') !== -1) cantAcompanantes = 2;
      for (var a = 0; a < cantAcompanantes; a++) {
        agregarAcompanante(nombre, fase);
      }

      // Agregar a Comunicaciones si no existe
      var id = obtenerIdAsistente(correo);
      sincronizarComunicaciones(correo, nombre, telefono, contacto, id);

      agregados++;
    } else {
      // Segunda o tercera fila del mismo correo (compra adicional)
      vistos[correo].boletos++;
      // Agregar acompañante por cada boleto adicional
      agregarAcompanante(vistos[correo].nombre, fase);
    }

    Utilities.sleep(300);
  }

  SpreadsheetApp.getUi().alert('Listo. Se sincronizaron ' + agregados + ' persona(s). Revisa Asistencia para ver los acompañantes agregados.');
}

/* ── Admin actions (desde panel) ─────────────────────────── */
function handleAdminAction(data) {
  var sub = data.sub || '';

  if (sub === 'test_email')                  return adminTestEmail();
  if (sub === 'checkin')                     return adminCheckin(data.id);
  if (sub === 'update_correo')               return adminUpdateCorreo(data.id, data.correo_nuevo);
  if (sub === 'registrar_acompanante')       return adminRegistrarAcompanante(data.id, data.nombre_nuevo, data.correo_nuevo);
  if (sub === 'enviar_bienvenida')           return adminEnviarBienvenida(data.id);
  if (sub === 'update_nombre')               return adminUpdateNombre(data.id, data.nombre_nuevo);
  if (sub === 'registro_manual')             return adminRegistroManual(data.nombre, data.correo, data.telefono, data.fase);
  if (sub === 'test_email_acompanante')      return adminTestEmailAcompanante();
  if (sub === 'estado_correos')              return adminEstadoCorreos();
  if (sub === 'programar_indicaciones')      return adminProgramarIndicaciones();
  if (sub === 'cancelar_indicaciones')       return adminCancelarIndicaciones();
  if (sub === 'enviar_indicaciones_masivo')  return adminEnviarIndicacionesMasivo();
  if (sub === 'programar_recordatorio')      return adminProgramarRecordatorio();
  if (sub === 'cancelar_recordatorio')       return adminCancelarRecordatorio();
  if (sub === 'enviar_recordatorio_masivo')  return adminEnviarRecordatorioMasivo();
  if (sub === 'programar_agradecimiento')    return adminProgramarAgradecimiento();
  if (sub === 'cancelar_agradecimiento')     return adminCancelarAgradecimiento();
  if (sub === 'enviar_agradecimiento_masivo')return adminEnviarAgradecimientoMasivo();
  if (sub === 'enviar_qr_masivo_gs')        return ContentService.createTextOutput(JSON.stringify(enviarCorreosQRMasivo()||{result:'ok'})).setMimeType(ContentService.MimeType.JSON);

  return ContentService.createTextOutput(JSON.stringify({ error: 'Acción desconocida' })).setMimeType(ContentService.MimeType.JSON);
}

function adminCheckin(id) {
  if (!id) return jsErr('ID requerido');
  var sheet = getAsistenciaSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString().trim() === id) {
      var nombre = data[i][1];
      if (data[i][4] === '✓') return jsOk({ yaRegistrado: true, nombre: nombre });

      // Gate: requiere encuesta previa
      var ss2 = SpreadsheetApp.getActiveSpreadsheet();
      var prevSheet = ss2.getSheetByName('Encuesta Previa');
      var tienePrevia = false;
      if (prevSheet) {
        var prevData = prevSheet.getDataRange().getValues();
        for (var k = 1; k < prevData.length; k++) {
          if ((prevData[k][0]||'').toString().trim() === id) { tienePrevia = true; break; }
        }
      }
      if (!tienePrevia) {
        return ContentService.createTextOutput(JSON.stringify({ error: 'encuesta_previa_pendiente', nombre: nombre })).setMimeType(ContentService.MimeType.JSON);
      }

      sheet.getRange(i+1,5).setValue('✓');
      sheet.getRange(i+1,6).setValue(new Date());
      return jsOk({ nombre: nombre, fase: data[i][3] });
    }
  }
  return jsErr('ID no encontrado');
}

function adminUpdateCorreo(id, correoNuevo) {
  if (!id || !correoNuevo) return jsErr('Faltan datos');
  correoNuevo = correoNuevo.toLowerCase().trim();

  // Actualizar en Asistencia
  var asi = getAsistenciaSheet();
  var aData = asi.getDataRange().getValues();
  var nombre = '';
  for (var i = 1; i < aData.length; i++) {
    if ((aData[i][0]||'').toString().trim() === id) {
      asi.getRange(i+1,3).setValue(correoNuevo);
      nombre = aData[i][1];
      break;
    }
  }

  // Actualizar o crear en Comunicaciones
  var com = getComunicacionesSheet();
  var cData = com.getDataRange().getValues();
  var filaComm = null;
  for (var j = 1; j < cData.length; j++) {
    if ((cData[j][0]||'').toLowerCase().trim() === correoNuevo) { filaComm = j+1; break; }
  }
  if (!filaComm) {
    com.appendRow([correoNuevo, nombre, '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
    filaComm = com.getLastRow();
  } else {
    if (!com.getRange(filaComm,2).getValue()) com.getRange(filaComm,2).setValue(nombre);
  }

  return jsOk({ correo: correoNuevo });
}

function adminRegistrarAcompanante(id, nombreNuevo, correoNuevo) {
  if (!id || !correoNuevo) return jsErr('Faltan datos');
  correoNuevo = correoNuevo.toLowerCase().trim();

  // Actualizar nombre y correo en Asistencia
  var asi = getAsistenciaSheet();
  var aData = asi.getDataRange().getValues();
  var fase = '';
  var nombreFinal = nombreNuevo || '';
  for (var i = 1; i < aData.length; i++) {
    if ((aData[i][0]||'').toString().trim() === id) {
      if (nombreNuevo) asi.getRange(i+1,2).setValue(nombreNuevo);
      asi.getRange(i+1,3).setValue(correoNuevo);
      fase = aData[i][3];
      break;
    }
  }

  // Crear en Comunicaciones
  var com = getComunicacionesSheet();
  var fila = findRowByEmailInSheet(com, correoNuevo);
  if (!fila) {
    com.appendRow([correoNuevo, nombreFinal, '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
    fila = com.getLastRow();
  }

  // Enviar correo de bienvenida todo-en-uno
  enviarCorreoBienvenidaAcompanante(nombreFinal, correoNuevo, fase, id);

  // Marcar correo de confirmación como enviado en Comunicaciones (col G)
  com.getRange(fila, 7).setValue('Sí');

  return jsOk({ result: 'ok', correo: correoNuevo });
}

function adminEnviarBienvenida(id) {
  if (!id) return jsErr('ID requerido');
  var asi = getAsistenciaSheet();
  var aData = asi.getDataRange().getValues();
  for (var i = 1; i < aData.length; i++) {
    if ((aData[i][0]||'').toString().trim() === id) {
      var nombre = aData[i][1];
      var correo = aData[i][2];
      var fase   = aData[i][3];
      if (!correo) return jsErr('Sin correo registrado');
      enviarCorreoBienvenidaAcompanante(nombre, correo, fase, id);
      var com = getComunicacionesSheet();
      var fila = findRowByEmailInSheet(com, correo);
      if (fila) com.getRange(fila,7).setValue('Sí');
      return jsOk({ result: 'ok' });
    }
  }
  return jsErr('ID no encontrado');
}

/* Correo único todo-en-uno para acompañantes */
function enviarCorreoBienvenidaAcompanante(nombre, correo, fase, idUnico) {
  var p      = nombre ? nombre.split(' ')[0] : 'Hola';
  var hubUrl = 'https://reinventabymarymendez.com.mx/hub?id=' + idUnico;

  var html = _headerCorreo()
    + '<div style="padding:2.2rem 2.6rem 2rem;">'

    // Chip: invitada especial
    + '<div style="display:inline-block;background:rgba(198,165,106,.12);border-left:2px solid #C6A56A;padding:.4rem .75rem;font-size:.65rem;letter-spacing:.13em;text-transform:uppercase;color:#2A0F25;margin-bottom:1.4rem;">Invitada especial · REINVENTA</div>'

    // Encabezado
    + '<h1 style="font-family:Georgia,serif;font-weight:400;font-size:1.5rem;line-height:1.35;color:#2A0F25;margin:0 0 1rem;">Hola, ' + p + '.<br>Qué gusto que nos acompañas este sábado.</h1>'
    + '<p style="font-size:.92rem;line-height:1.7;color:#4a3545;margin:0 0 1.8rem;">Alguien muy especial reservó un lugar para ti en el taller <em>Lo que tu imagen comunica</em> con Mary Méndez. Estamos felices de tenerte el <strong>sábado 15 de agosto</strong>.</p>'

    // Bloque urgente: encuesta previa (lo más importante)
    + '<div style="border:2px solid #C6A56A;padding:1.4rem 1.6rem;margin-bottom:1.8rem;background:rgba(198,165,106,.07);">'
    + '<p style="font-size:.6rem;letter-spacing:.16em;text-transform:uppercase;color:#8F7383;margin:0 0 .5rem;">Antes de llegar — paso obligatorio</p>'
    + '<p style="font-family:Georgia,serif;font-size:1.1rem;color:#2A0F25;margin:0 0 .8rem;">Contesta la encuesta previa desde tu espacio personal.</p>'
    + '<p style="font-size:.87rem;color:#4a3545;line-height:1.7;margin:0 0 1rem;">Es una encuesta breve que Mary revisa personalmente para personalizar la experiencia de cada asistente. <strong>Sin contestarla no podrás hacer check-in el día del evento</strong>, así que te pedimos que la completes antes del sábado.</p>'
    + '<a href="' + hubUrl + '" style="display:block;background:#C6A56A;color:#2A0F25;text-align:center;padding:.9rem 1.2rem;text-decoration:none;font-size:.8rem;letter-spacing:.1em;text-transform:uppercase;font-weight:600;margin-bottom:.6rem;">Ir a mi espacio y contestar la encuesta →</a>'
    + '<p style="font-size:.7rem;color:#8F7383;text-align:center;margin:0;">Aquí también encontrarás tu QR de entrada y toda la info del día</p>'
    + '</div>'

    // Detalles del evento (breve)
    + '<div style="border:1px solid rgba(42,15,37,.1);padding:1.2rem 1.4rem;margin-bottom:1.8rem;">'
    + '<p style="font-size:.6rem;letter-spacing:.16em;text-transform:uppercase;color:#8F7383;margin:0 0 .8rem;">Detalles del evento</p>'
    + '<p style="font-size:.88rem;color:#2A0F25;margin:0 0 .3rem;font-weight:600;">Lo que tu imagen comunica — REINVENTA</p>'
    + '<p style="font-size:.85rem;color:#4a3545;margin:0 0 .25rem;">Sábado 15 de agosto de 2026 · 10:00–12:00 pm</p>'
    + '<p style="font-size:.85rem;color:#4a3545;margin:0;">The University Club of Mexico<br>Av. Paseo de la Reforma 150, Juárez, CDMX</p>'
    + '</div>'

    // Cierre cálido
    + '<p style="font-size:.9rem;line-height:1.7;color:#4a3545;margin:0 0 .5rem;">¡Te esperamos el sábado! Cualquier duda, escríbenos a este correo.</p>'
    + '<p style="font-size:.9rem;color:#4a3545;margin:0;">Con cariño,</p>'
    + '<p style="font-family:Georgia,serif;font-size:.95rem;color:#2A0F25;margin:.2rem 0 0;">Mary y el equipo de REINVENTA</p>'

    + '</div>'
    + _footerCorreo(correo);

  MailApp.sendEmail({
    to: correo,
    bcc: 'alopez@alumbrastudios.com',
    name: 'Reinventa by Mary Méndez',
    subject: '¡Te esperamos este sábado! — Un paso antes de llegar',
    htmlBody: html
  });
}

/* ── Estado de correos (para el tab Correos del admin) ─────── */
function adminEstadoCorreos() {
  var com  = getComunicacionesSheet();
  var asi  = getAsistenciaSheet();
  var cData = com.getDataRange().getValues();
  var aData = asi.getDataRange().getValues();

  // Check-in set por correo
  var checkinSet = {};
  for (var i = 1; i < aData.length; i++) {
    var c = (aData[i][2]||'').toLowerCase().trim();
    if (c && aData[i][4] === '✓') checkinSet[c] = true;
  }

  var total = 0, recEnv = 0, qrEnv = 0, agrEnv = 0, indEnv = 0, conCheckin = 0;
  for (var j = 1; j < cData.length; j++) {
    var correo = (cData[j][0]||'').toLowerCase().trim();
    if (!correo) continue;
    total++;
    if ((cData[j][9] ||'').toString() === 'Sí') recEnv++;
    if ((cData[j][12]||'').toString() === 'Sí') qrEnv++;
    if ((cData[j][15]||'').toString() === 'Sí') agrEnv++;
    if ((cData[j][18]||'').toString() === 'Sí') indEnv++; // col S = indicaciones
    if (checkinSet[correo]) conCheckin++;
  }

  // Estado de triggers guardado en PropertiesService (no requiere scope extra)
  var props = PropertiesService.getScriptProperties();
  var trigRec = props.getProperty('trigger_recordatorio') === 'true';
  var trigAgr = props.getProperty('trigger_agradecimiento') === 'true';
  var trigInd = props.getProperty('trigger_indicaciones') === 'true';

  // Confirmación: col G (índice 6) = 'Sí' cuando se envió bienvenida
  var confEnv = 0;
  for (var k = 1; k < cData.length; k++) {
    if ((cData[k][6]||'').toString() === 'Sí') confEnv++;
  }

  return jsOk({
    total: total,
    confirmacion:   { enviados: confEnv, pendientes: total - confEnv },
    indicaciones:   { enviados: indEnv,  pendientes: total - indEnv,  programado: trigInd, fecha: 'Jue 13 ago 2026 · 12:00 pm' },
    recordatorio:   { enviados: recEnv,  pendientes: total - recEnv,  programado: trigRec, fecha: 'Vie 14 ago 2026 · 12:00 pm' },
    qr:             { enviados: qrEnv,   pendientes: total - qrEnv,   programado: true,    fecha: 'Sáb 15 ago 2026 · 8:00 am' },
    agradecimiento: { enviados: agrEnv,  pendientes: conCheckin - agrEnv, conCheckin: conCheckin, programado: trigAgr, fecha: 'Lun 17 ago 2026 · 10:00 am' }
  });
}

/* ── Recordatorio masivo ───────────────────────────────────── */
// Función ejecutada por el trigger automático
function triggerRecordatorioMasivo() { adminEnviarRecordatorioMasivo(); }

function adminEnviarRecordatorioMasivo() {
  var com   = getComunicacionesSheet();
  var asi   = getAsistenciaSheet();
  var cData = com.getDataRange().getValues();
  var aData = asi.getDataRange().getValues();

  var idPorCorreo = {};
  for (var i = 1; i < aData.length; i++) {
    var c = (aData[i][2]||'').toLowerCase().trim();
    if (c) idPorCorreo[c] = aData[i][0];
  }

  var enviados = 0;
  for (var j = 1; j < cData.length; j++) {
    var correo  = (cData[j][0]||'').toLowerCase().trim();
    var nombre  = cData[j][1] || '';
    var yaEnv   = (cData[j][9]||'').toString();
    if (!correo || yaEnv === 'Sí') continue;
    var id = idPorCorreo[correo] || '';
    try {
      enviarCorreoRecordatorio(nombre, correo, id);
      com.getRange(j+1, 10).setValue('Sí'); // col J
      enviados++;
      Utilities.sleep(800);
    } catch(e) { Logger.log('Error recordatorio ' + correo + ': ' + e); }
  }
  Logger.log('Recordatorios enviados: ' + enviados);
  return jsOk({ enviados: enviados });
}

function adminProgramarRecordatorio() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'triggerRecordatorioMasivo') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('triggerRecordatorioMasivo')
    .timeBased().at(new Date('2026-08-14T18:00:00')).create(); // UTC = 12pm Mexico City (UTC-6)
  PropertiesService.getScriptProperties().setProperty('trigger_recordatorio', 'true');
  return jsOk({ programado: true, fecha: 'Vie 14 ago 2026 · 12:00 pm' });
}

function adminCancelarRecordatorio() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'triggerRecordatorioMasivo') { ScriptApp.deleteTrigger(t); n++; }
  });
  PropertiesService.getScriptProperties().deleteProperty('trigger_recordatorio');
  return jsOk({ cancelado: true, eliminados: n });
}

/* ── Indicaciones masivo (jueves 13 ago 12pm) ──────────────── */
function triggerIndicacionesMasivo() { adminEnviarIndicacionesMasivo(); }

function adminEnviarIndicacionesMasivo() {
  var com   = getComunicacionesSheet();
  var asi   = getAsistenciaSheet();
  var cData = com.getDataRange().getValues();
  var aData = asi.getDataRange().getValues();

  var idPorCorreo = {};
  for (var i = 1; i < aData.length; i++) {
    var c = (aData[i][2]||'').toLowerCase().trim();
    if (c) idPorCorreo[c] = aData[i][0];
  }

  var enviados = 0;
  for (var j = 1; j < cData.length; j++) {
    var correo  = (cData[j][0]||'').toLowerCase().trim();
    var nombre  = cData[j][1] || '';
    var yaEnv   = (cData[j][18]||'').toString(); // col S (índice 18)
    if (!correo || yaEnv === 'Sí') continue;
    var id = idPorCorreo[correo] || '';
    try {
      enviarCorreoIndicaciones(nombre, correo, id);
      com.getRange(j+1, 19).setValue('Sí'); // col S
      enviados++;
      Utilities.sleep(800);
    } catch(e) { Logger.log('Error indicaciones ' + correo + ': ' + e); }
  }
  Logger.log('Indicaciones enviadas: ' + enviados);
  return jsOk({ enviados: enviados });
}

function adminProgramarIndicaciones() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'triggerIndicacionesMasivo') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('triggerIndicacionesMasivo')
    .timeBased().at(new Date('2026-08-13T18:00:00')).create(); // UTC = 12pm Mexico City (UTC-6)
  PropertiesService.getScriptProperties().setProperty('trigger_indicaciones', 'true');
  return jsOk({ programado: true, fecha: 'Jue 13 ago 2026 · 12:00 pm' });
}

function adminCancelarIndicaciones() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'triggerIndicacionesMasivo') { ScriptApp.deleteTrigger(t); n++; }
  });
  PropertiesService.getScriptProperties().deleteProperty('trigger_indicaciones');
  return jsOk({ cancelado: true, eliminados: n });
}

/* ── Agradecimiento masivo (solo asistentes con check-in) ──── */
function triggerAgradecimientoMasivo() { adminEnviarAgradecimientoMasivo(); }

function adminEnviarAgradecimientoMasivo() {
  var com   = getComunicacionesSheet();
  var asi   = getAsistenciaSheet();
  var cData = com.getDataRange().getValues();
  var aData = asi.getDataRange().getValues();

  var idPorCorreo = {}, checkinSet = {};
  for (var i = 1; i < aData.length; i++) {
    var c = (aData[i][2]||'').toLowerCase().trim();
    if (c) {
      idPorCorreo[c] = aData[i][0];
      if (aData[i][4] === '✓') checkinSet[c] = true;
    }
  }

  var enviados = 0;
  for (var j = 1; j < cData.length; j++) {
    var correo = (cData[j][0]||'').toLowerCase().trim();
    var nombre = cData[j][1] || '';
    var yaEnv  = (cData[j][15]||'').toString();
    if (!correo || yaEnv === 'Sí') continue;
    if (!checkinSet[correo]) continue; // Solo con check-in
    var id = idPorCorreo[correo] || '';
    try {
      enviarCorreoAgradecimiento(nombre, correo, id);
      com.getRange(j+1, 16).setValue('Sí'); // col P
      enviados++;
      Utilities.sleep(800);
    } catch(e) { Logger.log('Error agradecimiento ' + correo + ': ' + e); }
  }
  Logger.log('Agradecimientos enviados: ' + enviados);
  return jsOk({ enviados: enviados });
}

function adminProgramarAgradecimiento() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'triggerAgradecimientoMasivo') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('triggerAgradecimientoMasivo')
    .timeBased().at(new Date('2026-08-17T16:00:00')).create(); // UTC = 10am Mexico City (UTC-6)
  PropertiesService.getScriptProperties().setProperty('trigger_agradecimiento', 'true');
  return jsOk({ programado: true, fecha: 'Lun 17 ago 2026 · 10:00 am' });
}

function adminCancelarAgradecimiento() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'triggerAgradecimientoMasivo') { ScriptApp.deleteTrigger(t); n++; }
  });
  PropertiesService.getScriptProperties().deleteProperty('trigger_agradecimiento');
  return jsOk({ cancelado: true, eliminados: n });
}

function adminUpdateNombre(id, nombreNuevo) {
  if (!id || !nombreNuevo) return jsErr('Faltan datos');
  var asi = getAsistenciaSheet();
  var data = asi.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][0]||'').toString().trim() === id) {
      asi.getRange(i+1,2).setValue(nombreNuevo);
      var correo = (data[i][2]||'').toLowerCase().trim();
      if (correo) {
        var com = getComunicacionesSheet();
        var fila = findRowByEmailInSheet(com, correo);
        if (fila) com.getRange(fila,2).setValue(nombreNuevo);
      }
      return jsOk({ nombre: nombreNuevo });
    }
  }
  return jsErr('ID no encontrado');
}

function adminRegistroManual(nombre, correo, telefono, fase) {
  if (!nombre || !correo) return jsErr('Nombre y correo requeridos');
  correo = correo.toLowerCase().trim();
  fase = fase || 'Cortesía';

  var reg = getSheet();
  reg.appendRow([
    new Date(), nombre, correo, telefono||'', '', '', fase, '0.00', new Date(), '✓', 'MANUAL-' + new Date().getTime(), 'admin manual',
    '', '', '', '', ''
  ]);

  actualizarAsistencia(correo, nombre, fase);
  var id = obtenerIdAsistente(correo);
  sincronizarComunicaciones(correo, nombre, telefono||'', '', id);

  return jsOk({ result: 'ok', id: id });
}

function adminTestEmailAcompanante() {
  enviarCorreoBienvenidaAcompanante('Valeria García (PRUEBA acompañante)', 'mejoracontinua@caceca.org', 'Early Bird x2', 'RNV-001');
  return jsOk({ result: 'ok' });
}

function adminTestEmail() {
  enviarCorreoConfirmacion('Valeria García (PRUEBA)', 'mejoracontinua@caceca.org', 'Early Bird', 'RNV-001');
  return jsOk({ result: 'ok' });
}

function jsOk(obj)  { return ContentService.createTextOutput(JSON.stringify(obj || {result:'ok'})).setMimeType(ContentService.MimeType.JSON); }
function jsErr(msg) { return ContentService.createTextOutput(JSON.stringify({error: msg})).setMimeType(ContentService.MimeType.JSON); }
