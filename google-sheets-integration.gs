/**
 * REINVENTA by Mary Méndez — Apps Script
 *
 * Maneja dos tipos de entrada:
 *   1. POST desde el popup de la landing (JSON con nombre, correo, etc.)
 *   2. Webhook de Stripe (evento checkout.session.completed)
 *
 * ESTRUCTURA DEL SHEET (fila 1 = encabezados):
 *   A  Fecha registro
 *   B  Nombre
 *   C  Correo
 *   D  WhatsApp
 *   E  Contacto preferido
 *   F  ¿Qué busca?
 *   G  Fase de compra
 *   H  Monto pagado (MXN)
 *   I  Fecha de pago
 *   J  Pagó ✓
 *   K  Stripe Payment ID
 *   L  Origen
 *   M  Acepta marketing
 *   N  Autoriza uso de imagen
 *   O  Correo confirmación enviado
 *   P  Canal UTM (utm_source)
 *   Q  ¿Cómo se enteró?
 *   R  ¿A qué se dedica?
 *
 * INSTRUCCIONES DE DESPLIEGUE:
 *   1. Pega este código en Apps Script (Extensiones > Apps Script).
 *   2. Guarda y haz "Nueva implementación" > Aplicación web.
 *      - Ejecutar como: Yo
 *      - Acceso: Cualquier usuario
 *   3. Copia la URL /exec — es la misma que ya tienes en la landing
 *      y que pondrás como webhook en Stripe.
 *   4. En Stripe: Developers > Webhooks > Add endpoint
 *      - URL: la misma URL /exec
 *      - Evento: checkout.session.completed
 */

var SHEET_NAME = 'Registros'; // Cambia si tu hoja tiene otro nombre

/* ── Correos de notificación ────────────────────────────────── */
var EMAILS_NOTIFICACION = ['mejoracontinua@caceca.org', 'alopez@alumbrastudios.com'];
var LIMITE_EARLY_BIRD   = 10;
var LIMITE_TOTAL        = 40;

function doPost(e) {
  try {
    var raw = e.postData.contents;
    var data = JSON.parse(raw);

    // Detectar si es un webhook de Stripe (tiene objeto "type")
    if (data.type && data.type === 'checkout.session.completed') {
      return handleStripeWebhook(data);
    }

    // Si no, es el popup de la landing
    return handleFormSubmit(data);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/* ── Popup de la landing ─────────────────────────────────────── */
function handleFormSubmit(data) {
  var sheet = getSheet();
  var correo = (data.correo || '').toLowerCase().trim();

  // Buscar si ya existe ese correo para no duplicar
  var existingRow = findRowByEmail(sheet, correo);

  if (existingRow) {
    // Actualizar datos del popup si ya existe (vino antes de Stripe)
    sheet.getRange(existingRow, 2).setValue(data.nombre      || sheet.getRange(existingRow, 2).getValue());
    sheet.getRange(existingRow, 4).setValue(data.whatsapp    || sheet.getRange(existingRow, 4).getValue());
    sheet.getRange(existingRow, 5).setValue(data.contacto    || sheet.getRange(existingRow, 5).getValue());
    sheet.getRange(existingRow, 6).setValue(data.transformar || sheet.getRange(existingRow, 6).getValue());
    sheet.getRange(existingRow, 12).setValue('landing + stripe');
    sheet.getRange(existingRow, 13).setValue(data.marketing       || '');
    sheet.getRange(existingRow, 14).setValue(data.autoriza_imagen || '');
    if (data.canal_utm)      sheet.getRange(existingRow, 16).setValue(data.canal_utm);
    if (data.como_se_entero) sheet.getRange(existingRow, 17).setValue(data.como_se_entero);
    if (data.ocupacion)      sheet.getRange(existingRow, 18).setValue(data.ocupacion);
  } else {
    sheet.appendRow([
      new Date(),               // A Fecha registro
      data.nombre     || '',    // B Nombre
      correo,                   // C Correo
      data.whatsapp   || '',    // D WhatsApp
      data.contacto   || '',    // E Contacto preferido
      data.transformar|| '',    // F ¿Qué busca?
      data.servicio   || '',    // G Fase
      '',                       // H Monto
      '',                       // I Fecha pago
      '',                       // J Pagó
      '',                       // K Stripe ID
      data.origen     || 'landing', // L Origen
      data.marketing       || '',   // M Acepta marketing
      data.autoriza_imagen || '',   // N Autoriza imagen
      '',                           // O Correo confirmación enviado
      data.canal_utm       || '',   // P Canal UTM
      data.como_se_entero  || '',   // Q ¿Cómo se enteró?
      data.ocupacion       || ''    // R ¿A qué se dedica?
    ]);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ result: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ── Webhook de Stripe ───────────────────────────────────────── */
function handleStripeWebhook(event) {
  var session = event.data.object;
  var correo  = (session.customer_details && session.customer_details.email
                  ? session.customer_details.email
                  : (session.customer_email || '')).toLowerCase().trim();
  var nombre  = session.customer_details && session.customer_details.name
                  ? session.customer_details.name : '';
  var monto   = session.amount_total ? (session.amount_total / 100).toFixed(2) : '';
  var fecha   = session.created ? new Date(session.created * 1000) : new Date();
  var stripeId = session.id || '';

  // Detectar fase por monto
  var fase = 'Taller';
  if      (monto == '1300.00') fase = 'Early Bird';
  else if (monto == '1500.00') fase = 'Preventa';
  else if (monto == '1700.00') fase = 'Últimos lugares';

  /* ── Correo de confirmación — se envía primero para garantizar entrega ── */
  if (correo) {
    enviarCorreoConfirmacion(nombre, correo, fase);
  }

  var sheet = getSheet();

  /* Si este Stripe ID ya está registrado, ignorar el reintento */
  if (stripeId && findRowByStripeId(sheet, stripeId)) {
    return ContentService.createTextOutput(JSON.stringify({ result: 'duplicate' })).setMimeType(ContentService.MimeType.JSON);
  }

  var existingRow = findRowByEmail(sheet, correo);

  if (existingRow) {
    var yaPago = sheet.getRange(existingRow, 10).getValue() === '✓';

    if (!yaPago) {
      // Existe por popup, aún sin pago — completar normalmente
      if (nombre) sheet.getRange(existingRow, 2).setValue(nombre);
      sheet.getRange(existingRow, 7).setValue(fase);
      sheet.getRange(existingRow, 8).setValue(monto);
      sheet.getRange(existingRow, 9).setValue(fecha);
      sheet.getRange(existingRow, 10).setValue('✓');
      sheet.getRange(existingRow, 11).setValue(stripeId);
      sheet.getRange(existingRow, 12).setValue('landing + stripe');
    } else {
      // Ya tiene ✓ — crear nueva fila para este pago adicional
      sheet.appendRow([
        new Date(), nombre, correo, '', '', '', fase, monto, fecha, '✓', stripeId, 'stripe directo'
      ]);
    }
  } else {
    // Pagó directo sin pasar por el popup
    sheet.appendRow([
      new Date(),  // A Fecha registro
      nombre,      // B Nombre
      correo,      // C Correo
      '',          // D WhatsApp (no disponible)
      '',          // E Contacto preferido
      '',          // F ¿Qué busca?
      fase,        // G Fase
      monto,       // H Monto
      fecha,       // I Fecha pago
      '✓',         // J Pagó
      stripeId,    // K Stripe ID
      'stripe directo'  // L Origen
    ]);
  }

  /* ── Revisar hito sold out ───────────────────────────────────── */
  var sheet2 = getSheet();
  var totalPagos = contarPagosSheet(sheet2);
  if (totalPagos === LIMITE_TOTAL) {
    var props2 = PropertiesService.getScriptProperties();
    if (!props2.getProperty('sold_out_enviado')) {
      notificarCupoAgotado(sheet2);
      props2.setProperty('sold_out_enviado', 'true');
    }
  }

  /* ── Marcar correo enviado en columna O ──────────────────────── */
  if (correo) {
    marcarCorreoEnviado(sheet, correo);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ result: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ── Correo de confirmación post-pago ───────────────────────── */
function enviarCorreoConfirmacion(nombre, correo, fase) {
  var primerNombre = nombre ? nombre.split(' ')[0] : 'Hola';

  var calLink = 'https://calendar.google.com/calendar/render?action=TEMPLATE'
    + '&text=Taller+de+imagen+y+liderazgo+%E2%80%94+REINVENTA'
    + '&dates=20260815T160000Z/20260815T180000Z'
    + '&details=Taller+Lo+que+tu+imagen+comunica+%7C+REINVENTA+by+Mary+M%C3%A9ndez'
    + '&location=The+University+Club+of+Mexico%2C+Av.+Paseo+de+la+Reforma+150%2C+Ju%C3%A1rez%2C+CDMX';

  var asunto = 'Tu lugar en el taller está confirmado ✦ REINVENTA';

  var html = '<div style="background:#E8E2DB;padding:2rem 1rem;font-family:\'Gill Sans\',Calibri,\'Segoe UI\',sans-serif;">'
    + '<div style="max-width:540px;margin:0 auto;background:#EFE9E2;box-shadow:0 4px 40px rgba(42,15,37,.13);">'

    // Header
    + '<div style="background:#2A0F25;padding:2rem 2.4rem 1.6rem;text-align:center;">'
    + '<span style="font-family:Georgia,serif;font-weight:400;font-size:1rem;letter-spacing:.22em;text-transform:uppercase;color:#C6A56A;display:block;margin-bottom:.2rem;">Reinventa</span>'
    + '<span style="font-size:.62rem;letter-spacing:.16em;text-transform:uppercase;color:rgba(198,165,106,.5);">by Mary Méndez</span>'
    + '</div>'
    + '<div style="height:2px;background:#C6A56A;opacity:.45;"></div>'

    // Body
    + '<div style="padding:2.2rem 2.6rem 2rem;">'

    // Badge
    + '<div style="display:inline-block;background:rgba(42,15,37,.07);border-left:2px solid #C6A56A;padding:.4rem .75rem;font-size:.65rem;letter-spacing:.13em;text-transform:uppercase;color:#2A0F25;margin-bottom:1.4rem;">Pago confirmado &middot; ' + fase + '</div>'

    // Saludo
    + '<h1 style="font-family:Georgia,serif;font-weight:400;font-size:1.5rem;line-height:1.35;color:#2A0F25;margin:0 0 1rem;">Tu lugar está reservado,<br>' + primerNombre + '.</h1>'
    + '<p style="font-size:.92rem;line-height:1.7;color:#4a3545;margin:0 0 1.8rem;">Nos da mucho gusto tenerte en el taller. Mary estará encantada de acompañarte en este proceso. Guarda la fecha en tu calendario para que no se te pase ningún detalle.</p>'

    // Tarjeta evento
    + '<div style="border:1px solid rgba(42,15,37,.12);padding:1.4rem 1.6rem;margin-bottom:1.6rem;">'
    + '<p style="font-size:.6rem;letter-spacing:.16em;text-transform:uppercase;color:#8F7383;margin:0 0 .8rem;">Detalles del evento</p>'
    + '<p style="font-family:Georgia,serif;font-size:1.05rem;font-weight:400;color:#2A0F25;margin:0 0 .1rem;">Taller de imagen y liderazgo</p>'
    + '<p style="font-size:.78rem;color:#8F7383;font-style:italic;margin:0 0 1rem;">Lo que tu imagen comunica</p>'
    + '<p style="font-size:.85rem;color:#4a3545;margin:0 0 .5rem;">📅 &nbsp;Sábado 15 de agosto de 2026 &middot; 10:00–12:00 pm</p>'
    + '<p style="font-size:.85rem;color:#4a3545;margin:0;">📍 &nbsp;The University Club of Mexico<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#8F7383;font-size:.8rem;">Av. Paseo de la Reforma 150, Juárez, CDMX</span></p>'
    + '</div>'

    // Botón Google Calendar
    + '<a href="' + calLink + '" style="display:block;background:#2A0F25;color:#EFE9E2;text-align:center;padding:.9rem 1.2rem;text-decoration:none;font-size:.8rem;letter-spacing:.1em;text-transform:uppercase;margin-bottom:1.6rem;">&#128197; &nbsp;Agregar a Google Calendar</a>'

    // Divider
    + '<hr style="border:none;border-top:1px solid rgba(42,15,37,.1);margin:0 0 1.4rem;" />'

    + '<p style="font-size:.87rem;line-height:1.7;color:#4a3545;margin:0 0 .5rem;">Si tienes alguna pregunta antes del taller, puedes escribirle directamente a Mary. Nos vemos el 15 de agosto.</p>'
    + '<p style="font-family:Georgia,serif;font-size:.98rem;color:#2A0F25;margin:.8rem 0 .1rem;">Mary Méndez</p>'
    + '<p style="font-size:.72rem;color:#8F7383;margin:0;">Consultora de imagen y liderazgo</p>'
    + '</div>'

    // Footer
    + '<div style="background:#2A0F25;padding:1.1rem 2rem;text-align:center;">'
    + '<p style="font-size:.63rem;letter-spacing:.07em;color:rgba(198,165,106,.5);line-height:1.7;margin:0;">REINVENTA by Mary Méndez &middot; Ciudad de México<br>'
    + 'Este correo fue enviado a ' + correo + ' porque realizaste un pago.</p>'
    + '</div>'
    + '</div></div>';

  MailApp.sendEmail({
    to: correo,
    bcc: 'alopez@alumbrastudios.com',
    name: 'Reinventa by Mary Méndez',
    subject: asunto,
    htmlBody: html
  });
}

/* ── Correo de prueba a mejoracontinua@caceca.org ───────────── */
function enviarCorreoPrueba() {
  enviarCorreoConfirmacion('Valeria García', 'mejoracontinua@caceca.org', 'Early Bird');
}

/* ── Debug: prueba escritura de columnas P, Q, R ────────────── */
function probarColumnasPQR() {
  var sheet = getSheet();
  handleFormSubmit({
    nombre:          'TEST columnas PQR',
    correo:          'test-pqr-debug@reinventa.mx',
    whatsapp:        '5500000000',
    contacto:        'WhatsApp',
    transformar:     'Prueba debug',
    servicio:        'Taller de imagen y liderazgo — Preventa',
    marketing:       'No',
    autoriza_imagen: 'Sí',
    origen:          'debug',
    canal_utm:       'instagram',
    como_se_entero:  'Otro: boca a boca debug',
    ocupacion:       'Otro: coach de negocios debug'
  });
  Logger.log('✅ Fila de prueba insertada. Revisa el sheet.');
}

/* ── Enviar correo de confirmación a quienes no lo han recibido ─ */
function enviarConfirmacionExistentes() {
  var sheet = getSheet();
  var data  = sheet.getDataRange().getValues();
  var enviados = 0;
  var correosYaEnviados = {}; // evita duplicados por filas repetidas del mismo correo

  for (var i = 1; i < data.length; i++) {
    var pago      = data[i][9];  // J — Pagó ✓
    var nombre    = data[i][1];  // B
    var correo    = (data[i][2] || '').toLowerCase().trim(); // C
    var fase      = data[i][6];  // G
    var yaEnviado = data[i][14]; // O — Correo confirmación enviado

    if (pago === '✓' && correo && yaEnviado !== 'Sí' && !correosYaEnviados[correo]) {
      enviarCorreoConfirmacion(nombre, correo, fase);
      sheet.getRange(i + 1, 15).setValue('Sí');
      correosYaEnviados[correo] = true;
      enviados++;
      Utilities.sleep(1000);
    }
  }

  Logger.log('Correos enviados: ' + enviados);
}

/* ── Marcar columna O en webhook automático ──────────────────── */
function marcarCorreoEnviado(sheet, correo) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][2] || '').toLowerCase().trim() === correo.toLowerCase().trim()) {
      sheet.getRange(i + 1, 15).setValue('Sí');
      return;
    }
  }
}

/* ── Notificaciones ──────────────────────────────────────────── */

function notificarEarlyBirdAgotado(totalActual) {
  var asunto = '🎉 REINVENTA — Early Bird agotado (10/10 lugares)';
  var cuerpo = '¡Se agotaron los 10 lugares de Early Bird!\n\n'
    + 'Lugares vendidos en total: ' + totalActual + ' de ' + LIMITE_TOTAL + '\n'
    + 'A partir de ahora el link /reservar redirige automáticamente a Preventa ($1,500 MXN).\n\n'
    + 'Puedes revisar todos los registros en tu hoja de cálculo.';
  enviarCorreo(asunto, cuerpo);
}

function notificarCupoAgotado(sheet) {
  var totalPagos = contarPagosSheet(sheet);
  var eb  = contarFaseSheet(sheet, 'Early Bird');
  var pre = contarFaseSheet(sheet, 'Preventa');
  var fin = contarFaseSheet(sheet, 'Últimos lugares');
  var recaudado = (eb * 1300) + (pre * 1500) + (fin * 1700);

  var asunto = '🏆 REINVENTA — ¡Cupo completo! 40/40 lugares vendidos';
  var cuerpo = '¡SOLD OUT! Se vendieron los 40 lugares del taller.\n\n'
    + '— Early Bird:    ' + eb  + ' personas ($1,300)\n'
    + '— Preventa:      ' + pre + ' personas ($1,500)\n'
    + '— Precio Final:  ' + fin + ' personas ($1,700)\n\n'
    + 'Total recaudado estimado: $' + recaudado.toLocaleString('es-MX') + ' MXN\n\n'
    + 'El link /reservar ya muestra "cupo agotado" automáticamente.';
  enviarCorreo(asunto, cuerpo);
}

function notificarCambioDeFase(fase) {
  var textos = {
    preventa: {
      asunto: '📅 REINVENTA — Hoy arranca Preventa',
      cuerpo: 'Hoy 22 de julio arranca la fase de Preventa ($1,500 MXN).\n\n'
        + 'El link /reservar ya redirige automáticamente a Preventa.\n'
        + 'Es buen momento para activar comunicación en redes y WhatsApp.'
    },
    precio_final: {
      asunto: '📅 REINVENTA — Hoy arranca Precio Final',
      cuerpo: 'Hoy 10 de agosto arranca la fase de Precio Final ($1,700 MXN).\n\n'
        + 'El link /reservar ya redirige automáticamente a Precio Final.\n'
        + 'Quedan 5 días para el cierre de ventas (15 ago).'
    }
  };
  var t = textos[fase];
  if (t) enviarCorreo(t.asunto, t.cuerpo);
}

/* Disparadores de fecha — ejecutar UNA sola vez para registrarlos */
function configurarTriggersDeFecha() {
  // Eliminar triggers anteriores de estas funciones para no duplicar
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (fn === 'triggerPreventa' || fn === 'triggerPrecioFinal') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 22 jul 2026 08:00 CDMX — inicio Preventa
  var fechaPreventa = new Date('2026-07-22T08:00:00-06:00');
  ScriptApp.newTrigger('triggerPreventa')
    .timeBased()
    .at(fechaPreventa)
    .create();

  // 10 ago 2026 08:00 CDMX — inicio Precio Final
  var fechaFinal = new Date('2026-08-10T08:00:00-06:00');
  ScriptApp.newTrigger('triggerPrecioFinal')
    .timeBased()
    .at(fechaFinal)
    .create();
}

function triggerPreventa() {
  // Solo enviar si el Early Bird NO se agotó antes (si ya se agotó, ya llegó ese correo)
  var sheet = getSheet();
  var totalEarlyBird = contarFaseSheet(sheet, 'Early Bird');
  if (totalEarlyBird < LIMITE_EARLY_BIRD) {
    notificarCambioDeFase('preventa');
  }
}
function triggerPrecioFinal() { notificarCambioDeFase('precio_final'); }

function enviarCorreo(asunto, cuerpo) {
  EMAILS_NOTIFICACION.forEach(function(email) {
    MailApp.sendEmail(email, asunto, cuerpo);
  });
}

/* ── Helpers de conteo en Sheet ──────────────────────────────── */
function contarPagosSheet(sheet) {
  var data = sheet.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    if (data[i][9] === '✓') count++; // columna J
  }
  return count;
}

function contarFaseSheet(sheet, fase) {
  var data = sheet.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    if (data[i][9] === '✓' && data[i][6] === fase) count++; // J=pagó, G=fase
  }
  return count;
}

/* ── Helpers ─────────────────────────────────────────────────── */
function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.getActiveSheet();
  return sheet;
}

function findRowByEmail(sheet, correo) {
  if (!correo) return null;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][2] || '').toString().toLowerCase().trim() === correo) {
      return i + 1;
    }
  }
  return null;
}

/* Busca si un Stripe ID ya existe en columna K */
function findRowByStripeId(sheet, stripeId) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][10] || '').toString().trim() === stripeId) return i + 1;
  }
  return null;
}

