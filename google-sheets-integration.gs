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
    sheet.getRange(existingRow, 2).setValue(data.nombre    || sheet.getRange(existingRow, 2).getValue());
    sheet.getRange(existingRow, 4).setValue(data.whatsapp  || sheet.getRange(existingRow, 4).getValue());
    sheet.getRange(existingRow, 5).setValue(data.contacto  || sheet.getRange(existingRow, 5).getValue());
    sheet.getRange(existingRow, 6).setValue(data.transformar || sheet.getRange(existingRow, 6).getValue());
    sheet.getRange(existingRow, 12).setValue('landing + stripe');
  } else {
    sheet.appendRow([
      new Date(),        // A Fecha registro
      data.nombre     || '',
      correo,            // C Correo
      data.whatsapp   || '',
      data.contacto   || '',
      data.transformar|| '',
      data.servicio   || '',  // G Fase (ej. "Early Bird")
      '',                // H Monto
      '',                // I Fecha pago
      '',                // J Pagó
      '',                // K Stripe ID
      data.origen     || 'landing'
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

  var sheet = getSheet();
  var existingRow = findRowByEmail(sheet, correo);

  if (existingRow) {
    // Ya existe por popup — completar con datos de pago
    if (nombre) sheet.getRange(existingRow, 2).setValue(nombre);
    sheet.getRange(existingRow, 7).setValue(fase);
    sheet.getRange(existingRow, 8).setValue(monto);
    sheet.getRange(existingRow, 9).setValue(fecha);
    sheet.getRange(existingRow, 10).setValue('✓');
    sheet.getRange(existingRow, 11).setValue(stripeId);
    sheet.getRange(existingRow, 12).setValue('landing + stripe');
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

  return ContentService
    .createTextOutput(JSON.stringify({ result: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
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
      return i + 1; // fila 1-indexed
    }
  }
  return null;
}
