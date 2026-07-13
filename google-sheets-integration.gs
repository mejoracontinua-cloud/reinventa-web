/**
 * REINVENTA by Mary Méndez — Integración del formulario de captación con Google Sheets
 *
 * PASO 1. Crea una hoja de cálculo nueva en Google Sheets (ej. "REINVENTA — Leads").
 *
 * PASO 2. En la fila 1 agrega estos encabezados, en este orden:
 *   Fecha | Nombre | Correo | WhatsApp | Servicio | Qué quiere transformar | Cuándo | Contacto preferido | Origen
 *
 * PASO 3. Ve a Extensiones > Apps Script. Borra el código de ejemplo (myFunction)
 *         y pega TODO este archivo en su lugar.
 *
 * PASO 4. Guarda el proyecto (Ctrl/Cmd+S). Ponle un nombre, ej. "REINVENTA Form Handler".
 *
 * PASO 5. Haz clic en "Implementar" (Deploy) > "Nueva implementación".
 *         - Tipo: "Aplicación web" (Web app).
 *         - Descripción: la que quieras.
 *         - Ejecutar como: "Yo" (tu cuenta).
 *         - Quién tiene acceso: "Cualquier usuario" (Anyone).
 *         Haz clic en "Implementar" y autoriza los permisos que pida Google
 *         (te va a advertir que es un script no verificado — es tuyo, es seguro,
 *         solo confirma "Ir a REINVENTA Form Handler (no seguro)" > Permitir).
 *
 * PASO 6. Copia la URL que te entrega (termina en /exec) y pásamela.
 *         Yo la coloco en la constante FORM_ENDPOINT del sitio.
 *
 * Nota: cada vez que edites este script después de la primera implementación,
 * debes crear una "Nueva implementación" otra vez (o gestionar versiones)
 * para que los cambios queden activos en la URL pública.
 */

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = JSON.parse(e.postData.contents);

  sheet.appendRow([
    new Date(),
    data.nombre || '',
    data.correo || '',
    data.whatsapp || '',
    data.servicio || '',
    data.transformar || '',
    data.cuando || '',
    data.contacto || '',
    data.origen || ''
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ result: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}
