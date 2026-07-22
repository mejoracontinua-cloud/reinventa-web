/**
 * /api/entrada — Verifica y registra asistencia por ID único
 * GET ?id=RNV-001 → consulta sheet, marca asistencia, devuelve nombre y fase
 */

const SHEETS_ENDPOINT = 'https://script.google.com/macros/s/AKfycbxY1AuqgR3sonF2MhxsphCVWHQr5pTJg-Qs_xmEHEFnTaK4Q6y_ivFXrhfHUW69or7ymA/exec';

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');

  const { id } = req.query;
  if (!id) return res.status(400).json({ error: 'ID requerido' });

  try {
    const url = `${SHEETS_ENDPOINT}?action=entrada&id=${encodeURIComponent(id)}`;
    const r = await fetch(url);
    const data = await r.json();
    return res.status(200).json(data);
  } catch (err) {
    return res.status(500).json({ error: 'Error de conexión con el sheet' });
  }
};
