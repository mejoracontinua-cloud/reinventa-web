/**
 * /api/admin — Panel de administración de REINVENTA
 * GET ?key=PASSWORD → devuelve todos los asistentes con su estado
 * Contraseña: variable ADMIN_KEY o por defecto "reinventa2026"
 */

const SHEETS_ENDPOINT = 'https://script.google.com/macros/s/AKfycbxY1AuqgR3sonF2MhxsphCVWHQr5pTJg-Qs_xmEHEFnTaK4Q6y_ivFXrhfHUW69or7ymA/exec';
const DEFAULT_KEY = 'reinventa2026';

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');

  const { key } = req.query;
  const adminKey = process.env.ADMIN_KEY || DEFAULT_KEY;

  if (key !== adminKey) return res.status(401).json({ error: 'Acceso no autorizado' });

  try {
    const url = `${SHEETS_ENDPOINT}?action=admin`;
    const r = await fetch(url);
    const data = await r.json();
    return res.status(200).json(data);
  } catch (err) {
    return res.status(500).json({ error: 'Error de conexión con el sheet' });
  }
};
