/**
 * /api/encuesta — Guarda respuestas de encuesta de satisfacción
 * POST { id, calificacion, comentario }
 */

const SHEETS_ENDPOINT = 'https://script.google.com/macros/s/AKfycbxY1AuqgR3sonF2MhxsphCVWHQr5pTJg-Qs_xmEHEFnTaK4Q6y_ivFXrhfHUW69or7ymA/exec';

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Método no permitido' });

  const { id, calificacion, comentario } = req.body || {};
  if (!id) return res.status(400).json({ error: 'ID requerido' });

  try {
    const r = await fetch(SHEETS_ENDPOINT, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: 'encuesta', id, calificacion, comentario })
    });
    const data = await r.json();
    return res.status(200).json(data);
  } catch (err) {
    return res.status(500).json({ error: 'Error de conexión' });
  }
};
