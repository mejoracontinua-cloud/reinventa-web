/**
 * /api/meta/whatsapp/state — Genera un state CSRF firmado para Embedded Signup
 *
 * POST (sin body) → { state, appId, redirectUri }
 *
 * El state es un token HMAC-SHA256 stateless:
 *   state = base64url(JSON{ts, nonce}) + "." + hmac_sha256(payload, META_STATE_SECRET)
 *
 * No requiere base de datos ni sesiones. Se valida criptográficamente en /callback.
 *
 * Variables de entorno requeridas:
 *   META_APP_ID       — App ID de Meta (público, aparece en OAuth URLs)
 *   META_STATE_SECRET — Secreto para firmar el state (mínimo 32 chars aleatorios)
 */

const crypto = require('crypto');

const REDIRECT_URI = 'https://www.reinventabymarymendez.com.mx/api/meta/whatsapp/callback';

module.exports = async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store, private');
  res.setHeader('Access-Control-Allow-Origin', 'https://www.reinventabymarymendez.com.mx');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Método no permitido' });

  const appId = process.env.META_APP_ID;
  const stateSecret = process.env.META_STATE_SECRET;

  if (!appId || !stateSecret) {
    console.error('[WA state] Faltan META_APP_ID o META_STATE_SECRET en env vars');
    return res.status(500).json({ error: 'Servidor no configurado para WhatsApp' });
  }

  // Payload: timestamp + nonce aleatorio (16 bytes hex)
  const payload = Buffer.from(
    JSON.stringify({ ts: Date.now(), nonce: crypto.randomBytes(16).toString('hex') })
  ).toString('base64url');

  // Firma HMAC-SHA256
  const signature = crypto.createHmac('sha256', stateSecret).update(payload).digest('hex');

  const state = `${payload}.${signature}`;

  return res.status(200).json({
    state,
    appId,
    redirectUri: REDIRECT_URI,
  });
};
