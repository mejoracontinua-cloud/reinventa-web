/**
 * /api/meta/whatsapp/callback — OAuth callback de Meta WhatsApp Embedded Signup
 *
 * GET ?code=...&state=...  → intercambia code por token, obtiene WABA ID / Phone Number ID
 * GET ?error=...            → error del flujo OAuth de Meta
 *
 * Flujo:
 *   1. Valida state HMAC (generado por /api/meta/whatsapp/state)
 *   2. Intercambia code por access_token (solo server-side, NUNCA llega al browser)
 *   3. Consulta Graph API para obtener WABA ID y Phone Number ID
 *   4. Loguea IDs en consola Vercel para configuración manual
 *   5. Redirige a /admin?whatsapp=connected  (o =error si falla)
 *
 * Variables de entorno requeridas:
 *   META_APP_ID       — App ID de Meta
 *   META_APP_SECRET   — App Secret de Meta (NUNCA exponer al frontend)
 *   META_STATE_SECRET — Mismo secreto usado en /state para verificar HMAC
 *   META_WABA_ID      — (Opcional) Si ya está configurado, consulta phone numbers
 *
 * NO retorna tokens ni secrets al browser en ningún caso.
 * NO modifica ninguna funcionalidad existente del proyecto.
 */

const crypto = require('crypto');

const GRAPH_VERSION = 'v21.0';
const REDIRECT_URI  = 'https://www.reinventabymarymendez.com.mx/api/meta/whatsapp/callback';
const STATE_TTL_MS  = 10 * 60 * 1000; // 10 minutos

// ─── Validación de state HMAC-SHA256 ────────────────────────────────────────

function verifyState(state) {
  const secret = process.env.META_STATE_SECRET;
  if (!secret || !state || typeof state !== 'string') return false;

  const dot = state.lastIndexOf('.');
  if (dot === -1 || dot === state.length - 1) return false;

  const payload   = state.slice(0, dot);
  const signature = state.slice(dot + 1);

  // 1. Verificar firma (comparación timing-safe, previene timing attacks)
  const expected = crypto.createHmac('sha256', secret).update(payload).digest('hex');
  try {
    if (
      signature.length !== expected.length ||
      !crypto.timingSafeEqual(Buffer.from(signature, 'hex'), Buffer.from(expected, 'hex'))
    ) {
      return false;
    }
  } catch {
    return false;
  }

  // 2. Verificar que no ha expirado (máx 10 minutos)
  try {
    const decoded = JSON.parse(Buffer.from(payload, 'base64url').toString('utf8'));
    if (!decoded.ts || typeof decoded.ts !== 'number') return false;
    if (Date.now() - decoded.ts > STATE_TTL_MS) return false;
  } catch {
    return false;
  }

  return true;
}

// ─── Handler principal ───────────────────────────────────────────────────────

module.exports = async function handler(req, res) {
  // Anti-cache: nunca cachear un callback OAuth
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, private');
  res.setHeader('Pragma', 'no-cache');
  res.setHeader('X-Content-Type-Options', 'nosniff');

  if (req.method !== 'GET') return res.status(405).end();

  const { code, state, error, error_description } = req.query;

  // ── 1. Error explícito de Meta ──────────────────────────────────────────
  if (error) {
    // No logueamos error_description completo para evitar fugas de datos sensibles
    console.error(`[WA OAuth] Error de Meta: ${error}`);
    return redirect(res, '/admin?whatsapp=error');
  }

  // ── 2. Validar state CSRF antes de cualquier otra acción ────────────────
  if (!verifyState(state)) {
    console.error('[WA OAuth] State inválido o expirado — posible CSRF o flujo reiniciado');
    return redirect(res, '/admin?whatsapp=error');
  }

  // ── 3. Verificar que llegó el code ─────────────────────────────────────
  if (!code || typeof code !== 'string') {
    console.error('[WA OAuth] No se recibió code en el callback');
    return redirect(res, '/admin?whatsapp=error');
  }

  // ── 4. Verificar variables de entorno requeridas ────────────────────────
  const appId     = process.env.META_APP_ID;
  const appSecret = process.env.META_APP_SECRET;

  if (!appId || !appSecret) {
    console.error('[WA OAuth] CRÍTICO: Faltan META_APP_ID o META_APP_SECRET en Vercel env vars');
    return redirect(res, '/admin?whatsapp=error');
  }

  try {
    // ── 5. Intercambiar code por access_token (solo server-side) ──────────
    const tokenParams = new URLSearchParams({
      client_id:     appId,
      client_secret: appSecret,
      code:          code,
      redirect_uri:  REDIRECT_URI,
    });

    const tokenRes  = await fetch(
      `https://graph.facebook.com/${GRAPH_VERSION}/oauth/access_token?${tokenParams}`,
      { method: 'GET' }
    );
    const tokenData = await tokenRes.json();

    if (!tokenRes.ok || tokenData.error || !tokenData.access_token) {
      const msg = tokenData.error?.message ?? `HTTP ${tokenRes.status}`;
      console.error(`[WA OAuth] Token exchange fallido: ${msg}`);
      return redirect(res, '/admin?whatsapp=error');
    }

    // NUNCA logueamos el access_token completo
    console.log(`[WA OAuth] ✓ Token exchange exitoso. Tipo: ${tokenData.token_type ?? 'desconocido'}`);

    const token = tokenData.access_token;

    // ── 6. Obtener WhatsApp Business Accounts del usuario conectado ────────
    const bizRes  = await fetch(
      `https://graph.facebook.com/${GRAPH_VERSION}/me/businesses` +
      `?fields=id,name,whatsapp_business_accounts{id,name,currency,timezone_id}` +
      `&access_token=${token}`
    );
    const bizData = await bizRes.json();

    // Loguear IDs (no tokens) para que los copies en Vercel env vars
    if (bizData.data?.length) {
      bizData.data.forEach(biz => {
        console.log(`[WA OAuth] 📋 Business: ${biz.name} (ID: ${biz.id})`);
        const wabas = biz.whatsapp_business_accounts?.data;
        if (wabas?.length) {
          wabas.forEach(waba => {
            console.log(`[WA OAuth]   └─ WABA ID: ${waba.id} | ${waba.name}`);
          });
        }
      });
    } else {
      console.warn('[WA OAuth] No se encontraron WhatsApp Business Accounts en esta conexión');
      if (bizData.error) {
        console.error(`[WA OAuth] Error Graph API: ${bizData.error.message}`);
      }
    }

    // ── 7. Si ya tenemos META_WABA_ID configurado, listar números de teléfono ──
    const wabaId = process.env.META_WABA_ID;
    if (wabaId) {
      const phonesRes  = await fetch(
        `https://graph.facebook.com/${GRAPH_VERSION}/${wabaId}/phone_numbers` +
        `?fields=id,display_phone_number,verified_name,quality_rating` +
        `&access_token=${token}`
      );
      const phonesData = await phonesRes.json();

      if (phonesData.data?.length) {
        phonesData.data.forEach(phone => {
          console.log(
            `[WA OAuth]   📱 Phone Number ID: ${phone.id}` +
            ` | ${phone.display_phone_number} (${phone.verified_name})`
          );
        });
      } else if (phonesData.error) {
        console.warn(`[WA OAuth] No se pudieron obtener phone numbers: ${phonesData.error.message}`);
      }
    }

    // ── 8. Instrucción para configuración post-conexión ─────────────────────
    console.log('[WA OAuth] ──────────────────────────────────────────────────');
    console.log('[WA OAuth] PRÓXIMOS PASOS: Copia los IDs anteriores en Vercel:');
    console.log('[WA OAuth]   Dashboard → Settings → Environment Variables:');
    console.log('[WA OAuth]   META_WABA_ID         = <WABA ID del log anterior>');
    console.log('[WA OAuth]   META_PHONE_NUMBER_ID = <Phone Number ID del log anterior>');
    console.log('[WA OAuth]   META_WHATSAPP_TOKEN  = <System User Token permanente>');
    console.log('[WA OAuth] ──────────────────────────────────────────────────');

    return redirect(res, '/admin?whatsapp=connected');

  } catch (err) {
    // Error inesperado: logueamos solo el mensaje, nunca el stack completo en producción
    console.error(`[WA OAuth] Error inesperado: ${err.message}`);
    return redirect(res, '/admin?whatsapp=error');
  }
};

// ─── Helper: redirect limpio ─────────────────────────────────────────────────
function redirect(res, location) {
  res.writeHead(302, { Location: location });
  res.end();
}
