/**
 * /reservar — Función serverless de Vercel
 *
 * Revisa fechas + lugares vendidos en Stripe y redirige
 * al Payment Link correcto. URL pública para compartir:
 *   reinventabymarymendez.com.mx/reservar
 */

const LINKS = {
  early_bird:   'https://buy.stripe.com/aFa6oG2Mi3ED2505w76Ri00',
  preventa:     'https://buy.stripe.com/7sY8wO86C4IHcJE1fR6Ri01',
  precio_final: 'https://buy.stripe.com/14A8wO86C3ED7pk8Ij6Ri02'
};

/* Montos en centavos (MXN × 100) */
const MONTO_EARLY_BIRD   = 130000;
const MONTO_PREVENTA     = 150000;
const MONTO_PRECIO_FINAL = 170000;
const MONTOS_TALLER      = [MONTO_EARLY_BIRD, MONTO_PREVENTA, MONTO_PRECIO_FINAL];

const LIMITE_EARLY_BIRD = 10;
const LIMITE_TOTAL      = 40;

module.exports = async function handler(req, res) {
  try {
    const key = process.env.STRIPE_SECRET_KEY;
    if (!key) throw new Error('Sin clave de Stripe');

    /* ── Fecha actual en CDMX ─────────────────────────────── */
    const cdmx = new Date(new Date().toLocaleString('en-US', { timeZone: 'America/Mexico_City' }));
    const hoy  = new Date(cdmx.getFullYear(), cdmx.getMonth(), cdmx.getDate());

    const D_PREVENTA = new Date(2026, 6, 22); /* 22 jul */
    const D_FINAL    = new Date(2026, 7, 10); /* 10 ago */
    const D_CIERRE   = new Date(2026, 7, 15); /* 15 ago */

    if (hoy >= D_CIERRE) {
      return res.redirect(302, '/taller-imagen-personal?cupo=agotado');
    }

    /* ── Conteo de pagos en Stripe ───────────────────────── */
    const [earlyBirdCount, totalCount] = await Promise.all([
      contarPagos(key, MONTO_EARLY_BIRD),
      contarPagos(key, MONTOS_TALLER)
    ]);

    /* ── Cupo total agotado ───────────────────────────────── */
    if (totalCount >= LIMITE_TOTAL) {
      return res.redirect(302, '/taller-imagen-personal?cupo=agotado');
    }

    /* ── Decidir fase ────────────────────────────────────── */
    const earlyBirdAgotado = earlyBirdCount >= LIMITE_EARLY_BIRD;

    /* Determinar fase y redirigir a página de reserva con el parámetro */
    var fase;
    if (!earlyBirdAgotado && hoy < D_PREVENTA) fase = 'early_bird';
    else if (hoy < D_FINAL)                    fase = 'preventa';
    else                                        fase = 'precio_final';

    return res.redirect(302, '/reservar.html?fase=' + fase);

  } catch (err) {
    /* Si algo falla, manda a la landing para no dejar a nadie varada */
    console.error('Error /reservar:', err.message);
    return res.redirect(302, '/taller-imagen-personal');
  }
}

/* ── Cuenta checkout sessions completadas por monto ──────────── */
async function contarPagos(key, montos) {
  const filtro = Array.isArray(montos) ? montos : [montos];
  let count = 0;
  let hasMore = true;
  let startingAfter = null;

  while (hasMore) {
    let url = 'https://api.stripe.com/v1/checkout/sessions?status=complete&limit=100';
    if (startingAfter) url += `&starting_after=${startingAfter}`;

    const resp = await fetch(url, {
      headers: { Authorization: `Bearer ${key}` }
    });

    if (!resp.ok) throw new Error(`Stripe API error: ${resp.status}`);
    const data = await resp.json();

    for (const session of data.data) {
      if (filtro.includes(session.amount_total)) count++;
    }

    hasMore = data.has_more;
    if (hasMore) startingAfter = data.data[data.data.length - 1].id;
  }

  return count;
}
