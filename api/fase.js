/**
 * /api/fase — Devuelve la fase activa en JSON
 * La landing lo consulta al cargar para actualizar botones en tiempo real.
 */

const MONTO_EARLY_BIRD   = 130000;
const MONTO_PREVENTA     = 150000;
const MONTO_PRECIO_FINAL = 170000;
const MONTOS_TALLER      = [MONTO_EARLY_BIRD, MONTO_PREVENTA, MONTO_PRECIO_FINAL];

const LIMITE_EARLY_BIRD = 10;
const LIMITE_TOTAL      = 40;

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Cache-Control', 'no-store');

  try {
    const key = process.env.STRIPE_SECRET_KEY;
    if (!key) throw new Error('Sin clave de Stripe');

    const cdmx = new Date(new Date().toLocaleString('en-US', { timeZone: 'America/Mexico_City' }));
    const hoy  = new Date(cdmx.getFullYear(), cdmx.getMonth(), cdmx.getDate());

    const D_PREVENTA = new Date(2026, 6, 22);
    const D_FINAL    = new Date(2026, 7, 10);
    const D_CIERRE   = new Date(2026, 7, 15);

    if (hoy >= D_CIERRE) {
      return res.json({ fase: 'agotado' });
    }

    const [earlyBirdCount, totalCount] = await Promise.all([
      contarPagos(key, MONTO_EARLY_BIRD),
      contarPagos(key, MONTOS_TALLER)
    ]);

    if (totalCount >= LIMITE_TOTAL) {
      return res.json({ fase: 'agotado' });
    }

    const earlyBirdAgotado = earlyBirdCount >= LIMITE_EARLY_BIRD;

    let fase;
    if (!earlyBirdAgotado && hoy < D_PREVENTA) fase = 'early_bird';
    else if (hoy < D_FINAL)                    fase = 'preventa';
    else                                        fase = 'precio_final';

    return res.json({ fase, earlyBirdCount, totalCount });

  } catch (err) {
    console.error('Error /api/fase:', err.message);
    return res.status(500).json({ error: 'error' });
  }
}

async function contarPagos(key, montos) {
  const filtro = Array.isArray(montos) ? montos : [montos];
  let count = 0, hasMore = true, startingAfter = null;
  while (hasMore) {
    let url = 'https://api.stripe.com/v1/checkout/sessions?status=complete&limit=100';
    if (startingAfter) url += `&starting_after=${startingAfter}`;
    const resp = await fetch(url, { headers: { Authorization: `Bearer ${key}` } });
    if (!resp.ok) throw new Error(`Stripe API error: ${resp.status}`);
    const data = await resp.json();
    for (const s of data.data) {
      if (filtro.includes(s.amount_total)) count++;
    }
    hasMore = data.has_more;
    if (hasMore) startingAfter = data.data[data.data.length - 1].id;
  }
  return count;
}
