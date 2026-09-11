// Proxy serverless (Vercel) → Apps Script.
// El navegador llama a /api/sheet (MISMO origen), así ningún bloqueador de
// anuncios/privacidad puede cortar la conexión. Este servidor reenvía la
// consulta a tu Google Sheet vía Apps Script e inyecta el token del lado
// servidor (no queda expuesto en el HTML público).

const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzOrjuXb77KdjbqsmcnQosKTK8PbXmbLqikLrQXRXQr0axXC3RcZvRaQXOWgPL-rFU/exec";
const TOKEN = "boda-cj-2026";

module.exports = async (req, res) => {
  try {
    const src = (req.method === "POST" && req.body) ? req.body : (req.query || {});
    const action = String(src.action || (req.query && req.query.action) || "");
    let payload = src.payload;
    if (payload == null && req.query) payload = req.query.payload;
    if (payload == null) payload = [];
    if (typeof payload !== "string") payload = JSON.stringify(payload);

    const qs = new URLSearchParams({ action, token: TOKEN, payload });
    const upstream = await fetch(APPS_SCRIPT_URL + "?" + qs.toString(), {
      method: "GET",
      redirect: "follow",
    });
    const text = await upstream.text();

    res.setHeader("Content-Type", "application/json; charset=utf-8");
    res.setHeader("Cache-Control", "no-store");
    try {
      JSON.parse(text);           // Apps Script devolvió JSON válido
      return res.status(200).send(text);
    } catch (e) {
      // Apps Script devolvió HTML (p. ej. página de error / no autorizado)
      return res.status(200).json({ error: "Apps Script no devolvió JSON (revisá el deployment)" });
    }
  } catch (e) {
    res.setHeader("Content-Type", "application/json; charset=utf-8");
    return res.status(200).json({ error: String((e && e.message) || e) });
  }
};
