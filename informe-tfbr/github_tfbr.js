// Guarda y lee en GitHub para informe-tfbr: los 4 maestros (.xlsx) y el estado.json con
// el historial de cierres. Mismo patrón que informe-c/github_c.js (lecturas sin caché,
// reintento en 409 releyendo el sha real), generalizado a 4 archivos en vez de 1-2.

const GHT_SETTINGS_KEY = "informe_tfbr_gh_settings";
const GHT_CARPETA_DEFECTO = "informe-tfbr";
const GHT_ARCHIVO_ESTADO = "estado_tfbr.json";

// archivoId -> nombre de archivo persistido en GitHub
const GHT_ARCHIVOS_MAESTRO = {
  balance_mensual_ars: "base_bm_ars.xlsx",
  balance_mensual_brl: "base_bm_brl.xlsx",
  balance_acumulado_ars: "base_ba_ars.xlsx",
  balance_acumulado_brl: "base_ba_brl.xlsx",
};

function loadGhtSettings() {
  try { return JSON.parse(localStorage.getItem(GHT_SETTINGS_KEY) || "{}"); }
  catch (e) { return {}; }
}
function saveGhtSettings(s) { localStorage.setItem(GHT_SETTINGS_KEY, JSON.stringify(s)); }
function hasGhtSettings() { const s = loadGhtSettings(); return !!(s.token && s.repo); }

function ghtRuta(nombre) {
  const c = (loadGhtSettings().carpeta || GHT_CARPETA_DEFECTO).replace(/^\/+|\/+$/g, "");
  return c ? `${c}/${nombre}` : nombre;
}

function ghtUtf8ToBase64(str) {
  const bytes = new TextEncoder().encode(str);
  let b = "";
  bytes.forEach(x => { b += String.fromCharCode(x); });
  return btoa(b);
}
function ghtBase64ToUtf8(b64) {
  const bin = atob(b64.replace(/\s/g, ""));
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return new TextDecoder().decode(bytes);
}
function ghtBufferABase64(buffer) {
  const bytes = new Uint8Array(buffer);
  let b = "";
  for (let i = 0; i < bytes.length; i += 0x8000) {
    b += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  }
  return btoa(b);
}
function ghtBase64ABuffer(b64) {
  const bin = atob(b64.replace(/\s/g, ""));
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes.buffer;
}

function ghtCabeceras() {
  const s = loadGhtSettings();
  if (!s.token || !s.repo) throw new Error("Falta configurar GitHub (token y repositorio).");
  return { "Authorization": `token ${s.token}`, "Accept": "application/vnd.github+json" };
}

async function ghtLeer(nombre) {
  const s = loadGhtSettings();
  const ruta = ghtRuta(nombre);
  const url = `https://api.github.com/repos/${s.repo}/contents/${encodeURI(ruta)}` +
              `?ref=${encodeURIComponent(s.rama || "main")}&_=${Date.now()}`;
  const res = await fetch(url, { headers: ghtCabeceras(), cache: "no-store" });
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`No pude leer ${ruta} de GitHub (${res.status}): ${await res.text()}`);
  const data = await res.json();
  return { contenidoBase64: data.content, sha: data.sha };
}

function ghtPut(ruta, contenidoBase64, sha, mensaje, rama) {
  const s = loadGhtSettings();
  const cuerpo = { message: mensaje, content: contenidoBase64, branch: rama };
  if (sha) cuerpo.sha = sha;
  return fetch(`https://api.github.com/repos/${s.repo}/contents/${encodeURI(ruta)}`, {
    method: "PUT",
    headers: { ...ghtCabeceras(), "Content-Type": "application/json" },
    body: JSON.stringify(cuerpo),
    cache: "no-store",
  });
}

async function ghtEscribir(nombre, contenidoBase64, sha, mensaje) {
  const s = loadGhtSettings();
  const ruta = ghtRuta(nombre);
  const rama = s.rama || "main";
  let res = await ghtPut(ruta, contenidoBase64, sha, mensaje, rama);
  if (res.status === 409) {
    const actual = await ghtLeer(nombre);
    res = await ghtPut(ruta, contenidoBase64, actual ? actual.sha : undefined, mensaje, rama);
    if (res.status === 409) {
      throw new Error(
        `${ruta} cambió en GitHub mientras se guardaba, así que no lo piso. ` +
        `Actualizá la página (Ctrl+Shift+R) y volvé a intentarlo. NO se guardó nada.`
      );
    }
  }
  if (!res.ok) throw new Error(`No pude guardar ${ruta} en GitHub (${res.status}): ${await res.text()}`);
  return await res.json();
}

async function ghtLeerEstado() {
  const r = await ghtLeer(GHT_ARCHIVO_ESTADO);
  return r ? { estado: JSON.parse(ghtBase64ToUtf8(r.contenidoBase64)), sha: r.sha } : null;
}
async function ghtLeerMaestro(archivoId) {
  const nombre = GHT_ARCHIVOS_MAESTRO[archivoId];
  if (!nombre) throw new Error(`archivoId desconocido: ${archivoId}`);
  const r = await ghtLeer(nombre);
  return r ? { buffer: ghtBase64ABuffer(r.contenidoBase64), sha: r.sha } : null;
}
async function ghtGuardarMaestro(archivoId, buffer, mensaje) {
  const nombre = GHT_ARCHIVOS_MAESTRO[archivoId];
  if (!nombre) throw new Error(`archivoId desconocido: ${archivoId}`);
  const sha = (await ghtLeer(nombre))?.sha;
  await ghtEscribir(nombre, ghtBufferABase64(buffer), sha, mensaje);
}
async function ghtGuardarEstado(estado, mensaje) {
  const sha = (await ghtLeer(GHT_ARCHIVO_ESTADO))?.sha;
  await ghtEscribir(GHT_ARCHIVO_ESTADO, ghtUtf8ToBase64(JSON.stringify(estado, null, 1)), sha, mensaje);
}

// Aprobación agrupada (decisión del usuario): guarda los 4 maestros + el estado en una
// sola operación lógica. El maestro más chico primero no importa acá como en informe-c
// (ahí el orden defendía contra un corte a mitad de subida); acá se guarda cada uno con
// su propio sha, así que un fallo parcial dejaría un mensaje claro de cuál archivo faltó
// y se puede reintentar solo ese.
async function ghtGuardarTodosLosMaestros({ buffers, estado, mensaje }) {
  const guardados = [];
  for (const archivoId of Object.keys(GHT_ARCHIVOS_MAESTRO)) {
    if (!buffers[archivoId]) continue;
    await ghtGuardarMaestro(archivoId, buffers[archivoId], mensaje);
    guardados.push(archivoId);
  }
  await ghtGuardarEstado(estado, mensaje);
  return guardados;
}

if (typeof module !== "undefined") {
  module.exports = {
    GHT_ARCHIVOS_MAESTRO,
    loadGhtSettings, saveGhtSettings, hasGhtSettings,
    ghtLeer, ghtEscribir, ghtLeerEstado, ghtLeerMaestro, ghtGuardarMaestro,
    ghtGuardarEstado, ghtGuardarTodosLosMaestros,
  };
}
