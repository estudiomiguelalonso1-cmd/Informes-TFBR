// Panel "Configurar cuentas": qué cuenta alimenta cada rótulo del Anexo II, y cómo cambiarlo.
//
// Lo importante de cómo funciona: el rótulo que muestra sale de LAS FÓRMULAS del archivo, no
// de una tabla. Las fórmulas son las que deciden dónde cae el importe; una tabla que dijera
// otra cosa daría la sensación de estar configurado cuando no lo está. Y cambiar el rótulo acá
// REESCRIBE la fórmula: saca la cuenta del renglón donde estaba y la engancha en el nuevo.
//
// Por eso también se ven las cuentas que no lee nadie y las que leen dos renglones: son las dos
// formas en que un importe se pierde o se cuenta doble sin que el balance deje de cerrar.
//
// Los cambios se acumulan en memoria sobre una copia del maestro y recién se suben a GitHub al
// apretar "Guardar". Hasta entonces no se toca nada.

const CFG_ESTADOS = {
  ok:      { texto: "Configurada",     clase: "ok" },
  sin:     { texto: "Sin rótulo",      clase: "bad" },
  varias:  { texto: "En varios",       clase: "bad" },
};

let cfgArchivo = null;     // archivoId que se está mirando
let cfgWb = null;          // copia del maestro sobre la que se edita
let cfgCambios = [];       // qué se hizo, para el commit y para mostrarlo
let cfgFiltro = "";
let cfgEditando = null;    // código de la cuenta cuyo rótulo se está eligiendo
let cfgSoloProblemas = true;

function cfgTexto(ws, r, c) {
  const v = ws.getCell(r, c).value;
  if (v == null) return "";
  if (typeof v === "object") return v.richText ? v.richText.map(t => t.text).join("") : "";
  return String(v);
}
function cfgNorm(t) {
  return String(t || "").normalize("NFD").replace(/[̀-ͯ]/g, "")
    .toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();
}

// El estado de cada cuenta, leyendo las fórmulas del Anexo II.
function cfgLeerCuentas(wb) {
  const layout = derivarLayoutSaldos(wb);
  const S = wb.getWorksheet(layout.sheet);
  const ax = wb.getWorksheet("Anexo II");
  const bloque = rtUbicarBloque(ax);

  const porFila = {};
  for (let r = layout.planDeCuentas.desde; r <= layout.planDeCuentas.hasta; r++) {
    const t = cfgTexto(S, r, layout.keyCol).trim();
    const m = /^\s*([\d.]+)\s*-?\s*(.*)$/.exec(t);
    if (!m) continue;
    porFila[r] = { cod: m[1].replace(/\./g, ""), nom: m[2].trim(), fila: r };
  }

  // qué renglón lee cada fila de SALDOS
  const lectores = {};
  const rotulos = [];
  for (let r = bloque.desde; r <= bloque.hasta; r++) {
    const rot = cfgTexto(ax, r, 2).trim();
    if (rot) rotulos.push({ fila: r, rotulo: rot });
    for (const c of [4, 5, 6]) {
      const v = ax.getCell(r, c).value;
      if (!v || typeof v !== "object" || typeof v.formula !== "string") continue;
      if (/^SUM\(/i.test(v.formula)) continue;
      const vistas = new Set();
      for (const m of v.formula.matchAll(/SALDOS!\$?[A-Z]{1,3}\$?(\d+)/g)) vistas.add(Number(m[1]));
      vistas.forEach(f => (lectores[f] = lectores[f] || []).push({ fila: r, rotulo: rot, col: c }));
    }
  }

  const cuentas = [];
  for (const info of Object.values(porFila)) {
    // Solo las cuentas de GASTO. El Anexo II es la apertura del gasto por centro de costo: una
    // cuenta de activo, de pasivo o de ingreso no va ahí, y listarlas como "sin rótulo" haría
    // parecer que faltan trece configuraciones cuando faltan cuatro.
    if (!/^42/.test(info.cod)) continue;
    if (!PLAN_CENTROS_COSTO.has(info.cod)) continue;
    const quien = lectores[info.fila] || [];
    const estado = quien.length === 0 ? "sin" : (quien.length > 1 ? "varias" : "ok");
    cuentas.push({ ...info, rotulos: quien, estado, saldo: cfgSaldoDe(S, info.fila, layout) });
  }
  const orden = { varias: 0, sin: 1, ok: 2 };
  cuentas.sort((a, b) => (orden[a.estado] - orden[b.estado]) || a.cod.localeCompare(b.cod));

  // Para elegir destino, un rótulo repetido tiene que aparecer UNA vez: algunos archivos traen
  // el mismo concepto en dos filas (una con cuenta y otra vacía), y un desplegable con el
  // nombre dos veces no deja saber cuál se está eligiendo. Queda la primera, que es la misma
  // que usa cfgMoverCuenta.
  const vistos = new Set();
  const paraElegir = rotulos.filter(r => {
    const k = cfgNorm(r.rotulo);
    if (!k || vistos.has(k)) return false;
    vistos.add(k);
    return true;
  }).sort((a, b) => a.rotulo.localeCompare(b.rotulo, "es"));

  return { cuentas, rotulos: paraElegir, todosLosRotulos: rotulos, layout, bloque };
}

function cfgSaldoDe(ws, fila, layout) {
  const v = ws.getCell(fila, layout.deudorCol).value;
  if (typeof v === "number") return v;
  if (v && typeof v === "object" && typeof v.result === "number") return v.result;
  return 0;
}

// Mueve una cuenta al rótulo elegido: la saca de donde esté y la engancha en el nuevo.
function cfgMoverCuenta(wb, cod, rotuloDestino) {
  const { cuentas, rotulos, bloque } = cfgLeerCuentas(wb);
  const cuenta = cuentas.find(c => c.cod === cod);
  if (!cuenta) throw new Error(`No encontré la cuenta ${cod} en SALDOS.`);
  const ax = wb.getWorksheet("Anexo II");

  const destino = rotulos.find(r => cfgNorm(r.rotulo) === cfgNorm(rotuloDestino));
  if (!destino) throw new Error(`No existe el rótulo "${rotuloDestino}" en este archivo.`);

  // la columna de centro de costo: la que ya usaba, o la de administración si no tenía
  const col = cuenta.rotulos.length ? cuenta.rotulos[0].col : 4;

  // sacarla de todos lados
  const sacadaDe = [];
  for (let r = bloque.desde; r <= bloque.hasta; r++) {
    for (const c of [4, 5, 6]) {
      const dir = `${String.fromCharCode(64 + c)}${r}`;
      if (rtQuitarTermino(ax, dir, cuenta.fila)) sacadaDe.push(cfgTexto(ax, r, 2).trim());
    }
  }

  // engancharla en el nuevo
  const colImp = bloque.colImporte || "C";
  const celda = ax.getCell(destino.fila, col);
  const v = celda.value;
  const previo = (v && typeof v === "object" && typeof v.formula === "string") ? v.formula : "";
  celda.value = { formula: previo ? `${previo}+SALDOS!${colImp}${cuenta.fila}`
                                  : `+SALDOS!${colImp}${cuenta.fila}` };

  return { cod, nombre: cuenta.nom, de: sacadaDe, a: destino.rotulo };
}

// ------------------------------------------------------------------- pantalla

async function abrirConfigCuentas() {
  mostrar("cardCuentas", true);
  const sel = document.getElementById("cfgArchivo");
  if (!sel.options.length) {
    for (const a of ARCHIVOS_TFBR) sel.add(new Option(a.label, a.id));
  }
  cfgArchivo = sel.value || ARCHIVOS_TFBR[0].id;
  sel.value = cfgArchivo;
  await cfgCargarArchivo();
  document.getElementById("cardCuentas").scrollIntoView({ behavior: "smooth" });
}

async function cfgCargarArchivo() {
  const sel = document.getElementById("cfgArchivo");
  cfgArchivo = sel.value;
  cfgCambios = [];
  cfgEditando = null;
  estadoUi("cfgStatus", "Cargando el maestro…", "");
  const cargado = App.maestrosCargados[cfgArchivo];
  if (!cargado) { estadoUi("cfgStatus", "Todavía no está cargado ese maestro.", "bad"); return; }
  // copia en memoria: lo que se edita acá no toca el maestro hasta guardar
  const buffer = await cargado.wb.xlsx.writeBuffer();
  cfgWb = new ExcelJS.Workbook();
  await cfgWb.xlsx.load(buffer);
  estadoUi("cfgStatus", "", "");
  cfgPintar();
}

function cfgPintar() {
  const cont = document.getElementById("cfgLista");
  if (!cfgWb) { cont.innerHTML = ""; return; }
  const { cuentas, rotulos } = cfgLeerCuentas(cfgWb);

  const problemas = cuentas.filter(c => c.estado !== "ok");
  document.getElementById("cfgResumen").innerHTML =
    `<b>${cuentas.length}</b> cuentas · <b>${problemas.length}</b> a revisar ` +
    `(${cuentas.filter(c => c.estado === "varias").length} en varios rótulos, ` +
    `${cuentas.filter(c => c.estado === "sin").length} sin rótulo)` +
    (cfgCambios.length ? ` · <b>${cfgCambios.length} cambio(s) sin guardar</b>` : "");

  const filtro = cfgNorm(cfgFiltro);
  const visibles = cuentas.filter(c => {
    if (cfgSoloProblemas && c.estado === "ok") return false;
    if (!filtro) return true;
    return cfgNorm(`${c.cod} ${c.nom}`).includes(filtro) ||
           c.rotulos.some(r => cfgNorm(r.rotulo).includes(filtro));
  });

  let html = "<table class='cfg'><thead><tr><th>Cuenta</th><th>Saldo</th>" +
             "<th>Rótulo del Anexo II</th><th></th></tr></thead><tbody>";
  for (const c of visibles) {
    const e = CFG_ESTADOS[c.estado];
    const rots = c.rotulos.length
      ? c.rotulos.map(r => r.rotulo || `(fila ${r.fila})`).join(" + ")
      : "—";
    html += `<tr>` +
      `<td><span class="mono">${c.cod}</span><br><span class="cfg-nom">${c.nom}</span></td>` +
      `<td class="mono cfg-saldo">${c.saldo.toFixed(2)}</td>` +
      `<td>${rots}<br><span class="status-msg ${e.clase} cfg-chip">${e.texto}</span></td>` +
      `<td><button class="cfg-btn" onclick="cfgElegir('${c.cod}')">Cambiar</button></td>` +
      `</tr>`;
    if (cfgEditando === c.cod) {
      const opts = rotulos.map(r => `<option value="${r.rotulo.replace(/"/g, "&quot;")}">${r.rotulo}</option>`).join("");
      html += `<tr class="cfg-editor"><td colspan="4">` +
        `Mover <b>${c.nom}</b> a: <select id="cfgDestino">${opts}</select> ` +
        `<button class="cfg-btn" onclick="cfgAplicar('${c.cod}')">Aplicar</button> ` +
        `<button class="cfg-btn" onclick="cfgElegir(null)">Cancelar</button>` +
        `</td></tr>`;
    }
  }
  html += "</tbody></table>";
  if (!visibles.length) html = "<p class='footer-note'>No hay cuentas que mostrar con este filtro.</p>";
  cont.innerHTML = html;

  document.getElementById("btnGuardarCuentas").disabled = cfgCambios.length === 0;
  const det = document.getElementById("cfgCambios");
  det.innerHTML = cfgCambios.length
    ? "<b>Cambios sin guardar:</b><ul>" + cfgCambios.map(c =>
        `<li class="footer-note">${c.nombre}: ${c.de.length ? c.de.join(", ") : "(sin rótulo)"} → <b>${c.a}</b></li>`).join("") + "</ul>"
    : "";
}

function cfgElegir(cod) { cfgEditando = cod; cfgPintar(); }

function cfgAplicar(cod) {
  const destino = document.getElementById("cfgDestino").value;
  try {
    const r = cfgMoverCuenta(cfgWb, cod, destino);
    cfgCambios.push(r);
    cfgEditando = null;
    estadoUi("cfgStatus", `"${r.nombre}" pasó a "${r.a}".`, "ok");
    cfgPintar();
  } catch (e) {
    estadoUi("cfgStatus", "No pude moverla: " + e.message, "bad");
  }
}

function cfgBuscar(v) { cfgFiltro = v; cfgPintar(); }
function cfgVerTodas(v) { cfgSoloProblemas = !v; cfgPintar(); }

async function guardarConfigCuentas() {
  if (!cfgCambios.length) return;
  document.getElementById("btnGuardarCuentas").disabled = true;
  mostrar("spinnerCuentas", true);
  try {
    const buffer = await cfgWb.xlsx.writeBuffer();
    const detalle = cfgCambios.map(c => `${c.cod} ${c.nombre} → ${c.a}`).join("; ");
    await ghtGuardarMaestro(cfgArchivo, buffer,
      `Configurar cuentas de ${cfgArchivo}: ${detalle}`.slice(0, 240));
    log(`Configuración guardada en ${cfgArchivo}: ${detalle}`);
    estadoUi("cfgStatus", `Guardado. ${cfgCambios.length} cambio(s).`, "ok");
    cfgCambios = [];
    // el maestro en memoria quedó viejo: se vuelve a leer de GitHub
    await revisarMaestrosExistentes();
    await cfgCargarArchivo();
  } catch (e) {
    estadoUi("cfgStatus", "No pude guardar: " + e.message, "bad");
    document.getElementById("btnGuardarCuentas").disabled = false;
  } finally {
    mostrar("spinnerCuentas", false);
  }
}

if (typeof module !== "undefined") {
  module.exports = { cfgLeerCuentas, cfgMoverCuenta, cfgNorm };
}
