// Panel "Configurar cuentas": qué rótulo del Anexo II lee cada cuenta, y cómo cambiarlo.
//
// Una sola lista para los cuatro balances. Los cuatro tienen que abrir el gasto con los mismos
// rótulos y la misma cuenta en cada uno, así que mostrarlos por separado sería mostrar cuatro
// veces la misma configuración — y un cambio hecho en uno solo los desalinearía, que es
// justamente lo que se viene arreglando. El cambio se aplica a los cuatro.
//
// Lo que se muestra sale de LAS FÓRMULAS, no de una tabla: son las fórmulas las que deciden
// dónde cae el importe. Pero de las fórmulas DESPUÉS de que el motor haga lo suyo, no de las
// del archivo guardado: el maestro de GitHub todavía tiene el cableado viejo, y la limpieza y
// la unificación se aplican en cada corrida. Mirando el archivo crudo, el panel denunciaba
// como problemas cosas que el informe generado ya no tiene.
//
// Cambiar el rótulo reescribe la fórmula: saca la cuenta del renglón donde estaba y la engancha
// en el nuevo. Los cambios quedan en memoria y recién se suben al apretar Guardar.

const CFG_ESTADOS = {
  ok:      { texto: "OK",                    clase: "ok"  },
  sin:     { texto: "Sin rótulo",            clase: "bad" },
  varias:  { texto: "En varios",             clase: "bad" },
  difiere: { texto: "Difiere entre archivos", clase: "bad" },
};

let cfgCopias = null;      // { archivoId: workbook } — copias en memoria, ya preparadas
let cfgCambios = [];
let cfgFiltro = "";
let cfgEditando = null;
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

// Cómo queda un archivo después de lo que el motor aplica en cada corrida. Es lo que hay que
// mirar: el maestro guardado todavía no lo tiene.
function cfgPrepararCopia(wb) {
  materializarFormulasCompartidas(wb);
  const limpieza = limpiarPlanDeCuentas(wb, () => {});
  const layout = derivarLayoutSaldos(wb);
  const plan = leerPlanDeCuentas(wb, layout).cuentas;
  aplicarRepuntesAnexo(wb, null, plan, () => {});
  unificarRotulosAnexo(wb, plan, () => {});
  return wb;
}

// El estado de un archivo: qué renglón lee cada cuenta de gasto.
function cfgEstadoDe(wb) {
  const layout = derivarLayoutSaldos(wb);
  const S = wb.getWorksheet(layout.sheet);
  const ax = wb.getWorksheet("Anexo II");
  const bloque = rtUbicarBloque(ax);

  const porFila = {};
  for (let r = layout.planDeCuentas.desde; r <= layout.planDeCuentas.hasta; r++) {
    const m = /^\s*([\d.]+)\s*-?\s*(.*)$/.exec(cfgTexto(S, r, layout.keyCol).trim());
    if (m) porFila[r] = { cod: m[1].replace(/\./g, ""), nom: m[2].trim(), fila: r };
  }

  const lectores = {}, rotulos = [];
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

  const porCuenta = {};
  for (const info of Object.values(porFila)) {
    // Solo cuentas de GASTO: el Anexo II es la apertura del gasto por centro de costo. Una de
    // activo, pasivo o ingreso no va ahí, y listarlas haría parecer que falta configurarlas.
    if (!/^42/.test(info.cod)) continue;
    if (!PLAN_CENTROS_COSTO.has(info.cod)) continue;
    const quien = lectores[info.fila] || [];
    porCuenta[info.cod] = {
      ...info,
      saldo: cfgNumero(S, info.fila, layout.deudorCol),
      rotulos: quien.map(q => q.rotulo || `(fila ${q.fila})`),
      col: quien.length ? quien[0].col : null,
    };
  }
  const vistos = new Set();
  const paraElegir = rotulos.filter(r => {
    const k = cfgNorm(r.rotulo);
    if (!k || vistos.has(k)) return false;
    vistos.add(k); return true;
  }).sort((a, b) => a.rotulo.localeCompare(b.rotulo, "es"));

  return { porCuenta, rotulos: paraElegir, bloque, layout };
}

function cfgNumero(ws, fila, col) {
  const v = ws.getCell(fila, col).value;
  if (typeof v === "number") return v;
  if (v && typeof v === "object" && typeof v.result === "number") return v.result;
  return 0;
}

// Junta los cuatro archivos en una sola vista.
function cfgVistaUnica() {
  const estados = {};
  for (const a of ARCHIVOS_TFBR) {
    if (cfgCopias[a.id]) estados[a.id] = cfgEstadoDe(cfgCopias[a.id]);
  }
  const ids = Object.keys(estados);
  if (!ids.length) return { filas: [], rotulos: [] };

  const codigos = new Set();
  ids.forEach(id => Object.keys(estados[id].porCuenta).forEach(c => codigos.add(c)));

  const filas = [];
  for (const cod of codigos) {
    const enCada = ids.map(id => ({ id, c: estados[id].porCuenta[cod] })).filter(x => x.c);
    const nom = enCada[0].c.nom;
    const saldos = {};
    enCada.forEach(x => { saldos[x.id] = x.c.saldo; });

    // el rótulo, visto en cada archivo
    const porArchivo = {};
    enCada.forEach(x => { porArchivo[x.id] = x.c.rotulos; });
    const firmas = new Set(enCada.map(x => x.c.rotulos.map(cfgNorm).sort().join("|")));

    let estado = "ok";
    const alguno = enCada[0].c.rotulos;
    if (enCada.some(x => x.c.rotulos.length === 0)) estado = "sin";
    if (enCada.some(x => x.c.rotulos.length > 1)) estado = "varias";
    if (firmas.size > 1) estado = "difiere";

    filas.push({
      cod, nom, estado, porArchivo, saldos,
      rotulos: alguno,
      enArchivos: enCada.map(x => x.id),
      saldoMax: Math.max(...Object.values(saldos).map(Math.abs)),
    });
  }
  const orden = { difiere: 0, varias: 1, sin: 2, ok: 3 };
  filas.sort((a, b) => (orden[a.estado] - orden[b.estado]) ||
                       (b.saldoMax - a.saldoMax) || a.cod.localeCompare(b.cod));

  // los rótulos para elegir: los que están en TODOS los archivos
  const listas = ids.map(id => new Set(estados[id].rotulos.map(r => cfgNorm(r.rotulo))));
  const rotulos = estados[ids[0]].rotulos
    .filter(r => listas.every(s => s.has(cfgNorm(r.rotulo))))
    .map(r => r.rotulo);

  return { filas, rotulos, estados };
}

// Mueve la cuenta al rótulo elegido, EN LOS CUATRO archivos.
function cfgMoverEnTodos(cod, rotuloDestino) {
  const hechos = [];
  for (const a of ARCHIVOS_TFBR) {
    const wb = cfgCopias[a.id];
    if (!wb) continue;
    const est = cfgEstadoDe(wb);
    const cuenta = est.porCuenta[cod];
    if (!cuenta) continue;                       // esa cuenta no está en este archivo
    const destino = est.rotulos.find(r => cfgNorm(r.rotulo) === cfgNorm(rotuloDestino));
    if (!destino) continue;
    const ax = wb.getWorksheet("Anexo II");
    const col = cuenta.col || 4;

    for (let r = est.bloque.desde; r <= est.bloque.hasta; r++) {
      for (const c of [4, 5, 6]) rtQuitarTermino(ax, `${String.fromCharCode(64 + c)}${r}`, cuenta.fila);
    }
    const colImp = est.bloque.colImporte || "C";
    const celda = ax.getCell(destino.fila, col);
    const v = celda.value;
    const previo = (v && typeof v === "object" && typeof v.formula === "string") ? v.formula : "";
    celda.value = { formula: previo ? `${previo}+SALDOS!${colImp}${cuenta.fila}`
                                    : `+SALDOS!${colImp}${cuenta.fila}` };
    hechos.push(a.id);
  }
  return hechos;
}

// ------------------------------------------------------------------- pantalla

// Se abre como ventana encima de la página, no como una card más: la lista es larga y
// desplegarla en el flujo obligaba a bajar hasta el fondo para volver al cierre. Al cerrarla,
// las copias preparadas y los cambios sin guardar quedan en memoria — reabrirla es instantáneo
// y no se pierde nada.
async function abrirConfigCuentas() {
  mostrar("ovCuentas", true);
  document.body.classList.add("sin-scroll");
  const b = document.getElementById("cfgBuscador");
  if (b) setTimeout(() => b.focus(), 0);
  if (!cfgCopias) await cfgCargar();
  else cfgPintar();
}

function cerrarConfigCuentas() {
  mostrar("ovCuentas", false);
  document.body.classList.remove("sin-scroll");
}

// Click en el fondo (no en la ventana) = cerrar.
function cfgFondo(ev) {
  if (ev.target && ev.target.id === "ovCuentas") cerrarConfigCuentas();
}

document.addEventListener("keydown", (ev) => {
  if (ev.key !== "Escape") return;
  const ov = document.getElementById("ovCuentas");
  if (ov && !ov.classList.contains("hidden")) cerrarConfigCuentas();
});

async function cfgCargar() {
  estadoUi("cfgStatus", "Preparando los cuatro archivos…", "");
  document.getElementById("cfgLista").innerHTML = "";
  cfgCopias = {};
  cfgCambios = [];
  cfgEditando = null;
  try {
    for (const a of ARCHIVOS_TFBR) {
      const cargado = App.maestrosCargados[a.id];
      if (!cargado) continue;
      const buffer = await cargado.wb.xlsx.writeBuffer();
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buffer);
      cfgCopias[a.id] = cfgPrepararCopia(wb);
    }
    estadoUi("cfgStatus", "", "");
    cfgPintar();
  } catch (e) {
    estadoUi("cfgStatus", "No pude preparar los archivos: " + e.message, "bad");
  }
}

function cfgPintar() {
  const cont = document.getElementById("cfgLista");
  if (!cfgCopias) { cont.innerHTML = ""; return; }
  const { filas, rotulos } = cfgVistaUnica();

  const cuenta = (e) => filas.filter(f => f.estado === e).length;
  const problemas = filas.filter(f => f.estado !== "ok").length;
  const partes = [`<b>${filas.length}</b> cuentas`];
  if (problemas) {
    const d = [];
    if (cuenta("difiere")) d.push(`${cuenta("difiere")} difieren entre archivos`);
    if (cuenta("varias")) d.push(`${cuenta("varias")} en varios rótulos`);
    if (cuenta("sin")) d.push(`${cuenta("sin")} sin rótulo`);
    partes.push(`<b>${problemas}</b> a revisar — ${d.join(", ")}`);
  } else {
    partes.push("todas configuradas");
  }
  if (cfgCambios.length) partes.push(`<b>${cfgCambios.length} sin guardar</b>`);
  document.getElementById("cfgResumen").innerHTML = partes.join(" · ");

  // Buscar mira SIEMPRE las 90 cuentas, esté o no tildado "Ver todas". Con el filtro de
  // problemas por delante, buscar "sueldos" con todo configurado no devolvía nada: la cuenta
  // existía pero estaba descartada antes de comparar el texto, y parecía que el buscador
  // estaba roto.
  const filtro = cfgNorm(cfgFiltro);
  const visibles = filas.filter(f => {
    if (filtro) {
      return cfgNorm(`${f.cod} ${f.nom}`).includes(filtro) ||
             f.rotulos.some(r => cfgNorm(r).includes(filtro));
    }
    return !(cfgSoloProblemas && f.estado === "ok");
  });

  let html = "<table class='cfg'><thead><tr><th>Cuenta</th>" +
             "<th>Rótulo del Anexo II</th><th></th></tr></thead><tbody>";
  for (const f of visibles) {
    const e = CFG_ESTADOS[f.estado];
    const chip = `<span class="cfg-chip ${e.clase}">${e.texto}</span>`;
    let rots;
    if (f.estado === "difiere") {
      rots = `<span class="cfg-rot">—${chip}</span>` +
        ARCHIVOS_TFBR.filter(a => f.porArchivo[a.id])
          .map(a => `<span class="cfg-porarch"><b>${a.label}</b> · ${f.porArchivo[a.id].join(" + ") || "sin rótulo"}</span>`)
          .join("");
    } else {
      rots = `<span class="cfg-rot">${f.rotulos.length ? f.rotulos.join(" + ") : "—"}` +
             `${f.estado === "ok" ? "" : chip}</span>`;
    }
    html += `<tr>` +
      `<td><span class="mono">${f.cod}</span><span class="cfg-nom">${f.nom}</span></td>` +
      `<td>${rots}</td>` +
      `<td><button class="cfg-btn" onclick="cfgElegir('${f.cod}')">Cambiar</button></td>` +
      `</tr>`;
    if (cfgEditando === f.cod) {
      const opts = rotulos.map(r =>
        `<option value="${r.replace(/"/g, "&quot;")}">${r}</option>`).join("");
      html += `<tr class="cfg-editor"><td colspan="3">` +
        `Mover <b>${f.nom}</b> a: <select id="cfgDestino">${opts}</select> ` +
        `<button class="cfg-btn" onclick="cfgAplicar('${f.cod}')">Aplicar a los 4</button> ` +
        `<button class="cfg-btn" onclick="cfgElegir(null)">Cancelar</button>` +
        `</td></tr>`;
    }
  }
  html += "</tbody></table>";
  if (!visibles.length) {
    html = `<div class="cfg-vacio">${cfgFiltro
      ? `Ninguna de las ${filas.length} cuentas coincide con «${cfgFiltro}».`
      : "Todas las cuentas están configuradas. Marcá «Ver todas» para revisarlas."}</div>`;
  }
  cont.innerHTML = html;

  document.getElementById("btnGuardarCuentas").disabled = cfgCambios.length === 0;
  document.getElementById("cfgCambios").innerHTML = cfgCambios.length
    ? "<b>Cambios sin guardar:</b><ul>" + cfgCambios.map(c =>
        `<li class="footer-note">${c.nombre} → <b>${c.a}</b> (${c.archivos} archivo(s))</li>`).join("") + "</ul>"
    : "";
}

function cfgElegir(cod) { cfgEditando = cod; cfgPintar(); }

function cfgAplicar(cod) {
  const destino = document.getElementById("cfgDestino").value;
  try {
    const { filas } = cfgVistaUnica();
    const f = filas.find(x => x.cod === cod);
    const ids = cfgMoverEnTodos(cod, destino);
    cfgCambios.push({ cod, nombre: f ? f.nom : cod, a: destino, archivos: ids.length });
    cfgEditando = null;
    estadoUi("cfgStatus", `"${f ? f.nom : cod}" pasó a "${destino}" en ${ids.length} archivo(s).`, "ok");
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
    const detalle = cfgCambios.map(c => `${c.cod} ${c.nombre} → ${c.a}`).join("; ");
    const buffers = {};
    for (const a of ARCHIVOS_TFBR) {
      if (cfgCopias[a.id]) buffers[a.id] = await cfgCopias[a.id].xlsx.writeBuffer();
    }
    for (const [id, buf] of Object.entries(buffers)) {
      await ghtGuardarMaestro(id, buf, `Configurar cuentas: ${detalle}`.slice(0, 240));
      log(`Configuración guardada en ${id}`);
    }
    estadoUi("cfgStatus", `Guardado en los 4. ${cfgCambios.length} cambio(s).`, "ok");
    cfgCambios = [];
    await revisarMaestrosExistentes();
    cfgCopias = null;
    await cfgCargar();
  } catch (e) {
    estadoUi("cfgStatus", "No pude guardar: " + e.message, "bad");
    document.getElementById("btnGuardarCuentas").disabled = false;
  } finally {
    mostrar("spinnerCuentas", false);
  }
}

if (typeof module !== "undefined") {
  module.exports = { cfgEstadoDe, cfgPrepararCopia, cfgNorm };
}
