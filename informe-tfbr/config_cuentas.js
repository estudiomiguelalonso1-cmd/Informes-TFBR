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
let cfgHoja = "Anexo II";   // la solapa abierta

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
  // Sin esto, "- Proveedores" del Pasivo se vería leyendo 5 cuentas en vez de 31: el rango
  // todavía sin expandir sólo nombra sus extremos.
  expandirRangosSaldos(wb, layout, () => {});
  return wb;
}

// El estado de un archivo PARA UNA HOJA: qué renglón lee cada cuenta.
//
// Qué cuentas se listan depende de la hoja. En el Anexo II manda el plan oficial: son las que
// discriminan por centro de costo, ni una más. En las demás no hay una lista oficial, así que
// se usa el rubro — el Activo abre cuentas de activo, el Pasivo de pasivo — deducido de lo que
// la hoja ya lee, no escrito a mano.
function cfgEstadoDe(wb, hoja) {
  const layout = derivarLayoutSaldos(wb);
  const S = wb.getWorksheet(layout.sheet);
  const porFila = chPlanPorFila(wb, layout);
  const mapa = chMapaHoja(wb, layout, hoja);
  if (!mapa) return { porCuenta: {}, rotulos: [], layout, hoja };

  const lectores = {}, rotulos = [];
  for (const r of mapa.renglones) {
    if (r.rotulo) rotulos.push({ fila: r.fila, rotulo: r.rotulo });
    for (const f of r.filasSaldos) {
      (lectores[f] = lectores[f] || []).push({ fila: r.fila, rotulo: r.rotulo, col: r.col });
    }
  }

  const esAnexoII = /anexo\s*ii/i.test(hoja);
  const rubro = esAnexoII ? null : chRubroDeHoja(wb, layout, hoja, porFila);

  const porCuenta = {};
  for (const info of Object.values(porFila)) {
    if (esAnexoII) {
      // El Anexo II es la apertura del gasto por centro de costo: la lista la da el plan
      // oficial, no lo que el archivo tenga cableado hoy.
      if (!/^42/.test(info.cod)) continue;
      if (!PLAN_CENTROS_COSTO.has(info.cod)) continue;
    } else if (rubro) {
      // Las demás abren su propio rubro. Una cuenta de gasto no va en el Activo, y listarla
      // haría parecer que falta configurarla.
      if (String(info.cod)[0] !== rubro) continue;
    }
    const quien = lectores[info.fila] || [];
    porCuenta[info.cod] = {
      ...info,
      saldo: cfgNumero(S, info.fila, layout.deudorCol) - cfgNumero(S, info.fila, layout.acreedorCol),
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

  return { porCuenta, rotulos: paraElegir, layout, hoja };
}

function cfgNumero(ws, fila, col) {
  const v = ws.getCell(fila, col).value;
  if (typeof v === "number") return v;
  if (v && typeof v === "object" && typeof v.result === "number") return v.result;
  return 0;
}

// Junta los cuatro archivos en una sola vista.
function cfgVistaUnica(hoja = cfgHoja) {
  const estados = {};
  for (const a of ARCHIVOS_TFBR) {
    if (cfgCopias[a.id]) estados[a.id] = cfgEstadoDe(cfgCopias[a.id], hoja);
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

// Mueve la cuenta al rótulo elegido, EN LOS CUATRO archivos y en la hoja abierta.
function cfgMoverEnTodos(cod, rotuloDestino, hoja = cfgHoja) {
  const hechos = [], fallos = [], deducido = [];
  for (const a of ARCHIVOS_TFBR) {
    const wb = cfgCopias[a.id];
    if (!wb) continue;
    const layout = derivarLayoutSaldos(wb);
    const r = engancharEnHoja(wb, layout, hoja, cod, rotuloDestino);
    if (!r.hecho) {
      // Que la cuenta no esté en este archivo es normal y no se reporta como problema.
      if (!/no está en el plan/.test(r.motivo)) fallos.push({ archivo: a.label, motivo: r.motivo });
      continue;
    }
    hechos.push(a.id);
    if (r.signoDeducido) deducido.push(a.label);
  }
  return { hechos, fallos, deducido };
}

// Las hojas que se pueden configurar. Una hoja donde el rótulo no identifica el renglón queda
// afuera con su motivo a la vista, en vez de ofrecer un editor que haría cualquier cosa: el
// Anexo I es el caso — una fila es un bien y sus columnas son valor de origen, altas, bajas y
// amortización, que no son intercambiables.
let cfgHojasFuera = [];

function cfgHojasDisponibles() {
  const listas = [];
  const fuera = {};
  for (const a of ARCHIVOS_TFBR) {
    const wb = cfgCopias[a.id];
    if (!wb) continue;
    const layout = derivarLayoutSaldos(wb);
    const buenas = [];
    for (const h of chHojasConfigurables(wb, layout)) {
      const ed = chHojaEditable(wb, layout, h.hoja);
      if (ed.editable) buenas.push(h.hoja);
      else if (!fuera[h.hoja]) fuera[h.hoja] = { hoja: h.hoja, motivo: ed.motivo, archivo: a.label };
    }
    listas.push(buenas);
  }
  cfgHojasFuera = Object.values(fuera);
  if (!listas.length) return [];
  // Una hoja que no está en los cuatro no se puede configurar de una vez para todos: se
  // ofrece igual, pero el cambio sólo va a aplicar donde exista, y eso se avisa en la fila.
  const todas = [...new Set(listas.flat())];
  // El Anexo II primero: es la que más se usa.
  return todas.sort((x, y) => (/anexo\s*ii/i.test(y) ? 1 : 0) - (/anexo\s*ii/i.test(x) ? 1 : 0));
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
  soltarScrollSiNoQuedaVentana();
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

// Las solapas: una por hoja que engancha cuentas.
function cfgPintarSolapas() {
  const cont = document.getElementById("cfgSolapas");
  if (!cont) return;
  const hojas = cfgHojasDisponibles();
  if (!hojas.length) { cont.innerHTML = ""; return; }
  if (!hojas.some(h => cfgNorm(h) === cfgNorm(cfgHoja))) cfgHoja = hojas[0];
  cont.innerHTML = hojas.map(h =>
    `<button class="cfg-solapa${cfgNorm(h) === cfgNorm(cfgHoja) ? " activa" : ""}" ` +
    `onclick="cfgVerHoja('${h.replace(/'/g, "&#39;")}')">${h}</button>`).join("");
}

function cfgVerHoja(hoja) {
  if (cfgNorm(hoja) === cfgNorm(cfgHoja)) return;
  cfgHoja = hoja;
  cfgEditando = null;
  cfgPintar();
}

function cfgPintar() {
  const cont = document.getElementById("cfgLista");
  if (!cfgCopias) { cont.innerHTML = ""; return; }
  cfgPintarSolapas();
  const { filas, rotulos } = cfgVistaUnica();

  const cuenta = (e) => filas.filter(f => f.estado === e).length;
  const problemas = filas.filter(f => f.estado !== "ok").length;
  const partes = [`<b>${filas.length}</b> cuentas en <b>${cfgHoja}</b>`];
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

  const nota = document.getElementById("cfgNota");
  if (nota) {
    const fuera = cfgHojasFuera.filter(f => !cfgHojasDisponibles().some(h => cfgNorm(h) === cfgNorm(f.hoja)));
    nota.innerHTML = fuera.length
      ? fuera.map(f => `<span class="footer-note">${f.hoja} no se edita desde acá: ${f.motivo}.</span>`).join("")
      : "";
  }

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
             `<th>Renglón de ${cfgHoja}</th><th></th></tr></thead><tbody>`;
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
      : `Todas las cuentas de ${cfgHoja} están configuradas. Marcá «Ver todas» para revisarlas.`}</div>`;
  }
  cont.innerHTML = html;

  document.getElementById("btnGuardarCuentas").disabled = cfgCambios.length === 0;
  document.getElementById("cfgCambios").innerHTML = cfgCambios.length
    ? "<b>Cambios sin guardar:</b><ul>" + cfgCambios.map(c =>
        `<li class="footer-note">${c.nombre} → <b>${c.a}</b> en ${c.hoja} (${c.archivos} archivo(s))</li>`).join("") + "</ul>"
    : "";
}

function cfgElegir(cod) { cfgEditando = cod; cfgPintar(); }

function cfgAplicar(cod) {
  const destino = document.getElementById("cfgDestino").value;
  try {
    const { filas } = cfgVistaUnica();
    const f = filas.find(x => x.cod === cod);
    const nombre = f ? f.nom : cod;
    const r = cfgMoverEnTodos(cod, destino);

    if (!r.hechos.length) {
      estadoUi("cfgStatus",
        `No la moví en ningún archivo. ${r.fallos.map(x => `${x.archivo}: ${x.motivo}`).join("; ")}`, "bad");
      return;
    }

    cfgCambios.push({ cod, nombre, a: destino, hoja: cfgHoja, archivos: r.hechos.length });
    cfgEditando = null;

    // El signo deducido se avisa: cuando el renglón de destino estaba vacío, no había de dónde
    // leerlo y se usó el mayoritario de la hoja, que puede errarle.
    const partes = [`"${nombre}" pasó a "${destino}" en ${cfgHoja}, en ${r.hechos.length} archivo(s).`];
    if (r.deducido.length) {
      partes.push(`El renglón estaba vacío en ${r.deducido.join(", ")}, así que el signo se tomó ` +
                  `del resto de la hoja: verificá que el importe sume y no reste.`);
    }
    if (r.fallos.length) {
      partes.push(`No la pude mover en ${r.fallos.map(x => `${x.archivo} (${x.motivo})`).join("; ")}.`);
    }
    estadoUi("cfgStatus", partes.join(" "), r.fallos.length || r.deducido.length ? "" : "ok");
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
    const detalle = cfgCambios.map(c => `${c.cod} ${c.nombre} → ${c.a} (${c.hoja})`).join("; ");
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
