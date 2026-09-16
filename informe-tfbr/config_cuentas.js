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
  ok:       { texto: "OK",                     clase: "ok"    },
  sin:      { texto: "Sin rótulo",             clase: "bad"   },
  varias:   { texto: "En varios",              clase: "bad"   },
  difiere:  { texto: "Difiere entre archivos", clase: "bad"   },
  // La cuenta existe en el plan oficial pero ningún balance la tiene todavía. No es un
  // problema: se configura igual, y el día que el sumas y saldos la traiga el motor le crea
  // la fila y la engancha donde quedó dicho.
  fuera:    { texto: "Sin usar",               clase: "tenue" },
  // La persona decidió que no va a ningún renglón. No se vuelve a preguntar por ella.
  excluida: { texto: "No se usa",              clase: "tenue" },
};

// Los rubros del sumas y saldos, para agrupar la lista. Con 940 cuentas, una lista corrida
// no se puede leer.
const CFG_RUBROS = [
  { digito: "1", texto: "Activo" },
  { digito: "2", texto: "Pasivo" },
  { digito: "3", texto: "Patrimonio neto" },
  { digito: "4", texto: "Resultados" },
  { digito: "5", texto: "Orden" },
];
const cfgRubroDe = (cod) => String(cod).trim()[0];
const cfgNombreRubro = (d) => (CFG_RUBROS.find(r => r.digito === d) || { texto: "Otras" }).texto;

let cfgCopias = null;      // { archivoId: workbook } — copias en memoria, ya preparadas
let cfgCambios = [];
let cfgFiltro = "";
let cfgEditando = null;
let cfgHoja = "Anexo II";   // la solapa abierta

// El filtro de la lista. Antes era un casillero "Ver todas" que sólo distinguía entre "las que
// tienen problema" y "todas", y con 90 cuentas eso no alcanza: lo que se busca es "cuáles
// quedaron sin rótulo" o "cuáles difieren entre archivos", que son preguntas distintas y
// llevan a acciones distintas.
const CFG_FILTROS = [
  { id: "sin",    texto: "Sin rótulo", incluye: (f) => f.estado === "sin" },
  { id: "varias", texto: "En varios",  incluye: (f) => f.estado === "varias" || f.estado === "difiere" },
  { id: "ok",     texto: "Con rótulo", incluye: (f) => f.estado === "ok" },
  { id: "todas",  texto: "Todas",      incluye: () => true },
];
let cfgFiltroEstado = "sin";

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
// Tiene que quedar como lo deja el motor, paso por paso y en el mismo orden. Si el panel
// muestra otra cosa, la persona configura contra un estado que no existe: antes se veía
// "Autovia del Mercosur — sin rótulo" en los Mensuales cuando el motor ya la enganchaba en
// "- Proveedores", y los préstamos aparecían agrupados donde el motor ya los había abierto.
//
// Lo único que no se puede replicar acá es lo que depende del sumas y saldos del mes (el alta
// de cuentas nuevas y el valor de origen del Anexo I): el panel se abre sin haber cargado
// ningún export. Nada de eso cambia a qué renglón va una cuenta, que es lo que el panel muestra.
function cfgPrepararCopia(wb, archivoId) {
  materializarFormulasCompartidas(wb);
  limpiarPlanDeCuentas(wb, () => {});
  let layout = derivarLayoutSaldos(wb);
  const plan = leerPlanDeCuentas(wb, layout).cuentas;
  aplicarRepuntesAnexo(wb, null, plan, () => {});
  unificarRotulosAnexo(wb, plan, () => {});
  // Sin esto, "- Proveedores" del Pasivo se vería leyendo 5 cuentas en vez de 31: el rango
  // todavía sin expandir sólo nombra sus extremos.
  expandirRangosSaldos(wb, layout, () => {});
  aplicarRenombresRenglon(wb, layout, () => {});
  consolidarPrestamos(wb, layout, () => {});
  // Insertar filas mueve todo: el layout hay que volver a derivarlo.
  layout = derivarLayoutSaldos(wb);
  agregarRenglonesFaltantes(wb, layout, archivoId, () => {});
  layout = derivarLayoutSaldos(wb);
  aplicarAsignaciones(wb, layout, () => {});
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

  const esAnexoII = esAnexoIIHoja(hoja);
  // El Anexo II abre el gasto: su rubro es resultados, aunque la lista se filtre además por
  // el plan de centros de costo.
  const rubro = esAnexoII ? "4" : chRubroDeHoja(wb, layout, hoja, porFila);

  // Todas las cuentas, en todas las hojas. Antes cada solapa mostraba sólo las de su rubro —el
  // Activo las 1xx, el Pasivo las 2xx— y el Anexo II sólo las que el plan oficial marca con
  // centros de costo. Eso es el programa decidiendo qué se puede configurar y qué no, que es
  // justo lo que no tiene que hacer. La lista va agrupada por rubro y con el de la hoja
  // primero, así lo habitual queda arriba sin esconder el resto.
  const porCuenta = {};
  for (const info of Object.values(porFila)) {
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

  return { porCuenta, rotulos: paraElegir, layout, hoja, rubro };
}

const esAnexoIIHoja = (h) => /anexo\s*ii/i.test(String(h || ""));

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
  // Las del plan oficial que ningún balance tiene todavía. Aparecen para poder configurarlas
  // de antemano: así, cuando el sumas y saldos las traiga por primera vez, ya está decidido a
  // dónde van y nadie tiene que acordarse de nada.
  const rubro = ids.length ? estados[ids[0]].rubro : null;
  const excluidas = cuentasExcluidas(App.configuracion);
  if (typeof PLAN_OFICIAL === "object") {
    for (const [cod, nom] of Object.entries(PLAN_OFICIAL)) {
      if (codigos.has(cod)) continue;
      filas.push({
        cod, nom, estado: excluidas.has(cod) ? "excluida" : "fuera",
        porArchivo: {}, saldos: {}, rotulos: [], enArchivos: [], saldoMax: 0,
      });
    }
  }
  // Y las que están en los archivos pero la persona marcó como que no van a ningún renglón.
  for (const f of filas) if (excluidas.has(f.cod)) f.estado = "excluida";

  // Orden: primero el rubro de la hoja (en el Activo, las cuentas de activo), y dentro de cada
  // rubro las que hay que revisar antes que las que ya están bien.
  const orden = { difiere: 0, varias: 1, sin: 2, ok: 3, fuera: 4, excluida: 5 };
  const pesoRubro = (cod) => (rubro && cfgRubroDe(cod) === rubro) ? 0 : 1;
  filas.sort((a, b) =>
    (pesoRubro(a.cod) - pesoRubro(b.cod)) ||
    (cfgRubroDe(a.cod)).localeCompare(cfgRubroDe(b.cod)) ||
    (orden[a.estado] - orden[b.estado]) ||
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
      // El Anexo I no se configura desde acá en ningún archivo, ni siquiera en los que serían
      // editables: su valor de origen lo llena anexo_i.js desde las cuentas, y las columnas de
      // amortización se cargan a mano. Dejarlo como solapa mostraba las 140 cuentas de activo
      // del plan oficial como si les faltara rótulo, y ninguna va ahí.
      if (/anexo\s*i$/i.test(h.hoja)) {
        if (!fuera[h.hoja]) fuera[h.hoja] = { hoja: h.hoja, motivo: "el valor de origen sale de las cuentas y las amortizaciones se cargan a mano", archivo: a.label };
        continue;
      }
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
      cfgCopias[a.id] = cfgPrepararCopia(wb, a.id);
    }
    // Y lo último que hace el motor: que los cuatro lean las mismas cuentas. Necesita ver los
    // cuatro a la vez, así que va después del bucle.
    alinearHojas(
      ARCHIVOS_TFBR.filter(a => cfgCopias[a.id]).map(a => ({ id: a.id, label: a.label, wb: cfgCopias[a.id] })),
      () => {});
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

  const problemas = filas.filter(f => f.estado !== "ok").length;
  const partes = [`<b>${filas.length}</b> cuentas en <b>${cfgHoja}</b>`];
  partes.push(problemas ? `<b>${problemas}</b> a revisar` : "todas configuradas");
  if (cfgCambios.length) partes.push(`<b>${cfgCambios.length} sin guardar</b>`);
  document.getElementById("cfgResumen").innerHTML = partes.join(" \u00b7 ");

  const nota = document.getElementById("cfgNota");
  if (nota) {
    const fuera = cfgHojasFuera.filter(f => !cfgHojasDisponibles().some(h => cfgNorm(h) === cfgNorm(f.hoja)));
    nota.innerHTML = fuera.length
      ? fuera.map(f => `<span class="footer-note">${f.hoja} no se edita desde acá: ${f.motivo}.</span>`).join("")
      : "";
  }

  // Cada filtro lleva su número al lado: sin eso hay que ir tocándolos uno por uno para saber
  // si tienen algo. El que queda vacío se muestra igual, apagado, así los botones no se mueven
  // de lugar entre una hoja y otra.
  const cajaFiltros = document.getElementById("cfgFiltros");
  if (cajaFiltros) {
    if (!CFG_FILTROS.some(x => x.id === cfgFiltroEstado)) cfgFiltroEstado = "todas";
    cajaFiltros.innerHTML = CFG_FILTROS.map(x => {
      const n = filas.filter(x.incluye).length;
      return `<button class="cfg-filtro${x.id === cfgFiltroEstado ? " activo" : ""}` +
             `${n === 0 ? " vacio" : ""}" onclick="cfgVerEstado('${x.id}')">` +
             `${x.texto}<span class="cfg-filtro-n">${n}</span></button>`;
    }).join("");
  }

  // Buscar mira SIEMPRE todas las cuentas de la hoja, sin importar el filtro elegido. Con el
  // filtro por delante, buscar "sueldos" con todo configurado no devolvía nada: la cuenta
  // existía pero quedaba descartada antes de comparar el texto.
  const filtro = cfgNorm(cfgFiltro);
  const porEstado = CFG_FILTROS.find(x => x.id === cfgFiltroEstado) || CFG_FILTROS[CFG_FILTROS.length - 1];
  const visibles = filas.filter(f => {
    if (filtro) {
      return cfgNorm(`${f.cod} ${f.nom}`).includes(filtro) ||
             f.rotulos.some(r => cfgNorm(r).includes(filtro)) ||
             Object.values(f.porArchivo || {}).some(rs => rs.some(r => cfgNorm(r).includes(filtro)));
    }
    return porEstado.incluye(f);
  });

  // El pedazo que coincide con la búsqueda va marcado. Se ubica sobre el texto normalizado
  // para que "vacaciones" también marque "Vacaciónes", y se corta sobre el original para no
  // perder los acentos al mostrarlo.
  const resaltar = (t) => {
    const texto = String(t == null ? "" : t);
    if (!filtro) return texto;
    const i = cfgNorm(texto).indexOf(filtro);
    if (i < 0 || cfgNorm(texto).length !== texto.length) return texto;
    return `${texto.slice(0, i)}<mark>${texto.slice(i, i + filtro.length)}</mark>${texto.slice(i + filtro.length)}`;
  };

  // Agrupadas por rubro del sumas y saldos. Con 940 cuentas, una lista corrida no se lee.
  const porRubro = new Map();
  for (const f of visibles) {
    const d = cfgRubroDe(f.cod);
    if (!porRubro.has(d)) porRubro.set(d, []);
    porRubro.get(d).push(f);
  }
  // En el orden en que quedaron las filas: el rubro de la hoja primero.
  const rubrosOrdenados = [...porRubro.keys()];

  let html = "<table class='cfg'><thead><tr><th>Cuenta</th>" +
             `<th>Renglón de ${cfgHoja}</th><th></th></tr></thead><tbody>`;

  for (const d of rubrosOrdenados) {
    const delRubro = porRubro.get(d);
    html += `<tr class="cfg-rubro"><td colspan="3">` +
            `${cfgNombreRubro(d)}<span class="cfg-rubro-n">${delRubro.length}</span></td></tr>`;

    for (const f of delRubro) {
      const e = CFG_ESTADOS[f.estado];
      const chip = `<span class="cfg-chip ${e.clase}">${e.texto}</span>`;
      let rots;
      if (f.estado === "difiere") {
        rots = `<span class="cfg-rot">—${chip}</span>` +
          ARCHIVOS_TFBR.filter(a => f.porArchivo[a.id])
            .map(a => `<span class="cfg-porarch"><b>${a.label}</b> · ${f.porArchivo[a.id].join(" + ") || "sin rótulo"}</span>`)
            .join("");
      } else if (f.estado === "fuera" || f.estado === "excluida") {
        const dicho = (App.configuracion && App.configuracion.cuentas[f.cod]) || null;
        const destino = dicho && dicho.rotulo ? ` <span class="cfg-rot">${dicho.rotulo}</span>` : "";
        rots = `${chip}${destino}` +
          (f.estado === "fuera" && !destino
            ? `<span class="cfg-porarch">Ningún balance la tiene todavía. Si la configurás ahora, ` +
              `se engancha sola cuando el sumas y saldos la traiga.</span>`
            : "");
      } else {
        rots = `<span class="cfg-rot">${f.rotulos.length ? resaltar(f.rotulos.join(" + ")) : "—"}` +
               `${f.estado === "ok" ? "" : chip}</span>`;
      }

      const editando = cfgEditando === f.cod;
      const acciones = f.estado === "excluida"
        ? `<button class="cfg-btn" onclick="cfgReincorporar('${f.cod}')">Volver a usar</button>`
        : `<button class="cfg-btn" onclick="cfgElegir(${editando ? "null" : "'" + f.cod + "'"})">` +
          `${editando ? "Cerrar" : "Cambiar"}</button>`;

      html += `<tr${editando ? ' class="cfg-abierta"' : ""}${f.estado === "excluida" ? ' class="cfg-apagada"' : ""}>` +
        `<td><span class="mono">${resaltar(f.cod)}</span><span class="cfg-nom">${resaltar(f.nom)}</span></td>` +
        `<td>${rots}</td>` +
        `<td>${acciones}</td>` +
        `</tr>`;

      if (editando) {
        // Los rótulos que ya tienen cuentas van primero, con cuántas: en una lista de cien, el
        // que se busca casi siempre es uno que ya está en uso, y los vacíos son los que sobran
        // de limpiezas anteriores.
        const usados = {};
        filas.forEach(x => x.rotulos.forEach(r => { usados[cfgNorm(r)] = (usados[cfgNorm(r)] || 0) + 1; }));
        const dicho = (App.configuracion && App.configuracion.cuentas[f.cod]) || null;
        const actual = cfgNorm(f.rotulos[0] || (dicho && dicho.rotulo) || "");
        const opts = rotulos
          .map(r => ({ r, n: usados[cfgNorm(r)] || 0 }))
          .sort((a, b) => (b.n - a.n) || a.r.localeCompare(b.r, "es"))
          .map(x => `<option value="${x.r.replace(/"/g, "&quot;")}"` +
                    `${cfgNorm(x.r) === actual ? " selected" : ""}>` +
                    `${x.r}${x.n ? ` (${x.n})` : " — vacío"}</option>`)
          .join("");
        html += `<tr class="cfg-editor"><td colspan="3">` +
          `<div class="cfg-editor-caja">` +
          `<span>Mover <b>${f.nom}</b> a</span>` +
          `<select id="cfgDestino">${opts}</select>` +
          `<button class="cfg-btn primario" onclick="cfgAplicar('${f.cod}')">Aplicar a los 4</button>` +
          `<button class="cfg-btn" onclick="cfgExcluir('${f.cod}')">No usar esta cuenta</button>` +
          `<button class="cfg-btn" onclick="cfgElegir(null)">Cancelar</button>` +
          `</div></td></tr>`;
      }
    }
  }
  html += "</tbody></table>";

  if (!visibles.length) {
    html = `<div class="cfg-vacio">${cfgFiltro
      ? `Ninguna de las ${filas.length} cuentas de ${cfgHoja} coincide con «${cfgFiltro}».`
      : `No hay cuentas en «${porEstado.texto}».`}</div>`;
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

    // Queda anotado SIEMPRE, aunque no se haya podido mover en ningún archivo: si la cuenta
    // todavía no existe en ninguna hoja SALDOS, la decisión es justamente lo que hace que el
    // motor la enganche el día que aparezca.
    cfgAnotar(cod, { hoja: cfgHoja, rotulo: destino });

    if (!r.hechos.length) {
      const sinFila = f && (f.estado === "fuera" || f.estado === "excluida");
      estadoUi("cfgStatus", sinFila
        ? `"${nombre}" queda configurada en "${destino}". Ningún balance la tiene todavía: se ` +
          `engancha sola cuando el sumas y saldos la traiga.`
        : `No la moví en ningún archivo. ${r.fallos.map(x => `${x.archivo}: ${x.motivo}`).join("; ")}`,
        sinFila ? "ok" : "bad");
      cfgCambios.push({ cod, nombre, a: destino, hoja: cfgHoja, archivos: 0 });
      cfgEditando = null;
      cfgPintar();
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

function cfgBuscar(v) { cfgFiltro = v; cfgEditando = null; cfgPintar(); }

function cfgVerEstado(id) { cfgFiltroEstado = id; cfgEditando = null; cfgPintar(); }

// La decisión se anota en la configuración además de escribirse en los archivos. Las dos cosas
// hacen falta: la fórmula es lo que el Excel usa este mes, y la configuración es lo que hace
// que la decisión se vuelva a aplicar el mes que viene y en los archivos donde la cuenta
// todavía no existe.
function cfgAnotar(cod, decision) {
  if (!App.configuracion) App.configuracion = cgNormalizar(null);
  if (decision === null) delete App.configuracion.cuentas[cod];
  else App.configuracion.cuentas[cod] = decision;
}

// "No usar esta cuenta": queda anotado que no va a ningún renglón, y el control de cuentas sin
// destino deja de denunciarla. No se borra nada del Excel — si el sumas y saldos la trae, se
// pega igual, porque si no el Debe dejaría de dar igual que el Haber.
function cfgExcluir(cod) {
  const { filas } = cfgVistaUnica();
  const f = filas.find(x => x.cod === cod);
  const nombre = f ? f.nom : cod;
  cfgAnotar(cod, { excluida: true });
  cfgCambios.push({ cod, nombre, a: "no se usa", hoja: cfgHoja, archivos: 0, tipo: "excluir" });
  cfgEditando = null;
  estadoUi("cfgStatus", `"${nombre}" queda marcada como que no se usa. No se va a volver a ` +
                        `preguntar por ella ni a avisar que está sin rótulo.`, "ok");
  cfgPintar();
}

function cfgReincorporar(cod) {
  const { filas } = cfgVistaUnica();
  const f = filas.find(x => x.cod === cod);
  const nombre = f ? f.nom : cod;
  cfgAnotar(cod, null);
  cfgCambios.push({ cod, nombre, a: "vuelve a usarse", hoja: cfgHoja, archivos: 0, tipo: "reincorporar" });
  estadoUi("cfgStatus", `"${nombre}" vuelve a la lista. Asignale un renglón.`, "ok");
  cfgPintar();
}

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
    // Y las decisiones, que son las que se vuelven a aplicar el mes que viene.
    await guardarConfiguracion(App.configuracion, `Configurar cuentas: ${detalle}`.slice(0, 240));
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
