// Conecta la pantalla de informe-tfbr con motor_tfbr.js / parser_tfbr.js / fixes_tfbr.js /
// github_tfbr.js. No tiene lógica de negocio propia — solo lee archivos, llama al motor,
// y muestra lo que devuelve.

const ARCHIVOS_TFBR = [
  { id: "balance_mensual_ars", label: "Balance Mensual $", periodo: "mensual", campoSaldo: "saldo_ars" },
  { id: "balance_mensual_brl", label: "Balance Mensual R$", periodo: "mensual", campoSaldo: "saldo_brl" },
  { id: "balance_acumulado_ars", label: "Balance Acumulado $", periodo: "acumulado", campoSaldo: "saldo_ars" },
  { id: "balance_acumulado_brl", label: "Balance Acumulado R$", periodo: "acumulado", campoSaldo: "saldo_brl" },
];

const App = {
  altaBuffers: {},      // archivoId -> ArrayBuffer (subido en "primera vez", antes de guardarlo)
  maestrosCargados: {}, // archivoId -> { buffer, sha, wb? }  — el wb se arma recién al usarlo
  cuentasExport: {},    // "mensual"|"acumulado" -> cuentas parseadas
  resultados: {},       // archivoId -> { resumen, workbookBuffer }
  aprobadosBuffers: {}, // archivoId -> ArrayBuffer (subido en la revisión final)
  validaciones: {},     // archivoId -> resultado de validarRecalculado
  alineacion: null,     // resultado de alinearHojas sobre los 4 de esta corrida
  rotulosGuardados: {}, // codigo -> rótulo del Anexo II (null = sin rótulo, decidido)
  configuracion: null,  // lo que la persona configuró: ver config_guardada.js
  logLineas: [],
};

// La fecha de cierre que se carga en pantalla, en sus dos formas: el período (AAAA-MM), que es
// lo que usan el historial y el motor, y el día, que es lo que va escrito en los encabezados.
//
// Antes se cargaba el período a mano como texto ("2026-07") y el día se daba por sentado: el
// último del mes. Un cierre a mitad de mes no se podía expresar, y escribir el período a mano
// se presta a tipear "2026-7" o "07-2026" y que no ande.
function fechaDeCierre() {
  const v = (document.getElementById("fechaCierreInput") || {}).value || "";
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(v.trim());
  if (!m) return { periodo: "", dia: null, texto: "" };
  return { periodo: `${m[1]}-${m[2]}`, dia: Number(m[3]), texto: `${m[3]}/${m[2]}/${m[1]}` };
}

function periodoElegido() { return fechaDeCierre().periodo; }

// Debajo del campo se muestra qué período se va a cerrar, para que no haya dudas de que una
// fecha del 30/09 cierra septiembre.
function avisarPeriodo() {
  const el = document.getElementById("periodoAviso");
  if (!el) return;
  const f = fechaDeCierre();
  el.textContent = f.periodo
    ? `Se cierra el período ${f.periodo} y los informes van a decir ${f.texto}.`
    : "";
}

// Cómo se está entendiendo el tipo de cambio tipeado. Se muestra el número ya interpretado
// para que no haya que descubrir en el resultado que una coma se leyó como separador de miles.
function avisarTc() {
  const el = document.getElementById("tcAviso");
  if (!el) return;
  const escrito = (document.getElementById("tcCierreInput") || {}).value || "";
  if (!escrito.trim()) { el.textContent = ""; el.className = "footer-note"; return; }
  const n = ptNumeroTipeado(escrito);
  if (!n || n <= 0) {
    el.textContent = `No entiendo "${escrito.trim()}". Escribilo como 291,3301 o 291.3301.`;
    el.className = "footer-note aviso-mal";
    return;
  }
  el.textContent = `Se va a usar ${n}.`;
  el.className = "footer-note";
}

// Las dos librerías pesan 1,8 MB de los 2 MB de la página, y no hacen falta para mostrarla:
// ExcelJS recién se usa al procesar y XLSX al leer el export. Cargarlas al abrir obligaba a
// esperarlas en cada refresh aunque no se fuera a procesar nada.
//
// Se cargan una sola vez y se recuerda la promesa: dos llamadas simultáneas esperan la misma
// carga en vez de pedir el archivo dos veces.
const LIBRERIAS = {
  exceljs: { src: "vendor/exceljs.min.js", global: "ExcelJS" },
  xlsx: { src: "vendor/xlsx.full.min.js", global: "XLSX" },
};
const _libreriasPedidas = {};

function cargarLibreria(cual) {
  const lib = LIBRERIAS[cual];
  if (!lib) return Promise.reject(new Error(`Librería desconocida: ${cual}`));
  if (window[lib.global]) return Promise.resolve();
  if (_libreriasPedidas[cual]) return _libreriasPedidas[cual];

  _libreriasPedidas[cual] = new Promise((listo, error) => {
    const el = document.createElement("script");
    el.src = lib.src;
    el.onload = () => listo();
    el.onerror = () => {
      delete _libreriasPedidas[cual];   // que un fallo de red se pueda reintentar
      error(new Error(`No pude cargar ${lib.src}. Revisá la conexión y volvé a intentar.`));
    };
    document.head.appendChild(el);
  });
  return _libreriasPedidas[cual];
}

// El workbook de un maestro, armado recién cuando se lo necesita.
//
// Al abrir la página sólo se baja el archivo y se guarda el buffer: parsear los cuatro con
// ExcelJS para averiguar si existen costaba varios segundos en cada refresh, y el 90% de las
// veces la persona entra a mirar el historial o a configurar cuentas y no procesa nada.
async function maestroWb(archivoId) {
  const cargado = App.maestrosCargados[archivoId];
  if (!cargado) return null;
  if (cargado.wb) return cargado.wb;
  await cargarLibreria("exceljs");
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(cargado.buffer);
  cargado.wb = wb;
  return wb;
}

function mostrar(id, visible) {
  const el = document.getElementById(id);
  if (el) el.classList.toggle("hidden", !visible);
}

function log(msg) {
  App.logLineas.push(msg);
  const el = document.getElementById("logTexto");
  if (el) el.textContent = App.logLineas.join("\n");
  mostrar("cardLog", true);
  console.log(msg);
}

// Sin texto no hay mensaje: limpiar es limpiar. Antes dejaba el recuadro gris vacío —
// una barra sin contenido flotando debajo de los botones.
function estadoUi(elId, texto, clase) {
  const el = document.getElementById(elId);
  if (!el) return;
  el.innerHTML = texto
    ? `<div class="status-msg ${clase || ""}">${texto}</div>`
    : "";
}

// ------------------------------------------------------------ arranque / config

async function iniciar() {
  if (!hasGhtSettings()) {
    mostrar("cardSinConfig", true);
    return;
  }
  const s = loadGhtSettings();
  document.getElementById("cfgToken").value = s.token || "";
  document.getElementById("cfgRepo").value = s.repo || "";
  document.getElementById("cfgRama").value = s.rama || "main";
  document.getElementById("cfgCarpeta").value = s.carpeta || "informe-tfbr";
  await cargarRotulosGuardados();
  await revisarMaestrosExistentes();
}

// Igual que la de cuentas: ventana encima, no una card que alargue la página.
function abrirConfig() {
  mostrar("ovConfig", true);
  document.body.classList.add("sin-scroll");
  // Si ya está configurado, se muestra lo que hay: abrirla para cambiar la rama no tendría
  // que obligar a volver a pegar el token.
  const s = loadGhtSettings();
  if (s.token) document.getElementById("cfgToken").value = s.token;
  if (s.repo) document.getElementById("cfgRepo").value = s.repo;
  if (s.rama) document.getElementById("cfgRama").value = s.rama;
  if (s.carpeta) document.getElementById("cfgCarpeta").value = s.carpeta;
}

function cerrarConfig() {
  mostrar("ovConfig", false);
  soltarScrollSiNoQuedaVentana();
}

// El bloqueo del scroll de fondo se suelta cuando no queda NINGUNA ventana abierta: si no,
// cerrar una con la otra todavía arriba dejaba la página de atrás moviéndose.
function soltarScrollSiNoQuedaVentana() {
  const abierta = ["ovConfig", "ovCuentas"].some(id => {
    const el = document.getElementById(id);
    return el && !el.classList.contains("hidden");
  });
  if (!abierta) document.body.classList.remove("sin-scroll");
}

function cfgFondoGh(ev) {
  if (ev.target && ev.target.id === "ovConfig") cerrarConfig();
}

document.addEventListener("keydown", (ev) => {
  if (ev.key !== "Escape") return;
  const ov = document.getElementById("ovConfig");
  if (ov && !ov.classList.contains("hidden")) cerrarConfig();
});

function mostrarAyudaToken() { mostrar("ayudaToken", true); }

async function guardarConfig() {
  const s = {
    token: document.getElementById("cfgToken").value.trim(),
    repo: document.getElementById("cfgRepo").value.trim(),
    rama: document.getElementById("cfgRama").value.trim() || "main",
    carpeta: document.getElementById("cfgCarpeta").value.trim() || "informe-tfbr",
  };
  if (!s.token || !s.repo) {
    estadoUi("configStatus", "Falta el token o el repositorio.", "bad");
    return;
  }
  saveGhtSettings(s);
  mostrar("spinnerConfig", true);
  try {
    await ghtLeerEstado(); // si esto no tira, la conexión sirve (404 = repo ok, archivo nuevo)
    estadoUi("configStatus", "Conectado correctamente.", "ok");
    mostrar("cardSinConfig", false);
    await cargarRotulosGuardados();
    await revisarMaestrosExistentes();
  } catch (e) {
    estadoUi("configStatus", "No pude conectar: " + e.message, "bad");
  } finally {
    mostrar("spinnerConfig", false);
  }
}

// ------------------------------------------------------------ primera vez (alta de maestros)

// Las decisiones de rótulo tomadas en meses anteriores. Se leen una vez al arrancar: con
// ellas, una cuenta que ya se decidió no vuelve a preguntarse nunca.
async function cargarRotulosGuardados() {
  try {
    const r = await ghtLeerEstado();
    App.rotulosGuardados = (r && r.estado && r.estado.rotulosCuentas) || {};
    App.configuracion = cgNormalizar(r && r.estado && r.estado.configuracion);
  } catch (e) {
    App.rotulosGuardados = {};
    App.configuracion = cgNormalizar(null);
    log("No pude leer la configuración guardada: " + e.message);
  }
}

// Guarda la configuración en estado_tfbr.json, sin pisar el resto del estado.
async function guardarConfiguracion(config, mensaje) {
  const previo = await ghtLeerEstado();
  const estado = (previo && previo.estado) || { historial: [] };
  estado.configuracion = cgNormalizar(config);
  await ghtGuardarEstado(estado, mensaje);
  App.configuracion = estado.configuracion;
}

async function revisarMaestrosExistentes() {
  // Los cuatro a la vez, y sin parsearlos: acá sólo hace falta saber si están. Uno por uno y
  // pasándolos por ExcelJS eran cuatro viajes a GitHub en fila más cuatro parseos, en cada
  // refresh de la página. El workbook se arma recién cuando se usa (ver maestroWb).
  App.maestrosCargados = {};
  const bajados = await Promise.all(ARCHIVOS_TFBR.map(async (a) => {
    try { return { a, r: await ghtLeerMaestro(a.id) }; }
    catch (e) { log(`No pude leer el maestro de ${a.label}: ${e.message}`); return { a, r: null }; }
  }));

  const faltantes = [];
  for (const { a, r } of bajados) {
    if (!r) { faltantes.push(a); continue; }
    App.maestrosCargados[a.id] = { buffer: r.buffer, sha: r.sha };
  }
  if (faltantes.length) {
    pintarDropzonesAlta(faltantes);
    mostrar("cardAlta", true);
  } else {
    mostrar("cardExport", true);
  }
}

function pintarDropzonesAlta(faltantes) {
  const cont = document.getElementById("altaDropzones");
  cont.innerHTML = "";
  for (const a of faltantes) {
    const div = document.createElement("label");
    div.className = "dropzone";
    div.innerHTML = `
      <input type="file" accept=".xlsx" data-archivo="${a.id}">
      <div id="txtAlta_${a.id}">${a.label} — subir guardado como .xlsx</div>`;
    cont.appendChild(div);
    div.querySelector("input").addEventListener("change", (ev) => onAltaArchivo(a, ev));
  }
}

async function onAltaArchivo(a, ev) {
  const file = ev.target.files[0];
  if (!file) return;
  const txt = document.getElementById(`txtAlta_${a.id}`);
  if (!file.name.toLowerCase().endsWith(".xlsx")) {
    txt.textContent = `${a.label}: tiene que ser .xlsx (guardalo desde Excel primero).`;
    return;
  }
  const buffer = await file.arrayBuffer();
  try {
    await cargarLibreria("exceljs");
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    if (!wb.getWorksheet("SALDOS")) throw new Error("este archivo no tiene una hoja 'SALDOS'.");
    App.altaBuffers[a.id] = buffer;
    txt.textContent = `${a.label}: ${file.name} ✓`;
  } catch (e) {
    txt.textContent = `${a.label}: no pude leerlo (${e.message}).`;
    return;
  }
  document.getElementById("btnGuardarAlta").disabled =
    Object.keys(App.altaBuffers).length === 0;
}

async function guardarAlta() {
  mostrar("spinnerAlta", true);
  document.getElementById("btnGuardarAlta").disabled = true;
  try {
    for (const [archivoId, buffer] of Object.entries(App.altaBuffers)) {
      await ghtGuardarMaestro(archivoId, buffer, `Alta inicial: ${archivoId}`);
      log(`Guardado en GitHub: ${archivoId}`);
    }
    estadoUi("altaStatus", "Guardado. Cargando los maestros…", "ok");
    App.altaBuffers = {};
    await revisarMaestrosExistentes();
    mostrar("cardAlta", Object.keys(App.maestrosCargados).length < ARCHIVOS_TFBR.length);
  } catch (e) {
    estadoUi("altaStatus", "No pude guardar: " + e.message, "bad");
  } finally {
    mostrar("spinnerAlta", false);
  }
}

// ------------------------------------------------------------ carga del período

function revisarListoParaProcesar() {
  const periodo = periodoElegido();
  const tcCierre = document.getElementById("tcCierreInput").value.trim();
  const listo = /^\d{4}-\d{2}$/.test(periodo) && tcCierre !== "" &&
    App.cuentasExport.mensual && App.cuentasExport.acumulado;
  document.getElementById("btnProcesar").disabled = !listo;
}

async function onExportArchivo(periodo, ev) {
  const file = ev.target.files[0];
  if (!file) return;
  const txt = document.getElementById(periodo === "mensual" ? "txtMensual" : "txtAcumulado");
  try {
    await cargarLibreria("xlsx");
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, { type: "array", cellFormula: true });
    const ws = wb.Sheets["Sheet1"];
    if (!ws) throw new Error("el archivo no tiene una hoja 'Sheet1'.");
    const filas = XLSX.utils.sheet_to_json(ws, { header: 1 });
    const parsed = parseSumasYSaldosTFBR(filas, ws["!merges"]);
    App.cuentasExport[periodo] = parsed.cuentas;
    txt.textContent =
      `${file.name} ✓ — ${parsed.cuentas.length} cuentas, total $ ${parsed.totales.saldo_ars.toFixed(2)}`;

    // El tipo de cambio de cierre, si el export lo trae. Desde agosto de 2026 viene en el
    // encabezado ("TC 31/8  291.3301") y no hay razón para tipearlo. Se completa el campo y se
    // deja editable: si el export no lo trae, o si contaduría quiere usar otro, se escribe.
    if (parsed.tcCierre) {
      const campo = document.getElementById("tcCierreInput");
      const yaHabia = campo.value.trim();
      if (!yaHabia || Number(yaHabia) === parsed.tcCierre.valor) {
        campo.value = parsed.tcCierre.valor;
        avisarTc();
        txt.textContent += ` · TC ${parsed.tcCierre.valor} tomado del archivo`;
      } else {
        txt.textContent += ` · ⚠ el archivo trae TC ${parsed.tcCierre.valor} y en pantalla ` +
                           `hay ${yaHabia}: revisá cuál corresponde`;
      }
      revisarListoParaProcesar();
    }
    // El reporte corto de Onvio se come las cuentas de resultado que cierran en cero en pesos
    // pero no en reales, y el unico sintoma seria la diferencia de cambio saliendo corta.
    if (parsed.pareceReporteCorto) {
      txt.textContent += ` · ⚠ parece el reporte SIN las cuentas saldadas: pedilo en Onvio con ` +
                         `las cuentas en cero incluidas, si no se pierden las de resultado que ` +
                         `cierran en cero en pesos y no en reales`;
    }

    if (parsed.discrepanciasCapitulo.length) {
      txt.textContent += ` (⚠ ${parsed.discrepanciasCapitulo.length} discrepancia(s) de capítulo)`;
    }
  } catch (e) {
    txt.textContent = `${file.name}: no pude leerlo (${e.message}).`;
    App.cuentasExport[periodo] = null;
  }
  revisarListoParaProcesar();
}

document.addEventListener("DOMContentLoaded", () => {
  document.getElementById("fileMensual")?.addEventListener("change", (e) => onExportArchivo("mensual", e));
  document.getElementById("fileAcumulado")?.addEventListener("change", (e) => onExportArchivo("acumulado", e));
  document.getElementById("tcCierreInput")?.addEventListener("input", avisarTc);
  document.getElementById("fechaCierreInput")?.addEventListener("input", () => {
    avisarPeriodo();
    revisarListoParaProcesar();
  });
  document.getElementById("tcCierreInput")?.addEventListener("input", revisarListoParaProcesar);
  iniciar();
});

async function procesarPeriodo() {
  mostrar("spinnerProcesar", true);
  document.getElementById("btnProcesar").disabled = true;
  App.logLineas = [];
  try {
    await cargarLibreria("exceljs");
    // Se acepta "291,3301" y "291.3301": el TC se tipea con coma decimal y Number() de eso da
    // NaN, así que el cierre frenaba diciendo que faltaba el dato con el número en la pantalla.
    const escrito = document.getElementById("tcCierreInput").value.trim();
    const tcDelCierre = ptNumeroTipeado(escrito);
    if (!tcDelCierre || tcDelCierre <= 0) {
      throw new Error(escrito
        ? `No entiendo el tipo de cambio "${escrito}". Escribilo como 291,3301 o 291.3301.`
        : "Falta el tipo de cambio de cierre: sin él no se puede calcular el saldo en reales " +
          "de las cuentas de activo, pasivo y patrimonio.");
    }
    // La diferencia de cambio del mes la calcula el Mensual R$ —es el residuo de su plan— y la
    // usa después el Acumulado R$, que es el que tiene el cuadro de meses. Por eso viaja acá
    // afuera del bucle y no dentro del motor: son dos archivos distintos.
    let difCambioDelMes = null;

    for (const a of ARCHIVOS_TFBR) {
      log(`\n=== ${a.label} ===`);
      const cargado = App.maestrosCargados[a.id];
      if (!cargado) throw new Error(`Falta el maestro de ${a.label}.`);

      // se trabaja sobre una copia en memoria: si algo falla más adelante, el maestro
      // guardado en GitHub no se tocó
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(cargado.buffer);

      // Los errores que el archivo YA tenía antes de que lo tocáramos: son la línea de base
      // contra la que después se comparan los del archivo aprobado, para distinguir un
      // problema nuevo de uno viejo sin depender de una lista escrita a mano.
      // con la fórmula de cada una, para que un error que ya venía no cuente como nuevo
      // cuando el motor lo corre de fila (ver vtHuellaError)
      const erroresPrevios = celdasEnErrorDetalle(wb);

      // El saldo en reales se calcula acá, no se toma del export: las cuentas de activo,
      // pasivo y patrimonio salen del saldo en pesos dividido el tipo de cambio de cierre.
      // Ver convertirSaldosEnReales. Hace falta el TC, así que va después de cargarlo.
      const cuentasExport = convertirSaldosEnReales(App.cuentasExport[a.periodo], tcDelCierre);

      // El Anexo I se carga desde el export ACUMULADO, y ese export tambien tiene que pasar por
      // la conversion. Sin esto, el Anexo I del Mensual R$ tomaba el "Saldo (R$)" crudo —el
      // valor a tipo de cambio historico— en vez del saldo en pesos dividido el TC de cierre:
      // el valor de origen de bienes de uso daba 1.239.980,86 donde el Acumulado R$ decia
      // 694.401,13, y el "Bienes de uso" del EESP terminaba en negativo. Para los dos archivos
      // en pesos no cambia nada: la conversion no toca el saldo en pesos.
      const cuentasAcumuladoConv = convertirSaldosEnReales(App.cuentasExport.acumulado, tcDelCierre);
      const { resumen, planDeCuentas, escritas } =
        procesarMaestroTFBR({ wb, cuentasExport, campoSaldo: a.campoSaldo, archivoId: a.id,
                              rotulosGuardados: App.rotulosGuardados,
                              configuracion: App.configuracion,
                              // El Anexo I es información acumulada en los cuatro archivos,
                              // así que sale del export acumulado aunque el archivo sea mensual.
                              cuentasAcumulado: cuentasAcumuladoConv,
                              log });

      // el TC de cierre y la cifra de dif de cambio solo existen en uno de los 4 archivos:
      // escribirDatosDelPeriodo se fija solo si este los tiene, y avisa lo que no pudo cargar
      const periodoDatos = escribirDatosDelPeriodo(wb, {
        periodo: fechaDeCierre().periodo,
        diaCierre: fechaDeCierre().dia,
        tcCierre: document.getElementById("tcCierreInput").value.trim(),
        escritas,
        difCambioDelMes,
      }, log);

      // El residuo del Mensual R$ ES la diferencia de cambio del mes (ver pfResiduoDelPlan).
      if (a.id === "balance_mensual_brl" && periodoDatos.residuo !== null &&
          periodoDatos.residuo !== undefined) {
        difCambioDelMes = periodoDatos.residuo;
        App.difCambioDelMes = difCambioDelMes;
      }

      aplicarFixesAprobados(wb, a.id, log);

      const pendientes = pendientesManuales(wb, a.label);

      const outBuffer = await wb.xlsx.writeBuffer();
      App.resultados[a.id] = {
        resumen, periodoDatos, planDeCuentas, escritas, erroresPrevios, pendientes,
        // El workbook queda vivo: si hay cuentas nuevas que enganchar, la respuesta se
        // escribe sobre ESTE libro y recién ahí se vuelve a generar el .xlsx. Volver a
        // cargarlo desde el buffer costaría medio segundo por archivo y cuatro veces más
        // memoria por nada.
        wb,
        workbookBuffer: outBuffer,
      };
    }
    // Con los cuatro ya procesados, que lean las mismas cuentas en cada hoja. Va acá y no en
    // el motor porque necesita ver los cuatro a la vez: lo que decide dónde va una cuenta es
    // dónde la pusieron los otros archivos.
    App.alineacion = alinearHojas(
      ARCHIVOS_TFBR.map(a => ({ id: a.id, label: a.label, wb: App.resultados[a.id] && App.resultados[a.id].wb }))
        .filter(x => x.wb),
      log);
    // Los .xlsx se vuelven a generar: el enganche recién hecho tiene que estar en lo que se baja.
    for (const a of ARCHIVOS_TFBR) {
      const r = App.resultados[a.id];
      if (r && r.wb) r.workbookBuffer = await r.wb.xlsx.writeBuffer();
    }

    pintarCuentasNuevas();
    pintarResultado();
    mostrar("cardResultado", true);
    document.getElementById("cardResultado").scrollIntoView({ behavior: "smooth" });
  } catch (e) {
    estadoUi("exportStatus", "No pude procesar: " + e.message, "bad");
    log("ERROR: " + e.message);
  } finally {
    mostrar("spinnerProcesar", false);
    document.getElementById("btnProcesar").disabled = false;
  }
}

// Lo que la alineación entre los cuatro archivos hizo y lo que no pudo.
function pintarAlineacion() {
  const cont = document.getElementById("alineacionResumen");
  if (!cont) return;
  const a = App.alineacion;
  if (!a) { cont.innerHTML = ""; return; }

  let html = "";
  if (a.enganchadas.length) {
    const porHoja = {};
    a.enganchadas.forEach(x => { porHoja[x.hoja] = (porHoja[x.hoja] || 0) + 1; });
    html += `<p class="footer-note">✓ ${a.enganchadas.length} cuenta(s) se engancharon para que ` +
      `los cuatro archivos lean lo mismo (` +
      Object.entries(porHoja).map(([h, n]) => `${h}: ${n}`).join(", ") + `).</p>`;
  }
  // Una cuenta que cada archivo pone en un renglón distinto no se resuelve sola: elegir uno
  // sería mover plata de un renglón del balance a otro sin que nadie lo apruebe.
  if (a.discrepan.length) {
    html += `<div class="aviso-plata"><b>${a.discrepan.length} cuenta(s) están en renglones ` +
      `distintos según el archivo.</b> No elijo por mi cuenta: definilas en Configurar cuentas.<ul>` +
      a.discrepan.map(d =>
        `<li><span class="mono">${d.cod}</span> en ${d.hoja} — ${d.detalle}</li>`).join("") +
      `</ul></div>`;
  }
  if (a.pendientes.length) {
    html += a.pendientes.map(p =>
      `<p class="footer-note">⚠ ${p.archivo}: ${p.cod} no se pudo poner en ${p.hoja} → ` +
      `"${p.rotulo}" — ${p.motivo}.</p>`).join("");
  }
  cont.innerHTML = html;
}

// ------------------------------------------------- cuentas nuevas: ¿a qué rótulo van?

// Junta las cuentas sueltas de los cuatro archivos en una sola lista. Una cuenta nueva
// aparece en los cuatro, y los cuatro tienen que engancharla al mismo rótulo — preguntar
// cuatro veces lo mismo sería la forma más rápida de que los archivos queden desalineados.
function nuevasPendientes() {
  const porCod = {};
  for (const a of ARCHIVOS_TFBR) {
    const r = App.resultados[a.id];
    if (!r || !r.resumen.sinRotulo) continue;
    for (const c of r.resumen.sinRotulo) {
      if (!porCod[c.cod]) porCod[c.cod] = { cod: c.cod, nom: c.nom, archivos: [] };
      porCod[c.cod].archivos.push(a.id);
    }
  }
  return Object.values(porCod).sort((x, y) => x.cod.localeCompare(y.cod));
}

// Los rótulos que existen en LOS CUATRO. Ofrecer uno que solo tiene un archivo dejaría la
// cuenta enganchada en unos y suelta en otros.
function rotulosComunes() {
  const listas = [];
  for (const a of ARCHIVOS_TFBR) {
    const r = App.resultados[a.id];
    if (r && r.wb) listas.push(cnRotulosDisponibles(r.wb));
  }
  if (!listas.length) return [];
  const norm = (t) => String(t).normalize("NFD").replace(/[̀-ͯ]/g, "").toUpperCase().trim();
  const restantes = listas.slice(1).map(l => new Set(l.map(norm)));
  return listas[0].filter(r => restantes.every(s => s.has(norm(r))));
}

function pintarCuentasNuevas() {
  const pendientes = nuevasPendientes();
  const cont = document.getElementById("nuevasLista");
  if (!pendientes.length) {
    mostrar("cardNuevas", false);
    const auto = ARCHIVOS_TFBR.reduce((n, a) => {
      const r = App.resultados[a.id];
      return Math.max(n, (r && r.resumen.rotulosAuto) ? r.resumen.rotulosAuto.length : 0);
    }, 0);
    if (auto) log(`  ${auto} cuenta(s) se engancharon solas con decisiones de meses anteriores.`);
    return;
  }

  const rotulos = rotulosComunes();
  const opts = `<option value="">— dejar sin rótulo —</option>` +
    rotulos.map(r => `<option value="${r.replace(/"/g, "&quot;")}">${r}</option>`).join("");

  cont.innerHTML = pendientes.map(p => `
    <div class="nueva-fila">
      <div>
        <span class="mono">${p.cod}</span>
        <span class="cfg-nom">${p.nom}</span>
        <span class="footer-note" style="margin:0;">Aparece en ${p.archivos.length} de ${ARCHIVOS_TFBR.length} archivos</span>
      </div>
      <select id="nueva_${p.cod}">${opts}</select>
    </div>`).join("");

  document.getElementById("nuevasResumen").innerHTML =
    `<b>${pendientes.length}</b> cuenta(s) de gasto no las lee ningún renglón del Anexo II. ` +
    `Mientras sigan así, su importe entra en SALDOS pero no llega al estado de resultados.`;
  estadoUi("nuevasStatus", "", "");
  mostrar("cardNuevas", true);
}

// Aplica lo elegido a los cuatro archivos y guarda la decisión, para no volver a preguntar.
// "Dejar sin rótulo" también es una decisión y también se guarda: hay cuentas que a propósito
// no van a ningún renglón, y volver a preguntar por ellas todos los meses sería ruido.
async function aplicarCuentasNuevas() {
  const btn = document.getElementById("btnAplicarNuevas");
  btn.disabled = true;
  mostrar("spinnerNuevas", true);
  try {
    const pendientes = nuevasPendientes();
    const enganchadas = [], sinRotulo = [];

    for (const p of pendientes) {
      const rotulo = document.getElementById(`nueva_${p.cod}`).value;
      App.rotulosGuardados[p.cod] = rotulo || null;
      if (!rotulo) { sinRotulo.push(p); continue; }
      const donde = [];
      for (const a of ARCHIVOS_TFBR) {
        const r = App.resultados[a.id];
        if (!r || !r.wb) continue;
        const res = engancharCuentaEnRotulo(r.wb, p.cod, rotulo);
        if (res.hecho) donde.push(a.label);
        else log(`  ${a.label}: no pude enganchar ${p.cod} — ${res.motivo}`);
      }
      enganchadas.push({ ...p, rotulo, archivos: donde });
      log(`  ${p.cod} ${p.nom} → "${rotulo}" en ${donde.length} archivo(s).`);
    }

    // Los .xlsx se vuelven a generar: los que se bajen ahora ya traen el enganche.
    for (const a of ARCHIVOS_TFBR) {
      const r = App.resultados[a.id];
      if (r && r.wb) r.workbookBuffer = await r.wb.xlsx.writeBuffer();
    }

    // La decisión queda en GitHub. Si esto falla, lo aplicado a los archivos sigue valiendo:
    // se avisa que el mes que viene va a volver a preguntar, y no se pierde el trabajo.
    let aviso = "";
    try {
      const previo = await ghtLeerEstado();
      const estado = (previo && previo.estado) || { historial: [] };
      estado.rotulosCuentas = { ...(estado.rotulosCuentas || {}), ...App.rotulosGuardados };
      await ghtGuardarEstado(estado, `Rótulos de cuentas nuevas: ${pendientes.map(p => p.cod).join(", ")}`);
    } catch (e) {
      aviso = ` No pude guardar la decisión en GitHub (${e.message}), así que el mes que viene ` +
              `va a volver a preguntar. Los archivos de este mes sí quedaron bien.`;
    }

    // Con los cuatro ya procesados, que lean las mismas cuentas en cada hoja. Va acá y no en
    // el motor porque necesita ver los cuatro a la vez: lo que decide dónde va una cuenta es
    // dónde la pusieron los otros archivos.
    App.alineacion = alinearHojas(
      ARCHIVOS_TFBR.map(a => ({ id: a.id, label: a.label, wb: App.resultados[a.id] && App.resultados[a.id].wb }))
        .filter(x => x.wb),
      log);
    // Los .xlsx se vuelven a generar: el enganche recién hecho tiene que estar en lo que se baja.
    for (const a of ARCHIVOS_TFBR) {
      const r = App.resultados[a.id];
      if (r && r.wb) r.workbookBuffer = await r.wb.xlsx.writeBuffer();
    }

    pintarCuentasNuevas();
    pintarResultado();
    estadoUi("nuevasStatus",
      `${enganchadas.length} cuenta(s) enganchadas` +
      (sinRotulo.length ? `, ${sinRotulo.length} quedaron sin rótulo a propósito` : "") +
      `. Volvé a bajar los borradores: los de antes no tienen el cambio.` + aviso,
      aviso ? "" : "ok");
  } catch (e) {
    estadoUi("nuevasStatus", "No pude aplicarlo: " + e.message, "bad");
    log("ERROR: " + e.message);
  } finally {
    mostrar("spinnerNuevas", false);
    btn.disabled = false;
  }
}

// ------------------------------------------------------------ resultado / descargas

function pintarResultado() {
  pintarAlineacion();
  const cont = document.getElementById("resultadoResumen");
  cont.innerHTML = "";
  for (const a of ARCHIVOS_TFBR) {
    const r = App.resultados[a.id];
    if (!r) continue;
    const s = r.resumen;
    const badgeClase = s.sinMapear.length ? "bad" : "ok";
    const badgeTxto = s.sinMapear.length ? `${s.sinMapear.length} sin mapear` : "OK";
    const div = document.createElement("div");
    div.style.marginBottom = "14px";
    let extra = "";
    if (s.sinMapear.length) {
      extra += `<br><span class="footer-note">Sin mapear (no entran en ningún total): ` +
        s.sinMapear.map(c => `${c.codigo} ${c.nombre}`).join(", ") + `</span>`;
    }
    for (const alta of (s.altas || [])) {
      extra += `<br><span class="footer-note">➕ Alta: <b>${alta.clave}</b> (fila ${alta.fila})` +
        (alta.gemela
          ? ` — ya estaba con el código mal (<code>${alta.gemela}</code>); se movieron ` +
            `${alta.repuntadas.length} referencia(s) a la fila nueva: ${alta.repuntadas.join(", ")}`
          : "") + `</span>`;
    }
    // Plata de este mes que no llega a ninguna hoja. Se muestra arriba de todo lo demás y
    // con el importe: es lo único del resumen que significa que falta plata en el informe.
    if ((s.sinDestino || []).length) {
      const suma = s.sinDestino.reduce((t, c) => t + c.importe, 0);
      extra += `<div class="aviso-plata"><b>${s.sinDestino.length} cuenta(s) con ` +
        `${suma.toFixed(2)} que ninguna hoja lee.</b> El importe entra en SALDOS —el balance ` +
        `cierra igual— pero no llega a ningún estado.<ul>` +
        s.sinDestino.map(c =>
          `<li><span class="mono">${c.cod}</span> ${c.nom} — <b>${c.importe.toFixed(2)}</b></li>`
        ).join("") + `</ul></div>`;
    }
    for (const r of (s.rangosExpandidos || [])) {
      extra += `<br><span class="footer-note">✓ ${r.donde}: el rango ${r.rango} pasó a nombrar ` +
        `sus ${r.cuentas} cuentas una por una.</span>`;
    }
    for (const r of (s.rangosSalteados || [])) {
      extra += `<br><span class="footer-note">⚠ ${r.donde}: no expandí el rango — ${r.motivo}.</span>`;
    }
    for (const se of (s.sinEnganchar || [])) {
      extra += `<br><span class="footer-note">⚠ <b>${se.clave}</b> quedó en SALDOS pero sin ` +
        `línea en el Anexo II: hay que agregarla a mano o su importe no llega al estado de ` +
        `resultados.</span>`;
    }
    for (const d of (s.duplicadas || [])) {
      const etiqueta = {
        nombre_distinto: "dos cuentas distintas con el mismo código",
        texto_identico: "doble conteo",
        texto_distinto: "una de las dos nunca levanta su importe",
      }[d.tipo] || d.tipo;
      extra += `<br><span class="footer-note">⚠ <b>${d.codigo}</b> (filas ` +
        `${d.filas.map(f => f.fila).join(" y ")}) — ${etiqueta}: ${d.motivo}</span>`;
    }
    for (const h of (r.periodoDatos ? r.periodoDatos.hecho : [])) {
      extra += `<br><span class="footer-note">✓ ${h}</span>`;
    }
    for (const p of (r.periodoDatos ? r.periodoDatos.pendiente : [])) {
      extra += `<br><span class="footer-note">⚠ ${p}</span>`;
    }
    div.innerHTML = `
      <b>${a.label}</b> <span class="badge ${badgeClase}">${badgeTxto}</span><br>
      <span class="footer-note">
        ${s.cuentasEscritas} de ${s.cuentasExport} cuentas escritas · total $ ${s.totalEscrito.toFixed(2)}
      </span>${extra}`;
    cont.appendChild(div);
  }

  const desc = document.getElementById("resultadoDescargas");
  desc.innerHTML = "";
  for (const a of ARCHIVOS_TFBR) {
    const r = App.resultados[a.id];
    if (!r) continue;
    const btn = document.createElement("button");
    btn.className = "secondary";
    btn.textContent = "Descargar borrador — " + a.label;
    btn.onclick = () => descargarBorrador(a);
    desc.appendChild(btn);
  }
}

function bajarComoArchivo(buffer, nombre) {
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const el = document.createElement("a");
  el.href = url;
  el.download = nombre;
  el.click();
  URL.revokeObjectURL(url);
}

function descargarBorrador(a) {
  const r = App.resultados[a.id];
  if (!r) return;
  bajarComoArchivo(r.workbookBuffer, `${a.label} - BORRADOR.xlsx`);
}

// El checklist de lo que hay que completar a mano, con la ubicación de cada cosa. Va en el
// paso de revisión y no antes: es justo lo que hay que hacer en Excel con el borrador abierto.
function pintarChecklistManual() {
  const cont = document.getElementById("cierreChecklist");
  if (!cont) return;
  let html = "<p class='footer-note'>Antes de aprobar, con cada borrador abierto en Excel, " +
             "completá lo que la app no toca:</p>";
  for (const a of ARCHIVOS_TFBR) {
    const r = App.resultados[a.id];
    if (!r || !r.pendientes) continue;
    html += `<div style="margin-top:12px;"><b>${a.label}</b><ul style="margin:6px 0;">`;
    for (const p of r.pendientes.puntos) {
      html += `<li class="footer-note">${p.que} — <i>${p.donde}</i></li>`;
    }
    for (const pend of (r.periodoDatos ? r.periodoDatos.pendiente : [])) {
      html += `<li class="footer-note">${pend}</li>`;
    }
    html += "</ul></div>";
  }
  cont.innerHTML = html;
}

function irARevision() {
  pintarChecklistManual();
  pintarDropzonesCierre();
  mostrar("cardCierre", true);
  document.getElementById("cardCierre").scrollIntoView({ behavior: "smooth" });
}

// ------------------------------------------------------------ revisión y aprobación agrupada

function pintarDropzonesCierre() {
  const cont = document.getElementById("cierreDropzones");
  cont.innerHTML = "";
  for (const a of ARCHIVOS_TFBR) {
    const div = document.createElement("label");
    div.className = "dropzone";
    div.innerHTML = `
      <input type="file" accept=".xlsx" data-archivo="${a.id}">
      <div id="txtCierre_${a.id}">${a.label} — subir revisado y guardado en Excel</div>`;
    cont.appendChild(div);
    div.querySelector("input").addEventListener("change", (ev) => onCierreArchivo(a, ev));
  }
}

async function onCierreArchivo(a, ev) {
  const file = ev.target.files[0];
  if (!file) return;
  const txt = document.getElementById(`txtCierre_${a.id}`);
  try {
    await cargarLibreria("exceljs");
    const buffer = await file.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    if (!wb.getWorksheet("SALDOS")) throw new Error("no tiene una hoja 'SALDOS'.");

    // Recién acá los controles valen: este archivo ya pasó por Excel, así que las fórmulas
    // traen los números de este mes y no los del anterior.
    const r = App.resultados[a.id];
    if (!r) throw new Error("todavía no se procesó este archivo en esta corrida.");
    const v = validarRecalculado(wb, {
      planDeCuentas: r.planDeCuentas,
      escritas: r.escritas,
      erroresPrevios: r.erroresPrevios,
    });
    App.validaciones[a.id] = v;

    const lineas = v.controles.map(c =>
      `${c.pasa ? "✓" : (c.soloAviso ? "⚠" : "✗")} ${c.nombre} — ${c.detalle}`).join("<br>");
    txt.innerHTML = `<b>${a.label}: ${file.name}</b><br>` +
      `<span class="footer-note">${lineas}</span>`;

    if (v.pasa) {
      App.aprobadosBuffers[a.id] = buffer;
    } else {
      delete App.aprobadosBuffers[a.id];
    }
  } catch (e) {
    txt.textContent = `${a.label}: no pude leerlo (${e.message}).`;
    delete App.aprobadosBuffers[a.id];
    delete App.validaciones[a.id];
  }

  const todosOk = Object.keys(App.aprobadosBuffers).length === ARCHIVOS_TFBR.length;
  document.getElementById("btnCerrarMes").disabled = !todosOk;
  const fallan = ARCHIVOS_TFBR.filter(x => App.validaciones[x.id] && !App.validaciones[x.id].pasa);
  if (fallan.length) {
    estadoUi("cierreStatus",
      `No se puede cerrar el mes: ${fallan.map(x => x.label).join(", ")} ` +
      `${fallan.length === 1 ? "no pasa" : "no pasan"} los controles. Revisá el detalle arriba.`, "bad");
  } else if (todosOk) {
    estadoUi("cierreStatus", "Los 4 archivos pasaron los controles.", "ok");
  }
}

async function cerrarMes() {
  mostrar("spinnerCierre", true);
  document.getElementById("btnCerrarMes").disabled = true;
  try {
    const periodo = periodoElegido();
    const previo = await ghtLeerEstado();
    const estado = (previo && previo.estado) || { historial: [] };
    estado.periodoActual = periodo;
    estado.historial = estado.historial || [];
    estado.historial.push({
      periodo,
      fecha: new Date().toISOString(),
      tcCierre: document.getElementById("tcCierreInput").value.trim(),
      resumen: Object.fromEntries(
        ARCHIVOS_TFBR.map(a => [a.id, App.resultados[a.id] ? App.resultados[a.id].resumen : null])
      ),
    });

    const guardados = await ghtGuardarTodosLosMaestros({
      buffers: App.aprobadosBuffers,
      estado,
      mensaje: `Cierre ${periodo}`,
    });

    estadoUi("cierreStatus", `Mes ${periodo} cerrado. Guardados: ${guardados.join(", ")}.`, "ok");
    App.aprobadosBuffers = {};
    App.altaBuffers = {};
    App.resultados = {};
    App.cuentasExport = {};
    App.validaciones = {};
    await revisarMaestrosExistentes();
  } catch (e) {
    estadoUi("cierreStatus", "No pude cerrar el mes: " + e.message, "bad");
  } finally {
    mostrar("spinnerCierre", false);
  }
}

// ------------------------------------------------------------ confirmar e historial

// "Confirmar informes": archiva en GitHub los 4 informes de este período, con su tipo de
// cambio y su fecha, y los deja descargables desde el historial.
//
// Se archiva el archivo REVISADO cuando ya se subió en el paso de aprobación, y si no, el
// borrador que generó la app. La diferencia importa: el borrador trae las fórmulas pero
// todavía con los números en caché del mes anterior — Excel los recalcula al abrirlo, así
// que el archivo sirve igual, pero no es el que alguien revisó. El historial dice cuál es.
//
// Esto NO pisa los maestros: eso lo sigue haciendo "Guardar y cerrar el mes", que antes
// exige que los 4 pasen los controles. Confirmar es guardar una copia del mes; cerrar es
// mover la base de la que parte el mes que viene.
async function confirmarInformes() {
  const btn = document.getElementById("btnConfirmarInformes");
  const periodo = periodoElegido();
  if (!periodo) {
    estadoUi("confirmarStatus", "Falta el período: sin él no sé bajo qué nombre archivarlos.", "bad");
    return;
  }
  const hay = ARCHIVOS_TFBR.filter(a => App.aprobadosBuffers[a.id] || App.resultados[a.id]);
  if (!hay.length) {
    estadoUi("confirmarStatus", "No hay informes generados para confirmar.", "bad");
    return;
  }

  btn.disabled = true;
  mostrar("spinnerConfirmar", true);
  try {
    const mensaje = `Informes ${periodo}`;
    const archivos = {};
    for (const a of hay) {
      const revisado = App.aprobadosBuffers[a.id];
      const buffer = revisado || App.resultados[a.id].workbookBuffer;
      const g = await ghtGuardarInforme(periodo, a.id, buffer, mensaje);
      archivos[a.id] = { ...g, label: a.label, origen: revisado ? "revisado" : "borrador" };
      log(`Informe archivado: ${g.ruta}`);
    }

    const previo = await ghtLeerEstado();
    const estado = (previo && previo.estado) || { historial: [] };
    estado.historial = estado.historial || [];
    const entrada = {
      periodo,
      fecha: new Date().toISOString(),
      tcCierre: document.getElementById("tcCierreInput").value.trim(),
      difCambioMes: App.difCambioDelMes === null || App.difCambioDelMes === undefined
        ? "" : App.difCambioDelMes.toFixed(2),
      archivos,
      resumen: Object.fromEntries(
        ARCHIVOS_TFBR.map(a => [a.id, App.resultados[a.id] ? App.resultados[a.id].resumen : null])
      ),
    };
    // Confirmar dos veces el mismo período actualiza la entrada en vez de duplicarla: pasa
    // cada vez que se corrige algo y se vuelve a confirmar.
    const i = estado.historial.findIndex(h => h.periodo === periodo);
    if (i >= 0) estado.historial[i] = entrada; else estado.historial.push(entrada);
    estado.periodoActual = periodo;
    await ghtGuardarEstado(estado, mensaje);

    const cuantosRevisados = Object.values(archivos).filter(x => x.origen === "revisado").length;
    estadoUi("confirmarStatus",
      `${hay.length} informe(s) de ${periodo} guardados en el historial` +
      (cuantosRevisados === hay.length ? " (los revisados)."
        : cuantosRevisados ? ` (${cuantosRevisados} revisado(s), el resto borradores).`
        : " (borradores: todavía no subiste los revisados)."), "ok");
  } catch (e) {
    estadoUi("confirmarStatus", "No pude guardarlos: " + e.message, "bad");
    log("ERROR: " + e.message);
  } finally {
    mostrar("spinnerConfirmar", false);
    btn.disabled = false;
  }
}

async function mostrarHistorial() {
  const cont = document.getElementById("historialLista");
  cont.innerHTML = "<p class='footer-note'>Cargando…</p>";
  mostrar("ovHistorial", true);
  document.body.classList.add("sin-scroll");
  try {
    const r = await ghtLeerEstado();
    const historial = (r && r.estado && r.estado.historial) || [];
    if (!historial.length) {
      cont.innerHTML = "<p class='footer-note'>Todavía no hay ningún período confirmado.</p>";
      return;
    }
    cont.innerHTML = historial.slice().reverse().map(pintarEntradaHistorial).join("");
  } catch (e) {
    cont.innerHTML = `<p class="footer-note">No pude leer el historial: ${e.message}</p>`;
  }
}

function pintarEntradaHistorial(h) {
  const fecha = h.fecha ? new Date(h.fecha).toLocaleString("es-AR") : "—";
  const datos = [`Confirmado el ${fecha}`];
  if (h.tcCierre) datos.push(`TC de cierre <b>${h.tcCierre}</b>`);
  if (h.difCambioMes) datos.push(`Dif. de cambio <b>${h.difCambioMes}</b>`);

  // Las entradas viejas se guardaron antes de que se archivaran los archivos: se muestran
  // igual, con sus datos, y se avisa por qué no tienen nada para bajar.
  const archivos = h.archivos || {};
  const ids = ARCHIVOS_TFBR.map(a => a.id).filter(id => archivos[id]);
  const descargas = ids.length
    ? `<div class="hist-descargas">` + ids.map(id => {
        const f = archivos[id];
        const ruta = String(f.ruta).replace(/'/g, "&#39;");
        const nom = `${f.label || id} - ${h.periodo}.xlsx`.replace(/'/g, "&#39;");
        return `<button class="cfg-btn" onclick="descargarDelHistorial('${ruta}', '${nom}', this)">` +
               `${f.label || id}${f.origen === "borrador" ? " (borrador)" : ""}</button>`;
      }).join("") + `</div>`
    : `<div class="footer-note">Este período se cerró antes de que se archivaran los ` +
      `informes, así que no hay archivos para descargar.</div>`;

  return `<div class="hist-item">
      <div class="hist-cab"><b>${h.periodo}</b><span class="hist-datos">${datos.join(" · ")}</span></div>
      ${descargas}
    </div>`;
}

async function descargarDelHistorial(ruta, nombre, btn) {
  const antes = btn.textContent;
  btn.disabled = true;
  btn.textContent = "Bajando…";
  try {
    const buffer = await ghtLeerInforme(ruta);
    if (!buffer) throw new Error("ya no está en GitHub.");
    bajarComoArchivo(buffer, nombre);
  } catch (e) {
    estadoUi("historialStatus", `No pude bajar ${nombre}: ${e.message}`, "bad");
  } finally {
    btn.disabled = false;
    btn.textContent = antes;
  }
}

function cerrarHistorial() {
  mostrar("ovHistorial", false);
  soltarScrollSiNoQuedaVentana();
}

function cfgFondoHist(ev) {
  if (ev.target && ev.target.id === "ovHistorial") cerrarHistorial();
}

document.addEventListener("keydown", (ev) => {
  if (ev.key !== "Escape") return;
  const ov = document.getElementById("ovHistorial");
  if (ov && !ov.classList.contains("hidden")) cerrarHistorial();
});
