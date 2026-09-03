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
  maestrosCargados: {}, // archivoId -> { wb: ExcelJS.Workbook, sha }  (para la corrida del mes)
  cuentasExport: {},    // "mensual"|"acumulado" -> cuentas parseadas
  resultados: {},       // archivoId -> { resumen, workbookBuffer }
  aprobadosBuffers: {}, // archivoId -> ArrayBuffer (subido en la revisión final)
  validaciones: {},     // archivoId -> resultado de validarRecalculado
  logLineas: [],
};

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

function estadoUi(elId, texto, clase) {
  const el = document.getElementById(elId);
  if (!el) return;
  el.innerHTML = `<div class="status-msg ${clase || ""}">${texto}</div>`;
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
  await revisarMaestrosExistentes();
}

function abrirConfig() { mostrar("cardConfig", true); }
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
    await revisarMaestrosExistentes();
  } catch (e) {
    estadoUi("configStatus", "No pude conectar: " + e.message, "bad");
  } finally {
    mostrar("spinnerConfig", false);
  }
}

// ------------------------------------------------------------ primera vez (alta de maestros)

async function revisarMaestrosExistentes() {
  const faltantes = [];
  for (const a of ARCHIVOS_TFBR) {
    const r = await ghtLeerMaestro(a.id);
    if (!r) { faltantes.push(a); continue; }
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(r.buffer);
    App.maestrosCargados[a.id] = { wb, sha: r.sha };
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
  const periodo = document.getElementById("periodoInput").value.trim();
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
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, { type: "array", cellFormula: true });
    const ws = wb.Sheets["Sheet1"];
    if (!ws) throw new Error("el archivo no tiene una hoja 'Sheet1'.");
    const filas = XLSX.utils.sheet_to_json(ws, { header: 1 });
    const parsed = parseSumasYSaldosTFBR(filas, ws["!merges"]);
    App.cuentasExport[periodo] = parsed.cuentas;
    txt.textContent =
      `${file.name} ✓ — ${parsed.cuentas.length} cuentas, total $ ${parsed.totales.saldo_ars.toFixed(2)}`;
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
  document.getElementById("periodoInput")?.addEventListener("input", revisarListoParaProcesar);
  document.getElementById("tcCierreInput")?.addEventListener("input", revisarListoParaProcesar);
  iniciar();
});

async function procesarPeriodo() {
  mostrar("spinnerProcesar", true);
  document.getElementById("btnProcesar").disabled = true;
  App.logLineas = [];
  try {
    for (const a of ARCHIVOS_TFBR) {
      log(`\n=== ${a.label} ===`);
      const cargado = App.maestrosCargados[a.id];
      if (!cargado) throw new Error(`Falta el maestro de ${a.label}.`);

      // se trabaja sobre una copia en memoria: si algo falla más adelante, el maestro
      // guardado en GitHub no se tocó
      const buffer = await cargado.wb.xlsx.writeBuffer();
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buffer);

      // Los errores que el archivo YA tenía antes de que lo tocáramos: son la línea de base
      // contra la que después se comparan los del archivo aprobado, para distinguir un
      // problema nuevo de uno viejo sin depender de una lista escrita a mano.
      const erroresPrevios = buscarCeldasEnError(wb);

      const cuentasExport = App.cuentasExport[a.periodo];
      const { resumen, planDeCuentas, escritas } =
        procesarMaestroTFBR({ wb, cuentasExport, campoSaldo: a.campoSaldo, log });

      // el TC de cierre y la cifra de dif de cambio solo existen en uno de los 4 archivos:
      // escribirDatosDelPeriodo se fija solo si este los tiene, y avisa lo que no pudo cargar
      const periodoDatos = escribirDatosDelPeriodo(wb, {
        periodo: document.getElementById("periodoInput").value.trim(),
        tcCierre: document.getElementById("tcCierreInput").value.trim(),
        difCambioMes: document.getElementById("difCambioInput").value.trim(),
      }, log);

      aplicarFixesAprobados(wb, a.id, log);

      const outBuffer = await wb.xlsx.writeBuffer();
      App.resultados[a.id] = {
        resumen, periodoDatos, planDeCuentas, escritas, erroresPrevios,
        workbookBuffer: outBuffer,
      };
    }
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

// ------------------------------------------------------------ resultado / descargas

function pintarResultado() {
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

function descargarBorrador(a) {
  const r = App.resultados[a.id];
  if (!r) return;
  const blob = new Blob([r.workbookBuffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const el = document.createElement("a");
  el.href = url;
  el.download = `${a.label} - BORRADOR.xlsx`;
  el.click();
  URL.revokeObjectURL(url);
}

function irARevision() {
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
    const periodo = document.getElementById("periodoInput").value.trim();
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

// ------------------------------------------------------------ historial

async function mostrarHistorial() {
  const cont = document.getElementById("historialLista");
  cont.innerHTML = "Cargando…";
  mostrar("cardHistorial", true);
  try {
    const r = await ghtLeerEstado();
    if (!r || !r.estado.historial || !r.estado.historial.length) {
      cont.innerHTML = "<p class='footer-note'>Todavía no hay ningún mes cerrado.</p>";
      return;
    }
    cont.innerHTML = r.estado.historial
      .slice()
      .reverse()
      .map(h => `<div style="padding:8px 0; border-bottom:1px solid var(--borde);">
          <b>${h.periodo}</b> — cerrado el ${new Date(h.fecha).toLocaleString("es-AR")}
          (TC ${h.tcCierre || "?"})
        </div>`)
      .join("");
  } catch (e) {
    cont.innerHTML = `<p class="footer-note">No pude leer el historial: ${e.message}</p>`;
  }
}
