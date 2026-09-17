// Anexo I — Bienes de uso: el valor de origen sale de las cuentas, en los cuatro archivos.
//
// Qué pasaba. De los cuatro, sólo el Acumulado R$ tenía el valor de origen cableado contra
// SALDOS. Los otros tres lo traían escrito a mano —y el Mensual R$ ni eso: estaba TODO en
// cero, así que su "Bienes de uso" del EESP daba cero menos la amortización del mes, en
// negativo. Un número tipeado no se actualiza nunca: el día que cambie el valor de un bien,
// el anexo sigue mostrando el del año pasado y nada avisa.
//
// El mapeo NO se inventa: es el que el Acumulado R$ ya tenía, confirmado por contaduría.
// Dos renglones agrupan dos cuentas cada uno, y la suma da exacto contra lo que hoy está
// escrito a mano: Hardware = LX300 + Multifunción = 349.454,86; Instalaciones = Aire
// acondicionado + Calefón = 9.934.406,71.
//
// Fórmula o número, según el archivo. En los Acumulados el sumas y saldos trae las cuentas de
// bienes de uso, así que el renglón las lee y se actualiza solo. En los Mensuales NO las trae
// —el export mensual sólo tiene las cuentas que se movieron, y un rodado no se mueve— así que
// ahí el renglón leería cero. Tampoco se pueden pegar en la zona de pegado del mensual: sus
// saldos son acumulados (202 millones) y descuadrarían el Debe = Haber, que en el mensual
// cierra sobre el movimiento del mes. Por eso en los Mensuales el importe se escribe como
// número, tomado del mismo export acumulado que ya se carga. El valor es el mismo; lo que
// cambia es de dónde lo saca el archivo.

const ANEXO_I_RENGLONES = [
  { rotulo: "Software",         cuentas: ["1210100100"] },
  { rotulo: "Hardware",         cuentas: ["1210300100", "1210300200"] },
  { rotulo: "Computadoras",     cuentas: ["1210300300"] },
  { rotulo: "Equipos",          cuentas: ["1210400100"] },
  { rotulo: "Muebles y Utiles", cuentas: ["1210200300"] },
  { rotulo: "Instalaciones",    cuentas: ["1210500100", "1210500200"] },
  { rotulo: "Rodados",          cuentas: ["1210700100"] },
];

// Cuentas que el Anexo I lee y NO tendría que leer, porque ya las cuenta otra hoja.
//
// El Anexo I del Mensual $ tenía "I18 = +SALDOS!B159", que es 4212500000 AMORTIZACIONES — la
// misma cuenta que el Anexo II ya toma como gasto. Contada dos veces: una llega al Activo por
// el Anexo I y otra al resultado por el Anexo II, y el balance en pesos no cierra. Los otros
// tres archivos no la tienen ahí; era una diferencia de ese archivo, no un criterio.
//
// Se busca por la cuenta que lee, no por la dirección: el motor inserta filas y la celda se
// mueve. Si el archivo ya no la tiene, no hace nada.
const ANEXO_I_A_QUITAR = [
  { cuenta: "4212500000", porque: "el Anexo II ya la toma como gasto" },
];

// Saca del Anexo I las referencias que están de más. Devuelve qué sacó, para el log.
function quitarDuplicadosAnexoI(wb, layout, log = () => {}) {
  const ax = wb.getWorksheet("Anexo I");
  if (!ax) return [];
  const plan = leerPlanDeCuentas(wb, layout).cuentas;
  const sacados = [];

  for (const x of ANEXO_I_A_QUITAR) {
    const cuenta = plan[x.cuenta];
    if (!cuenta) continue;
    const filas = [cuenta].concat(cuenta.otrasFilas || []).map(f => f.fila);
    ax.eachRow({ includeEmpty: false }, (row, r) => {
      row.eachCell({ includeEmpty: false }, (cell, c) => {
        const v = cell.value;
        if (!v || typeof v !== "object" || typeof v.formula !== "string") return;
        for (const f of filas) {
          if (!rtQuitarTermino(ax, cell.address, f)) continue;
          sacados.push({ celda: cell.address, cuenta: x.cuenta, antes: v.formula });
          log(`  Anexo I ${cell.address}: le saqué ${x.cuenta} — ${x.porque}, y acá se contaba ` +
              `dos veces (era "${v.formula}").`);
        }
      });
    });
  }
  return sacados;
}

// El "Bienes de uso" del EESP tiene que salir del Anexo I, en los cuatro archivos.
//
// De los cuatro, el Mensual $ era el unico que no lo hacia: su EESP leia "+SALDOS!E47", que es
// un subtotal de amortizaciones, y mostraba -2.485.298,32 —un activo en negativo— mientras su
// propio Anexo I decia 82.447.582,24. Los otros tres ya leen "+'Anexo I'!K<total>".
//
// No se escribe la direccion: el motor inserta filas y las celdas se mueven. Se busca el
// renglon por su texto en el EESP y el total por la columna de neto resultante del anexo.

function aiColumnaNeto(ax, ubic) {
  // El neto resultante es la ultima columna con importe en los renglones del anexo.
  let ultima = null;
  for (let c = ubic.colConcepto + 1; c <= ubic.colConcepto + 16; c++) {
    for (const fila of Object.values(ubic.filas)) {
      const v = ax.getCell(fila, c).value;
      const esImporte = typeof v === "number" ||
        (v && typeof v === "object" && typeof v.formula === "string");
      if (esImporte) { ultima = c; break; }
    }
  }
  return ultima;
}

function aiFilaTotal(ax, ubic, colNeto) {
  const ultimoRenglon = Math.max(...Object.values(ubic.filas));
  for (let r = ultimoRenglon + 1; r <= ultimoRenglon + 6; r++) {
    const f = ctFormulaDe(ax, r, colNeto);
    if (f && /^SUM\(/i.test(f.trim())) return r;
  }
  return null;
}

// Devuelve lo que hizo, o null si no habia nada que cambiar.
function apuntarBienesDeUsoAlAnexo(wb, ubic, log = () => {}) {
  const ax = wb.getWorksheet("Anexo I");
  const es = wb.getWorksheet("EESP");
  if (!ax || !es || !ubic) return null;

  const colNeto = aiColumnaNeto(ax, ubic);
  if (colNeto === null) return null;
  const filaTotal = aiFilaTotal(ax, ubic, colNeto);
  if (filaTotal === null) {
    log("  ⚠ EESP: no encontre la fila de total del Anexo I; deje 'Bienes de uso' como estaba.");
    return null;
  }

  const destino = `+'Anexo I'!${ctColNumeroALetra(colNeto)}${filaTotal}`;

  for (let r = 1; r <= es.rowCount; r++) {
    for (let c = 1; c <= 8; c++) {
      // Por prefijo, no por igualdad: cada archivo lo rotula distinto —"Bienes de uso",
      // "Bienes de uso (Anexo I)" y, en el Mensual R$, "Bienes de uso (Anexo l )" con ele.
      if (!aiNorm(aiTexto(es, r, c)).startsWith("BIENES DE USO")) continue;
      // El importe es la primera celda a la derecha del rotulo que tenga formula o numero.
      for (let k = c + 1; k <= c + 8; k++) {
        const v = es.getCell(r, k).value;
        const esImporte = typeof v === "number" ||
          (v && typeof v === "object" && typeof v.formula === "string");
        if (!esImporte) continue;
        const antes = (v && typeof v === "object" && v.formula) ? "+" + v.formula.replace(/^\+/, "") : String(v);
        if (antes === destino) return null;
        es.getCell(r, k).value = { formula: destino };
        log(`  EESP ${ctColNumeroALetra(k)}${r} "Bienes de uso": ahora lee ${destino} ` +
            `(antes "${antes}").`);
        return { celda: ctColNumeroALetra(k) + r, antes, ahora: destino };
      }
    }
  }
  return null;
}

function aiNorm(t) {
  return String(t || "").normalize("NFD").replace(/[̀-ͯ]/g, "")
    .toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();
}

function aiTexto(ws, r, c) {
  const v = ws.getCell(r, c).value;
  if (v == null) return "";
  if (typeof v === "object") return v.richText ? v.richText.map(t => t.text).join("") : "";
  return String(v);
}

// Dónde está cada renglón del Anexo I y en qué columna va el valor de origen.
//
// Se deduce, no se escribe: los cuatro archivos tienen el anexo en filas distintas (el
// Mensual $ arranca en la 11 y el Acumulado $ en la 14). El concepto es la primera columna
// con el texto del renglón; el valor de origen, la primera columna de importe a su derecha.
function aiUbicar(ax) {
  let colConcepto = null, colOrigen = null;
  const filas = {};

  for (let r = 1; r <= ax.rowCount; r++) {
    for (let c = 1; c <= 4; c++) {
      const t = aiTexto(ax, r, c).trim();
      if (!t) continue;
      const match = ANEXO_I_RENGLONES.find(x => aiNorm(x.rotulo) === aiNorm(t));
      if (!match) continue;
      if (colConcepto === null) colConcepto = c;
      if (c !== colConcepto) continue;
      filas[aiNorm(t)] = r;
    }
  }
  if (colConcepto === null) return null;

  // La columna del valor de origen: la primera a la derecha del concepto que tenga número o
  // fórmula en alguno de los renglones encontrados.
  for (let c = colConcepto + 1; c <= colConcepto + 4 && colOrigen === null; c++) {
    for (const fila of Object.values(filas)) {
      const v = ax.getCell(fila, c).value;
      const esImporte = typeof v === "number" ||
        (v && typeof v === "object" && typeof v.formula === "string");
      if (esImporte) { colOrigen = c; break; }
    }
  }
  if (colOrigen === null) return null;
  return { colConcepto, colOrigen, filas };
}

// Completa el valor de origen. Devuelve lo que hizo, para el log y la pantalla.
function completarAnexoI(wb, layout, { escritas, cuentasAcumulado, campoSaldo }, log = () => {}) {
  const ax = wb.getWorksheet("Anexo I");
  if (!ax) return { hechos: [], salteados: [], modo: null };

  const duplicados = quitarDuplicadosAnexoI(wb, layout, log);

  const ubic = aiUbicar(ax);
  if (!ubic) {
    log("  ⚠ Anexo I: no ubiqué los renglones de bienes de uso; lo dejé como estaba.");
    return { hechos: [], salteados: [], modo: null };
  }

  const plan = leerPlanDeCuentas(wb, layout).cuentas;
  const todasLasCuentas = ANEXO_I_RENGLONES.flatMap(r => r.cuentas);

  // ¿Este archivo tiene los bienes de uso en su zona de pegado? Si sí, el renglón puede leer
  // las cuentas; si no, hay que escribir el número. No se decide por el nombre del archivo:
  // se mira qué se pegó en esta corrida.
  const pegadas = escritas || {};
  const conDato = todasLasCuentas.filter(c => pegadas[c] !== undefined).length;
  const porFormula = conDato === todasLasCuentas.length;

  const porCodigo = {};
  for (const c of (cuentasAcumulado || [])) porCodigo[c.codigo] = c;

  const hechos = [], salteados = [];
  for (const reng of ANEXO_I_RENGLONES) {
    const fila = ubic.filas[aiNorm(reng.rotulo)];
    if (!fila) { salteados.push({ ...reng, motivo: "el archivo no tiene ese renglón" }); continue; }

    if (porFormula) {
      const faltan = reng.cuentas.filter(c => !plan[c]);
      if (faltan.length) {
        salteados.push({ ...reng, motivo: `${faltan.join(", ")} no está(n) en el plan de SALDOS` });
        continue;
      }
      const terminos = reng.cuentas.map(c => ctFormulaNetaAnexo(layout, plan[c].fila)).join("");
      ax.getCell(fila, ubic.colOrigen).value = { formula: terminos.replace(/^\+/, "+") };
      hechos.push({ rotulo: reng.rotulo, fila, modo: "fórmula", cuentas: reng.cuentas });
    } else {
      const faltan = reng.cuentas.filter(c => !porCodigo[c]);
      if (faltan.length) {
        salteados.push({ ...reng, motivo: `${faltan.join(", ")} no vino(ieron) en el sumas y saldos acumulado` });
        continue;
      }
      const importe = reng.cuentas.reduce((s, c) => s + (porCodigo[c][campoSaldo] || 0), 0);
      ax.getCell(fila, ubic.colOrigen).value = importe;
      hechos.push({ rotulo: reng.rotulo, fila, modo: "valor", cuentas: reng.cuentas, importe });
    }
  }

  const bienesDeUso = apuntarBienesDeUsoAlAnexo(wb, ubic, log);

  const modo = porFormula ? "fórmula" : "valor";
  if (hechos.length) {
    log(`  Anexo I: ${hechos.length} renglón(es) de bienes de uso cargados por ${modo} ` +
        `(valor de origen, columna ${ctColNumeroALetra(ubic.colOrigen)}).`);
  }
  for (const s of salteados) log(`  ⚠ Anexo I "${s.rotulo}": ${s.motivo}.`);
  return { hechos, salteados, modo, duplicados, bienesDeUso };
}

if (typeof module !== "undefined") {
  const cfg = require("./config_tfbr.js");
  global.leerPlanDeCuentas = cfg.leerPlanDeCuentas;
  global.ctFormulaNetaAnexo = cfg.ctFormulaNetaAnexo;
  global.ctColNumeroALetra = cfg.ctColNumeroALetra;
  global.ctFormulaDe = cfg.ctFormulaDe;
  global.rtQuitarTermino = require("./rotulos_anexo.js").rtQuitarTermino;
  module.exports = {
    ANEXO_I_RENGLONES, ANEXO_I_A_QUITAR, aiUbicar, completarAnexoI, quitarDuplicadosAnexoI,
    apuntarBienesDeUsoAlAnexo, aiColumnaNeto, aiFilaTotal,
  };
}
