// Prueba de la expansión de rangos, contra los 4 maestros de verdad.
//
// Lo que hay que demostrar es que el número NO cambia. Un rango y la suma explícita de sus
// filas dan lo mismo sólo si se expande exactamente el mismo conjunto de filas — ni una de
// más ni una de menos. Comparar los importes en caché no sirve: los maestros guardados traen
// el VLOOKUP sin resultado calculado, así que todo daría cero y la prueba pasaría sola.
//
// Entonces se compara el CONJUNTO DE FILAS que cada fórmula lee, antes y después. Si es el
// mismo conjunto, la suma es la misma, con cualquier dato que traiga el mes.
//
// Correr con: node informe-tfbr/test_rastreo.js

const path = require("path");
const ExcelJS = require("exceljs");

global.materializarFormulasCompartidas = require("./formula_utils.js").materializarFormulasCompartidas;
global.limpiarPlanDeCuentas = require("./limpieza_plan.js").limpiarPlanDeCuentas;
const cfg = require("./config_tfbr.js");
const rr = require("./rastreo_rangos.js");

const MAESTROS = {
  "Mensual $": "base_bm_ars.xlsx",
  "Mensual R$": "base_bm_brl.xlsx",
  "Acumulado $": "base_ba_ars.xlsx",
  "Acumulado R$": "base_ba_brl.xlsx",
};

// Las celdas de SALDOS que lee una fórmula, como conjunto de "COL<fila>". Los rangos se
// expanden acá también, para poder comparar un rango contra su versión escrita término a
// término.
function celdasQueLee(formula) {
  const celdas = new Set();
  let f = formula;
  for (const m of f.matchAll(/SALDOS!\$?([A-Z]{1,3})\$?(\d+)\s*:\s*\$?([A-Z]{1,3})\$?(\d+)/g)) {
    const col = m[1].toUpperCase();
    for (let r = Math.min(+m[2], +m[4]); r <= Math.max(+m[2], +m[4]); r++) celdas.add(`${col}${r}`);
  }
  f = f.replace(/SALDOS!\$?[A-Z]{1,3}\$?\d+\s*:\s*\$?[A-Z]{1,3}\$?\d+/g, "");
  for (const m of f.matchAll(/SALDOS!\$?([A-Z]{1,3})\$?(\d+)/g)) {
    celdas.add(`${m[1].toUpperCase()}${m[2]}`);
  }
  return celdas;
}

// Toda fórmula del libro que toque SALDOS, por dirección.
function formulasQueTocanSaldos(wb, hojaSaldos) {
  const mapa = {};
  for (const ws of wb.worksheets) {
    if (ws.name === hojaSaldos) continue;
    ws.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const v = cell.value;
        if (!v || typeof v !== "object" || typeof v.formula !== "string") return;
        if (!/SALDOS!/.test(v.formula)) return;
        mapa[`${ws.name}!${cell.address}`] = v.formula;
      });
    });
  }
  return mapa;
}

const mismoConjunto = (a, b) => a.size === b.size && [...a].every(x => b.has(x));

(async () => {
  let fallas = 0;
  const fallo = (msg) => { console.log(`   ✗ ${msg}`); fallas++; };

  for (const [nombre, archivo] of Object.entries(MAESTROS)) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path.join(__dirname, archivo));
    materializarFormulasCompartidas(wb);
    limpiarPlanDeCuentas(wb, () => {});
    const layout = cfg.derivarLayoutSaldos(wb);

    const antes = formulasQueTocanSaldos(wb, layout.sheet);
    const conRango = Object.entries(antes).filter(([, f]) => rr.RR_RE_RANGO.test(f));

    const res = rr.expandirRangosSaldos(wb, layout, () => {});
    const despues = formulasQueTocanSaldos(wb, layout.sheet);

    console.log(`\n== ${nombre}: ${conRango.length} fórmula(s) con rango, ` +
                `${res.expandidos.length} rango(s) expandido(s), ${res.salteados.length} salteado(s)`);

    // 1. Ninguna fórmula cambia de qué celdas lee.
    for (const [donde, formulaAntes] of Object.entries(antes)) {
      const a = celdasQueLee(formulaAntes);
      const b = celdasQueLee(despues[donde] || "");
      if (!mismoConjunto(a, b)) {
        const sobran = [...b].filter(x => !a.has(x)), faltan = [...a].filter(x => !b.has(x));
        fallo(`${donde} cambió lo que lee — sobran ${sobran.join(",") || "-"}, faltan ${faltan.join(",") || "-"}`);
      }
    }

    // 2. No queda ningún rango sin expandir (salvo los que se saltearon a propósito).
    const quedan = Object.entries(despues).filter(([, f]) => rr.RR_RE_RANGO.test(f));
    if (quedan.length !== res.salteados.length) {
      fallo(`quedaron ${quedan.length} fórmula(s) con rango y sólo ${res.salteados.length} fueron ` +
            `salteadas a propósito: ${quedan.map(([d]) => d).join(", ")}`);
    }

    // 3. Correrlo de nuevo no cambia nada: el motor lo aplica en cada corrida.
    const segunda = rr.expandirRangosSaldos(wb, layout, () => {});
    if (segunda.expandidos.length) {
      fallo(`la segunda pasada volvió a expandir ${segunda.expandidos.length} rango(s)`);
    }

    for (const e of res.expandidos) console.log(`   ${e.donde}: ${e.rango} → ${e.cuentas} cuentas`);
    for (const s of res.salteados) console.log(`   ⚠ salteado ${s.donde}: ${s.motivo}`);
  }

  console.log(fallas ? `\n✗ ${fallas} falla(s).` : "\n✓ Las 4 pasan: nadie cambió lo que lee.");
  process.exit(fallas ? 1 : 0);
})();
