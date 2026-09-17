// Que los maestros conserven la tipografía y los recuadros del informe original.
//
// Ya pasó una vez y no se vio hasta que el contador abrió el archivo: los maestros se armaron
// con una conversión que se quedó con los valores, las fórmulas y el formato de número pero
// tiró todo lo demás. Las 93 celdas en negrita y las 816 con borde del Mensual $ quedaron en
// cero y el libro entero salió en Calibri 12, donde el informe de verdad usa Arial 8 y Bookman
// Old Style. Los números estaban bien; el informe no se parecía al informe.
//
// Es el error que ninguna otra prueba agarra: no hay fórmula mal ni cuenta perdida, así que
// todo lo demás pasa en verde. Por eso este archivo mira solamente la ropa.
//
// El piso no son números inventados: son los que tiene cada informe original de julio 2026,
// que están en INFORMES BASE. Si un maestro cambia a propósito y mueve estos números, hay que
// actualizarlos acá a mano —a propósito, para que el cambio quede a la vista.
//
// Correr con: node informe-tfbr/test_estilo.js

const path = require("path");
const fs = require("fs");
const XLSX = require("xlsx");
const ExcelJS = require("exceljs");

global.XLSX = XLSX;
const P = require("./parser_tfbr.js");
const motor = require("./motor_tfbr.js");

const PERIODO = "2026-07";

// SALDOS no entra: es la hoja de trabajo, no se imprime, y está corrida de fila respecto del
// original (el maestro tiene arreglos que el original no tiene).
const NO_SE_IMPRIME = new Set(["SALDOS"]);

const MAESTROS = [
  { label: "Mensual $",    archivo: "base_bm_ars.xlsx", negritas: 93, bordes: 816,
    periodo: "mensual",   campo: "saldo_ars", id: "balance_mensual_ars" },
  { label: "Mensual R$",   archivo: "base_bm_brl.xlsx", negritas: 93, bordes: 810,
    periodo: "mensual",   campo: "saldo_brl", id: "balance_mensual_brl" },
  { label: "Acumulado $",  archivo: "base_ba_ars.xlsx", negritas: 96, bordes: 591,
    periodo: "acumulado", campo: "saldo_ars", id: "balance_acumulado_ars" },
  { label: "Acumulado R$", archivo: "base_ba_brl.xlsx", negritas: 96, bordes: 589,
    periodo: "acumulado", campo: "saldo_brl", id: "balance_acumulado_brl" },
];

function leerExport(periodo) {
  const buf = fs.readFileSync(path.join(__dirname, "..", "inputs", PERIODO, `sumas_y_saldos_${periodo}.xls`));
  const libro = XLSX.read(buf, { type: "buffer" });
  const ws = libro.Sheets[libro.SheetNames[0]];
  return P.parseSumasYSaldosTFBR(XLSX.utils.sheet_to_json(ws, { header: 1 }), ws["!merges"]).cuentas;
}

// Cuenta la ropa de las hojas que se imprimen.
function mirarFormato(wb) {
  let negritas = 0, bordes = 0, conFuente = 0;
  const sinFuente = [], calibri = [];
  for (const ws of wb.worksheets) {
    if (NO_SE_IMPRIME.has(ws.name)) continue;
    ws.eachRow({ includeEmpty: false }, (row, r) => {
      row.eachCell({ includeEmpty: false }, (cell, c) => {
        const f = cell.font;
        if (f && f.name) {
          conFuente++;
          if (f.name === FUENTE_DE_CONVERSION && calibri.length < 3) calibri.push(`${ws.name}!${ctLetra(c)}${r}`);
        } else if (sinFuente.length < 3) sinFuente.push(`${ws.name}!${ctLetra(c)}${r}`);
        if (f && f.bold) negritas++;
        const b = cell.border || {};
        if (["top", "bottom", "left", "right"].some(s => b[s] && b[s].style)) bordes++;
      });
    });
  }
  return { negritas, bordes, conFuente, sinFuente, calibri };
}

// La que pone Excel cuando una conversión perdió la fuente. Si aparece en una hoja que se
// imprime, algo la despintó.
const FUENTE_DE_CONVERSION = "Calibri";

(async () => {
  let fallas = 0;
  const fallo = (m) => { console.log(`   ✗ ${m}`); fallas++; };

  for (const m of MAESTROS) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path.join(__dirname, m.archivo));

    const f = mirarFormato(wb);
    const { negritas, bordes, conFuente, calibri } = f;

    console.log(`
== ${m.label}: ${negritas} negritas · ${bordes} con recuadro · ` +
                `${conFuente} celdas con tipografía propia`);

    if (calibri.length) {
      fallo(`${m.label}: hay celdas en ${FUENTE_DE_CONVERSION} en hojas que se imprimen ` +
            `(${calibri.join(", ")}) — el maestro perdió la tipografía del original`);
    }
    if (negritas < m.negritas) {
      fallo(`${m.label}: quedaron ${negritas} celdas en negrita y el original tiene ${m.negritas}`);
    }
    if (bordes < m.bordes) {
      fallo(`${m.label}: quedaron ${bordes} celdas con recuadro y el original tiene ${m.bordes}`);
    }
  }

  // Y que la corrida no despinte nada. Las filas que agrega el motor tienen que salir con la
  // tipografia de sus vecinas: una celda sin fuente la dibuja Excel con la del libro —Calibri
  // 11— y en medio de un cuadro en Arial 8 se ve de lejos.
  console.log("\n--- despues de procesar (export de julio 2026)");
  const exports_ = { mensual: leerExport("mensual"), acumulado: leerExport("acumulado") };

  for (const m of MAESTROS) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path.join(__dirname, m.archivo));
    motor.procesarMaestroTFBR({
      wb, cuentasExport: exports_[m.periodo], campoSaldo: m.campo, archivoId: m.id,
      cuentasAcumulado: exports_.acumulado, log: () => {},
    });
    const f = mirarFormato(wb);
    console.log(`   ${m.label}: ${f.negritas} negritas · ${f.bordes} con recuadro`);
    if (f.sinFuente.length) {
      fallo(`${m.label}: despues de procesar quedaron celdas sin tipografia ` +
            `(${f.sinFuente.join(", ")}) — el motor las creo y no les copio el formato`);
    }
    if (f.calibri.length) {
      fallo(`${m.label}: despues de procesar hay celdas en ${FUENTE_DE_CONVERSION} (${f.calibri.join(", ")})`);
    }
    if (f.negritas < m.negritas) {
      fallo(`${m.label}: despues de procesar quedaron ${f.negritas} negritas y habia ${m.negritas}`);
    }
    if (f.bordes < m.bordes) {
      fallo(`${m.label}: despues de procesar quedaron ${f.bordes} recuadros y habia ${m.bordes}`);
    }
  }

  console.log(fallas ? `\n✗ ${fallas} falla(s).` : "\n✓ Los 4 maestros conservan la estética del original.");
  process.exit(fallas ? 1 : 0);
})();

function ctLetra(n) {
  let s = "";
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = (n - 1 - r) / 26; }
  return s;
}
