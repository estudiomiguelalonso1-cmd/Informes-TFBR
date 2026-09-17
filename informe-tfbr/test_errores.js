// Ningún informe se entrega con un #REF! adentro.
//
// Los cuatro maestros arrastran fórmulas rotas de cierres viejos: "SUM(#REF!)" en una columna
// lateral del Pasivo y del Anexo II, "+#REF!+K13" en el EEPN de los acumulados y "+#REF!" en el
// Anexo I del Mensual R$, que además rompía el total de su columna.
//
// No las lee nadie, así que no mueven ningún número, pero se imprimen. Un informe contable con
// un #REF! a la vista no se puede entregar, y ese es todo el motivo de este test.
//
// También cuida lo contrario: que no se vacíe una celda rota que alguien SÍ lee. Ahí el número
// correcto no lo sabe el programa y vaciarla cambiaría el informe en silencio.
//
// Correr con: node informe-tfbr/test_errores.js

const path = require("path");
const fs = require("fs");
const XLSX = require("xlsx");
const ExcelJS = require("exceljs");

global.XLSX = XLSX;
const P = require("./parser_tfbr.js");
const motor = require("./motor_tfbr.js");
const le = require("./limpiar_errores.js");

const PERIODO = "2026-08";

const MAESTROS = [
  { label: "Mensual $",    archivo: "base_bm_ars.xlsx", periodo: "mensual",   campo: "saldo_ars", id: "balance_mensual_ars" },
  { label: "Mensual R$",   archivo: "base_bm_brl.xlsx", periodo: "mensual",   campo: "saldo_brl", id: "balance_mensual_brl" },
  { label: "Acumulado $",  archivo: "base_ba_ars.xlsx", periodo: "acumulado", campo: "saldo_ars", id: "balance_acumulado_ars" },
  { label: "Acumulado R$", archivo: "base_ba_brl.xlsx", periodo: "acumulado", campo: "saldo_brl", id: "balance_acumulado_brl" },
];

function leerExport(cual) {
  const f = path.join(__dirname, "..", "inputs", PERIODO, `sumas_y_saldos_${cual}_completo.xls`);
  const libro = XLSX.read(fs.readFileSync(f), { type: "buffer" });
  const ws = libro.Sheets[libro.SheetNames[0]];
  return P.parseSumasYSaldosTFBR(XLSX.utils.sheet_to_json(ws, { header: 1 }), ws["!merges"]).cuentas;
}

function rotas(wb) {
  const o = [];
  for (const ws of wb.worksheets) {
    ws.eachRow({ includeEmpty: false }, (row) => row.eachCell({ includeEmpty: false }, (cell) => {
      if (le.leRotaTexto(cell)) o.push(`${ws.name}!${cell.address} = ${cell.value.formula}`);
    }));
  }
  return o;
}

(async () => {
  let fallas = 0;
  const fallo = (m) => { console.log(`   ✗ ${m}`); fallas++; };

  // Una celda rota que alguien lee no se toca.
  {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Hoja");
    ws.getCell("A1").value = { formula: "SUM(#REF!)" };
    ws.getCell("B1").value = { formula: "+A1" };
    ws.getCell("A2").value = { formula: "SUM(#REF!)" };   // a ésta no la lee nadie
    const r = le.limpiarErrores(wb, () => {});
    if (r.vaciadas.length !== 1 || r.vaciadas[0].dir !== "A2") {
      fallo(`tendría que vaciar sólo A2 y vació ${r.vaciadas.map(x => x.dir).join(", ") || "nada"}`);
    }
    if (r.trabadas.length !== 1 || r.trabadas[0].dir !== "A1") {
      fallo("tendría que avisar que A1 está rota y la lee B1");
    }
    console.log("\n== la regla: una celda rota que alguien lee no se vacía, se avisa");
  }

  const exports_ = { mensual: leerExport("mensual"), acumulado: leerExport("acumulado") };
  for (const m of MAESTROS) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path.join(__dirname, m.archivo));
    const antes = rotas(wb);
    motor.procesarMaestroTFBR({
      wb, cuentasExport: exports_[m.periodo], campoSaldo: m.campo, archivoId: m.id,
      cuentasAcumulado: exports_.acumulado, log: () => {},
    });
    const despues = rotas(wb);
    console.log(`   ${m.label.padEnd(13)} ${antes.length} celda(s) con #REF! en el maestro → ${despues.length} después de procesar`);
    for (const x of despues) fallo(`${m.label}: quedó ${x}`);
  }

  console.log(fallas ? `\n✗ ${fallas} falla(s).` : "\n✓ Ningún informe queda con un #REF! adentro.");
  process.exit(fallas ? 1 : 0);
})();
