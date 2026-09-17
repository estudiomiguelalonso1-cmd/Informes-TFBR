// Ningún renglón puede leer media cuenta.
//
// En SALDOS cada cuenta ocupa dos columnas, la deudora y la acreedora, y su saldo es la resta
// de las dos. Un renglón que nombra una sola de las dos funciona mientras la cuenta caiga
// siempre del mismo lado, y el día que se da vuelta pierde el importe entero sin avisar.
//
// Pasó en agosto de 2026: "1140900900 SEGUROS A DEV. FEDERACIÓN" venía acreedora todos los
// meses, el renglón "Seguros a devengar" del Activo la leía como "-SALDOS!C28", y cuando quedó
// deudora por 866.748,53 ese renglón dio cero. El balance del Mensual $ descuadró 881.272,63,
// de los cuales 866.748,53 eran esta cuenta.
//
// Los cuatro maestros venían con 74 renglones así. Este test corre el motor y comprueba que
// después no queda ninguno.
//
// Correr con: node informe-tfbr/test_media_columna.js

const path = require("path");
const fs = require("fs");
const XLSX = require("xlsx");
const ExcelJS = require("exceljs");

global.XLSX = XLSX;
const P = require("./parser_tfbr.js");
const motor = require("./motor_tfbr.js");
const cfg = require("./config_tfbr.js");
const mc = require("./media_columna.js");

const PERIODO = "2026-07";
const REF = /SALDOS![A-Z$]+([0-9]+)/g;

const MAESTROS = [
  { label: "Mensual $",    archivo: "base_bm_ars.xlsx", periodo: "mensual",   campo: "saldo_ars", id: "balance_mensual_ars" },
  { label: "Mensual R$",   archivo: "base_bm_brl.xlsx", periodo: "mensual",   campo: "saldo_brl", id: "balance_mensual_brl" },
  { label: "Acumulado $",  archivo: "base_ba_ars.xlsx", periodo: "acumulado", campo: "saldo_ars", id: "balance_acumulado_ars" },
  { label: "Acumulado R$", archivo: "base_ba_brl.xlsx", periodo: "acumulado", campo: "saldo_brl", id: "balance_acumulado_brl" },
];

function leerExport(periodo) {
  const buf = fs.readFileSync(path.join(__dirname, "..", "inputs", PERIODO, `sumas_y_saldos_${periodo}.xls`));
  const libro = XLSX.read(buf, { type: "buffer" });
  const ws = libro.Sheets[libro.SheetNames[0]];
  return P.parseSumasYSaldosTFBR(XLSX.utils.sheet_to_json(ws, { header: 1 }), ws["!merges"]).cuentas;
}

// Cuentas del plan que alguna hoja lee por una sola de sus dos columnas.
function aMedias(wb) {
  const s = wb.getWorksheet("SALDOS");
  const L = cfg.derivarLayoutSaldos(wb);
  const cd = cfg.ctColNumeroALetra(L.deudorCol), ca = cfg.ctColNumeroALetra(L.acreedorCol);
  const filas = {};
  for (let r = L.planDeCuentas.desde; r <= L.planDeCuentas.hasta; r++) {
    let t = s.getCell(r, L.keyCol).value;
    if (t && typeof t === "object") t = t.richText ? t.richText.map(x => x.text).join("") : "";
    const txt = String(t || "").trim();
    if (!/^\d{10}\b/.test(txt)) continue;
    filas[r] = { nom: txt, cols: new Set(), donde: [] };
  }
  for (const ws of wb.worksheets) {
    if (ws.name === "SALDOS") continue;
    ws.eachRow({ includeEmpty: false }, (row) => row.eachCell({ includeEmpty: false }, (cell) => {
      const v = cell.value;
      if (!v || typeof v !== "object" || typeof v.formula !== "string") return;
      let m; REF.lastIndex = 0;
      while ((m = REF.exec(v.formula)) !== null) {
        const f = filas[Number(m[1])];
        if (!f) continue;
        f.cols.add(v.formula.slice(m.index).match(/SALDOS!\$?([A-Z]+)/)[1]);
        f.donde.push(ws.name + "!" + cell.address);
      }
    }));
  }
  const malas = [];
  for (const r of Object.keys(filas)) {
    const f = filas[r];
    if (!f.donde.length) continue;
    if (f.cols.has(cd) !== f.cols.has(ca)) {
      malas.push(`${f.nom.slice(0, 30)} ← ${[...new Set(f.donde)].join(", ")}`);
    }
  }
  return malas;
}

(async () => {
  let fallas = 0;
  const fallo = (m) => { console.log(`   ✗ ${m}`); fallas++; };

  // Primero, que la regla sea la que decimos: se agrega la columna que falta con el signo
  // cambiado, y no se toca lo que ya está completo ni los rangos.
  const esCuenta = () => true;
  const casos = [
    ["-SALDOS!C28", "-SALDOS!C28+SALDOS!B28"],
    ["+SALDOS!B24", "+SALDOS!B24-SALDOS!C24"],
    ["+SALDOS!B24-SALDOS!C24", null],                  // ya está completo
    ["-SUM(SALDOS!C53:C79)", null],                    // un rango toma la columna entera a propósito
    ["+SALDOS!B7+SALDOS!B8", "+SALDOS!B7+SALDOS!B8-SALDOS!C7-SALDOS!C8"],
  ];
  for (const [entra, sale] of casos) {
    const dio = mc.mcCompletar(entra, "B", "C", esCuenta);
    if (dio !== sale) fallo(`"${entra}" tendría que dar ${sale === null ? "null" : `"${sale}"`} y dio ${dio === null ? "null" : `"${dio}"`}`);
  }
  console.log(`\n== la regla: ${casos.length} caso(s) probados`);

  // Y después, que no quede ninguno en los cuatro archivos.
  const exports_ = { mensual: leerExport("mensual"), acumulado: leerExport("acumulado") };
  for (const m of MAESTROS) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path.join(__dirname, m.archivo));
    const antes = aMedias(wb);
    motor.procesarMaestroTFBR({
      wb, cuentasExport: exports_[m.periodo], campoSaldo: m.campo, archivoId: m.id,
      cuentasAcumulado: exports_.acumulado, log: () => {},
    });
    const despues = aMedias(wb);
    console.log(`   ${m.label.padEnd(13)} ${antes.length} renglón(es) a medias en el maestro → ${despues.length} después de procesar`);
    for (const x of despues.slice(0, 5)) fallo(`${m.label}: ${x}`);
  }

  console.log(fallas ? `\n✗ ${fallas} falla(s).` : "\n✓ Ningún renglón lee media cuenta.");
  process.exit(fallas ? 1 : 0);
})();
