// Prueba de la consolidación de préstamos, contra los 4 maestros y el export de julio.
//
// Lo que tiene que quedar demostrado, después de correr el motor:
//   1. Ningún renglón del Pasivo lee una cuenta de préstamo. Son cuentas 114 —de activo— y
//      en el Pasivo entraban restando: Cuba tiene 200.000 deudor y el Pasivo mostraba
//      -200.000. Se compensaba con el Activo, el balance cerraba, y encima la fila está
//      oculta, así que no había forma de verlo.
//   2. Cada préstamo que el archivo tiene en su plan está en "- Adelanto al personal".
//   3. Ese renglón existe con el mismo nombre en los cuatro, y no está oculto — si queda
//      oculto el importe entra en el subtotal pero no se ve.
//   4. Ningún préstamo queda leído dos veces.
//
// Correr con: node informe-tfbr/test_prestamos.js

const path = require("path");
const fs = require("fs");
const XLSX = require("xlsx");
const ExcelJS = require("exceljs");

global.XLSX = XLSX;
const P = require("./parser_tfbr.js");
const cfg = require("./config_tfbr.js");
const ch = require("./config_hojas.js");
const pr = require("./prestamos.js");
const motor = require("./motor_tfbr.js");

const PERIODO = "2026-07";
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

(async () => {
  let fallas = 0;
  const fallo = (m) => { console.log(`   ✗ ${m}`); fallas++; };
  const ex = { mensual: leerExport("mensual"), acumulado: leerExport("acumulado") };

  for (const m of MAESTROS) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path.join(__dirname, m.archivo));
    motor.procesarMaestroTFBR({
      wb, cuentasExport: ex[m.periodo], campoSaldo: m.campo, archivoId: m.id,
      cuentasAcumulado: ex.acumulado, log: () => {},
    });

    const layout = cfg.derivarLayoutSaldos(wb);
    const plan = ch.chPlanPorFila(wb, layout);
    const enPlan = pr.PRESTAMOS_AL_PERSONAL.filter(c =>
      Object.values(plan).some(x => x.cod === c));

    // Dónde quedó cada préstamo, hoja por hoja.
    const lugares = {};       // cod -> [ "Hoja: rotulo" ]
    for (const c of enPlan) lugares[c] = [];
    for (const h of ch.chHojasConfigurables(wb, layout)) {
      const mapa = ch.chMapaHoja(wb, layout, h.hoja);
      for (const r of mapa.renglones) {
        for (const f of r.filasSaldos) {
          const info = plan[f];
          if (info && lugares[info.cod]) lugares[info.cod].push(`${h.hoja}: "${r.rotulo}"`);
        }
      }
    }

    console.log(`\n== ${m.label} — ${enPlan.length} préstamo(s) en el plan`);

    // 1. Ninguno en el Pasivo.
    for (const [cod, sitios] of Object.entries(lugares)) {
      const enPasivo = sitios.filter(s => /^Pasivo/.test(s));
      if (enPasivo.length) {
        fallo(`${cod} sigue en el Pasivo: ${enPasivo.join(", ")} — es una cuenta de activo`);
      }
    }

    // 2. Cada uno en "- Adelanto al personal".
    for (const cod of enPlan) {
      const bien = lugares[cod].some(s =>
        ch.chNorm(s) === ch.chNorm(`${pr.PR_ROTULO}`.replace(/^-\s*/, "")) ||
        s.indexOf(pr.PR_ROTULO) >= 0);
      if (!bien) fallo(`${cod} no quedó en "${pr.PR_ROTULO}" (está en: ${lugares[cod].join(", ") || "ningún renglón"})`);
    }

    // 3. El renglón existe, se llama igual en los cuatro, y se ve.
    const mapaActivo = ch.chMapaHoja(wb, layout, "Activo");
    const destino = mapaActivo && mapaActivo.renglones.find(r => ch.chNorm(r.rotulo) === ch.chNorm(pr.PR_ROTULO));
    if (!destino) {
      fallo(`no existe el renglón "${pr.PR_ROTULO}" en el Activo`);
    } else {
      if (wb.getWorksheet("Activo").getRow(destino.fila).hidden) {
        fallo(`el renglón "${pr.PR_ROTULO}" quedó oculto: el importe entra en el subtotal pero no se ve`);
      }
      console.log(`   "${pr.PR_ROTULO}" en Activo fila ${destino.fila}, visible`);
    }

    // 4. Ninguno leído dos veces.
    for (const [cod, sitios] of Object.entries(lugares)) {
      if (sitios.length > 1) fallo(`${cod} lo leen ${sitios.length} renglones: ${sitios.join(", ")}`);
    }

    for (const cod of enPlan) {
      console.log(`     ${cod} -> ${lugares[cod].join(", ") || "SIN RENGLÓN"}`);
    }
  }

  console.log(fallas ? `\n✗ ${fallas} falla(s).` : "\n✓ Las 4 pasan: los préstamos van juntos en el Activo y ninguno quedó en el Pasivo.");
  process.exit(fallas ? 1 : 0);
})();
