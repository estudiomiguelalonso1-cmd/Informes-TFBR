// Prueba de los préstamos al personal, contra los 4 maestros y el export de julio.
//
// Lo que tiene que quedar demostrado:
//   1. Ninguno en el Pasivo. Son cuentas 114 —de activo— y en el Pasivo entraban restando:
//      Cuba tiene 200.000 deudor y el Pasivo mostraba -200.000. Se compensaba con el Activo,
//      el balance cerraba, y la fila está oculta, así que no había forma de verlo.
//   2. Cada préstamo en SU renglón, con el mismo nombre exacto en los cuatro. No agrupados:
//      los Acumulados los metían todos en "Adelanto al personal" y así no se pueden comparar
//      contra los Mensuales, que los abren uno por uno.
//   3. Los seis renglones existen en los cuatro archivos, incluso donde la cuenta todavía no
//      está en el plan — si no, el día que el sumas la traiga no habría dónde ponerla.
//   4. Ninguno leído dos veces.
//   5. El que tiene importe este mes, visible. Un renglón oculto suma en el subtotal pero no
//      sale en el informe.
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
    const res = motor.procesarMaestroTFBR({
      wb, cuentasExport: ex[m.periodo], campoSaldo: m.campo, archivoId: m.id,
      cuentasAcumulado: ex.acumulado, log: () => {},
    });

    const layout = cfg.derivarLayoutSaldos(wb);
    const plan = ch.chPlanPorFila(wb, layout);
    const porCodigo = {};
    for (const x of Object.values(plan)) porCodigo[x.cod] = x;

    // Dónde quedó cada préstamo.
    const lugares = {};
    for (const p of pr.PRESTAMOS_AL_PERSONAL) lugares[p.cuenta] = [];
    for (const h of ch.chHojasConfigurables(wb, layout)) {
      const mapa = ch.chMapaHoja(wb, layout, h.hoja);
      for (const r of mapa.renglones) {
        for (const f of r.filasSaldos) {
          const info = plan[f];
          if (info && lugares[info.cod]) lugares[info.cod].push({ hoja: h.hoja, rotulo: r.rotulo, fila: r.fila });
        }
      }
    }

    const mapaActivo = ch.chMapaHoja(wb, layout, "Activo");
    const enPlan = pr.PRESTAMOS_AL_PERSONAL.filter(p => porCodigo[p.cuenta]);
    console.log(`\n== ${m.label} — ${enPlan.length} de ${pr.PRESTAMOS_AL_PERSONAL.length} préstamos en el plan`);

    for (const p of pr.PRESTAMOS_AL_PERSONAL) {
      // 3. El renglón existe, con el nombre exacto.
      const reng = mapaActivo && mapaActivo.renglones.find(x => String(x.rotulo).trim() === p.rotulo);
      if (!reng) { fallo(`falta el renglón "${p.rotulo}" en el Activo`); continue; }

      if (!porCodigo[p.cuenta]) { console.log(`     ${p.rotulo} — renglón listo (la cuenta aún no está en el plan)`); continue; }

      const sitios = lugares[p.cuenta];

      // 1. Ninguno en el Pasivo.
      const enPasivo = sitios.filter(s => s.hoja === "Pasivo");
      if (enPasivo.length) {
        fallo(`${p.cuenta} sigue en el Pasivo: ${enPasivo.map(s => `"${s.rotulo}"`).join(", ")} — es una cuenta de activo`);
      }

      // 2. En SU renglón.
      const bien = sitios.some(s => s.hoja === "Activo" && String(s.rotulo).trim() === p.rotulo);
      if (!bien) {
        fallo(`${p.cuenta} no quedó en "${p.rotulo}" (está en: ` +
              `${sitios.map(s => `${s.hoja} "${s.rotulo}"`).join(", ") || "ningún renglón"})`);
      }

      // 4. Una sola vez.
      if (sitios.length > 1) {
        fallo(`${p.cuenta} lo leen ${sitios.length} renglones: ${sitios.map(s => `${s.hoja} "${s.rotulo}"`).join(", ")}`);
      }

      // 5. Con importe, visible.
      const importe = res.escritas[p.cuenta];
      const oculto = wb.getWorksheet("Activo").getRow(reng.fila).hidden;
      if (importe !== undefined && Math.abs(importe) > 0.005 && oculto) {
        fallo(`"${p.rotulo}" tiene ${importe.toFixed(2)} y la fila quedó oculta: suma en el subtotal pero no sale en el informe`);
      }
      console.log(`     ${p.rotulo.padEnd(28)} fila ${String(reng.fila).padStart(3)}` +
                  (importe !== undefined && Math.abs(importe) > 0.005 ? `  ${importe.toFixed(2)}` : "  (sin importe)") +
                  (oculto ? "  [oculta]" : ""));
    }
  }

  console.log(fallas ? `\n✗ ${fallas} falla(s).` : "\n✓ Las 4 pasan: cada préstamo en su renglón del Activo, ninguno en el Pasivo.");
  process.exit(fallas ? 1 : 0);
})();
