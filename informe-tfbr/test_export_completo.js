// El export de Onvio hay que pedirlo CON las cuentas saldadas.
//
// El reporte corto omite las cuentas con saldo cero, y mide el cero en pesos. Hay cuentas de
// resultado que cierran en cero en pesos y NO en reales: los movimientos del año se
// convirtieron a tipos de cambio distintos y no se cancelan. En agosto de 2026 es
// "4212100000 FLETES", con $ 0,00 y R$ 3.356,67.
//
// Con el reporte corto esa plata no llega al informe en reales, y el único síntoma es la
// diferencia de cambio saliendo corta — nada dice que falta una cuenta. Por eso hay dos cosas
// que cuidar acá: que el parser reconozca el reporte corto para poder avisar, y que el largo
// no traiga de arrastre las 940 cuentas del plan a la zona de pegado.
//
// Los dos exports de agosto 2026 están en inputs/2026-08/.
//
// Correr con: node informe-tfbr/test_export_completo.js

const path = require("path");
const fs = require("fs");
const XLSX = require("xlsx");
const ExcelJS = require("exceljs");

global.XLSX = XLSX;
const P = require("./parser_tfbr.js");
const motor = require("./motor_tfbr.js");
const pf = require("./periodo_tfbr.js");

const DIR = path.join(__dirname, "..", "inputs", "2026-08");
const TC = 291.3301;
const FLETES = { codigo: "4212100000", brl: 3356.67 };

function leer(archivo) {
  const libro = XLSX.read(fs.readFileSync(path.join(DIR, archivo)), { type: "buffer" });
  const ws = libro.Sheets[libro.SheetNames[0]];
  return P.parseSumasYSaldosTFBR(XLSX.utils.sheet_to_json(ws, { header: 1 }), ws["!merges"]);
}

async function residuoDe(cuentas, maestro = "base_ba_brl.xlsx", id = "balance_acumulado_brl") {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path.join(__dirname, maestro));
  const { escritas } = motor.procesarMaestroTFBR({
    wb, cuentasExport: cuentas, campoSaldo: "saldo_brl",
    archivoId: id, cuentasAcumulado: cuentas, log: () => {},
  });
  const datos = pf.escribirDatosDelPeriodo(wb, {
    periodo: "2026-08", diaCierre: 31, tcCierre: String(TC), escritas,
  }, () => {});
  return { residuo: datos.residuo, escritas };
}

(async () => {
  let fallas = 0;
  const fallo = (m) => { console.log(`   ✗ ${m}`); fallas++; };

  const corto = leer("sumas_y_saldos_acumulado.xls");
  const largo = leer("sumas_y_saldos_acumulado_completo.xls");

  console.log(`\n== corto: ${corto.cuentas.length} cuentas, ${corto.cuentasEnCeroEnPesos} en cero` +
              ` · largo: ${largo.cuentas.length} cuentas, ${largo.cuentasEnCeroEnPesos} en cero`);

  // 1. Reconocerlos, que es lo que permite avisar.
  if (!corto.pareceReporteCorto) fallo("no reconocí el reporte corto");
  if (largo.pareceReporteCorto) fallo("tomé el reporte completo por corto");

  // 2. Los dos tienen que decir lo mismo de las cuentas que los dos traen.
  const porCodigo = {};
  for (const c of largo.cuentas) porCodigo[c.codigo] = c;
  let distintas = 0;
  for (const c of corto.cuentas) {
    const o = porCodigo[c.codigo];
    if (!o) { fallo(`${c.codigo} está en el corto y no en el completo`); continue; }
    if (Math.abs(o.saldo_ars - c.saldo_ars) > 0.02) distintas++;
  }
  if (distintas) fallo(`${distintas} cuenta(s) tienen distinto saldo en pesos entre los dos exports`);

  // 3. FLETES: el caso que motiva todo esto.
  const f = porCodigo[FLETES.codigo];
  if (!f) fallo(`el export completo no trae ${FLETES.codigo}`);
  else {
    console.log(`   ${FLETES.codigo} FLETES: $ ${f.saldo_ars.toFixed(2)} · R$ ${f.saldo_brl.toFixed(2)}`);
    if (Math.abs(f.saldo_ars) > 0.005) fallo("FLETES tendría que tener saldo cero en pesos");
    if (Math.abs(f.saldo_brl - FLETES.brl) > 0.02) fallo(`FLETES tendría que tener R$ ${FLETES.brl}`);
  }

  // 4. El resultado: con el completo entra, con el corto no, y la diferencia es su importe.
  const a = await residuoDe(P.convertirSaldosEnReales(corto.cuentas, TC));
  const b = await residuoDe(P.convertirSaldosEnReales(largo.cuentas, TC));
  console.log(`   residuo con el corto ${a.residuo.toFixed(2)} · con el completo ${b.residuo.toFixed(2)}` +
              ` · diferencia ${(b.residuo - a.residuo).toFixed(2)}`);

  if (a.escritas[FLETES.codigo] !== undefined) fallo("con el reporte corto FLETES no debería entrar");
  if (b.escritas[FLETES.codigo] === undefined) fallo("con el reporte completo FLETES tiene que entrar");
  if (Math.abs((b.residuo - a.residuo) - FLETES.brl) > 0.02) {
    fallo(`la diferencia de residuo tendría que ser exactamente FLETES (${FLETES.brl})`);
  }

  // 5. Y las 940 cuentas del plan no se pegan: las que no tienen saldo no aportan nada y
  //    taparían el aviso de cuentas sin mapear.
  const pegadas = Object.keys(b.escritas).length;
  console.log(`   cuentas pegadas con el completo: ${pegadas} (el export trae ${largo.cuentas.length})`);
  if (pegadas > corto.cuentas.length + 10) {
    fallo(`se pegaron ${pegadas} cuentas: las que están en cero no tendrían que pegarse`);
  }

  // 6. Lo mismo para el mensual. En agosto de 2026 las 13 cuentas que agrega el completo son
  //    todas patrimoniales con los pesos en cero, y el R$ de esas sale de los pesos dividido el
  //    TC, asi que dan cero y el numero no se mueve. Igual hay que pedirlo completo: el mes que
  //    aparezca una cuenta de resultado como FLETES, el corto se la come sin avisar.
  const mCorto = leer("sumas_y_saldos_mensual.xls");
  const mLargo = leer("sumas_y_saldos_mensual_completo.xls");
  console.log(`
== mensual — corto: ${mCorto.cuentas.length} cuentas · completo: ${mLargo.cuentas.length}`);

  if (!mCorto.pareceReporteCorto) fallo("no reconocí el mensual corto");
  if (mLargo.pareceReporteCorto) fallo("tomé el mensual completo por corto");

  const ma = await residuoDe(P.convertirSaldosEnReales(mCorto.cuentas, TC), "base_bm_brl.xlsx", "balance_mensual_brl");
  const mb = await residuoDe(P.convertirSaldosEnReales(mLargo.cuentas, TC), "base_bm_brl.xlsx", "balance_mensual_brl");
  console.log(`   diferencia de cambio del mes: corto ${ma.residuo.toFixed(2)} · completo ${mb.residuo.toFixed(2)}` +
              ` · pegadas ${Object.keys(mb.escritas).length}`);
  if (Math.abs(ma.residuo - mb.residuo) > 0.02) {
    fallo(`en agosto los dos exports tendrían que dar lo mismo y dan ${ma.residuo.toFixed(2)} y ` +
          `${mb.residuo.toFixed(2)}: apareció una cuenta que el corto se come`);
  }
  if (Object.keys(mb.escritas).length > mCorto.cuentas.length + 10) {
    fallo(`se pegaron ${Object.keys(mb.escritas).length} cuentas en el mensual`);
  }


  console.log(fallas ? `\n✗ ${fallas} falla(s).` : "\n✓ El export completo suma lo que falta y no arrastra lo que sobra.");
  process.exit(fallas ? 1 : 0);
})();
