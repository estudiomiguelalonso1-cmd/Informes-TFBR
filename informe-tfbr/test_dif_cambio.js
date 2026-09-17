// La diferencia de cambio, calculada en vez de tipeada.
//
// Es un residuo, no un dato que alguien busque en algún lado: las cuentas de activo, pasivo y
// patrimonio se convierten al tipo de cambio de cierre y las de resultado vienen a los suyos,
// así que la suma de todas no da cero. Lo que sobra ES la diferencia de cambio del período que
// cubre el archivo: la del MES en el Mensual R$ y la del AÑO en el Acumulado R$.
//
// La fila del mes del cuadro de meses —que solo existe en el Acumulado R$— se calcula como el
// residuo del año menos lo que ya suma el resto del cuadro. No se toma del Mensual R$, aunque
// sea lo natural, porque el export mensual de Onvio no trae los movimientos de las cuentas de
// patrimonio: en agosto de 2026 el acumulado dice que "3310000000 RNA EJERCICIO ANTERIOR" se
// movió -21.174,07 y el mensual no lo reporta. Con esas cuentas afuera el cuadro quedaba corto
// en 12.413,97 y el balance en reales no cerraba.
//
// La referencia es el cierre de julio 2026, que hizo contaduría a mano y no pasó por el
// sistema. Su Acumulado R$ tiene 107.528,26 en la línea del plan y 3.231,78 en la fila de JULIO
// del cuadro.
//
// Sobre los centavos: el sistema redondea el saldo de cada cuenta a centavos —es la regla de
// conversión, ver test_reales.js— y contaduría pega los saldos con todos los decimales. Sobre
// ~150 cuentas eso da 33 centavos de diferencia en el acumulado y 7 en el mensual. No es un
// error de método: es el redondeo, y por eso la comparación admite medio peso.
//
// Correr con: node informe-tfbr/test_dif_cambio.js

const path = require("path");
const fs = require("fs");
const XLSX = require("xlsx");
const ExcelJS = require("exceljs");

global.XLSX = XLSX;
const P = require("./parser_tfbr.js");
const motor = require("./motor_tfbr.js");
const pf = require("./periodo_tfbr.js");
const cfg = require("./config_tfbr.js");

const PERIODO = "2026-07";
const TC = 293.17575;            // el que tiene escrito el Acumulado R$ de julio
const CONTADURIA_ANIO = 107528.26;
const CONTADURIA_MES = 3231.78;
const CENTAVOS = 0.5;            // lo que puede mover el redondeo por cuenta

function leerExport(cual) {
  const f = path.join(__dirname, "..", "inputs", PERIODO, `sumas_y_saldos_${cual}_completo.xls`);
  const libro = XLSX.read(fs.readFileSync(f), { type: "buffer" });
  const ws = libro.Sheets[libro.SheetNames[0]];
  return P.parseSumasYSaldosTFBR(XLSX.utils.sheet_to_json(ws, { header: 1 }), ws["!merges"]).cuentas;
}

async function correr(maestro, id, cuentas, cuentasAcumulado) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path.join(__dirname, maestro));
  const { escritas } = motor.procesarMaestroTFBR({
    wb, cuentasExport: cuentas, campoSaldo: "saldo_brl", archivoId: id,
    cuentasAcumulado, log: () => {},
  });
  const datos = pf.escribirDatosDelPeriodo(wb, {
    periodo: PERIODO, diaCierre: 31, tcCierre: String(TC), escritas,
  }, () => {});
  return { wb, datos };
}

(async () => {
  let fallas = 0;
  const fallo = (m) => { console.log(`   ✗ ${m}`); fallas++; };

  const mensual = P.convertirSaldosEnReales(leerExport("mensual"), TC);
  const acumulado = P.convertirSaldosEnReales(leerExport("acumulado"), TC);

  // 1. El Mensual R$: su residuo es la diferencia del mes, y va a su propia línea.
  const m = await correr("base_bm_brl.xlsx", "balance_mensual_brl", mensual, acumulado);
  console.log(`\n== Mensual R$: residuo ${m.datos.residuo.toFixed(2)} ` +
              `(contaduría ${CONTADURIA_MES.toFixed(2)})`);
  if (Math.abs(m.datos.residuo - CONTADURIA_MES) > CENTAVOS) {
    fallo(`el residuo del Mensual R$ tendría que dar ${CONTADURIA_MES.toFixed(2)}`);
  }
  {
    const ws = m.wb.getWorksheet("SALDOS");
    const layout = cfg.derivarLayoutSaldos(m.wb);
    const linea = pf.pfUbicarLineaDifCambio(ws, layout);
    if (!linea) fallo("el Mensual R$ no tiene línea de diferencia de cambio");
    else {
      // La línea tiene que dejar el balance en cero: su aporte es deudora menos acreedora, y
      // eso tiene que ser justo lo contrario del residuo.
      const g = (col) => {
        const v = ws.getCell(linea.fila, col).value;
        if (typeof v === "number") return v;
        if (v && typeof v === "object" && (v.formula || v.sharedFormula)) {
          // la fórmula replica max(saldo,0) / max(-saldo,0)
          let s = ws.getCell(linea.fila, layout.saldoCol).value;
          if (s && typeof s === "object") s = s.result;
          if (typeof s !== "number") return 0;
          return col === layout.deudorCol ? Math.max(s, 0) : Math.max(-s, 0);
        }
        return 0;
      };
      const aporte = g(layout.deudorCol) - g(layout.acreedorCol);
      console.log(`   la línea aporta ${aporte.toFixed(2)} → el balance queda en ` +
                  `${(m.datos.residuo + aporte).toFixed(2)}`);
      if (Math.abs(m.datos.residuo + aporte) > 0.02) {
        fallo(`la línea no deja el balance en cero: queda ${(m.datos.residuo + aporte).toFixed(2)}`);
      }
    }
  }

  // 2. El Acumulado R$: su residuo es la del año, y la fila del mes sale de ahí.
  const a = await correr("base_ba_brl.xlsx", "balance_acumulado_brl", acumulado, acumulado);
  const ws = a.wb.getWorksheet("SALDOS");
  const layout = cfg.derivarLayoutSaldos(a.wb);
  const cuadro = pf.ubicarCuadroDifCambio(ws);
  const total = cuadro ? pf.pfTotalDelCuadro(ws, cuadro) : null;
  const filaJulio = cuadro ? cuadro.filas.find(f => f.mes === 7) : null;
  let vJulio = filaJulio ? ws.getCell(filaJulio.fila, cuadro.colValor).value : null;
  if (vJulio && typeof vJulio === "object") vJulio = vJulio.result;

  console.log(`\n== Acumulado R$: residuo ${a.datos.residuo.toFixed(2)} ` +
              `(contaduría ${CONTADURIA_ANIO.toFixed(2)})`);
  console.log(`   fila JULIO del cuadro ${typeof vJulio === "number" ? vJulio.toFixed(2) : vJulio} ` +
              `(contaduría ${CONTADURIA_MES.toFixed(2)}) · total del cuadro ` +
              `${total === null ? "—" : total.toFixed(2)}`);

  if (Math.abs(a.datos.residuo - CONTADURIA_ANIO) > CENTAVOS) {
    fallo(`el residuo del Acumulado R$ tendría que dar ${CONTADURIA_ANIO.toFixed(2)}`);
  }
  if (typeof vJulio !== "number" || Math.abs(vJulio - CONTADURIA_MES) > CENTAVOS) {
    fallo(`la fila JULIO tendría que dar ${CONTADURIA_MES.toFixed(2)} y dio ${vJulio}`);
  }
  // Lo que hace que el balance cierre: el cuadro tiene que dar lo mismo que las cuentas.
  if (total === null || Math.abs(total - a.datos.residuo) > 0.02) {
    fallo(`el cuadro suma ${total} y las cuentas dan ${a.datos.residuo.toFixed(2)}: no cierran`);
  }
  // Y ese mismo número tiene que estar en la línea del plan.
  const linea = pf.pfUbicarLineaDifCambio(ws, layout);
  let vLinea = linea ? ws.getCell(linea.fila, layout.acreedorCol).value : null;
  if (vLinea && typeof vLinea === "object") vLinea = vLinea.result;
  if (typeof vLinea !== "number" || Math.abs(vLinea - total) > 0.02) {
    fallo(`la línea quedó en ${vLinea} y el cuadro suma ${total}`);
  }

  console.log(fallas ? `\n✗ ${fallas} falla(s).` : "\n✓ La diferencia de cambio sale de las cuentas y el cuadro cierra contra ellas.");
  process.exit(fallas ? 1 : 0);
})();
