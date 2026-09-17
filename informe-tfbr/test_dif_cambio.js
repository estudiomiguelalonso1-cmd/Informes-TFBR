// La diferencia de cambio, calculada en vez de tipeada.
//
// Es un residuo, no un dato que alguien busque en algún lado: las cuentas de activo, pasivo y
// patrimonio se convierten al tipo de cambio de cierre y las de resultado vienen a los suyos,
// así que la suma de todas no da cero. Lo que sobra ES la diferencia de cambio del período, y
// por eso se puede calcular.
//
// La referencia son los informes de julio 2026 que hizo contaduría (INFORMES BASE), que no
// pasaron por el sistema:
//
//   Mensual R$   → residuo 3.231,78   = la fila JULIO del cuadro de meses del Acumulado R$
//   Acumulado R$ → residuo 107.528,26 = el total de ese cuadro y la línea del plan
//
// Lo que este test cuida es que el número salga de las cuentas y siga dando eso. Si alguna vez
// deja de dar, no es que el test esté viejo: es que el residuo cambió y hay que entender por
// qué antes de mover el número esperado.
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

const PERIODO = "2026-07";
const TC = 293.17575;          // el que tiene escrito el Acumulado R$ de julio

// Lo que puso contaduría en julio, y lo que da el sistema. No son el mismo número y la
// diferencia está entendida:
//
//   Mensual R$   contaduría 3.231,776840…   sistema 3.231,85   (7 centavos)
//     Contaduría pega los saldos con todos los decimales; el sistema los redondea a centavos,
//     que es como vienen en el sumas y saldos convertido (ver test_reales.js). Cada uno cierra
//     con su propio número: el residuo se calcula sobre lo que se pegó.
//
//   Acumulado R$ contaduría 107.528,26       sistema 104.171,92  (3.356,34)
//     El informe de julio tiene pegada "4212100000 FLETES" por 3.356,67 y el export acumulado
//     de julio no la trae. El balance cierra igual, pero el acumulado del año sale corto. Por
//     eso escribirDatosDelPeriodo avisa cuando el cuadro y las cuentas no dan lo mismo.
const CONTADURIA_MES = 3231.776840406703;
const CONTADURIA_ACUM = 107528.26;
const DEL_MES = 3231.85;
const RESIDUO_ACUM = 104171.92;
const FLETES = 3356.67;

function leerExport(periodo) {
  const buf = fs.readFileSync(path.join(__dirname, "..", "inputs", PERIODO, `sumas_y_saldos_${periodo}.xls`));
  const libro = XLSX.read(buf, { type: "buffer" });
  const ws = libro.Sheets[libro.SheetNames[0]];
  return P.parseSumasYSaldosTFBR(XLSX.utils.sheet_to_json(ws, { header: 1 }), ws["!merges"]).cuentas;
}

async function correr(archivo, periodo, campo, id, cuentasAcumulado, difCambioDelMes) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path.join(__dirname, archivo));
  const { escritas } = motor.procesarMaestroTFBR({
    wb, cuentasExport: periodo.cuentas, campoSaldo: campo, archivoId: id,
    cuentasAcumulado, log: () => {},
  });
  const datos = pf.escribirDatosDelPeriodo(wb, {
    periodo: PERIODO, diaCierre: 31, tcCierre: String(TC),
    escritas, difCambioDelMes,
  }, () => {});
  return { wb, datos };
}

(async () => {
  let fallas = 0;
  const fallo = (m) => { console.log(`   ✗ ${m}`); fallas++; };

  const mensual = P.convertirSaldosEnReales(leerExport("mensual"), TC);
  const acumulado = P.convertirSaldosEnReales(leerExport("acumulado"), TC);

  // 1. El Mensual R$ calcula la cifra del mes.
  const m = await correr("base_bm_brl.xlsx", { cuentas: mensual }, "saldo_brl",
                         "balance_mensual_brl", acumulado, null);
  console.log(`\n== Mensual R$: residuo ${m.datos.residuo === null ? "—" : m.datos.residuo.toFixed(2)} ` +
              `(contaduría puso ${CONTADURIA_MES.toFixed(2)}; ${Math.abs(DEL_MES - CONTADURIA_MES).toFixed(2)} de redondeo)`);
  if (m.datos.residuo === null || Math.abs(m.datos.residuo - DEL_MES) > 0.02) {
    fallo(`el residuo del Mensual R$ tendría que dar ${DEL_MES.toFixed(2)}`);
  }

  // Y lo escribe en su propia línea del plan, sin que nadie lo tipee.
  {
    const ws = m.wb.getWorksheet("SALDOS");
    const layout = require("./config_tfbr.js").derivarLayoutSaldos(m.wb);
    const linea = pf.pfUbicarLineaDifCambio(ws, layout);
    const v = linea ? ws.getCell(linea.fila, layout.acreedorCol).value : null;
    console.log(`   línea "DIFERENCIA DE CAMBIO" (fila ${linea ? linea.fila : "—"}): ${v}`);
    if (typeof v !== "number" || Math.abs(v - DEL_MES) > 0.02) {
      fallo(`la línea del Mensual R$ quedó en ${v} y tendría que decir ${DEL_MES.toFixed(2)}`);
    }
  }

  // 2. El Acumulado R$ usa esa cifra para el cuadro de meses, y su línea lleva el acumulado.
  const a = await correr("base_ba_brl.xlsx", { cuentas: acumulado }, "saldo_brl",
                         "balance_acumulado_brl", acumulado, m.datos.residuo);
  {
    const ws = a.wb.getWorksheet("SALDOS");
    const layout = require("./config_tfbr.js").derivarLayoutSaldos(a.wb);
    const cuadro = pf.ubicarCuadroDifCambio(ws);
    const total = cuadro ? pf.pfTotalDelCuadro(ws, cuadro) : null;
    const filaJulio = cuadro ? cuadro.filas.find(f => f.mes === 7) : null;
    const vJulio = filaJulio ? ws.getCell(filaJulio.fila, cuadro.colValor).value : null;
    const linea = pf.pfUbicarLineaDifCambio(ws, layout);
    const vLinea = linea ? ws.getCell(linea.fila, layout.acreedorCol).value : null;

    console.log(`\n== Acumulado R$: cuadro JULIO ${vJulio} · total ${total === null ? "—" : total.toFixed(2)} ` +
                `· línea ${vLinea}`);
    console.log(`   su propio residuo: ${a.datos.residuo === null ? "—" : a.datos.residuo.toFixed(2)}`);

    // La cifra del mes que calculó el Mensual R$ tiene que llegar tal cual al cuadro.
    if (typeof vJulio !== "number" || Math.abs(vJulio - DEL_MES) > 0.02) {
      fallo(`la fila JULIO del cuadro quedó en ${vJulio} y tendría que decir ${DEL_MES.toFixed(2)}`);
    }
    // Y la línea del plan, el total del cuadro.
    if (total === null || typeof vLinea !== "number" || Math.abs(vLinea - total) > 0.02) {
      fallo(`la línea quedó en ${vLinea} y el cuadro suma ${total}: tendrían que ser lo mismo`);
    }
    // El residuo de este archivo sale corto por FLETES, que el export no trae. Si algún día
    // el export la trae, esto falla y hay que actualizar los números de arriba.
    if (a.datos.residuo === null || Math.abs(a.datos.residuo - RESIDUO_ACUM) > 0.02) {
      fallo(`el residuo del Acumulado R$ dio ${a.datos.residuo} y se esperaba ${RESIDUO_ACUM.toFixed(2)}`);
    }
    const hueco = total - a.datos.residuo;
    console.log(`   cuadro − cuentas = ${hueco.toFixed(2)} (FLETES, que el export no trae: ${FLETES.toFixed(2)})`);
    // Y el aviso tiene que estar: es lo que hace visible el hueco.
    if (!a.datos.pendiente.some(t => /cuadro de dif de cambio suma/.test(t))) {
      fallo("no avisó que el cuadro y las cuentas no dan lo mismo");
    }
  }

  console.log(fallas ? `\n✗ ${fallas} falla(s).` : "\n✓ La diferencia de cambio sale de las cuentas.");
  process.exit(fallas ? 1 : 0);
})();
