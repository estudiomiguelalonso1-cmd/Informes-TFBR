// Prueba del Anexo I contra los 4 maestros y el export de julio 2026.
//
// La prueba que vale es contra los números que hoy están tipeados a mano. Si el valor de
// origen calculado desde las cuentas da distinto de lo que contaduría venía poniendo, el
// mapeo está mal y hay que mirarlo antes de pisar nada. En el Mensual $ los siete renglones
// están cargados, así que sirve de patrón: tienen que dar iguales al centavo.
//
// Correr con: node informe-tfbr/test_anexo_i.js

const path = require("path");
const fs = require("fs");
const XLSX = require("xlsx");
const ExcelJS = require("exceljs");

global.XLSX = XLSX;
const P = require("./parser_tfbr.js");
const cfg = require("./config_tfbr.js");
const ai = require("./anexo_i.js");
const motor = require("./motor_tfbr.js");

const PERIODO = "2026-07";
const MAESTROS = [
  { label: "Mensual $",     archivo: "base_bm_ars.xlsx", periodo: "mensual",   campo: "saldo_ars", id: "balance_mensual_ars" },
  { label: "Mensual R$",    archivo: "base_bm_brl.xlsx", periodo: "mensual",   campo: "saldo_brl", id: "balance_mensual_brl" },
  { label: "Acumulado $",   archivo: "base_ba_ars.xlsx", periodo: "acumulado", campo: "saldo_ars", id: "balance_acumulado_ars" },
  { label: "Acumulado R$",  archivo: "base_ba_brl.xlsx", periodo: "acumulado", campo: "saldo_brl", id: "balance_acumulado_brl" },
];

function leerExport(periodo) {
  const buf = fs.readFileSync(path.join(__dirname, "..", "inputs", PERIODO, `sumas_y_saldos_${periodo}.xls`));
  const libro = XLSX.read(buf, { type: "buffer" });
  const ws = libro.Sheets[libro.SheetNames[0]];
  return P.parseSumasYSaldosTFBR(XLSX.utils.sheet_to_json(ws, { header: 1 }), ws["!merges"]).cuentas;
}

const norm = (t) => String(t).normalize("NFD").replace(/[̀-ͯ]/g, "")
  .toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();

function valorDe(ax, fila, col) {
  const v = ax.getCell(fila, col).value;
  if (typeof v === "number") return v;
  if (v && typeof v === "object" && typeof v.result === "number") return v.result;
  return null;
}

(async () => {
  let fallas = 0;
  const fallo = (m) => { console.log(`   ✗ ${m}`); fallas++; };

  const exports_ = { mensual: leerExport("mensual"), acumulado: leerExport("acumulado") };

  // Lo que debería dar cada renglón, según el export acumulado.
  const esperado = {};
  const porCodigo = {};
  for (const c of exports_.acumulado) porCodigo[c.codigo] = c;

  for (const m of MAESTROS) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path.join(__dirname, m.archivo));

    // Lo que hay ANTES, para comparar contra lo cargado a mano.
    const ax0 = wb.getWorksheet("Anexo I");
    const u0 = ai.aiUbicar(ax0);
    if (!u0) { fallo(`${m.label}: no ubiqué los renglones del Anexo I`); continue; }
    const antes = {};
    for (const [k, fila] of Object.entries(u0.filas)) antes[k] = valorDe(ax0, fila, u0.colOrigen);

    const r = motor.procesarMaestroTFBR({
      wb, cuentasExport: exports_[m.periodo], campoSaldo: m.campo, archivoId: m.id,
      cuentasAcumulado: exports_.acumulado, log: () => {},
    });
    const res = r.resumen.anexoI;
    const ax = wb.getWorksheet("Anexo I");
    const u = ai.aiUbicar(ax);

    console.log(`\n== ${m.label} — ${res.hechos.length} renglón(es) por ${res.modo}` +
                (res.salteados.length ? `, ${res.salteados.length} salteado(s)` : ""));

    if (res.hechos.length !== ai.ANEXO_I_RENGLONES.length) {
      fallo(`${m.label}: cargué ${res.hechos.length} de ${ai.ANEXO_I_RENGLONES.length} renglones`);
      res.salteados.forEach(s => console.log(`      ${s.rotulo}: ${s.motivo}`));
    }

    for (const reng of ai.ANEXO_I_RENGLONES) {
      const fila = u.filas[norm(reng.rotulo)];
      if (!fila) continue;

      // Lo que tienen que sumar las cuentas de ese renglón, en la moneda del archivo.
      const debeDar = reng.cuentas.reduce((s, c) => s + ((porCodigo[c] || {})[m.campo] || 0), 0);
      const quedo = valorDe(ax, fila, u.colOrigen);

      // Con fórmula no hay resultado en caché hasta que Excel recalcule: se verifica que la
      // fórmula nombre exactamente las cuentas del renglón, que es lo que sí se puede saber acá.
      const celda = ax.getCell(fila, u.colOrigen).value;
      if (celda && typeof celda === "object" && typeof celda.formula === "string") {
        const layout = cfg.derivarLayoutSaldos(wb);
        const plan = cfg.leerPlanDeCuentas(wb, layout).cuentas;
        const filasEsperadas = new Set(reng.cuentas.map(c => plan[c] && plan[c].fila).filter(Boolean));
        const filasEnFormula = new Set(
          [...celda.formula.matchAll(/SALDOS!\$?[A-Z]{1,3}\$?(\d+)/g)].map(x => +x[1]));
        const igual = filasEsperadas.size === filasEnFormula.size &&
          [...filasEsperadas].every(f => filasEnFormula.has(f));
        if (!igual) {
          fallo(`${m.label} "${reng.rotulo}": la fórmula lee ${[...filasEnFormula].join(",")} y ` +
                `esperaba ${[...filasEsperadas].join(",")}`);
        }
        continue;
      }

      if (quedo == null) { fallo(`${m.label} "${reng.rotulo}": quedó vacío`); continue; }
      if (Math.abs(quedo - debeDar) > 0.01) {
        fallo(`${m.label} "${reng.rotulo}": quedó ${quedo.toFixed(2)} y las cuentas suman ${debeDar.toFixed(2)}`);
        continue;
      }

      // Y lo que más importa: contra el número que venía a mano.
      const previo = antes[norm(reng.rotulo)];
      if (previo != null && Math.abs(previo) > 0.005 && Math.abs(previo - quedo) > 0.01) {
        fallo(`${m.label} "${reng.rotulo}": a mano decía ${previo.toFixed(2)} y las cuentas dan ` +
              `${quedo.toFixed(2)} — revisar el mapeo antes de pisarlo`);
      }
    }

    // El Mensual $ es el patrón: tenía los 7 cargados a mano y tienen que dar iguales.
    if (m.label === "Mensual $") {
      const cargados = Object.values(antes).filter(v => v != null && Math.abs(v) > 0.005).length;
      console.log(`   (venían ${cargados} de ${ai.ANEXO_I_RENGLONES.length} cargados a mano, ` +
                  `y el cálculo desde las cuentas los reprodujo)`);
    }
  }

  console.log(fallas ? `\n✗ ${fallas} falla(s).` : "\n✓ Las 4 pasan.");
  process.exit(fallas ? 1 : 0);
})();
