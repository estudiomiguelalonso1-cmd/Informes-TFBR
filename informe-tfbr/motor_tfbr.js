// Motor de informe-tfbr: escribe SOLO la zona de pegado dentro de SALDOS (ver
// config_tfbr.js) y aplica los fixes aprobados. Todo lo demás (EESP, EERR, EEPN, Activo,
// Pasivo, Anexo I, Anexo II) se resuelve con las fórmulas que el maestro ya tiene, cuando
// se abre en Excel real — el motor nunca intenta recalcular nada (ver docs/formula_analysis.md
// y el plan de arquitectura: TFBR sigue el patrón de round-trip por Excel real, no el de
// calcular en JS).

// derivarLayoutSaldos / leerPlanDeCuentas / ctColNumeroALetra vienen de config_tfbr.js,
// cargado antes que este archivo (ver index.html) — se usan acá como globales, igual que
// motor_balances.js usa insertRowEn de formula_hojas.js. El require() de más abajo es solo
// para que los tests de Node (que no cargan <script> en orden) tengan lo mismo disponible.

// Empareja las cuentas del export contra el plan de cuentas del maestro POR CÓDIGO, nunca
// por el texto tal cual lo trae Onvio ("1120100100 - FORD", con guión) porque el VLOOKUP
// del maestro busca el texto EXACTO que la plantilla ya tiene ("1120100100  FORD", sin
// guión, doble espacio). Ver docs/formula_analysis.md, sección de riesgo de matching.
function emparejarConPlan(cuentasExport, planDeCuentas, campoSaldo) {
  const matcheadas = []; // [{texto, saldo}] listas para escribir en el staging
  const sinMapear = [];  // cuentas del export sin fila correspondiente en el maestro
  for (const c of cuentasExport) {
    const enPlan = planDeCuentas[c.codigo];
    if (!enPlan) { sinMapear.push(c); continue; }
    matcheadas.push({ texto: enPlan.texto, saldo: c[campoSaldo] });
  }
  return { matcheadas, sinMapear };
}

// Limpia el rango de staging completo (evita que queden filas de un mes anterior con más
// cuentas que el actual) y escribe las cuentas emparejadas, una por fila, empezando en la
// esquina superior izquierda del rango.
function escribirStaging(wb, layout, matcheadas, log = () => {}) {
  const ws = wb.getWorksheet(layout.sheet);
  const { colDesde, filaDesde, colHasta, filaHasta } = layout.stagingRange;
  const capacidad = filaHasta - filaDesde + 1;

  if (matcheadas.length > capacidad) {
    throw new Error(
      `El rango de staging (${ctColNumeroALetra(colDesde)}${filaDesde}:${ctColNumeroALetra(colHasta)}${filaHasta}, ` +
      `${capacidad} filas) no alcanza para las ${matcheadas.length} cuentas emparejadas. ` +
      "NO se escribió nada. Hay que ampliar el rango en la plantilla antes de seguir."
    );
  }

  for (let r = filaDesde; r <= filaHasta; r++) {
    ws.getCell(r, colDesde).value = null;
    ws.getCell(r, colDesde + 1).value = null;
  }
  matcheadas.forEach((m, i) => {
    ws.getCell(filaDesde + i, colDesde).value = m.texto;
    ws.getCell(filaDesde + i, colDesde + 1).value = m.saldo;
  });

  log(`  SALDOS: ${matcheadas.length} cuenta(s) escritas en el staging ` +
      `(${ctColNumeroALetra(colDesde)}${filaDesde}:${ctColNumeroALetra(colHasta)}${filaHasta}).`);
}

// Corre el proceso completo para UN archivo/moneda. No guarda el archivo (eso lo decide
// quien llama, según si va a pedir más pasos antes de bajar el .xlsx).
function procesarMaestroTFBR({ wb, cuentasExport, campoSaldo, log = () => {} }) {
  const layout = derivarLayoutSaldos(wb);
  const planDeCuentas = leerPlanDeCuentas(wb, layout);
  const { matcheadas, sinMapear } = emparejarConPlan(cuentasExport, planDeCuentas, campoSaldo);

  escribirStaging(wb, layout, matcheadas, log);

  if (sinMapear.length) {
    log(`  ⚠ ${sinMapear.length} cuenta(s) del export NO están en el plan de cuentas de este ` +
        `maestro y quedaron SIN incluir en ningún total: ` +
        sinMapear.map(c => `${c.codigo} (${c.nombre})`).join(", ") + ".");
  }

  const totalEscrito = matcheadas.reduce((s, m) => s + m.saldo, 0);
  const totalExport = cuentasExport.reduce((s, c) => s + c[campoSaldo], 0);
  log(`  Total escrito: ${totalEscrito.toFixed(2)} | Total del export: ${totalExport.toFixed(2)}` +
      (Math.abs(totalEscrito - totalExport) > 0.02
        ? " ⚠ NO COINCIDEN (revisar cuentas sin mapear arriba)."
        : " (coinciden)."));

  return {
    layout,
    resumen: {
      cuentasExport: cuentasExport.length,
      cuentasEscritas: matcheadas.length,
      sinMapear: sinMapear.map(c => ({ codigo: c.codigo, nombre: c.nombre, saldo: c[campoSaldo] })),
      totalEscrito,
      totalExport,
    },
  };
}

if (typeof module !== "undefined") {
  const cfg = require("./config_tfbr.js");
  global.derivarLayoutSaldos = cfg.derivarLayoutSaldos;
  global.leerPlanDeCuentas = cfg.leerPlanDeCuentas;
  global.ctColNumeroALetra = cfg.ctColNumeroALetra;
  module.exports = { emparejarConPlan, escribirStaging, procesarMaestroTFBR };
}
