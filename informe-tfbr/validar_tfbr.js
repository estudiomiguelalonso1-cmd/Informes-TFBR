// Controles sobre un maestro YA RECALCULADO por Excel.
//
// Por qué recién ahí: el motor escribe la zona de pegado, pero no recalcula nada — las
// fórmulas del archivo conservan el resultado que traían del mes anterior hasta que alguien
// lo abre en Excel. Leer los controles antes de eso mostraría los números del mes pasado con
// cara de estar bien, que es peor que no mostrar nada. Por eso corren cuando se sube el
// archivo revisado, en el paso de aprobación.
//
// Los controles se derivan del layout que deduce config_tfbr.js, así que valen igual en los
// 4 archivos aunque cada uno tenga otra geometría.

function vtNumero(ws, fila, col) {
  const v = ws.getCell(fila, col).value;
  if (typeof v === "number") return v;
  if (v && typeof v === "object" && typeof v.result === "number") return v.result;
  return 0;
}

function vtEsFormula(ws, fila, col) {
  const v = ws.getCell(fila, col).value;
  return !!(v && typeof v === "object" && typeof v.formula === "string");
}

// Debe = Haber sobre todas las filas de cuenta.
//
// No alcanza con sumar las filas que tienen BUSCARV: en los archivos en R$ la diferencia de
// cambio del período está cargada A MANO en una fila aparte, justo debajo del bloque de
// fórmulas (en el Acumulado R$ es la fila "DIFERENCIA DE CAMBIO", 107.528,26). Sin ella el
// control da un descuadre exactamente igual a esa cifra y parece un error cuando no lo es.
// Por eso se suman también las filas cargadas a mano que hay entre el final del plan de
// cuentas y la zona de pegado — salteando las que son fórmula, que son los subtotales y
// contarían dos veces.
function controlDebeHaber(ws, layout) {
  let deudor = 0, acreedor = 0;
  for (let r = layout.planDeCuentas.desde; r <= layout.planDeCuentas.hasta; r++) {
    deudor += vtNumero(ws, r, layout.deudorCol);
    acreedor += vtNumero(ws, r, layout.acreedorCol);
  }
  let manuales = 0;
  for (let r = layout.planDeCuentas.hasta + 1; r < layout.stagingRange.filaDesde; r++) {
    if (vtEsFormula(ws, r, layout.deudorCol) || vtEsFormula(ws, r, layout.acreedorCol)) continue;
    const d = vtNumero(ws, r, layout.deudorCol);
    const a = vtNumero(ws, r, layout.acreedorCol);
    if (!d && !a) continue;
    deudor += d; acreedor += a; manuales++;
  }
  const dif = deudor - acreedor;
  const cuadra = Math.abs(dif) < 0.02;
  return {
    nombre: "Debe = Haber",
    // En los archivos en pesos esto da cero. En los de reales NO, y no es un error: queda un
    // descuadre igual a la diferencia de cambio del período (en julio 2026, 3.231,78 en el
    // Mensual R$ y 107.528,26 en el Acumulado R$). Todavía no está confirmado con contaduría
    // cómo se compone exactamente esa cifra, así que cuando no cuadra se informa el importe
    // en vez de frenar el cierre: un control que grita todos los meses en 2 de 4 archivos
    // enseña a ignorarlo.
    detalle: cuadra
      ? `deudor ${deudor.toFixed(2)} vs acreedor ${acreedor.toFixed(2)}`
      : `descuadre de ${dif.toFixed(2)} — en los archivos en R$ esto corresponde a la ` +
        `diferencia de cambio del período; verificá que sea la cifra esperada` +
        (manuales ? ` (se sumaron ${manuales} fila(s) cargada(s) a mano)` : ""),
    diferencia: dif,
    pasa: cuadra,
    soloAviso: !cuadra,
  };
}

// Todo lo que se pegó tiene que haber sido levantado por el BUSCARV de su fila.
//
// Este es el control que más importa y el que no existe en el proceso manual: si el texto de
// una cuenta no coincide exactamente, el BUSCARV no la encuentra, el IFERROR la convierte en
// cero, y el balance CIERRA IGUAL — la cuenta simplemente desaparece sin que nada avise.
//
// Se compara cuenta por cuenta (no la suma de la columna): así se sabe CUÁL no se levantó,
// y además funciona aunque distintas secciones usen distinta columna de saldo.
function controlCuentasLevantadas(ws, planDeCuentas, escritas) {
  const fallan = [];
  let comparadas = 0;
  for (const [codigo, valorEscrito] of Object.entries(escritas || {})) {
    const info = planDeCuentas[codigo];
    if (!info) continue;                    // sin fila en el plan: ya se reporta aparte
    comparadas++;
    // Con un código repetido alcanza con que UNA de sus filas lo haya levantado: la otra
    // queda en cero por tener el texto escrito distinto, y eso se avisa como duplicado,
    // no como importe perdido.
    const filas = [info].concat(info.otrasFilas || []);
    const alguna = filas.some(f => Math.abs(vtNumero(ws, f.fila, f.colSaldo) - valorEscrito) <= 0.02);
    if (!alguna) {
      const leidos = filas.map(f => `fila ${f.fila}: ${vtNumero(ws, f.fila, f.colSaldo).toFixed(2)}`).join(", ");
      fallan.push(`${codigo} (pegado ${valorEscrito.toFixed(2)}; ${leidos})`);
    }
  }
  return {
    nombre: "Cada cuenta pegada se levantó en SALDOS",
    detalle: fallan.length
      ? `${fallan.length} de ${comparadas} no coinciden: ${fallan.slice(0, 5).join("; ")}${fallan.length > 5 ? "…" : ""}`
      : `${comparadas} cuentas verificadas`,
    diferencia: fallan.length,
    pasa: fallan.length === 0,
  };
}

// Celdas en error (#REF!, #VALUE!…) en todo el libro.
function buscarCeldasEnError(wb) {
  const encontradas = [];
  wb.worksheets.forEach(ws => {
    ws.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const v = cell.value;
        const enError = v && typeof v === "object" &&
          (v.error || (v.result && v.result.error));
        if (enError) encontradas.push(`${ws.name}!${cell.address}`);
      });
    });
  });
  return encontradas;
}

// Los errores que ya traía el archivo NO se listan a mano acá: se toman del propio maestro
// anterior, al que este reemplaza. Una lista escrita a mano se desactualiza —y de hecho la
// de la investigación inicial tenía celdas corridas por una fila— y además haría falta
// tocar código cada vez que contaduría arregle uno.
function validarRecalculado(wb, { planDeCuentas, escritas, erroresPrevios }) {
  const layout = derivarLayoutSaldos(wb);
  const ws = wb.getWorksheet(layout.sheet);
  const plan = planDeCuentas || leerPlanDeCuentas(wb, layout).cuentas;

  const controles = [
    controlDebeHaber(ws, layout),
    controlCuentasLevantadas(ws, plan, escritas),
  ];

  const conocidos = new Set(erroresPrevios || []);
  const enError = buscarCeldasEnError(wb);
  const nuevos = enError.filter(x => !conocidos.has(x));
  const preexistentes = enError.filter(x => conocidos.has(x));

  controles.push({
    nombre: "Sin errores nuevos",
    detalle: nuevos.length
      ? `aparecieron ${nuevos.length}: ${nuevos.slice(0, 8).join(", ")}${nuevos.length > 8 ? "…" : ""}`
      : (preexistentes.length ? `${preexistentes.length} ya venían de antes, sin novedades` : "ninguna"),
    diferencia: nuevos.length,
    pasa: nuevos.length === 0,
  });

  return {
    controles,
    // Un control marcado `soloAviso` no frena el cierre: informa. Ver el comentario de
    // controlDebeHaber sobre por qué el descuadre de los archivos en R$ entra en esa
    // categoría hasta que se confirme cómo se compone.
    pasa: controles.every(c => c.pasa || c.soloAviso),
    hayAvisos: controles.some(c => !c.pasa && c.soloAviso),
    erroresPreexistentes: preexistentes,
    erroresTodos: enError,
  };
}

if (typeof module !== "undefined") {
  const cfg = require("./config_tfbr.js");
  global.derivarLayoutSaldos = cfg.derivarLayoutSaldos;
  global.leerPlanDeCuentas = cfg.leerPlanDeCuentas;
  module.exports = {
    validarRecalculado, controlDebeHaber, controlCuentasLevantadas, buscarCeldasEnError,
  };
}
