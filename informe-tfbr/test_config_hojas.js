// Prueba del editor multihoja, contra los 4 maestros de verdad.
//
// Lo que más importa acá es EL SIGNO. El Activo muestra "deudora menos acreedora" porque un
// activo es deudor; el Pasivo y el EERR muestran lo contrario, porque una deuda y un ingreso
// son acreedores y ahí van en positivo. Enganchar una cuenta en el Pasivo con el signo del
// Anexo II le restaría el importe en vez de sumárselo: el renglón daría de menos justo el
// doble de la cuenta, el balance seguiría cerrando, y nada lo avisaría. Por eso se prueba
// desenganchando una cuenta de cada renglón de cada hoja y volviéndola a enganchar, y
// comparando la fórmula que queda contra la que había.
//
// Correr con: node informe-tfbr/test_config_hojas.js

const path = require("path");
const ExcelJS = require("exceljs");

global.materializarFormulasCompartidas = require("./formula_utils.js").materializarFormulasCompartidas;
global.limpiarPlanDeCuentas = require("./limpieza_plan.js").limpiarPlanDeCuentas;
const cfg = require("./config_tfbr.js");
const rr = require("./rastreo_rangos.js");
const ch = require("./config_hojas.js");
const rt = require("./rotulos_anexo.js");

const MAESTROS = {
  "Mensual $": "base_bm_ars.xlsx",
  "Mensual R$": "base_bm_brl.xlsx",
  "Acumulado $": "base_ba_ars.xlsx",
  "Acumulado R$": "base_ba_brl.xlsx",
};

// El aporte de un renglón a cada fila de SALDOS, expresado sobre el saldo con signo.
//
// SALDOS parte cada cuenta en columna deudora y acreedora, y la columna cruda del VLOOKUP es
// exactamente deudora menos acreedora. Entonces "+SALDOS!D11" (cruda) y "+SALDOS!B11-SALDOS!C11"
// son lo mismo, y hay que contarlos igual o la comparación denuncia diferencias que no existen.
function aporte(formula, layout) {
  const d = cfg.ctColNumeroALetra(layout.deudorCol);
  const a = cfg.ctColNumeroALetra(layout.acreedorCol);
  const cruda = cfg.ctColNumeroALetra(layout.saldoCol);
  // Se desarman los paréntesis precedidos de "-": "-(x+y)" es "-x-y".
  let f = String(formula || "").replace(/-\s*\(([^()]*)\)/g, (todo, dentro) =>
    dentro.replace(/([+-]?)\s*(SALDOS!)/g, (t, s, x) => (s === "-" ? "+" : "-") + x));
  f = f.replace(/\+?\s*\(([^()]*)\)/g, "+$1");

  const mapa = {};   // fila -> { b, c } coeficientes de la deudora y la acreedora
  const sumar = (fila, b, c) => {
    const e = mapa[fila] || (mapa[fila] = { b: 0, c: 0 });
    e.b += b; e.c += c;
  };
  for (const m of f.matchAll(/([+-]?)\s*SALDOS!\$?([A-Z]{1,3})\$?(\d+)/g)) {
    const s = m[1] === "-" ? -1 : 1;
    const col = m[2].toUpperCase(), fila = +m[3];
    if (col === d) sumar(fila, s, 0);
    else if (col === a) sumar(fila, 0, s);
    else if (col === cruda) sumar(fila, s, -s);      // la cruda es deudora menos acreedora
  }
  for (const k of Object.keys(mapa)) if (!mapa[k].b && !mapa[k].c) delete mapa[k];
  return mapa;
}

// La dirección con la que un renglón toma una cuenta: +1 la suma, -1 la resta.
function direccion(ap, fila) {
  const e = ap[fila];
  if (!e) return 0;
  // La deudora manda; si el renglón sólo mira la acreedora, el signo es el opuesto.
  if (e.b) return e.b > 0 ? 1 : -1;
  return e.c > 0 ? -1 : 1;
}

(async () => {
  let fallas = 0, probados = 0;
  const fallo = (msg) => { console.log(`   ✗ ${msg}`); fallas++; };

  for (const [nombre, archivo] of Object.entries(MAESTROS)) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path.join(__dirname, archivo));
    materializarFormulasCompartidas(wb);
    limpiarPlanDeCuentas(wb, () => {});
    const layout = cfg.derivarLayoutSaldos(wb);
    rr.expandirRangosSaldos(wb, layout, () => {});
    const plan = ch.chPlanPorFila(wb, layout);

    console.log(`\n== ${nombre}`);
    for (const h of ch.chHojasConfigurables(wb, layout)) {
      // Una hoja donde el rótulo no identifica el renglón no se edita desde el panel: se
      // informa y se sigue. El Anexo I es el caso — una fila es un bien y sus columnas son
      // valor de origen, altas, bajas y amortización, que no son intercambiables.
      const ed = ch.chHojaEditable(wb, layout, h.hoja);
      if (!ed.editable) {
        console.log(`   ${h.hoja.padEnd(10)} no editable — ${ed.motivo}`);
        continue;
      }
      const mapa = ch.chMapaHoja(wb, layout, h.hoja);
      let enHoja = 0, malSigno = 0, ambiguos = 0;

      for (const renglon of mapa.renglones) {
        if (!renglon.rotulo || renglon.vacio) continue;
        // Una cuenta del plan que este renglón lea, y la celda donde está: en el Anexo II un
        // renglón tiene hasta tres celdas de importe y la cuenta vive en una sola.
        const victima = renglon.filasSaldos.map(f => plan[f]).find(Boolean);
        if (!victima) continue;
        const suCelda = renglon.cols.find(c => {
          const v = mapa.ws.getCell(renglon.fila, c.col).value;
          return v && typeof v === "object" && typeof v.formula === "string" &&
                 new RegExp("SALDOS!\\$?[A-Z]{1,3}\\$?" + victima.fila + "(?!\\d)").test(v.formula);
        });
        if (!suCelda) continue;
        const original = mapa.ws.getCell(renglon.fila, suCelda.col).value.formula;

        const antes = direccion(aporte(original, layout), victima.fila);

        // Se llama a la operación de verdad, sin vaciar nada antes: engancharEnHoja saca la
        // cuenta de donde esté y la vuelve a poner. Vaciando el renglón primero se probaba
        // otra cosa —el alta en un renglón vacío— y se perdía justo lo que hay que verificar,
        // que es que mover una cuenta no le cambie el signo.
        const res = ch.engancharEnHoja(wb, layout, h.hoja, victima.cod, renglon.rotulo);
        if (!res.hecho) {
          // Negarse a mover una cuenta a un rótulo que aparece dos veces es lo correcto: sin
          // más dato que el nombre, elegir uno de los dos sería adivinar.
          if (/aparece \d+ veces/.test(res.motivo)) { ambiguos++; continue; }
          fallo(`${h.hoja} "${renglon.rotulo}": ${res.motivo}`);
          continue;
        }

        const nueva = mapa.ws.getCell(res.fila, res.col).value;
        const ahora = direccion(aporte(nueva && nueva.formula, layout), victima.fila);
        probados++; enHoja++;
        // Se compara la DIRECCIÓN, no los términos: el reenganche escribe siempre el par
        // completo (deudora menos acreedora) aunque el original mirara sólo una de las dos,
        // y eso es a propósito — un gasto que queda acreedor no tiene que mostrarse en cero.
        // Lo que no puede cambiar es si la cuenta suma o resta.
        if (ahora !== antes) {
          malSigno++;
          fallo(`${h.hoja}!${renglon.cols[0].dir} "${renglon.rotulo}" — ${victima.cod} pasó de ` +
                `${antes > 0 ? "sumar" : "restar"} a ${ahora > 0 ? "sumar" : (ahora ? "restar" : "no estar")}` +
                `\n       antes: ${original}\n       después: ${(nueva && nueva.formula) || "(vacía)"}`);
        }
      }
      console.log(`   ${h.hoja.padEnd(10)} ${String(enHoja).padStart(3)} renglón(es) probado(s)` +
                  (ambiguos ? `, ${ambiguos} rótulo(s) repetido(s) que se negó a mover` : "") +
                  (malSigno ? `  ✗ ${malSigno} con el signo cambiado` : "  ✓"));
    }
  }

  console.log(fallas
    ? `\n✗ ${fallas} falla(s) sobre ${probados} renglones.`
    : `\n✓ ${probados} renglones de las 4 planillas: desenganchar y volver a enganchar deja la misma cuenta con el mismo signo.`);
  process.exit(fallas ? 1 : 0);
})();
