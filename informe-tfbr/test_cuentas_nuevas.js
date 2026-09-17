// Prueba del enganche de cuentas nuevas, contra los 4 maestros de verdad.
//
// Qué prueba, y por qué esto y no otra cosa. Una cuenta nueva sin enganchar no rompe nada
// visible: el balance cierra igual y el importe simplemente no está en el estado de
// resultados. Los dos errores posibles al engancharla son igual de silenciosos —
// engancharla dos veces (el gasto se cuenta doble) o engancharla mal (el importe queda en
// cero). Ninguno tira un error; hay que ir a buscarlos.
//
// Cómo. Se toma una cuenta que HOY está bien enganchada, se la desengancha para simular una
// cuenta nueva, y se comprueba que (1) el sistema la detecta como suelta, (2) al engancharla
// de vuelta no aparecen dobles lecturas, (3) engancharla otra vez no la duplica, y (4) la
// fórmula que queda lee la columna deudora Y la acreedora — un gasto que en el mes quedó
// acreedor vive entero en la segunda, y leer solo la primera lo mostraría en cero.
//
// Correr con: node informe-tfbr/test_cuentas_nuevas.js

const path = require("path");
const ExcelJS = require("exceljs");

global.materializarFormulasCompartidas = require("./formula_utils.js").materializarFormulasCompartidas;
global.limpiarPlanDeCuentas = require("./limpieza_plan.js").limpiarPlanDeCuentas;
global.aplicarRepuntesAnexo = require("./repuntes_anexo.js").aplicarRepuntesAnexo;
global.unificarRotulosAnexo = require("./rotulos_unificados.js").unificarRotulosAnexo;
const cfg = require("./config_tfbr.js");
const cn = require("./cuentas_nuevas.js");
const rt = require("./rotulos_anexo.js");
const { PLAN_CENTROS_COSTO } = require("./plan_oficial.js");

const MAESTROS = {
  "Mensual $": "base_bm_ars.xlsx",
  "Mensual R$": "base_bm_brl.xlsx",
  "Acumulado $": "base_ba_ars.xlsx",
  "Acumulado R$": "base_ba_brl.xlsx",
};

// Cada referencia del Anexo II a una celda de SALDOS, contada. La clave es la CELDA, no la
// fila: SALDOS parte cada cuenta en columna deudora y acreedora, y un renglón que lee las dos
// no está contando dos veces — está tomando el importe con signo.
function referencias(wb) {
  const ax = wb.getWorksheet("Anexo II");
  const mapa = cn.cnMapaAnexo(wb);
  const cuenta = {};
  for (let r = mapa.bloque.desde; r <= mapa.bloque.hasta; r++) {
    for (const col of [4, 5, 6]) {
      const v = ax.getCell(r, col).value;
      if (!v || typeof v !== "object" || typeof v.formula !== "string") continue;
      if (/^SUM\(/i.test(v.formula)) continue;
      for (const m of v.formula.matchAll(/SALDOS!(\$?[A-Z]{1,3}\$?\d+)/g)) {
        const k = m[1].replace(/\$/g, "");
        cuenta[k] = (cuenta[k] || 0) + 1;
      }
    }
  }
  return cuenta;
}
const dobles = (refs) => Object.entries(refs).filter(([, n]) => n > 1);
const total = (refs) => Object.values(refs).reduce((a, b) => a + b, 0);

function texto(ws, r, c) {
  const v = ws.getCell(r, c).value;
  if (v == null) return "";
  if (typeof v === "object") return v.richText ? v.richText.map(t => t.text).join("") : "";
  return String(v);
}

// Deja el maestro como lo deja el motor antes de tocar el Anexo II.
async function prepararMaestro(archivo) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path.join(__dirname, archivo));
  materializarFormulasCompartidas(wb);
  limpiarPlanDeCuentas(wb, () => {});
  const layout = cfg.derivarLayoutSaldos(wb);
  const plan = cfg.leerPlanDeCuentas(wb, layout).cuentas;
  aplicarRepuntesAnexo(wb, null, plan, () => {});
  unificarRotulosAnexo(wb, layout, plan, () => {});
  return { wb, layout };
}

function rotuloQueLee(wb, fila) {
  const ax = wb.getWorksheet("Anexo II");
  const mapa = cn.cnMapaAnexo(wb);
  const re = new RegExp("SALDOS!\\$?[A-Z]{1,3}\\$?" + fila + "(?!\\d)");
  for (const r of mapa.rotulos) {
    for (const col of [4, 5, 6]) {
      const v = ax.getCell(r.fila, col).value;
      if (v && typeof v === "object" && typeof v.formula === "string" && re.test(v.formula)) {
        return r.rotulo;
      }
    }
  }
  return null;
}

(async () => {
  let fallas = 0;
  const fallo = (msg) => { console.log(`   ✗ ${msg}`); fallas++; };

  for (const [nombre, archivo] of Object.entries(MAESTROS)) {
    const { wb, layout } = await prepararMaestro(archivo);
    const ax = wb.getWorksheet("Anexo II");
    const S = wb.getWorksheet(layout.sheet);
    const base = referencias(wb);

    console.log(`\n== ${nombre}: ${total(base)} referencias, ${dobles(base).length} doble(s), ` +
                `${cn.cuentasSinRotulo(wb).length} sin rótulo`);

    // Una cuenta de gasto por centro de costo que hoy SÍ está enganchada.
    const mapa = cn.cnMapaAnexo(wb);
    let victima = null;
    for (let r = layout.planDeCuentas.desde; r <= layout.planDeCuentas.hasta && !victima; r++) {
      const m = /^\s*([\d.]+)\s*-?\s*(.*)$/.exec(texto(S, r, layout.keyCol).trim());
      if (!m || !mapa.leidas.has(r)) continue;
      const cod = m[1].replace(/\./g, "");
      if (/^42/.test(cod) && PLAN_CENTROS_COSTO.has(cod)) victima = { cod, nom: m[2].trim(), fila: r };
    }
    if (!victima) { fallo("no encontré ninguna cuenta enganchada para la prueba"); continue; }

    const rotulo = rotuloQueLee(wb, victima.fila);

    // Simular la cuenta nueva: sacarla de donde está.
    for (let r = mapa.bloque.desde; r <= mapa.bloque.hasta; r++) {
      for (const col of [4, 5, 6]) {
        rt.rtQuitarTermino(ax, `${String.fromCharCode(64 + col)}${r}`, victima.fila);
      }
    }

    if (!cn.cuentasSinRotulo(wb).some(s => s.cod === victima.cod)) {
      fallo(`${victima.cod} quedó sin ningún renglón que la lea y NO la detectó como suelta`);
      continue;
    }

    const res = cn.engancharCuentaEnRotulo(wb, victima.cod, rotulo);
    if (!res.hecho) { fallo(`no pude engancharla: ${res.motivo}`); continue; }

    const despues = referencias(wb);
    if (dobles(despues).length > dobles(base).length) {
      fallo(`engancharla creó doble lectura: ${dobles(despues).map(([k]) => k).join(", ")}`);
    }
    if (cn.cuentasSinRotulo(wb).length) {
      fallo(`después de engancharla siguen quedando ${cn.cuentasSinRotulo(wb).length} sueltas`);
    }

    // Engancharla de nuevo al mismo rótulo no tiene que sumar el importe dos veces.
    cn.engancharCuentaEnRotulo(wb, victima.cod, rotulo);
    if (total(referencias(wb)) !== total(despues)) {
      fallo("engancharla dos veces al mismo rótulo la duplicó");
    }

    // La fórmula tiene que leer la columna deudora Y la acreedora.
    const esperada = cfg.ctFormulaNetaAnexo(layout, victima.fila);
    const formula = (() => {
      const re = new RegExp("SALDOS!\\$?[A-Z]{1,3}\\$?" + victima.fila + "(?!\\d)");
      const m2 = cn.cnMapaAnexo(wb);
      for (const r of m2.rotulos) {
        for (const col of [4, 5, 6]) {
          const v = ax.getCell(r.fila, col).value;
          if (v && typeof v === "object" && typeof v.formula === "string" && re.test(v.formula)) return v.formula;
        }
      }
      return null;
    })();
    if (!formula || !formula.includes(esperada)) {
      fallo(`la fórmula quedó "${formula}" y esperaba que contuviera "${esperada}"`);
    }

    if (!fallas) {
      console.log(`   ✓ ${victima.cod} ${victima.nom}: detectada suelta, reenganchada a ` +
                  `"${rotulo}" con ${esperada}, sin duplicar`);
    }
  }

  console.log(fallas ? `\n✗ ${fallas} falla(s).` : "\n✓ Las 4 pasan.");
  process.exit(fallas ? 1 : 0);
})();
