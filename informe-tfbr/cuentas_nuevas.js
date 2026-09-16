// Cuentas nuevas: a qué rótulo del Anexo II va cada una.
//
// El problema que resuelve. Cuando en el sumas y saldos aparece una cuenta de gasto que el
// maestro no tenía, el motor la da de alta en SALDOS y su importe entra en los totales de esa
// hoja. Pero el Anexo II es la apertura del gasto por centro de costo, y ahí cada renglón suma
// cuentas nombradas una por una: si nadie nombra a la cuenta nueva, su importe no llega al
// Anexo II ni al estado de resultados. El balance CIERRA IGUAL — el importe simplemente no
// está. Hasta ahora eso se avisaba en el log y había que arreglarlo a mano en Excel.
//
// Cómo se decide. Se pregunta una sola vez, y la respuesta se aplica a los cuatro archivos:
// los cuatro tienen que abrir el gasto con los mismos rótulos y la misma cuenta en cada uno.
//
// Dónde queda la decisión. En dos lugares, a propósito:
//  1. En la fórmula del Anexo II del propio archivo. Ese es el mecanismo de verdad: al mes
//     siguiente la cuenta ya está enganchada y no hay nada que preguntar.
//  2. En estado_tfbr.json, como respaldo. Si por lo que sea la cuenta vuelve a aparecer
//     suelta —se restauró un maestro viejo, se corrigió un código, alguien rehízo el Anexo—
//     se re-engancha sola, sin volver a preguntar.

function cnTexto(ws, r, c) {
  const v = ws.getCell(r, c).value;
  if (v == null) return "";
  if (typeof v === "object") return v.richText ? v.richText.map(t => t.text).join("") : "";
  return String(v);
}

function cnNorm(t) {
  return String(t || "").normalize("NFD").replace(/[̀-ͯ]/g, "")
    .toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();
}

// Qué filas de SALDOS lee cada renglón del Anexo II, y qué rótulos hay para elegir.
function cnMapaAnexo(wb) {
  const ax = wb.getWorksheet("Anexo II");
  if (!ax) return null;
  const bloque = rtUbicarBloque(ax);
  if (bloque.desde == null || bloque.hasta == null) return null;

  const leidas = new Set();          // filas de SALDOS que ya tiene alguien
  const rotulos = [];                // [{fila, rotulo, cols:[...]}]
  for (let r = bloque.desde; r <= bloque.hasta; r++) {
    const rot = cnTexto(ax, r, 2).trim();
    const cols = [];
    for (const c of [4, 5, 6]) {
      const v = ax.getCell(r, c).value;
      if (!v || typeof v !== "object" || typeof v.formula !== "string") continue;
      if (/^SUM\(/i.test(v.formula)) continue;
      let tiene = false;
      for (const m of v.formula.matchAll(/SALDOS!\$?[A-Z]{1,3}\$?(\d+)/g)) {
        leidas.add(Number(m[1])); tiene = true;
      }
      if (tiene) cols.push(c);
    }
    if (rot && !/^total|^conceptos/i.test(rot)) rotulos.push({ fila: r, rotulo: rot, cols });
  }
  return { ax, bloque, leidas, rotulos };
}

// Los rótulos que se le pueden ofrecer, sin repetidos y en orden alfabético.
function cnRotulosDisponibles(wb) {
  const mapa = cnMapaAnexo(wb);
  if (!mapa) return [];
  const vistos = new Set();
  return mapa.rotulos
    .filter(r => { const k = cnNorm(r.rotulo); if (vistos.has(k)) return false; vistos.add(k); return true; })
    .map(r => r.rotulo)
    .sort((a, b) => a.localeCompare(b, "es"));
}

// Engancha la cuenta `cod` al renglón `rotulo`, en los términos del Anexo II de ESTE archivo.
//
// Antes de engancharla se la saca de cualquier otro renglón que la venga sumando: si no, el
// importe se contaría dos veces y el Anexo II daría más que SALDOS.
//
// La columna de centro de costo (D administración, E comercialización, F financieros) no se
// pregunta: se toma la que el renglón de destino ya usa. Un rótulo que hoy suma solo en
// comercialización es un concepto de comercialización, y poner la cuenta nueva en otra columna
// abriría una apertura que ese renglón nunca tuvo. Si el renglón está vacío (un rótulo
// disponible, todavía sin cuentas) se usa D.
function engancharCuentaEnRotulo(wb, cod, rotulo) {
  const layout = derivarLayoutSaldos(wb);
  const plan = leerPlanDeCuentas(wb, layout).cuentas;
  const cuenta = plan[cod];
  if (!cuenta) return { hecho: false, motivo: `${cod} no está en el plan de cuentas de este archivo` };

  const mapa = cnMapaAnexo(wb);
  if (!mapa) return { hecho: false, motivo: "este archivo no tiene un Anexo II legible" };

  const destino = mapa.rotulos.find(r => cnNorm(r.rotulo) === cnNorm(rotulo));
  if (!destino) return { hecho: false, motivo: `el rótulo "${rotulo}" no existe en este archivo` };

  for (let r = mapa.bloque.desde; r <= mapa.bloque.hasta; r++) {
    for (const c of [4, 5, 6]) {
      rtQuitarTermino(mapa.ax, `${String.fromCharCode(64 + c)}${r}`, cuenta.fila);
    }
  }

  const col = destino.cols.length ? destino.cols[0] : 4;
  const termino = ctFormulaNetaAnexo(layout, cuenta.fila);   // deudora menos acreedora
  const celda = mapa.ax.getCell(destino.fila, col);
  const v = celda.value;
  const previo = (v && typeof v === "object" && typeof v.formula === "string") ? v.formula : "";
  celda.value = { formula: previo ? `${previo}${termino}` : termino };
  return { hecho: true, fila: destino.fila, col, rotulo: destino.rotulo };
}

// Las cuentas de gasto que están en SALDOS y no las lee ningún renglón del Anexo II.
//
// Se miran SOLO las que discriminan por centro de costo según el plan oficial: el Anexo II
// abre el gasto, y una cuenta de activo, pasivo o ingreso no va ahí. Preguntar por ellas
// haría parecer que falta configurar algo que no falta.
function cuentasSinRotulo(wb) {
  const mapa = cnMapaAnexo(wb);
  if (!mapa) return [];
  const layout = derivarLayoutSaldos(wb);
  const S = wb.getWorksheet(layout.sheet);
  const sueltas = [];
  for (let r = layout.planDeCuentas.desde; r <= layout.planDeCuentas.hasta; r++) {
    const m = /^\s*([\d.]+)\s*-?\s*(.*)$/.exec(cnTexto(S, r, layout.keyCol).trim());
    if (!m) continue;
    const cod = m[1].replace(/\./g, "");
    if (!/^42/.test(cod)) continue;
    if (!PLAN_CENTROS_COSTO.has(cod)) continue;
    if (mapa.leidas.has(r)) continue;
    sueltas.push({ cod, nom: m[2].trim(), fila: r });
  }
  return sueltas;
}

// Aplica las decisiones ya tomadas en meses anteriores. Devuelve las que enganchó, para el log.
// Lo que no pudo engancharse no se fuerza: vuelve a la lista de preguntas.
function aplicarRotulosGuardados(wb, guardados, log = () => {}) {
  const hechos = [];
  if (!guardados) return hechos;
  for (const suelta of cuentasSinRotulo(wb)) {
    const rotulo = guardados[suelta.cod];
    if (!rotulo) continue;
    const r = engancharCuentaEnRotulo(wb, suelta.cod, rotulo);
    if (!r.hecho) continue;
    hechos.push({ cod: suelta.cod, nom: suelta.nom, rotulo: r.rotulo });
    log(`  ${suelta.cod} ${suelta.nom}: se enganchó a "${r.rotulo}" con la decisión guardada.`);
  }
  return hechos;
}

if (typeof module !== "undefined") {
  const cfg = require("./config_tfbr.js");
  global.derivarLayoutSaldos = cfg.derivarLayoutSaldos;
  global.leerPlanDeCuentas = cfg.leerPlanDeCuentas;
  global.ctFormulaNetaAnexo = cfg.ctFormulaNetaAnexo;
  const rt = require("./rotulos_anexo.js");
  global.rtUbicarBloque = rt.rtUbicarBloque;
  global.rtQuitarTermino = rt.rtQuitarTermino;
  global.PLAN_CENTROS_COSTO = require("./plan_oficial.js").PLAN_CENTROS_COSTO;
  module.exports = {
    cnMapaAnexo, cnRotulosDisponibles, engancharCuentaEnRotulo,
    cuentasSinRotulo, aplicarRotulosGuardados,
  };
}
