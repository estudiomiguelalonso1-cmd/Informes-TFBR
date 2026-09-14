// Pone el plan de cuentas de la hoja SALDOS de acuerdo con el plan oficial del sistema
// contable (plan_oficial.js).
//
// El problema que resuelve: los códigos de SALDOS se tipearon a mano y decenas quedaron con
// UN DÍGITO DE MENOS — 9 en vez de 10, siempre un cero. "422040000 CUSTODIA" donde el plan
// dice "4220400000 CUSTODIA". Cuando después alguien dio de alta la cuenta con el código
// correcto, la fila vieja se quedó ahí: de ahí salen los códigos repetidos.
//
// Una cuenta con el código mal NUNCA levanta importe, porque el motor empareja por código
// contra lo que manda Onvio. El importe entra en cero y el balance cierra igual.
//
// Se corrige en dos pasos, en este orden:
//   1. arreglarCodigos     — al código que no está en el plan se le prueba agregar un cero en
//                            cada posición; si alguna combinación existe en el plan Y el
//                            nombre coincide, ese es el código bueno.
//   2. fusionarDuplicados  — si al corregir quedaron dos filas con el mismo código, se juntan:
//                            lo que referenciaba a la que sobra se repunta a la que queda, y
//                            recién ahí se borra.
//
// Nada se hace por criterio: cada cambio se valida contra el plan oficial por código + nombre.

function lpNormaliza(t) {
  return String(t || "").normalize("NFD").replace(/[̀-ͯ]/g, "")
    .toUpperCase()
    .replace(/[.,](?=\d)/g, "")        // "LEY 25,413" y "LEY 25413" son lo mismo
    .replace(/[^A-Z0-9]+/g, " ").trim();
}

// Compara dos nombres de cuenta tolerando lo que varía al tipearlos: singular/plural
// ("SINIESTRO" / "SINIESTROS"), y las preposiciones ("GASTOS DE CAPACITACION" /
// "GASTOS CAPACITACIÓN"). No tolera palabras distintas: "GASTOS TELEFÓNICOS" y
// "GASTOS EN EQ. TELEFÓNICOS" NO son lo mismo para esta función, y tienen que quedar
// a la vista en vez de fusionarse solos.
const LP_VACIAS = ["DE", "DEL", "Y", "LA", "EL", "LOS", "LAS", "A", "EN"];
function lpPalabras(t) {
  return new Set(lpNormaliza(t).split(" ")
    .filter(w => w && !LP_VACIAS.includes(w))
    .map(w => w.replace(/(ES|S)$/, "")));
}
function lpMismoNombre(a, b) {
  const A = lpPalabras(a), B = lpPalabras(b);
  if (A.size !== B.size || !A.size) return false;
  for (const w of A) if (!B.has(w)) return false;
  return true;
}

// El código que el plan le da a una cuenta con este nombre. Solo sirve si hay UNO.
function lpCodigoPorNombre(nombre, plan) {
  const hits = [];
  for (const [cod, nom] of Object.entries(plan)) if (lpMismoNombre(nom, nombre)) hits.push(cod);
  return hits.length === 1 ? hits[0] : null;
}

function lpNombreDe(texto) {
  const m = /^\s*([\d.]+)\s*-?\s*(.*)$/.exec(String(texto).trim());
  return m ? m[2].trim() : String(texto).trim();
}
function lpCodigoDe(texto) {
  const m = /^\s*([\d.]+)/.exec(String(texto).trim());
  return m ? m[1].replace(/\./g, "") : null;
}

// El código correcto para una fila cuyo código no está en el plan: se prueba insertar un cero
// en cada posición y se exige que el NOMBRE coincida con el del plan. Sin esa segunda
// condición, un código corto podría "arreglarse" hacia una cuenta que no tiene nada que ver.
function lpCodigoCorregido(codigo, nombre, plan) {
  if (!codigo || plan[codigo]) return null;
  // Se deduplican: donde el código ya trae ceros, insertar el cero en varias posiciones
  // distintas da exactamente el mismo resultado ("422040000" -> "4220400000" desde la
  // posición 5 a la 9). Eso es UN candidato, no cinco.
  const candidatos = new Set();
  for (let i = 1; i <= codigo.length; i++) {
    const c = codigo.slice(0, i) + "0" + codigo.slice(i);
    if (plan[c] && lpMismoNombre(plan[c], nombre)) candidatos.add(c);
  }
  // Si quedan dos códigos distintos no se elige: se avisa. No vale adivinar con el plan.
  return candidatos.size === 1 ? [...candidatos][0] : null;
}

// Paso 1. Devuelve { corregidos, sinExplicar } y deja el workbook con los códigos arreglados.
function arreglarCodigos(wb, layout, plan, log = () => {}) {
  const ws = wb.getWorksheet(layout.sheet);
  const corregidos = [], sinExplicar = [];
  for (let r = layout.planDeCuentas.desde; r <= layout.planDeCuentas.hasta; r++) {
    const texto = String(ws.getCell(r, layout.keyCol).value || "").trim();
    if (!texto) continue;
    const cod = lpCodigoDe(texto);
    if (!cod || plan[cod]) continue;
    const nom = lpNombreDe(texto);
    const bueno = lpCodigoCorregido(cod, nom, plan);
    if (!bueno) { sinExplicar.push({ fila: r, codigo: cod, nombre: nom }); continue; }
    const nuevo = texto.replace(/^\s*[\d.]+/, bueno);
    ws.getCell(r, layout.keyCol).value = nuevo;
    corregidos.push({ fila: r, de: cod, a: bueno, nombre: nom });
    log(`  ${layout.sheet}!${r}: ${cod} → ${bueno}  (${nom})`);
  }
  return { corregidos, sinExplicar };
}

// Paso 1b. Una fila cuyo código SÍ está en el plan pero con OTRO nombre: el código es de otra
// cuenta. Si el plan tiene un código para el nombre de esta fila, y es uno solo, se lo pone.
// Es lo que destraba los códigos compartidos por dos cuentas reales: "4211100000 SERVICIOS DE
// LIMPIEZA" convive con "4211100000 REDONDEO" porque el plan dice que 4211100000 es REDONDEO
// y que limpieza es 4211600000.
function reasignarPorNombre(wb, layout, plan, log = () => {}) {
  const ws = wb.getWorksheet(layout.sheet);
  const reasignados = [], ambiguos = [];
  for (let r = layout.planDeCuentas.desde; r <= layout.planDeCuentas.hasta; r++) {
    const texto = String(ws.getCell(r, layout.keyCol).value || "").trim();
    if (!texto) continue;
    const cod = lpCodigoDe(texto);
    if (!cod || !plan[cod]) continue;
    const nom = lpNombreDe(texto);
    if (lpMismoNombre(plan[cod], nom)) continue;          // el nombre es el que corresponde
    const bueno = lpCodigoPorNombre(nom, plan);
    if (!bueno || bueno === cod) {
      ambiguos.push({ fila: r, codigo: cod, nombre: nom, oficial: plan[cod] });
      continue;
    }
    ws.getCell(r, layout.keyCol).value = texto.replace(/^\s*[\d.]+/, bueno);
    reasignados.push({ fila: r, de: cod, a: bueno, nombre: nom, ocupadoPor: plan[cod] });
    log(`  ${layout.sheet}!${r}: ${cod} → ${bueno}  ("${nom}"; ${cod} es "${plan[cod]}")`);
  }
  return { reasignados, ambiguos };
}

// Paso 2. Junta las filas que quedaron con el mismo código Y el mismo nombre: son la misma
// cuenta cargada dos veces. Las que comparten código pero tienen nombres distintos NO se
// tocan: ahí hay dos cuentas reales y cuál sobra es una decisión contable.
function fusionarDuplicados(wb, layout, planDeCuentas, log = () => {}) {
  const fusionadas = [], trabadas = [];
  const grupos = {};
  for (const [codigo, info] of Object.entries(planDeCuentas)) {
    const filas = [info].concat(info.otrasFilas || []);
    if (filas.length < 2) continue;
    grupos[codigo] = filas;
  }
  // De abajo hacia arriba: cada borrado corre las filas de abajo.
  const ordenados = Object.entries(grupos)
    .map(([codigo, filas]) => ({ codigo, filas: filas.slice().sort((a, b) => a.fila - b.fila) }))
    .sort((a, b) => b.filas[b.filas.length - 1].fila - a.filas[a.filas.length - 1].fila);

  for (const { codigo, filas } of ordenados) {
    // Se fusionan solo si TODAS coinciden entre sí. Cuando el plan oficial tiene ese código,
    // además tienen que coincidir con el nombre oficial: si una no coincide, esa fila es otra
    // cuenta mal codificada y la decide el paso 1b, no una fusión.
    const base = lpNombreDe(filas[0].texto);
    const todasIguales = filas.every(f => lpMismoNombre(base, lpNombreDe(f.texto)));
    const oficial = typeof PLAN_OFICIAL !== "undefined" ? PLAN_OFICIAL[codigo] : null;
    if (!todasIguales || (oficial && !lpMismoNombre(oficial, base))) {
      trabadas.push({ codigo, motivo: oficial
          ? `el plan dice que ${codigo} es "${oficial}"`
          : "el mismo código lo usan cuentas con nombres distintos",
        filas: filas.map(f => ({ fila: f.fila, texto: f.texto })) });
      continue;
    }
    // se queda la primera; las demás se repuntan hacia ella y se borran
    const queda = filas[0];
    for (let i = filas.length - 1; i >= 1; i--) {
      const sobra = filas[i];
      const movidas = repuntarGemela(wb, layout.sheet, sobra.fila, queda.fila, () => {});
      const culpables = quienReferenciaLaFila(wb, layout.sheet, sobra.fila);
      if (culpables.length) {
        trabadas.push({ codigo, motivo: `todavía la referencian ${culpables.length} fórmula(s)`,
          filas: [{ fila: sobra.fila, texto: sobra.texto }] });
        continue;
      }
      borrarFilaEn(wb, layout.sheet, sobra.fila);
      fusionadas.push({ codigo, nombre: lpNombreDe(sobra.texto), filaBorrada: sobra.fila,
        filaQueQueda: queda.fila, referenciasMovidas: movidas.length || 0 });
      log(`  ${codigo} "${lpNombreDe(sobra.texto)}": fila ${sobra.fila} fusionada en la ${queda.fila}` +
          (movidas.length ? ` (${movidas.length} referencia(s) repuntadas)` : ""));
    }
  }
  return { fusionadas, trabadas };
}

// Corre los dos pasos. Devuelve todo lo que hizo y lo que no pudo, para mostrarlo en pantalla.
function limpiarPlanDeCuentas(wb, log = () => {}) {
  let layout = derivarLayoutSaldos(wb);
  const paso1 = arreglarCodigos(wb, layout, PLAN_OFICIAL, log);

  layout = derivarLayoutSaldos(wb);
  const paso1b = reasignarPorNombre(wb, layout, PLAN_OFICIAL, log);

  // los códigos cambiaron: hay que releer el plan antes de buscar duplicados
  layout = derivarLayoutSaldos(wb);
  let { cuentas } = leerPlanDeCuentas(wb, layout);
  const paso2 = fusionarDuplicados(wb, layout, cuentas, log);

  if (paso2.fusionadas.length) {
    layout = derivarLayoutSaldos(wb);
    ({ cuentas } = leerPlanDeCuentas(wb, layout));
  }

  // control final: cuentas del plan oficial que llevan centro de costo y no están en SALDOS
  const faltantes = [];
  for (const cod of PLAN_CENTROS_COSTO) if (!cuentas[cod]) faltantes.push(cod);

  return { ...paso1, ...paso1b, ...paso2, faltantesConCentroCosto: faltantes, layout, cuentas };
}

if (typeof module !== "undefined") {
  const cfg = require("./config_tfbr.js");
  const fh = require("./formula_hojas.js");
  const ic = require("./insertar_cuenta.js");
  const po = require("./plan_oficial.js");
  global.derivarLayoutSaldos = cfg.derivarLayoutSaldos;
  global.leerPlanDeCuentas = cfg.leerPlanDeCuentas;
  global.borrarFilaEn = fh.borrarFilaEn;
  global.quienReferenciaLaFila = fh.quienReferenciaLaFila;
  global.repuntarGemela = ic.repuntarGemela;
  global.PLAN_OFICIAL = po.PLAN_OFICIAL;
  global.PLAN_CENTROS_COSTO = po.PLAN_CENTROS_COSTO;
  module.exports = {
    limpiarPlanDeCuentas, arreglarCodigos, reasignarPorNombre, fusionarDuplicados,
    lpCodigoCorregido, lpCodigoPorNombre, lpMismoNombre,
    lpNombreDe, lpCodigoDe, lpNormaliza,
  };
}
