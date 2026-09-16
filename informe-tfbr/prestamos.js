// Los préstamos al personal: cada uno en su renglón, con el mismo nombre en los cuatro.
//
// Las seis cuentas de préstamo son 114xxxxxxx — todas de ACTIVO: es plata que la empresa
// prestó, no que debe. Cada archivo las trataba distinto, y dos las tenían en el PASIVO:
//
//   Mensual $     renglones sueltos en el Activo (Paccielo, Velasco, Alves, Cuba)
//   Mensual R$    renglones sueltos, y Carlos Furlong en el PASIVO
//   Acumulado $   todas juntas en "Adelanto al personal", y Cuba y Furlong en el PASIVO
//   Acumulado R$  todas juntas en "Prestamos al personal", y Cuba y Vera en el PASIVO
//
// Los renglones del Pasivo son el problema de fondo: los alimenta una cuenta de activo, así
// que el importe entra restando. Cuba tiene saldo deudor de 200.000 y el Pasivo mostraba
// -200.000. Se compensaba con el Activo, así que el balance cerraba y nada avisaba. Encima
// esas filas están ocultas, con lo cual el número no se ve ni abriendo el informe.
//
// Y los rótulos estaban cruzados: el renglón "Préstamo Lorena Alves" del Pasivo sumaba
// PRESTAMO CUBA, y el "Prestamo Carlos Furlong" del Acumulado R$ sumaba PRESTAMO VERA.
//
// Qué hace esto. Le da a cada préstamo su propio renglón en el Activo, con el mismo nombre en
// los cuatro archivos, creándolo donde no exista, y los saca del Pasivo. Juntarlos todos en
// "Adelanto al personal" —como estaban los Acumulados— haría que los informes no se puedan
// comparar contra los Mensuales, que los abren uno por uno.
//
// El orden importa: se crean en el orden de la lista, cada uno después del anterior, así los
// cuatro archivos terminan con los seis renglones en la misma secuencia.

const PRESTAMOS_AL_PERSONAL = [
  { cuenta: "1140600300", rotulo: "- Préstamo A Velasco" },
  { cuenta: "1140600700", rotulo: "- Préstamo Carlos Furlong" },
  { cuenta: "1140601100", rotulo: "- Préstamo Lorena Alves" },
  { cuenta: "1140601300", rotulo: "- Préstamo Cuba" },
  { cuenta: "1140601400", rotulo: "- Préstamo Vera" },
  { cuenta: "1140601500", rotulo: "- Préstamo Paccielo" },
];

const PR_HOJA = "Activo";
const PR_HOJAS_A_LIMPIAR = ["Pasivo"];

// Dónde arranca el bloque de préstamos si el archivo no tiene ninguno todavía. Son renglones
// de "otros créditos", así que van con los adelantos. Se prueban varios porque no todos los
// archivos tienen los mismos.
const PR_ANCLAS = ["- Adelanto al personal", "- Prestamos al personal", "- Adelantos PDT"];

function prNormalizarNombres(wb, layout, plan, log) {
  // Un renglón que ya lee la cuenta pero se llama distinto se renombra, en vez de crear otro:
  // si no, el archivo terminaría con "- Prestamo Velasco" vacío y "- Préstamo A Velasco" al lado.
  const mapa = chMapaHoja(wb, layout, PR_HOJA);
  if (!mapa) return [];
  const hechos = [];
  for (const p of PRESTAMOS_AL_PERSONAL) {
    const cuenta = plan[p.cuenta];
    if (!cuenta) continue;
    const filas = [cuenta].concat(cuenta.otrasFilas || []).map(f => f.fila);
    const reng = mapa.renglones.find(r => r.filasSaldos.some(f => filas.includes(f)));
    if (!reng || !reng.rotulo) continue;
    if (String(reng.rotulo).trim() === p.rotulo) continue;   // texto exacto: ver config_hojas
    // Sólo si ese renglón es de ESE préstamo y de nada más. El renglón agrupado de los
    // Acumulados ("Adelanto al personal") lee varias cuentas y no se puede renombrar.
    if (reng.filasSaldos.length > 1) continue;
    if (mapa.renglones.some(r => r !== reng && String(r.rotulo).trim() === p.rotulo)) continue;
    const celda = chCeldaDelRotulo(mapa.ws, reng.fila, reng.cols[0].col);
    if (!celda) continue;
    const antes = reng.rotulo;
    mapa.ws.getCell(celda.fila, celda.col).value = p.rotulo;
    reng.rotulo = p.rotulo;
    hechos.push({ de: antes, a: p.rotulo, fila: reng.fila });
    log(`  ${PR_HOJA} fila ${reng.fila}: "${antes}" pasa a llamarse "${p.rotulo}".`);
  }
  return hechos;
}

function prVaciarEnHoja(wb, layout, nombreHoja, filas) {
  const mapa = chMapaHoja(wb, layout, nombreHoja);
  if (!mapa) return [];
  const sacados = [];
  for (const r of mapa.renglones) {
    for (const c of r.cols) {
      for (const f of filas) {
        if (rtQuitarTermino(mapa.ws, c.dir, f)) sacados.push({ hoja: nombreHoja, fila: r.fila, rotulo: r.rotulo });
      }
    }
  }
  return sacados;
}

function consolidarPrestamos(wb, layout, log = () => {}) {
  let plan = leerPlanDeCuentas(wb, layout).cuentas;
  const presentes = PRESTAMOS_AL_PERSONAL.filter(p => plan[p.cuenta]);

  const renombrados = prNormalizarNombres(wb, layout, plan, log);

  // Fuera del Pasivo primero: si algo fallara después, no queda contado dos veces.
  const sacadasDelPasivo = [];
  for (const hoja of PR_HOJAS_A_LIMPIAR) {
    const filas = presentes.flatMap(p => [plan[p.cuenta]].concat(plan[p.cuenta].otrasFilas || []).map(f => f.fila));
    for (const s of prVaciarEnHoja(wb, layout, hoja, filas)) {
      sacadasDelPasivo.push(s);
      log(`  ${s.hoja} fila ${s.fila} ("${s.rotulo}"): le saqué un préstamo — es una cuenta de ` +
          `activo y ahí entraba restando.`);
    }
  }

  // Cada uno su renglón, en orden, encadenando el ancla para que queden seguidos.
  //
  // Se crean LOS SEIS en todos los archivos, incluso donde la cuenta todavía no está en el
  // plan de SALDOS. Los cuatro informes tienen que tener los mismos renglones: si el renglón
  // se creara sólo donde hay cuenta, el día que el sumas y saldos traiga ese préstamo el
  // archivo no tendría dónde ponerlo y habría que acordarse de agregarlo a mano.
  const creados = [];
  let anclaPrevia = null;
  for (const p of PRESTAMOS_AL_PERSONAL) {
    const anclas = anclaPrevia ? [anclaPrevia].concat(PR_ANCLAS) : PR_ANCLAS;
    const r = crearRenglon(wb, layout, PR_HOJA, p.rotulo, anclas, log);
    if (r.hecho) creados.push({ rotulo: p.rotulo, fila: r.fila });
    else if (!r.yaEstaba) log(`  ⚠ Préstamos: no pude crear "${p.rotulo}" — ${r.motivo}.`);
    anclaPrevia = p.rotulo;
    // Insertar filas mueve el plan de cuentas: hay que releerlo.
    plan = leerPlanDeCuentas(wb, layout).cuentas;
  }

  const movidas = [];
  for (const p of presentes) {
    const res = engancharEnHoja(wb, layout, PR_HOJA, p.cuenta, p.rotulo);
    if (!res.hecho) { log(`  ⚠ Préstamos: ${p.cuenta} — ${res.motivo}.`); continue; }
    movidas.push({ cod: p.cuenta, rotulo: p.rotulo, fila: res.fila });
  }

  if (movidas.length) {
    log(`  Préstamos: ${movidas.length} cuenta(s), cada una en su renglón del ${PR_HOJA}` +
        (creados.length ? `, ${creados.length} renglón(es) creado(s)` : "") + ".");
  }
  return { movidas, sacadasDelPasivo, creados, renombrados };
}

if (typeof module !== "undefined") {
  const cfg = require("./config_tfbr.js");
  global.leerPlanDeCuentas = cfg.leerPlanDeCuentas;
  const ch = require("./config_hojas.js");
  global.chMapaHoja = ch.chMapaHoja;
  global.chNorm = ch.chNorm;
  global.chCeldaDelRotulo = ch.chCeldaDelRotulo;
  global.engancharEnHoja = ch.engancharEnHoja;
  global.crearRenglon = require("./renglones_faltantes.js").crearRenglon;
  global.rtQuitarTermino = require("./rotulos_anexo.js").rtQuitarTermino;
  module.exports = { PRESTAMOS_AL_PERSONAL, consolidarPrestamos };
}
