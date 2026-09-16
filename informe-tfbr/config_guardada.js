// Las decisiones de configuración, guardadas fuera del código.
//
// Por qué existe. Hasta acá, cada decisión —a qué renglón va una cuenta, cómo se llama un
// renglón, qué renglón hay que crear— vivía en una tabla dentro de un archivo .js. Eso
// significa que para cambiar cualquier cosa hay que editar el código, y entonces el sistema no
// es autónomo: la persona que lo usa no puede decidir nada sin que alguien toque el programa.
//
// Ahora las decisiones viven en estado_tfbr.json, junto al historial, y el motor las aplica en
// cada corrida. El panel las escribe. El código queda con el CÓMO (cómo se engancha una cuenta,
// cómo se crea un renglón) y los datos con el QUÉ.
//
// Forma del objeto:
//
//   {
//     "cuentas": {
//       "4222700000": { "hoja": "Anexo II", "rotulo": "Comisiones" },
//       "1140601500": { "excluida": true }
//     },
//     "renglones": [ { "hoja": "Activo", "rotulo": "- Préstamo Vera", "despuesDe": "- Adelantos PDT" } ],
//     "renombres": [ { "hoja": "Activo", "cuenta": "1160100000", "a": "- Fondo común de inversión BBVA" } ]
//   }
//
// Qué significa "excluida": que esa cuenta NO va a ningún renglón y que no hay que volver a
// preguntar por ella. No quiere decir que no se pegue — si el sumas y saldos la trae, se pega
// igual, porque si no el Debe dejaría de dar igual que el Haber. Quiere decir que quedó así a
// propósito, y el control de cuentas sin destino no la denuncia.

const CG_VACIA = { cuentas: {}, renglones: [], renombres: [] };

function cgNormalizar(bruto) {
  const c = bruto && typeof bruto === "object" ? bruto : {};
  return {
    cuentas: (c.cuentas && typeof c.cuentas === "object") ? c.cuentas : {},
    renglones: Array.isArray(c.renglones) ? c.renglones : [],
    renombres: Array.isArray(c.renombres) ? c.renombres : [],
  };
}

// Lo que el motor tiene que aplicar, en el orden en que tiene que aplicarlo.
//
// Crear renglones primero: una cuenta no se puede mandar a un renglón que todavía no existe.
// Renombrar después, porque un renombre puede ser justamente el que le da a un renglón el
// nombre con el que las cuentas lo van a buscar. Y las cuentas al final.
function aplicarConfiguracion(wb, layout0, config, log = () => {}) {
  const cfg = cgNormalizar(config);
  let layout = layout0;
  const creados = [], renombrados = [], asignadas = [], salteadas = [];

  for (const r of cfg.renglones) {
    if (!r || !r.hoja || !r.rotulo) continue;
    const res = crearRenglon(wb, layout, r.hoja, r.rotulo, r.despuesDe || [], log);
    if (res.yaEstaba) continue;
    if (!res.hecho) { salteadas.push({ que: `renglón "${r.rotulo}" en ${r.hoja}`, motivo: res.motivo }); continue; }
    creados.push({ hoja: r.hoja, rotulo: r.rotulo, fila: res.fila });
    // Insertar filas mueve todo lo de abajo: el layout hay que volver a derivarlo.
    layout = derivarLayoutSaldos(wb);
  }

  for (const r of cfg.renombres) {
    if (!r || !r.hoja || !r.cuenta || !r.a) continue;
    const res = cgRenombrarRenglonDe(wb, layout, r.hoja, r.cuenta, r.a, log);
    if (res.hecho) renombrados.push(res);
    else if (res.motivo) salteadas.push({ que: `renombre de ${r.cuenta} en ${r.hoja}`, motivo: res.motivo });
  }

  const plan = leerPlanDeCuentas(wb, layout).cuentas;
  for (const [cod, decision] of Object.entries(cfg.cuentas)) {
    if (!decision || decision.excluida) continue;
    if (!decision.hoja || !decision.rotulo) continue;
    if (!plan[cod]) continue;                 // no está en este archivo: el alta la crea cuando venga
    const res = engancharEnHoja(wb, layout, decision.hoja, cod, decision.rotulo);
    if (!res.hecho) { salteadas.push({ que: `${cod} → "${decision.rotulo}" en ${decision.hoja}`, motivo: res.motivo }); continue; }
    asignadas.push({ cod, hoja: decision.hoja, rotulo: decision.rotulo, fila: res.fila });
  }

  if (creados.length) log(`  Configuración: ${creados.length} renglón(es) creado(s).`);
  if (renombrados.length) log(`  Configuración: ${renombrados.length} renglón(es) renombrado(s).`);
  if (asignadas.length) log(`  Configuración: ${asignadas.length} cuenta(s) enganchada(s) según lo configurado.`);
  for (const s of salteadas) log(`  ⚠ Configuración: ${s.que} — ${s.motivo}.`);

  return { creados, renombrados, asignadas, salteadas, layout };
}

// Renombra el renglón que lee esa cuenta. Se busca por la cuenta y no por el nombre viejo
// porque el nombre es justamente lo que cambia de un archivo a otro.
function cgRenombrarRenglonDe(wb, layout, hoja, cuentaCod, nombreNuevo, log = () => {}) {
  const plan = leerPlanDeCuentas(wb, layout).cuentas;
  const cuenta = plan[cuentaCod];
  if (!cuenta) return {};                      // este archivo no tiene esa cuenta
  const mapa = chMapaHoja(wb, layout, hoja);
  if (!mapa) return {};

  const filas = [cuenta].concat(cuenta.otrasFilas || []).map(f => f.fila);
  const destino = mapa.renglones.find(x => x.filasSaldos.some(f => filas.includes(f)));
  if (!destino) return {};
  if (String(destino.rotulo).trim() === nombreNuevo) return {};   // ya se llama así

  // Un renglón que agrupa varias cuentas no lleva el nombre de una sola: escondería a las otras.
  if (destino.filasSaldos.length > 1) {
    return { motivo: `el renglón agrupa ${destino.filasSaldos.length} cuentas y no puede llamarse como una` };
  }
  if (mapa.renglones.some(x => x !== destino && String(x.rotulo).trim() === nombreNuevo)) {
    return { motivo: `"${nombreNuevo}" ya existe en esa hoja` };
  }

  const celda = chCeldaDelRotulo(mapa.ws, destino.fila, destino.cols[0].col);
  if (!celda) return { motivo: "no ubiqué la celda del rótulo" };
  const antes = destino.rotulo;
  mapa.ws.getCell(celda.fila, celda.col).value = nombreNuevo;
  log(`  ${hoja} fila ${destino.fila}: "${antes}" pasa a llamarse "${nombreNuevo}".`);
  return { hecho: true, hoja, fila: destino.fila, de: antes, a: nombreNuevo };
}

// Las cuentas que el usuario marcó como que no van a ningún renglón. El control de cuentas sin
// destino las saltea: quedaron así a propósito y avisar todos los meses enseña a ignorar el aviso.
function cuentasExcluidas(config) {
  const cfg = cgNormalizar(config);
  return new Set(Object.entries(cfg.cuentas)
    .filter(([, d]) => d && d.excluida)
    .map(([cod]) => cod));
}

if (typeof module !== "undefined") {
  const cfg = require("./config_tfbr.js");
  global.derivarLayoutSaldos = cfg.derivarLayoutSaldos;
  global.leerPlanDeCuentas = cfg.leerPlanDeCuentas;
  const ch = require("./config_hojas.js");
  global.chMapaHoja = ch.chMapaHoja;
  global.chCeldaDelRotulo = ch.chCeldaDelRotulo;
  global.engancharEnHoja = ch.engancharEnHoja;
  global.crearRenglon = require("./renglones_faltantes.js").crearRenglon;
  module.exports = {
    CG_VACIA, cgNormalizar, aplicarConfiguracion, cgRenombrarRenglonDe, cuentasExcluidas,
  };
}
