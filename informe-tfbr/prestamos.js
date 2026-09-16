// Los préstamos al personal, en un solo renglón del Activo, igual en los cuatro archivos.
//
// Qué estaba pasando. Las seis cuentas de préstamo son 114xxxxxxx — todas de ACTIVO, son
// plata que la empresa prestó, no que debe. Pero cada archivo las trataba distinto:
//
//   Mensual $     cuatro renglones sueltos en el Activo (Paccielo, Velasco, Alves, Cuba)
//   Mensual R$    cuatro renglones sueltos, y Carlos Furlong en el PASIVO
//   Acumulado $   agrupadas en "Adelanto al personal", y Cuba y Furlong también en el PASIVO
//   Acumulado R$  agrupadas en "Prestamos al personal", y Cuba y Vera también en el PASIVO
//
// Los renglones del Pasivo son el problema de fondo: los alimenta una cuenta de activo, así
// que el importe entra restando. Cuba tiene saldo deudor de 200.000 y el Pasivo mostraba
// -200.000. Se compensaba con el Activo, así que el balance cerraba y nada avisaba. Encima
// esas filas están ocultas, con lo cual el número no se ve ni abriendo el informe.
//
// Y los rótulos estaban cruzados: el renglón "Préstamo Lorena Alves" del Pasivo sumaba
// PRESTAMO CUBA, y el "Prestamo Carlos Furlong" del Acumulado R$ sumaba PRESTAMO VERA.
// Renombrarlos no alcanzaba: dejaba a la otra cuenta sin renglón.
//
// Qué hace esto. Manda las seis al renglón "- Adelanto al personal" del Activo, en los cuatro,
// y las saca del Pasivo. El renglón de destino se deja visible: si queda oculto, el importe
// entra en el subtotal pero no se ve, que es como empezó todo esto.

const PRESTAMOS_AL_PERSONAL = [
  "1140600300",   // PRESTAMO A. VELASCO
  "1140600700",   // PRESTAMO CARLOS FURLONG
  "1140601100",   // PRESTAMO LORENA ALVES
  "1140601300",   // PRESTAMO CUBA
  "1140601400",   // PRESTAMO VERA
  "1140601500",   // PRESTAMO PACCIELO
];

const PR_ROTULO = "- Adelanto al personal";
const PR_HOJA_DESTINO = "Activo";
const PR_HOJAS_A_LIMPIAR = ["Pasivo"];

// El renglón que va a juntarlos. Si el archivo ya tiene uno con ese nombre, ese; si no, se
// renombra el primero que hoy lee un préstamo. No se inserta una fila nueva: los cuatro
// archivos ya tienen dónde, sólo que con nombres distintos ("Prestamos al personal" en los
// R$), y meter filas mueve todas las fórmulas de abajo sin necesidad.
function prRenglonDestino(wb, layout, plan, log) {
  const mapa = chMapaHoja(wb, layout, PR_HOJA_DESTINO);
  if (!mapa) return null;

  const yaEsta = mapa.renglones.find(r => chNorm(r.rotulo) === chNorm(PR_ROTULO));
  if (yaEsta) return { mapa, renglon: yaEsta, renombrado: null };

  const filasPrestamo = PRESTAMOS_AL_PERSONAL
    .map(c => plan[c] && plan[c].fila).filter(Boolean);
  const candidato = mapa.renglones.find(r => r.filasSaldos.some(f => filasPrestamo.includes(f)));
  if (!candidato) return null;

  const celda = chCeldaDelRotulo(mapa.ws, candidato.fila, candidato.cols[0].col);
  if (!celda) return null;
  const antes = candidato.rotulo;
  mapa.ws.getCell(celda.fila, celda.col).value = PR_ROTULO;
  candidato.rotulo = PR_ROTULO;
  log(`  ${PR_HOJA_DESTINO} fila ${candidato.fila}: "${antes}" pasa a llamarse "${PR_ROTULO}" ` +
      `para que los cuatro archivos junten los préstamos en el mismo renglón.`);
  return { mapa, renglon: candidato, renombrado: antes };
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
  const plan = leerPlanDeCuentas(wb, layout).cuentas;
  const presentes = PRESTAMOS_AL_PERSONAL.filter(c => plan[c]);
  if (!presentes.length) return { movidas: [], sacadasDelPasivo: [], destino: null };

  const destino = prRenglonDestino(wb, layout, plan, log);
  if (!destino) {
    log(`  ⚠ Préstamos: no encontré en ${PR_HOJA_DESTINO} un renglón donde juntarlos; no toqué nada.`);
    return { movidas: [], sacadasDelPasivo: [], destino: null };
  }

  // Primero se los saca del Pasivo. Va antes del enganche para que, si algo falla después,
  // no quede el importe contado dos veces.
  const sacadasDelPasivo = [];
  for (const hoja of PR_HOJAS_A_LIMPIAR) {
    const filas = presentes.flatMap(c => [plan[c]].concat(plan[c].otrasFilas || []).map(f => f.fila));
    for (const s of prVaciarEnHoja(wb, layout, hoja, filas)) {
      sacadasDelPasivo.push(s);
      log(`  ${s.hoja} fila ${s.fila} ("${s.rotulo}"): le saqué un préstamo — es una cuenta de ` +
          `activo y ahí entraba restando.`);
    }
  }

  const movidas = [];
  for (const cod of presentes) {
    const r = engancharEnHoja(wb, layout, PR_HOJA_DESTINO, cod, PR_ROTULO);
    if (!r.hecho) { log(`  ⚠ Préstamos: ${cod} — ${r.motivo}.`); continue; }
    movidas.push({ cod, fila: r.fila });
  }

  // El renglón que los junta tiene que verse. Si queda oculto, el importe entra en el
  // subtotal pero no aparece en el informe impreso.
  const fila = wb.getWorksheet(PR_HOJA_DESTINO).getRow(destino.renglon.fila);
  if (fila.hidden) { fila.hidden = false; log(`  ${PR_HOJA_DESTINO} fila ${destino.renglon.fila}: estaba oculta, la dejé visible.`); }

  if (movidas.length) {
    log(`  Préstamos: ${movidas.length} cuenta(s) al renglón "${PR_ROTULO}" del ${PR_HOJA_DESTINO}.`);
  }
  return { movidas, sacadasDelPasivo, destino: destino.renglon.fila, renombrado: destino.renombrado };
}

if (typeof module !== "undefined") {
  const cfg = require("./config_tfbr.js");
  global.leerPlanDeCuentas = cfg.leerPlanDeCuentas;
  const ch = require("./config_hojas.js");
  global.chMapaHoja = ch.chMapaHoja;
  global.chNorm = ch.chNorm;
  global.chCeldaDelRotulo = ch.chCeldaDelRotulo;
  global.engancharEnHoja = ch.engancharEnHoja;
  global.rtQuitarTermino = require("./rotulos_anexo.js").rtQuitarTermino;
  module.exports = { PRESTAMOS_AL_PERSONAL, PR_ROTULO, consolidarPrestamos };
}
