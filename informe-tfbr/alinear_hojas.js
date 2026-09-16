// Que los cuatro informes lean las mismas cuentas en cada hoja.
//
// El problema. Cada archivo se fue cableando por su cuenta a lo largo de los años, así que la
// misma hoja nombra cuentas distintas en cada uno. El caso que lo muestra: AMX ARGENTINA está
// en "- Proveedores" del Pasivo en tres archivos y en ninguno del Mensual R$ — R$ 12,19 que
// entran en SALDOS y no llegan a ningún estado. El balance cierra igual, y no hay forma de
// notarlo mirando el informe.
//
// Por qué se compara entre archivos y no contra una tabla. Una tabla de "esta cuenta va en
// este renglón" habría que mantenerla a mano y se desactualizaría con el primer cambio. Los
// otros tres archivos ya dicen dónde va cada cuenta: si tres la ponen en "- Proveedores", el
// cuarto también. Es la misma idea que el panel de configuración, pero automática.
//
// Qué NO hace:
//  - No inventa renglones. Si el archivo no tiene el rótulo donde los otros la ponen, no la
//    engancha en cualquier lado: lo informa y queda para decidir.
//  - No toca hojas donde el rótulo no identifica el renglón (el Anexo I; ver chHojaEditable).
//  - No toca cuentas que ese archivo no tiene en su plan. No es un error: el plan de cada
//    archivo tiene las cuentas que su sumas y saldos alguna vez trajo, y el motor da de alta
//    lo que falte cuando aparezca.
//  - No decide por mayoría cuando los archivos discrepan entre sí: si una cuenta está en dos
//    renglones distintos según el archivo, no elige uno — lo informa. Adivinar ahí es mover
//    plata de un renglón del balance a otro sin que nadie lo haya aprobado.

function alinearHojas(copias, log = () => {}) {
  const vivos = (copias || []).filter(c => c && c.wb);
  if (vivos.length < 2) return { enganchadas: [], pendientes: [], discrepan: [] };

  // Estado de cada archivo: qué lee cada hoja, qué rótulos tiene, y qué cuentas hay en su plan.
  const estado = vivos.map(c => {
    const layout = derivarLayoutSaldos(c.wb);
    const plan = chPlanPorFila(c.wb, layout);
    const porCodigo = {};
    for (const x of Object.values(plan)) porCodigo[x.cod] = x;
    const hojas = {};
    for (const h of chHojasConfigurables(c.wb, layout)) {
      if (!chHojaEditable(c.wb, layout, h.hoja).editable) continue;
      // El Anexo I queda afuera: su valor de origen lo llena anexo_i.js desde las cuentas, y
      // las columnas de amortización se cargan a mano por decisión de contaduría. Alinearlo
      // contra los otros archivos le metería las cuentas de amortización acumulada en el
      // renglón del bien, que es otra columna.
      if (/anexo\s*i$/i.test(h.hoja)) continue;
      const mapa = chMapaHoja(c.wb, layout, h.hoja);
      const lee = {}, rotulos = new Set();
      for (const r of mapa.renglones) {
        if (r.rotulo) rotulos.add(chNorm(r.rotulo));
        for (const f of r.filasSaldos) {
          const info = plan[f];
          if (info) lee[info.cod] = r.rotulo;
        }
      }
      hojas[h.hoja] = { lee, rotulos };
    }
    return { ...c, layout, porCodigo, hojas };
  });

  const nombresDeHoja = [...new Set(estado.flatMap(e => Object.keys(e.hojas)))];
  const enganchadas = [], pendientes = [], discrepan = [];

  for (const hoja of nombresDeHoja) {
    const conLaHoja = estado.filter(e => e.hojas[hoja]);
    if (conLaHoja.length < 2) continue;

    const codigos = new Set();
    for (const e of conLaHoja) Object.keys(e.hojas[hoja].lee).forEach(c => codigos.add(c));

    for (const cod of codigos) {
      // ¿Todos los que la leen la ponen en el mismo renglón?
      const rotulos = new Set();
      for (const e of conLaHoja) {
        const r = e.hojas[hoja].lee[cod];
        if (r) rotulos.add(chNorm(r));
      }
      if (rotulos.size > 1) {
        const detalle = conLaHoja.filter(e => e.hojas[hoja].lee[cod])
          .map(e => `${e.label}: "${e.hojas[hoja].lee[cod]}"`).join(" ; ");
        discrepan.push({ hoja, cod, detalle });
        log(`  ⚠ ${hoja}: ${cod} está en renglones distintos según el archivo (${detalle}). ` +
            `No elijo por mi cuenta.`);
        continue;
      }
      const rotulo = conLaHoja.map(e => e.hojas[hoja].lee[cod]).find(Boolean);
      if (!rotulo) continue;

      for (const e of conLaHoja) {
        if (e.hojas[hoja].lee[cod]) continue;              // ya la lee
        if (!e.porCodigo[cod]) continue;                    // no está en su plan: nada que enganchar
        if (!e.hojas[hoja].rotulos.has(chNorm(rotulo))) {
          pendientes.push({ hoja, cod, rotulo, archivo: e.label, motivo: `no tiene el renglón "${rotulo}"` });
          continue;
        }
        const r = engancharEnHoja(e.wb, e.layout, hoja, cod, rotulo);
        if (!r.hecho) { pendientes.push({ hoja, cod, rotulo, archivo: e.label, motivo: r.motivo }); continue; }
        e.hojas[hoja].lee[cod] = rotulo;
        enganchadas.push({ hoja, cod, rotulo, archivo: e.label, nombre: e.porCodigo[cod].nom });
      }
    }
  }

  for (const x of enganchadas) {
    log(`  ${x.archivo}: ${x.cod} ${x.nombre} se enganchó en ${x.hoja} → "${x.rotulo}", ` +
        `como en los otros archivos.`);
  }
  for (const x of pendientes) {
    log(`  ⚠ ${x.archivo}: no pude poner ${x.cod} en ${x.hoja} → "${x.rotulo}" — ${x.motivo}.`);
  }
  return { enganchadas, pendientes, discrepan };
}

if (typeof module !== "undefined") {
  const cfg = require("./config_tfbr.js");
  global.derivarLayoutSaldos = cfg.derivarLayoutSaldos;
  const ch = require("./config_hojas.js");
  global.chPlanPorFila = ch.chPlanPorFila;
  global.chHojasConfigurables = ch.chHojasConfigurables;
  global.chHojaEditable = ch.chHojaEditable;
  global.chMapaHoja = ch.chMapaHoja;
  global.chNorm = ch.chNorm;
  global.engancharEnHoja = ch.engancharEnHoja;
  module.exports = { alinearHojas };
}
