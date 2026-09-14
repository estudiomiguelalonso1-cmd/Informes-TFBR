// Las decisiones de contaduría sobre los códigos usados por dos cuentas distintas
// (documento "Cuentas TFBR", septiembre 2026: en rojo, la fila que se elimina).
//
// Se identifican por CÓDIGO + NOMBRE, nunca por número de fila. El documento vino con las
// filas del archivo de julio, pero las filas cambian entre un mes y otro —y vuelven a cambiar
// cada vez que se da de alta una cuenta— así que aplicarlas por número borraría otra cosa.
//
// Aplicarlas es idempotente: si la fila ya no está, no hay nada que hacer.

// Dos formas de resolver un código repetido:
//   borrarNombre    -> la fila sobra, se elimina (si nadie la referencia).
//   recodificarNombre + codigoNuevo -> la fila es una cuenta legítima que quedó cargada con
//                      el código de otra. No se borra: se le corrige el código.
const DECISIONES_DUPLICADOS = [
  // ADELANTO VIAJE no es una cuenta repetida: es una cuenta con el código mal tipeado. Los dos
  // archivos en pesos coinciden en que su código es 4226000000 (y que 4223600000 es DEUDORES
  // INCOBRABLES y 4225000000 es CESIÓN DE DERECHOS); en los dos archivos en reales se cargó con
  // el código de la cuenta vecina, y por eso choca. Borrarla —como decía la lectura inicial del
  // documento— dejaría a los archivos en reales sin el concepto "Adelanto Viaje" del Anexo II y
  // sin ninguna fila para 4226000000, así que el mes que Onvio mande movimiento ahí entraría
  // como cuenta sin mapear. Corregir el código resuelve el duplicado sin perder nada, y de paso
  // arregla un error activo: hoy 4223600000 resuelve a la fila de ADELANTO VIAJE (la primera de
  // las dos), así que un importe de DEUDORES INCOBRABLES se reportaría en la línea equivocada.
  { archivo: "balance_mensual_brl",   codigo: "4223600000", recodificarNombre: "ADELANTO VIAJE", codigoNuevo: "4226000000" },
  { archivo: "balance_acumulado_brl", codigo: "4225000000", recodificarNombre: "ADELANTO VIAJE", codigoNuevo: "4226000000" },

  { archivo: "balance_acumulado_ars", codigo: "4211100000", borrarNombre: "SERVICIOS DE LIMPIEZA", queda: "REDONDEO" },
  { archivo: "balance_acumulado_ars", codigo: "4211200000", borrarNombre: "GASTOS TELEFONICOS", queda: "GASTOS EN EQ. TELEFONICOS" },
  { archivo: "balance_acumulado_ars", codigo: "4211700000", borrarNombre: "GASTOS CAPACITACION", queda: "GASTOS DE CAPACITACION" },
  { archivo: "balance_acumulado_ars", codigo: "4230200000", borrarNombre: "IMP. A LOS DEBITOS Y CREDITOS LEY 25,413", queda: "IMP. A LOS DEBITOS" },
  { archivo: "balance_acumulado_ars", codigo: "4212800000", borrarNombre: "TASAS AFIP",         queda: "TASA AFIP" },
  { archivo: "balance_acumulado_brl", codigo: "4230200000", borrarNombre: "IMP. A LOS DEBITOS", queda: "IMP. A LOS DEBITOS Y CREDITOS LEY 25,413" },
];

function ddNormaliza(t) {
  return String(t || "")
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();
}

// El nombre de una fila del plan, sin el código.
function ddNombreDe(texto) {
  const m = /^\s*(\d{6,})\s*-?\s*(.*)$/.exec(String(texto).trim());
  return ddNormaliza(m ? m[2] : texto);
}

// Busca la fila que hay que borrar: misma cuenta (código) y el nombre que marcaron en rojo.
// Devuelve null si ya no está (por ejemplo, porque se aplicó en una corrida anterior).
function ddBuscarFila(planDeCuentas, decision) {
  const objetivo = ddNormaliza(decision.borrarNombre || decision.recodificarNombre);
  const info = planDeCuentas[decision.codigo];
  if (!info) return null;
  const candidatas = [info].concat(info.otrasFilas || []);
  const encontrada = candidatas.find(f => ddNombreDe(f.texto) === objetivo);
  return encontrada || null;
}

// Le corrige el código a una cuenta que quedó cargada con el de otra.
//
// Se reescribe SOLO el código, dejando intacto el resto del texto (el separador del maestro es
// un doble espacio, y el BUSCARV de la fila busca el texto exacto). No hace falta tocar ninguna
// fórmula: cada fila usa su propia celda como clave de búsqueda, así que al corregir el texto
// la búsqueda se corrige sola, y la fila no se mueve, así que nada de lo que la referencia
// cambia de lugar.
function ddRecodificar(wb, layout, fila, codigoNuevo, log = () => {}) {
  const ws = wb.getWorksheet(layout.sheet);
  const celda = ws.getCell(fila.fila, layout.keyCol);
  const textoViejo = String(fila.texto);
  const textoNuevo = textoViejo.replace(/^\s*[\d.]+/, codigoNuevo);
  if (textoNuevo === textoViejo) return null;   // ya estaba corregido
  celda.value = textoNuevo;
  log(`  Recodificada ${layout.sheet}!${fila.fila} — "${textoViejo}" → "${textoNuevo}".`);
  return { fila: fila.fila, textoViejo, textoNuevo };
}

// Aplica las decisiones que correspondan a este archivo.
//
// Borrar una fila referenciada por alguna hoja la dejaría en #REF!, así que esas NO se tocan
// y se devuelven como pendientes: hay que definir antes a qué cuenta pasa a leer esa línea.
// Cuatro de las ocho están en ese caso, y en dos de ellas la línea del Anexo II ya venía
// leyendo una cuenta que no es la suya (la línea "Gastos Telefonico" lee SERVICIOS DE
// LIMPIEZA, y "Gastos obra social" lee GASTOS TELEFONICOS).
function aplicarDecisionesDuplicados(wb, archivoId, layout, planDeCuentas, log = () => {}) {
  const borradas = [];
  const recodificadas = [];
  const pendientes = [];
  const noEstaban = [];

  const mias = DECISIONES_DUPLICADOS.filter(d => d.archivo === archivoId);

  // Las recodificaciones van primero y aparte: no mueven filas, así que no dependen del orden
  // ni invalidan los números de fila que usa el borrado de más abajo.
  for (const d of mias.filter(x => x.recodificarNombre)) {
    const f = ddBuscarFila(planDeCuentas, d);
    if (!f) { noEstaban.push(d); continue; }
    // Si el código nuevo ya lo usa otra fila, corregirlo crearía un duplicado nuevo en vez de
    // resolver uno. No se fuerza: se avisa y se deja como está.
    const yaExiste = planDeCuentas[d.codigoNuevo];
    if (yaExiste && yaExiste.fila !== f.fila) {
      pendientes.push({
        codigo: d.codigo, nombre: d.recodificarNombre, fila: f.fila, referencias: [],
        motivo: `no la recodifiqué a ${d.codigoNuevo}: ese código ya lo usa la fila ` +
                `${yaExiste.fila} ("${yaExiste.texto}"), así que el cambio crearía un ` +
                `duplicado nuevo en vez de resolver este.`,
      });
      log(`  ⚠ ${d.codigo} "${d.recodificarNombre}" NO se recodificó: ${d.codigoNuevo} ya está en uso.`);
      continue;
    }
    const r = ddRecodificar(wb, layout, f, d.codigoNuevo, log);
    if (r) recodificadas.push({ codigo: d.codigo, codigoNuevo: d.codigoNuevo, ...r });
    else noEstaban.push(d);
  }
  // De abajo hacia arriba: cada borrado corre las filas de abajo.
  const conFila = mias.filter(d => d.borrarNombre)
    .map(d => ({ d, f: ddBuscarFila(planDeCuentas, d) }))
    .filter(x => { if (!x.f) noEstaban.push(x.d); return !!x.f; })
    .sort((a, b) => b.f.fila - a.f.fila);

  for (const { d, f } of conFila) {
    const culpables = quienReferenciaLaFila(wb, layout.sheet, f.fila);
    if (culpables.length) {
      pendientes.push({
        codigo: d.codigo, nombre: d.borrarNombre, fila: f.fila,
        referencias: culpables,
        motivo: `la referencian ${culpables.length} fórmula(s) y quedarían en #REF!: ` +
                culpables.slice(0, 3).join(" | ") +
                `. Antes de borrarla hay que definir qué cuenta pasa a leer esa línea.`,
      });
      log(`  ⚠ ${d.codigo} "${d.borrarNombre}" (fila ${f.fila}) NO se borró: ` +
          `la referencian ${culpables.map(c => c.split(" = ")[0]).join(", ")}.`);
      continue;
    }
    const modificadas = borrarFilaEn(wb, layout.sheet, f.fila);
    borradas.push({ codigo: d.codigo, nombre: d.borrarNombre, fila: f.fila, queda: d.queda });
    log(`  Borrada ${layout.sheet}!${f.fila} — ${d.codigo} "${d.borrarNombre}" (queda "${d.queda}"). ` +
        `${modificadas} referencia(s) reacomodadas.`);
  }

  return { borradas, recodificadas, pendientes, noEstaban };
}

if (typeof module !== "undefined") {
  const fh = require("./formula_hojas.js");
  global.borrarFilaEn = fh.borrarFilaEn;
  global.quienReferenciaLaFila = fh.quienReferenciaLaFila;
  module.exports = {
    DECISIONES_DUPLICADOS, aplicarDecisionesDuplicados, ddBuscarFila, ddNombreDe,
  };
}
