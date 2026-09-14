// Las decisiones de contaduría sobre los códigos usados por dos cuentas distintas
// (documento "Cuentas TFBR", septiembre 2026: en rojo, la fila que se elimina).
//
// Se identifican por CÓDIGO + NOMBRE, nunca por número de fila. El documento vino con las
// filas del archivo de julio, pero las filas cambian entre un mes y otro —y vuelven a cambiar
// cada vez que se da de alta una cuenta— así que aplicarlas por número borraría otra cosa.
//
// Aplicarlas es idempotente: si la fila ya no está, no hay nada que hacer.

const DECISIONES_DUPLICADOS = [
  { archivo: "balance_mensual_brl",   codigo: "4223600000", borrarNombre: "ADELANTO VIAJE",   queda: "DEUDORES INCOBRABLES" },
  { archivo: "balance_acumulado_ars", codigo: "4211100000", borrarNombre: "SERVICIOS DE LIMPIEZA", queda: "REDONDEO" },
  { archivo: "balance_acumulado_ars", codigo: "4211200000", borrarNombre: "GASTOS TELEFONICOS", queda: "GASTOS EN EQ. TELEFONICOS" },
  { archivo: "balance_acumulado_ars", codigo: "4211700000", borrarNombre: "GASTOS CAPACITACION", queda: "GASTOS DE CAPACITACION" },
  { archivo: "balance_acumulado_ars", codigo: "4230200000", borrarNombre: "IMP. A LOS DEBITOS Y CREDITOS LEY 25,413", queda: "IMP. A LOS DEBITOS" },
  { archivo: "balance_acumulado_ars", codigo: "4212800000", borrarNombre: "TASAS AFIP",         queda: "TASA AFIP" },
  { archivo: "balance_acumulado_brl", codigo: "4225000000", borrarNombre: "ADELANTO VIAJE",     queda: "CESION DE DERECHOS" },
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
  const objetivo = ddNormaliza(decision.borrarNombre);
  const info = planDeCuentas[decision.codigo];
  if (!info) return null;
  const candidatas = [info].concat(info.otrasFilas || []);
  const encontrada = candidatas.find(f => ddNombreDe(f.texto) === objetivo);
  return encontrada || null;
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
  const pendientes = [];
  const noEstaban = [];

  const mias = DECISIONES_DUPLICADOS.filter(d => d.archivo === archivoId);
  // De abajo hacia arriba: cada borrado corre las filas de abajo.
  const conFila = mias
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

  return { borradas, pendientes, noEstaban };
}

if (typeof module !== "undefined") {
  const fh = require("./formula_hojas.js");
  global.borrarFilaEn = fh.borrarFilaEn;
  global.quienReferenciaLaFila = fh.quienReferenciaLaFila;
  module.exports = {
    DECISIONES_DUPLICADOS, aplicarDecisionesDuplicados, ddBuscarFila, ddNombreDe,
  };
}
