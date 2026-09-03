// Borra la fila sobrante de un código repetido, cuando se puede hacer sin romper nada.
//
// La regla, decidida con la usuaria: de las dos filas se borra la que NO está referenciada
// por ninguna hoja, y queda la que sí (borrar la referenciada dejaría esa línea del Anexo II
// o del EERR apuntando a una fila que ya no existe). Si ninguna de las dos está referenciada,
// da igual cuál: se deja la primera.
//
// Lo que NO se resuelve solo y queda para decidir:
//   - el mismo código usado por dos cuentas DISTINTAS: borrar una elimina una cuenta real;
//   - las dos filas referenciadas por líneas distintas: cuál sobrevive cambia en qué línea
//     del estado se reporta el importe;
//   - la fila muerta referenciada: hay que repuntar esa referencia antes, y a dónde va es
//     una decisión contable.
//
// Después de borrar se compara, referencia por referencia, a qué CUENTA llega cada fórmula
// (no a qué celda: las direcciones se corren solas al borrar). Si alguna termina apuntando a
// otra cuenta, se considera fallida la operación y no se guarda el archivo.

function rdElegirFilaABorrar(caso) {
  if (caso.tipo === "nombre_distinto") {
    return { fila: null, motivo: "son dos cuentas distintas con el mismo código" };
  }

  const [a, b] = caso.filas;
  const aRef = a.refs.length > 0, bRef = b.refs.length > 0;

  if (aRef && bRef) {
    const celdasA = a.refs.map(r => r.hoja + "!" + r.celda);
    const celdasB = b.refs.map(r => r.hoja + "!" + r.celda);
    const compartidas = celdasA.filter(x => celdasB.includes(x));
    if (compartidas.length) {
      // El caso más claro de doble conteo: una sola fórmula suma las dos filas, así que el
      // mismo importe entra dos veces en la MISMA línea del estado. Arreglarlo es sacar uno
      // de los dos términos de esa fórmula, y eso ya es tocar el estado: se avisa.
      return {
        fila: null,
        motivo: `la misma fórmula (${compartidas.join(", ")}) suma las dos filas, así que el ` +
                `importe entra dos veces en esa línea. Hay que sacar uno de los dos términos ` +
                `de la fórmula antes de borrar la fila`,
      };
    }
    return {
      fila: null,
      motivo: `las dos filas están referenciadas por líneas distintas ` +
              `(${celdasA.join(", ")} y ${celdasB.join(", ")}): cuál sobrevive cambia ` +
              `en qué línea del estado se reporta el importe`,
    };
  }

  if (caso.tipo === "texto_distinto") {
    const muerta = caso.filas.find(f => !f.matchea);
    const viva = caso.filas.find(f => f.matchea);
    if (muerta && viva && muerta.refs.length) {
      return {
        fila: null,
        motivo: `la fila muerta (${muerta.fila}) está referenciada por ` +
                `${muerta.refs.map(r => r.hoja + "!" + r.celda).join(", ")}: hay que repuntar ` +
                `esa referencia antes de borrarla, y a dónde va es una decisión contable`,
      };
    }
    if (muerta && viva) {
      return { fila: muerta.fila, motivo: `la fila ${muerta.fila} nunca levanta su importe y no la referencia nadie` };
    }
    // ninguna matchea en esta corrida: se van igual las dos a la misma cuenta, se deja la primera
    return { fila: b.fila, motivo: `ninguna matchea este mes y ninguna está referenciada; se deja la primera (${a.fila})` };
  }

  // texto idéntico: las dos levantan el mismo importe
  if (aRef) return { fila: b.fila, motivo: `la fila ${b.fila} no la referencia nadie; queda la ${a.fila}, que sí` };
  if (bRef) return { fila: a.fila, motivo: `la fila ${a.fila} no la referencia nadie; queda la ${b.fila}, que sí` };
  return { fila: b.fila, motivo: `ninguna de las dos está referenciada; se deja la primera (${a.fila})` };
}

// Resuelve todos los casos que se puedan de un maestro ya abierto. Devuelve qué borró, qué
// quedó pendiente, y el resultado de la verificación.
function resolverDuplicados(wb, layout, casos, log = () => {}) {
  const antes = fotoDeReferencias(wb, layout);

  const aBorrar = [];
  const pendientes = [];
  for (const caso of casos) {
    const { fila, motivo } = rdElegirFilaABorrar(caso);
    if (fila === null) {
      pendientes.push({ codigo: caso.codigo, tipo: caso.tipo, motivo, filas: caso.filas.map(f => f.fila) });
    } else {
      aBorrar.push({ codigo: caso.codigo, fila, motivo });
    }
  }

  // De abajo hacia arriba: borrar una fila corre todo lo de abajo, así que hacerlo al revés
  // invalidaría los números de fila de los borrados siguientes.
  aBorrar.sort((x, y) => y.fila - x.fila);

  const borradas = [];
  for (const b of aBorrar) {
    const culpables = quienReferenciaLaFila(wb, layout.sheet, b.fila);
    if (culpables.length) {
      // no debería pasar (se eligió una fila sin referencias), pero si pasa no se fuerza
      pendientes.push({
        codigo: b.codigo, tipo: "referencia_inesperada", filas: [b.fila],
        motivo: `al ir a borrarla aparecieron referencias que no estaban en el análisis: ${culpables.slice(0, 3).join(" | ")}`,
      });
      continue;
    }
    const modificadas = borrarFilaEn(wb, layout.sheet, b.fila);
    borradas.push({ ...b, modificadas });
    log(`  Borrada ${layout.sheet}!${b.fila} (${b.codigo}): ${b.motivo}. ${modificadas} referencia(s) reacomodadas.`);
  }

  const despues = fotoDeReferencias(wb, layout);
  const cambios = compararFotos(antes, despues);

  return { borradas, pendientes, verificacion: { cambios, ok: cambios.length === 0 } };
}

if (typeof module !== "undefined") {
  const dup = require("./duplicados_tfbr.js");
  const fh = require("./formula_hojas.js");
  global.fotoDeReferencias = dup.fotoDeReferencias;
  global.compararFotos = dup.compararFotos;
  global.borrarFilaEn = fh.borrarFilaEn;
  global.quienReferenciaLaFila = fh.quienReferenciaLaFila;
  module.exports = { resolverDuplicados, rdElegirFilaABorrar };
}
