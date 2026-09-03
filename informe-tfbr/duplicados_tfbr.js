// Códigos de cuenta repetidos dentro de SALDOS: los clasifica y, cuando se puede probar que
// es seguro, borra la fila sobrante.
//
// No todos los repetidos son el mismo problema, y por eso no se pueden tratar igual:
//
//   texto idéntico   las dos filas encuentran el mismo importe en la zona de pegado, así que
//                    el importe entra DOS VECES en los subtotales por rango, y a veces también
//                    en dos líneas distintas del Anexo II o del EERR. Es un doble conteo.
//
//   texto distinto   (un espacio de más, un acento) sólo una de las dos matchea; la otra queda
//                    en cero para siempre. Si nadie la referencia, sobra y se puede borrar. Si
//                    alguien la referencia, esa línea viene mostrando cero.
//
//   nombre distinto  el mismo código usado por DOS CUENTAS DIFERENTES (en el Mensual R$,
//                    4223600000 es "ADELANTO VIAJE" en una fila y "DEUDORES INCOBRABLES" en
//                    otra). Borrar cualquiera elimina una cuenta real: no se toca, se avisa.
//
// Sólo se borra la fila que cumple TODAS estas condiciones: el nombre coincide con el de su
// par, su texto NO es el que matchea contra la zona de pegado (o sea, está muerta), y ninguna
// fórmula del libro la referencia de forma suelta. Todo lo demás requiere una decisión
// contable y queda para revisar.

const DUP_RE_RANGO = /\$?([A-Z]{1,3})\$?(\d+)\s*:\s*\$?([A-Z]{1,3})\$?(\d+)/g;
const DUP_RE_CUENTA = /^\s*(\d{6,})\s*(.*)$/;

function dupTexto(ws, fila, col) {
  const v = ws.getCell(fila, col).value;
  if (v === null || v === undefined) return "";
  if (typeof v === "object") {
    if (v.richText) return v.richText.map(t => t.text).join("");
    return "";
  }
  return String(v);
}

// El nombre de la cuenta, sin el código y normalizado, para poder comparar "GASTOS DE
// CAPACITACION" con "GASTOS CAPACITACIÓN" y decidir si son la misma cuenta escrita distinto
// o dos cuentas diferentes.
function dupNombreNormalizado(texto) {
  const m = DUP_RE_CUENTA.exec(String(texto).trim());
  const nombre = m ? m[2] : String(texto);
  return nombre
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();
}

// Las celdas que referencian esta fila de SALDOS con una referencia suelta (no dentro de un
// rango). Son las que quedarían en #REF! si la fila se borrara.
function dupReferenciasSueltas(wb, hoja, fila) {
  const reSuelta = new RegExp("(?:'?" + hoja + "'?!)?\\$?[A-Z]{1,3}\\$?" + fila + "(?!\\d)");
  const reNombra = new RegExp("'?" + hoja + "'?!");
  const encontradas = [];
  for (const ws of wb.worksheets) {
    const esLaHoja = ws.name === hoja;
    ws.eachRow({ includeEmpty: false }, (row, r) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const v = cell.value;
        if (!(v && typeof v === "object" && typeof v.formula === "string")) return;
        if (!reNombra.test(v.formula) && !esLaHoja) return;
        if (esLaHoja && r === fila) return;
        const sinRangos = v.formula.replace(DUP_RE_RANGO, " ");
        if (reSuelta.test(sinRangos)) {
          encontradas.push({ hoja: ws.name, celda: cell.address, formula: v.formula });
        }
      });
    });
  }
  return encontradas;
}

// ¿El texto de esta fila matchea contra la zona de pegado? Es lo que decide cuál de las dos
// levanta el importe. Se compara igual que el BUSCARV: texto exacto.
function dupTextoEstaEnStaging(ws, layout, texto) {
  for (let r = layout.stagingRange.filaDesde; r <= layout.stagingRange.filaHasta; r++) {
    if (dupTexto(ws, r, layout.stagingRange.colDesde) === texto) return true;
  }
  return false;
}

function clasificarDuplicados(wb, layout, duplicadas) {
  const ws = wb.getWorksheet(layout.sheet);
  const casos = [];

  for (const d of duplicadas) {
    const filas = [d.filaPrevia, d.fila].map(fila => {
      const texto = dupTexto(ws, fila, layout.keyCol);
      return {
        fila,
        texto,
        nombre: dupNombreNormalizado(texto),
        matchea: dupTextoEstaEnStaging(ws, layout, texto),
        refs: dupReferenciasSueltas(wb, layout.sheet, fila),
      };
    });

    const [a, b] = filas;
    let tipo, sePuedeBorrar = null, motivo;

    if (a.nombre !== b.nombre) {
      tipo = "nombre_distinto";
      motivo = `El código lo usan dos cuentas diferentes ("${a.texto.trim()}" y ` +
               `"${b.texto.trim()}"). Borrar cualquiera elimina una cuenta real del balance.`;
    } else if (a.texto === b.texto) {
      tipo = "texto_identico";
      motivo = `Las dos filas tienen el texto idéntico, así que las dos levantan el mismo ` +
               `importe y se cuenta dos veces` +
               (a.refs.length && b.refs.length
                 ? ` (además cada una alimenta líneas distintas: ${a.refs.map(r => r.hoja + "!" + r.celda).join(", ")} ` +
                   `y ${b.refs.map(r => r.hoja + "!" + r.celda).join(", ")}).`
                 : `.`) +
               ` Cuál de las dos sobra es una decisión contable.`;
    } else {
      // texto distinto: una matchea y la otra no
      const muerta = filas.find(f => !f.matchea);
      const viva = filas.find(f => f.matchea);
      tipo = "texto_distinto";
      if (!muerta || !viva) {
        motivo = `Ninguna de las dos matchea contra la zona de pegado en esta corrida, así que ` +
                 `no se puede saber cuál queda viva. Se revisa a mano.`;
      } else if (muerta.refs.length) {
        motivo = `La fila ${muerta.fila} nunca levanta su importe (su texto no matchea) pero ` +
                 `${muerta.refs.map(r => r.hoja + "!" + r.celda).join(", ")} la referencia: ` +
                 `esa línea viene mostrando cero. Hay que repuntarla a la fila ${viva.fila} ` +
                 `antes de borrar nada.`;
      } else {
        sePuedeBorrar = muerta.fila;
        motivo = `La fila ${muerta.fila} nunca levanta su importe y no la referencia ninguna ` +
                 `fórmula: sobra. Queda la ${viva.fila}, que es la que matchea.`;
      }
    }

    casos.push({ codigo: d.codigo, tipo, filas, sePuedeBorrar, motivo });
  }
  return casos;
}

// Foto de a qué CUENTA apunta cada referencia del libro, para poder comparar antes y después
// de borrar filas. Es la verificación de que lo que queda sigue apuntando al lugar correcto:
// las direcciones cambian (todo lo de abajo sube una fila), pero la cuenta a la que llegan
// tiene que ser exactamente la misma.
function fotoDeReferencias(wb, layout) {
  const hoja = layout.sheet;
  const ws = wb.getWorksheet(hoja);
  const reNombra = new RegExp("'?" + hoja + "'?!");
  const reRef = new RegExp("(?:'?" + hoja + "'?!)\\$?([A-Z]{1,3})\\$?(\\d+)(?!\\d)", "g");
  const foto = {};
  for (const w of wb.worksheets) {
    w.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const v = cell.value;
        if (!(v && typeof v === "object" && typeof v.formula === "string")) return;
        if (!reNombra.test(v.formula)) return;
        const sinRangos = v.formula.replace(DUP_RE_RANGO, " ");
        reRef.lastIndex = 0;
        let m;
        const apunta = [];
        while ((m = reRef.exec(sinRangos))) {
          apunta.push(dupTexto(ws, parseInt(m[2], 10), layout.keyCol).trim());
        }
        if (apunta.length) foto[w.name + "!" + cell.address] = apunta;
      });
    });
  }
  return foto;
}

function compararFotos(antes, despues) {
  const cambios = [];
  for (const clave of Object.keys(antes)) {
    const a = antes[clave], b = despues[clave];
    if (!b) { cambios.push(`${clave}: ya no referencia SALDOS`); continue; }
    if (a.join("|") !== b.join("|")) {
      cambios.push(`${clave}: apuntaba a [${a.join(", ")}] y ahora a [${b.join(", ")}]`);
    }
  }
  return cambios;
}

if (typeof module !== "undefined") {
  module.exports = {
    clasificarDuplicados, fotoDeReferencias, compararFotos,
    dupNombreNormalizado, dupReferenciasSueltas,
  };
}
