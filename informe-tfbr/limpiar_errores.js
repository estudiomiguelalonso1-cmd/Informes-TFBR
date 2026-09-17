// Las celdas rotas que el informe arrastra de cierres viejos.
//
// Los cuatro maestros vienen con fórmulas que dicen "#REF!": alguien borró en su momento las
// filas o columnas que sumaban y Excel dejó la referencia rota para siempre. Son "SUM(#REF!)"
// en una columna lateral del Pasivo y del Anexo II, "+#REF!+K13" en el EEPN de los acumulados
// y "+#REF!" en el Anexo I del Mensual R$.
//
// No las lee nadie —ninguna otra celda las referencia— así que no mueven ningún número. Pero se
// imprimen, y un informe contable con un #REF! a la vista no se puede entregar.
//
// Qué se hace y por qué es seguro. Una fórmula que dice #REF! perdió el rango que sumaba y no
// hay forma de saber qué sumaba: no se puede reparar, sólo vaciar. Y vaciarla es exactamente lo
// que tienen los otros informes en esa misma celda —el Pasivo y el Anexo II la tienen vacía en
// los otros tres archivos, el EEPN la tiene vacía en los dos mensuales—, así que el resultado
// es que los cuatro queden iguales, que es como tienen que estar.
//
// El caso del Anexo I del Mensual R$ es el mismo por otro camino: "D17 = +#REF!" rompía además
// el total "D18 = SUM(D11:D17)". Los dos acumulados tienen 0,00 en esa celda y el resto de esa
// misma columna del Mensual R$ está vacío, así que vaciarla la deja como sus vecinas y el total
// pasa a dar cero. El "Bienes de uso" del informe no se mueve: sigue siendo la amortización del
// mes, igual que en el Mensual $ (-2.485.298,32 en pesos, -8.531,00 al tipo de cambio).
//
// Lo que NO se toca: una celda rota que alguien lee de forma directa. Ahí vaciarla cambiaría el
// número de quien la lee, y el número correcto no lo sabe el programa. Queda el aviso con la
// dirección exacta.
//
// Tampoco se toca una celda que da error por otro motivo (#DIV/0!, #N/A): ésas suelen arreglarse
// solas cuando se arregla lo que leen —el total del Anexo I es el ejemplo— y Excel las recalcula
// al abrir el archivo.

function leRotaTexto(cell) {
  const v = cell.value;
  if (!v || typeof v !== "object" || typeof v.formula !== "string") return false;
  return v.formula.includes("#REF!");
}

function leEsError(cell) {
  const v = cell.value;
  if (!v || typeof v !== "object") return null;
  if (v.error) return v.error;
  if (v.result && v.result.error) return v.result.error;
  return null;
}

// Una referencia con hoja adelante ('Anexo II'!D17, SALDOS!B7) y una sin ella (D17).
const LE_RE_REF = /(?:'([^']+)'|([A-Za-z0-9 _.À-ſ]+))!\$?([A-Z]{1,3})\$?(\d+)(?![0-9])/g;
const LE_RE_LOCAL = /\$?([A-Z]{1,3})\$?(\d+)(?![0-9])/g;

// ¿La referencia que empieza en `i` es parte de un rango (D11:D17)?
//
// Importa para decidir si se puede vaciar. Un rango no se rompe porque una de sus celdas quede
// vacía: pasa a sumar cero y el total sigue funcionando. Una referencia directa sí.
function leEnRango(formula, i, largo) {
  return formula[i + largo] === ":" || formula[i - 1] === ":";
}

// Quién referencia una celda de forma directa, desde cualquier hoja. Se compara hoja y dirección
// sin armar una expresión regular con el nombre de la hoja adentro: un nombre con paréntesis o
// acento rompería el patrón.
function leLectores(wb, hoja, dir) {
  const lectores = [];
  const buscado = (hoja + "!" + dir).toUpperCase();
  const anotar = (ws, cell) => {
    const x = ws.name + "!" + cell.address;
    if (!lectores.includes(x)) lectores.push(x);
  };

  for (const ws of wb.worksheets) {
    ws.eachRow({ includeEmpty: false }, (row) => row.eachCell({ includeEmpty: false }, (cell) => {
      const v = cell.value;
      if (!v || typeof v !== "object" || typeof v.formula !== "string") return;
      if (ws.name === hoja && cell.address === dir) return;
      const f = v.formula;

      let m; LE_RE_REF.lastIndex = 0;
      while ((m = LE_RE_REF.exec(f)) !== null) {
        if (((m[1] || m[2]) + "!" + m[3] + m[4]).toUpperCase() !== buscado) continue;
        if (leEnRango(f, m.index, m[0].length)) continue;
        anotar(ws, cell);
      }

      if (ws.name !== hoja) return;
      let q; LE_RE_LOCAL.lastIndex = 0;
      while ((q = LE_RE_LOCAL.exec(f)) !== null) {
        if ((q[1] + q[2]).toUpperCase() !== dir.toUpperCase()) continue;
        if (f[q.index - 1] === "!") continue;              // ésa es de otra hoja, ya la vimos
        if (leEnRango(f, q.index, q[0].length)) continue;
        anotar(ws, cell);
      }
    }));
  }
  return lectores;
}

function limpiarErrores(wb, log = () => {}) {
  const vaciadas = [], trabadas = [], otros = [], recalculadas = [];

  for (const ws of wb.worksheets) {
    const aVaciar = [];
    ws.eachRow({ includeEmpty: false }, (row) => row.eachCell({ includeEmpty: false }, (cell) => {
      if (!leRotaTexto(cell)) return;
      aVaciar.push({ dir: cell.address, formula: cell.value.formula });
    }));
    for (const x of aVaciar) {
      const lectores = leLectores(wb, ws.name, x.dir);
      if (lectores.length) { trabadas.push({ hoja: ws.name, ...x, lectores }); continue; }
      ws.getCell(x.dir).value = null;
      vaciadas.push({ hoja: ws.name, ...x });
    }
  }

  // Lo que quedo en error y NO es una referencia rota: casi siempre es un total que daba error
  // solo porque una celda suya estaba rota. Se le tira el valor viejo guardado —la formula no se
  // toca— para que Excel lo recalcule al abrir en vez de dejar el error escrito en el archivo.
  for (const ws of wb.worksheets) {
    ws.eachRow({ includeEmpty: false }, (row) => row.eachCell({ includeEmpty: false }, (cell) => {
      const e = leEsError(cell);
      if (!e || leRotaTexto(cell)) return;
      const f = (cell.value && typeof cell.value === "object" && cell.value.formula) || "";
      if (f) { cell.value = { formula: f }; recalculadas.push({ hoja: ws.name, dir: cell.address, error: e }); return; }
      otros.push({ hoja: ws.name, dir: cell.address, error: e, formula: f });
    }));
  }

  if (vaciadas.length) {
    log(`  ${vaciadas.length} celda(s) con referencias rotas (#REF!) vaciadas: ` +
        vaciadas.map(x => `${x.hoja}!${x.dir}`).join(", ") + ".");
  }
  for (const t of trabadas) {
    log(`  ⚠ ${t.hoja}!${t.dir} dice "${t.formula}" y la leen ${t.lectores.join(", ")}: ` +
        `no la vacío porque cambiaría un número. Hay que arreglarla a mano.`);
  }
  if (recalculadas.length) {
    log(`  ${recalculadas.length} celda(s) quedaron con el error viejo guardado y se dejaron ` +
        `para que Excel las recalcule: ` + recalculadas.map(x => `${x.hoja}!${x.dir}`).join(", ") + ".");
  }
  return { vaciadas, trabadas, otros, recalculadas };
}

if (typeof module !== "undefined") {
  module.exports = { limpiarErrores, leLectores, leRotaTexto, leEsError, leEnRango };
}
