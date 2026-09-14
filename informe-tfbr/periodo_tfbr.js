// Los datos del período que NO vienen del export de Onvio y hoy se tipean a mano:
// el tipo de cambio de cierre y la cifra mensual del cuadro "Explicación dif de cambio".
//
// Los dos viven solo en el Balance Acumulado R$ (los otros tres archivos no los tienen).
// Igual que el resto del motor, las celdas no están hardcodeadas: se ubican leyendo el
// archivo, porque una dirección anotada a mano se desactualiza sin avisar — de hecho, la
// primera investigación de estos archivos tenía mal la celda del TC por ese motivo.

const MESES_ES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
                  "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];

// "D" a partir del número de columna, para poder nombrar celdas concretas en los avisos.
// Local a este archivo (config_tfbr.js tiene ctColNumeroALetra, pero periodo_tfbr.js no
// depende de él en Node y no vale la pena acoplarlos por un aviso).
function pfColLetra(n) {
  let s = "";
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = (n - r - 1) / 26; }
  return s;
}

function pfTexto(ws, fila, col) {
  const v = ws.getCell(fila, col).value;
  if (v === null || v === undefined) return "";
  if (typeof v === "object") {
    if (v.richText) return v.richText.map(t => t.text).join("");
    if (v.result !== undefined && typeof v.result !== "object") return String(v.result);
    return "";
  }
  return String(v);
}

// El TC se reconoce por su etiqueta ("T.C. Reales al 31/07/2026"), y el valor es la celda
// de al lado. Devuelve null si este archivo no tiene TC — que es el caso de 3 de los 4.
function ubicarTcCierre(ws) {
  for (let r = 1; r <= Math.min(ws.rowCount, 6); r++) {
    for (let c = 1; c <= Math.min(ws.columnCount, 15); c++) {
      const t = pfTexto(ws, r, c).trim();
      if (!/^T\.?\s*C\.?\b/i.test(t)) continue;
      const valor = ws.getCell(r, c + 1).value;
      if (typeof valor !== "number") continue;
      return { filaEtiqueta: r, colEtiqueta: c, colValor: c + 1, etiqueta: t };
    }
  }
  return null;
}

// El cuadro arranca con el título "EXPLICACIÓN DIF DE CAMBIO" y sigue con una fila por mes,
// más una de acumulado, y cierra con el SUM. Cada fila se identifica por su texto, así que
// no hace falta saber de antemano qué fila es cada mes (ni suponer que agosto es la
// siguiente a julio: la fila de acumulado está justo en el medio).
function ubicarCuadroDifCambio(ws) {
  let filaTitulo = null, colTitulo = null;
  for (let r = 1; r <= ws.rowCount && filaTitulo === null; r++) {
    for (let c = 1; c <= Math.min(ws.columnCount, 15); c++) {
      const t = pfTexto(ws, r, c).trim().toUpperCase();
      if (t.startsWith("EXPLICACI") && t.includes("DIF") && t.includes("CAMBIO")) {
        filaTitulo = r; colTitulo = c; break;
      }
    }
  }
  if (filaTitulo === null) return null;

  const colValor = colTitulo + 1;
  const filas = [];
  for (let r = filaTitulo + 1; r <= Math.min(ws.rowCount, filaTitulo + 30); r++) {
    // La fila del total lleva el SUM y NO lleva etiqueta, así que se mira antes que el
    // texto: si se saltea por no tener etiqueta, el recorrido sigue de largo y termina
    // leyendo la zona de pegado como si fueran meses.
    const v = ws.getCell(r, colValor).value;
    if (v && typeof v === "object" && typeof v.formula === "string") {
      return pfCerrarCuadro({ filaTitulo, colEtiqueta: colTitulo, colValor, filas,
                              filaTotal: r, formulaTotal: v.formula });
    }
    const t = pfTexto(ws, r, colTitulo).trim();
    if (!t) continue;
    const mes = MESES_ES.findIndex(m => t.toUpperCase().startsWith(m)) + 1;
    filas.push({ fila: r, etiqueta: t, mes: mes || null });
  }
  return pfCerrarCuadro({ filaTitulo, colEtiqueta: colTitulo, colValor, filas,
                          filaTotal: null, formulaTotal: null });
}

// Dónde habría que insertar la fila de un mes que todavía no está: justo debajo del último
// mes cargado. NO al final del cuadro — después del último mes viene la línea de acumulado
// semestral, y el SUM del total la incluye. Una fila agregada debajo de esa línea cae fuera
// del rango del SUM (Excel solo lo estira cuando se inserta DENTRO), así que el importe no
// sumaría y el único síntoma sería el tie-out de más abajo dando distinto, sin decir por qué.
function pfCerrarCuadro(cuadro) {
  const meses = cuadro.filas.filter(f => f.mes !== null);
  cuadro.ultimoMes = meses.length ? meses[meses.length - 1] : null;
  // las filas sin mes que quedaron DESPUÉS del último mes (la de acumulado semestral)
  cuadro.filasNoMes = cuadro.ultimoMes
    ? cuadro.filas.filter(f => f.mes === null && f.fila > cuadro.ultimoMes.fila)
    : cuadro.filas.filter(f => f.mes === null);
  cuadro.filaAInsertar = cuadro.ultimoMes ? cuadro.ultimoMes.fila + 1 : null;
  return cuadro;
}

function pfUltimoDiaDelMes(anio, mes) {
  return new Date(anio, mes, 0).getDate();
}

// El texto que se lee en el checklist, con el Excel abierto al lado. Dice la fila exacta y
// contra qué etiquetas verificar que cayó en el lugar correcto, porque una instrucción del
// tipo "agregala al cuadro" lleva justo al error que se quiere evitar.
function pfAvisoFaltaFilaMes(cuadro, mes, valor) {
  const nombreMes = MESES_ES[mes - 1];
  const col = pfColLetra(cuadro.colValor);
  let msg = `El cuadro "Explicación dif de cambio" no tiene fila para ${nombreMes}: hay que ` +
            `agregarla a mano en el Excel`;

  if (cuadro.filaAInsertar && cuadro.filasNoMes.length) {
    // el caso de hoy: después del último mes viene la línea de acumulado, y el total la suma
    msg += `. Insertá una fila en la ${cuadro.filaAInsertar}, debajo de ` +
           `"${cuadro.ultimoMes.etiqueta}" y ARRIBA de "${cuadro.filasNoMes[0].etiqueta}"` +
           (cuadro.filaTotal
             ? `, así el total (${col}${cuadro.filaTotal}${cuadro.formulaTotal ? ` = ${cuadro.formulaTotal}` : ""}) la toma solo`
             : "") +
           `. Si la agregás debajo de "${cuadro.filasNoMes[0].etiqueta}" queda FUERA del total ` +
           `y el importe no suma`;
  } else if (cuadro.filaAInsertar) {
    msg += `. Va en la fila ${cuadro.filaAInsertar}, debajo de "${cuadro.ultimoMes.etiqueta}"` +
           (cuadro.filaTotal ? `, dentro del rango que suma ${col}${cuadro.filaTotal}` : "");
  }

  if (valor !== null && valor !== undefined) {
    msg += `. El importe que cargaste (${valor}) NO se escribió: ponelo en esa fila`;
  }

  // Si entre el último mes del cuadro y el que se está cerrando hay un hueco, lo más probable
  // es que se haya elegido mal el período (el maestro de cada mes sale del anterior, así que
  // saltear uno no debería poder pasar). Sin este aviso, el texto de arriba manda a insertar
  // el mes nuevo pegado al último cargado, tapando el hueco en vez de mostrarlo.
  if (cuadro.ultimoMes && cuadro.ultimoMes.mes && mes > cuadro.ultimoMes.mes + 1) {
    const faltan = [];
    for (let m = cuadro.ultimoMes.mes + 1; m < mes; m++) faltan.push(MESES_ES[m - 1]);
    msg += ` ⚠ OJO: el cuadro llega hasta ${MESES_ES[cuadro.ultimoMes.mes - 1]} y falta(n) ` +
           `${faltan.join(", ")} en el medio. Revisá que el período sea el correcto antes de ` +
           `insertar nada`;
  }
  return msg + ".";
}

// Escribe en el maestro lo que corresponda de este período. Devuelve qué escribió y qué no,
// para que la pantalla lo muestre: un dato que no se pudo escribir tiene que quedar a la
// vista como paso manual pendiente, no desaparecer.
function escribirDatosDelPeriodo(wb, { periodo, tcCierre, difCambioMes }, log = () => {}) {
  const ws = wb.getWorksheet("SALDOS");
  const [anioStr, mesStr] = String(periodo).split("-");
  const anio = parseInt(anioStr, 10);
  const mes = parseInt(mesStr, 10);
  const hecho = [];
  const pendiente = [];

  const tc = ubicarTcCierre(ws);
  if (tc && tcCierre !== null && tcCierre !== undefined && tcCierre !== "") {
    ws.getCell(tc.filaEtiqueta, tc.colValor).value = Number(tcCierre);
    // la etiqueta lleva la fecha de cierre: si no se actualiza, el reporte muestra el mes pasado
    const dia = pfUltimoDiaDelMes(anio, mes);
    const fecha = `${String(dia).padStart(2, "0")}/${String(mes).padStart(2, "0")}/${anio}`;
    ws.getCell(tc.filaEtiqueta, tc.colEtiqueta).value = `T.C. Reales al ${fecha}`;
    hecho.push(`TC de cierre ${tcCierre} escrito en SALDOS (con su etiqueta al ${fecha})`);
    log(`  TC de cierre: ${tcCierre} (etiqueta "T.C. Reales al ${fecha}").`);
  }

  const cuadro = ubicarCuadroDifCambio(ws);
  if (cuadro) {
    const fila = cuadro.filas.find(f => f.mes === mes);
    const valor = difCambioMes === "" || difCambioMes === null || difCambioMes === undefined
      ? null : Number(difCambioMes);
    if (fila && valor !== null) {
      ws.getCell(fila.fila, cuadro.colValor).value = valor;
      hecho.push(`Diferencia de cambio de ${fila.etiqueta}: ${valor}`);
      log(`  Explicación dif de cambio: ${valor} en "${fila.etiqueta}".`);
    } else if (!fila) {
      // La fila no se inserta sola: agregar un mes al cuadro es una decisión contable (qué
      // período abarca, y si el acumulado semestral de abajo se recalcula), no de formato.
      // Pero el aviso sí tiene que decir DÓNDE va, porque el lugar no es el intuitivo: el
      // final del cuadro es el lugar equivocado (ver pfCerrarCuadro).
      pendiente.push(pfAvisoFaltaFilaMes(cuadro, mes, valor));
      log(`  ⚠ Sin fila para ${MESES_ES[mes - 1]} en el cuadro de dif de cambio: queda pendiente a mano.`);
    } else if (valor === null) {
      pendiente.push(
        `Falta la cifra de diferencia de cambio de ${MESES_ES[mes - 1]} (fila "${fila.etiqueta}").`
      );
      log(`  ⚠ No se cargó la diferencia de cambio de ${MESES_ES[mes - 1]}: queda pendiente a mano.`);
    }
  }

  return { hecho, pendiente, tieneTc: !!tc, tieneCuadro: !!cuadro };
}

if (typeof module !== "undefined") {
  module.exports = {
    MESES_ES, ubicarTcCierre, ubicarCuadroDifCambio, escribirDatosDelPeriodo,
    pfAvisoFaltaFilaMes, pfColLetra,
  };
}
