// Los datos del período que NO vienen del export de Onvio y hoy se tipean a mano:
// el tipo de cambio de cierre y la cifra mensual del cuadro "Explicación dif de cambio".
//
// Los dos viven solo en el Balance Acumulado R$ (los otros tres archivos no los tienen).
// Igual que el resto del motor, las celdas no están hardcodeadas: se ubican leyendo el
// archivo, porque una dirección anotada a mano se desactualiza sin avisar — de hecho, la
// primera investigación de estos archivos tenía mal la celda del TC por ese motivo.

const MESES_ES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
                  "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];

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
      return { filaTitulo, colEtiqueta: colTitulo, colValor, filas, filaTotal: r };
    }
    const t = pfTexto(ws, r, colTitulo).trim();
    if (!t) continue;
    const mes = MESES_ES.findIndex(m => t.toUpperCase().startsWith(m)) + 1;
    filas.push({ fila: r, etiqueta: t, mes: mes || null });
  }
  return { filaTitulo, colEtiqueta: colTitulo, colValor, filas, filaTotal: null };
}

function pfUltimoDiaDelMes(anio, mes) {
  return new Date(anio, mes, 0).getDate();
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
      // No se inserta la fila sola: dónde va exactamente (antes o después de la fila de
      // acumulado, que está en el medio del cuadro) cambia lo que suma el total, y eso es
      // una decisión contable, no de formato.
      pendiente.push(
        `El cuadro "Explicación dif de cambio" no tiene fila para ${MESES_ES[mes - 1]}: ` +
        `hay que agregarla a mano en el Excel (hoy llega hasta "${cuadro.filas.length ? cuadro.filas[cuadro.filas.length - 1].etiqueta : "—"}").`
      );
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
  };
}
