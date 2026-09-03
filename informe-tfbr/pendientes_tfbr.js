// El checklist de lo que la app NO completa y hay que cargar a mano en Excel.
//
// No es una lista escrita a mano de direcciones de celda: esas se desactualizan y, en estos
// archivos, ya se demostró que las anotadas en la investigación inicial venían corridas. Acá
// cada punto se ubica por su ETIQUETA dentro de la hoja, así sigue valiendo aunque las filas
// se muevan (y de hecho se movieron al resolver los códigos repetidos).
//
// Lo que se busca es lo que el proceso manual ya hacía: movimientos de patrimonio neto,
// altas y bajas de bienes de uso, el reparto por centro de costo del Anexo II y el impuesto
// a las ganancias en los acumulados.

function pdTexto(ws, fila, col) {
  const v = ws.getCell(fila, col).value;
  if (v === null || v === undefined) return "";
  if (typeof v === "object") {
    if (v.richText) return v.richText.map(t => t.text).join("");
    if (v.result !== undefined && typeof v.result !== "object") return String(v.result);
    return "";
  }
  return String(v);
}

function pdNormaliza(t) {
  return String(t || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toUpperCase().replace(/\s+/g, " ").trim();
}

// Busca una fila por su etiqueta dentro de una hoja. Devuelve {fila, col, texto} o null.
function pdBuscarEtiqueta(ws, fragmento, maxCol = 12) {
  const buscado = pdNormaliza(fragmento);
  for (let r = 1; r <= ws.rowCount; r++) {
    for (let c = 1; c <= Math.min(ws.columnCount, maxCol); c++) {
      const t = pdNormaliza(pdTexto(ws, r, c));
      if (t && t.includes(buscado)) return { fila: r, col: c, texto: pdTexto(ws, r, c).trim() };
    }
  }
  return null;
}

function pdValor(ws, fila, col) {
  const v = ws.getCell(fila, col).value;
  if (typeof v === "number") return v;
  if (v && typeof v === "object" && typeof v.result === "number") return v.result;
  return null;
}

// Impuesto a las ganancias: en los acumulados es un número tipeado (o directamente vacío,
// que es peor porque no se distingue de un cero intencional). Se informa qué tiene hoy.
function pdImpuestoGanancias(wb) {
  const ws = wb.getWorksheet("EERR");
  if (!ws) return null;
  const et = pdBuscarEtiqueta(ws, "IMPUESTO A LAS GANANCIAS");
  if (!et) return null;
  // el importe está a la derecha de la etiqueta
  for (let c = et.col + 1; c <= Math.min(ws.columnCount, et.col + 6); c++) {
    const celda = ws.getCell(et.fila, c);
    const v = celda.value;
    if (v === null || v === undefined) continue;
    const esFormula = typeof v === "object" && typeof v.formula === "string";
    return {
      celda: `EERR!${celda.address}`,
      valor: pdValor(ws, et.fila, c),
      esFormula,
    };
  }
  return { celda: `EERR!(fila ${et.fila})`, valor: null, esFormula: false };
}

// Las columnas del Anexo II donde se reparte el gasto por centro de costo. Se reconocen por
// el encabezado (ADM CENTRAL, PACHECO, CORDOBA, TRANSPORTES). Sólo existen en los mensuales.
function pdColumnasCentroDeCosto(wb) {
  const ws = wb.getWorksheet("Anexo II");
  if (!ws) return [];
  // "ADM CTRAL" tal como está escrito en el archivo, no "ADM CENTRAL": buscando el nombre
  // completo la columna J quedaba afuera del checklist sin que se notara.
  const nombres = ["ADM CTRAL", "PACHECO", "CORDOBA", "TRANSPORTES"];
  const encontradas = [];
  for (let r = 1; r <= Math.min(ws.rowCount, 20); r++) {
    for (let c = 1; c <= Math.min(ws.columnCount, 30); c++) {
      const t = pdNormaliza(pdTexto(ws, r, c));
      if (!t) continue;
      for (const n of nombres) {
        if (t.includes(n) && !encontradas.some(x => x.nombre === n)) {
          encontradas.push({ nombre: n, columna: ws.getColumn(c).letter, fila: r });
        }
      }
    }
  }
  return encontradas;
}

// Arma el checklist del archivo. Cada punto dice DÓNDE está, no sólo qué falta.
function pendientesManuales(wb, etiquetaArchivo) {
  const puntos = [];

  const eepn = wb.getWorksheet("EEPN");
  if (eepn) {
    const inicio = pdBuscarEtiqueta(eepn, "SALDOS AL INICIO");
    puntos.push({
      que: "Movimientos de patrimonio neto (aportes, distribuciones, ajustes)",
      donde: "hoja EEPN" + (inicio ? `, debajo de "${inicio.texto}" (fila ${inicio.fila})` : ""),
    });
  }

  const anexoI = wb.getWorksheet("Anexo I");
  if (anexoI) {
    puntos.push({
      que: "Altas y bajas de bienes de uso del período",
      donde: "hoja Anexo I, columnas de aumentos/bajas (las amortizaciones sí salen de SALDOS)",
    });
  }

  const cc = pdColumnasCentroDeCosto(wb);
  if (cc.length) {
    puntos.push({
      que: "Reparto de gastos por centro de costo",
      donde: `hoja Anexo II, columnas ${cc.map(x => x.columna + " (" + x.nombre + ")").join(", ")}`,
    });
  }

  const imp = pdImpuestoGanancias(wb);
  if (imp && !imp.esFormula) {
    puntos.push({
      que: "Impuesto a las ganancias" +
           (imp.valor === null ? " (hoy está vacío, no se distingue de un cero intencional)"
                               : ` (hoy dice ${imp.valor})`),
      donde: imp.celda,
    });
  }

  return { archivo: etiquetaArchivo, puntos };
}

if (typeof module !== "undefined") {
  module.exports = {
    pendientesManuales, pdBuscarEtiqueta, pdColumnasCentroDeCosto, pdImpuestoGanancias,
  };
}
