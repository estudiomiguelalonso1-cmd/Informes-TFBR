// Deriva el layout de la hoja SALDOS de un maestro TFBR leyendo sus propias fórmulas —
// mismo principio que informe-c/config_balances.js: la verdad es el archivo, no un JSON
// hardcodeado que se puede desincronizar (los 4 maestros de TFBR tienen 4 layouts
// DISTINTOS, confirmado en docs/formula_analysis.md y docs/defect_diagnosis.md, así que
// no hay una única geometría que asumir).
//
// A diferencia de SCA (paste-zone y hub en hojas separadas), en TFBR la zona de pegado
// vive DENTRO de la propia hoja SALDOS, más abajo del plan de cuentas: el patrón de cada
// fila de cuenta es
//   Deudor  = IF(<saldoCol><fila> > 0, <saldoCol><fila>, 0)
//   Acreedor = IF(<saldoCol><fila> < 0, -<saldoCol><fila>, 0)
//   Saldo   = IFERROR(VLOOKUP(<keyCol><fila>, <staging_range>, 2, FALSE), 0)
// y el rango de staging (`<staging_range>`) es el segundo argumento del VLOOKUP.

const CT_RE_VLOOKUP = /VLOOKUP\(\s*\$?([A-Z]{1,3})\$?(\d+)\s*,\s*\$([A-Z]{1,3})\$(\d+):\$([A-Z]{1,3})\$(\d+)\s*,\s*2\s*,\s*(?:FALSE|0)\s*\)/i;
const CT_RE_IF_POS = /^IF\(\$?([A-Z]{1,3})\$?(\d+)\s*>\s*0\s*,/i;
const CT_RE_IF_NEG = /^IF\(\$?([A-Z]{1,3})\$?(\d+)\s*<\s*0\s*,/i;
const CT_RE_CUENTA = /^\s*(\d{6,})/;

function ctColLetraANumero(letras) {
  return letras.toUpperCase().split("").reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0);
}
function ctColNumeroALetra(n) {
  let s = "";
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function ctFormulaDe(ws, fila, col) {
  const v = ws.getCell(fila, col).value;
  if (v && typeof v === "object" && typeof v.formula === "string") return v.formula;
  return null;
}

function ctTextoDe(ws, fila, col) {
  const v = ws.getCell(fila, col).value;
  if (v === null || v === undefined) return "";
  if (typeof v === "object") {
    if (v.richText) return v.richText.map(t => t.text).join("");
    return "";
  }
  return String(v);
}

// Recorre la hoja buscando la primera fila cuya fórmula matchea el patrón IFERROR(VLOOKUP(...)).
// Esa fórmula sola alcanza para derivar keyCol, saldoCol y staging_range de una sola vez.
function ctEncontrarPatronVlookup(ws) {
  for (let r = 1; r <= ws.rowCount; r++) {
    for (let c = 1; c <= Math.min(ws.columnCount, 15); c++) {
      const f = ctFormulaDe(ws, r, c);
      if (!f) continue;
      const m = CT_RE_VLOOKUP.exec(f);
      if (!m) continue;
      const [, keyColRef, keyFilaRef, stagDesdeCol, stagDesdeFila, stagHastaCol, stagHastaFila] = m;
      if (parseInt(keyFilaRef, 10) !== r) continue; // el VLOOKUP de una fila de cuenta se busca a si misma
      return {
        filaEjemplo: r,
        saldoCol: c,
        keyCol: ctColLetraANumero(keyColRef),
        stagingRange: {
          colDesde: ctColLetraANumero(stagDesdeCol), filaDesde: parseInt(stagDesdeFila, 10),
          colHasta: ctColLetraANumero(stagHastaCol), filaHasta: parseInt(stagHastaFila, 10),
        },
      };
    }
  }
  return null;
}

// Busca, en la misma fila de ejemplo, las columnas IF(saldoCol>0,...) / IF(saldoCol<0,...).
function ctEncontrarDeudorAcreedor(ws, fila, saldoCol) {
  let deudorCol = null, acreedorCol = null;
  for (let c = 1; c <= Math.min(ws.columnCount, 15); c++) {
    const f = ctFormulaDe(ws, fila, c);
    if (!f) continue;
    const mPos = CT_RE_IF_POS.exec(f);
    if (mPos && ctColLetraANumero(mPos[1]) === saldoCol && parseInt(mPos[2], 10) === fila) deudorCol = c;
    const mNeg = CT_RE_IF_NEG.exec(f);
    if (mNeg && ctColLetraANumero(mNeg[1]) === saldoCol && parseInt(mNeg[2], 10) === fila) acreedorCol = c;
  }
  return { deudorCol, acreedorCol };
}

// El rango de filas del plan de cuentas (para el matching por código): desde la primera
// fila con el patron VLOOKUP hasta la ultima antes de que empiece la zona de staging.
function ctRangoPlanDeCuentas(ws, keyCol, saldoCol, stagingRange) {
  let desde = null, hasta = null;
  const limiteFila = stagingRange.filaDesde - 1;
  for (let r = 1; r <= limiteFila; r++) {
    const f = ctFormulaDe(ws, r, saldoCol);
    if (!f || !CT_RE_VLOOKUP.test(f)) continue;
    if (desde === null) desde = r;
    hasta = r;
  }
  return { desde, hasta };
}

// Deriva el layout completo de una hoja SALDOS ya abierta con ExcelJS.
function derivarLayoutSaldos(wb) {
  const ws = wb.getWorksheet("SALDOS");
  if (!ws) throw new Error("El maestro no tiene una hoja 'SALDOS'.");

  const patron = ctEncontrarPatronVlookup(ws);
  if (!patron) {
    throw new Error(
      "No encontré ninguna fórmula IFERROR(VLOOKUP(...)) en 'SALDOS' para derivar el layout. " +
      "¿Es un maestro de TFBR sin tocar?"
    );
  }

  const { deudorCol, acreedorCol } = ctEncontrarDeudorAcreedor(ws, patron.filaEjemplo, patron.saldoCol);
  if (!deudorCol || !acreedorCol) {
    throw new Error(
      `Encontré el VLOOKUP en SALDOS!${ctColNumeroALetra(patron.saldoCol)}${patron.filaEjemplo} pero no ` +
      "las fórmulas IF(...>0,...) / IF(...<0,...) de Deudor/Acreedor en la misma fila."
    );
  }

  const plan = ctRangoPlanDeCuentas(ws, patron.keyCol, patron.saldoCol, patron.stagingRange);

  return {
    sheet: "SALDOS",
    keyCol: patron.keyCol,
    deudorCol,
    acreedorCol,
    saldoCol: patron.saldoCol,
    stagingRange: patron.stagingRange,
    planDeCuentas: plan, // {desde, hasta}: filas de la plantilla fija, para matching por codigo
  };
}

// Lee el texto clave de cada fila del plan de cuentas -> {codigo: {fila, texto}}.
// El texto (no solo el codigo) es lo que hay que volver a escribir en el staging para
// garantizar que el VLOOKUP matchee (ver la nota de diseño en docs/formula_analysis.md
// sobre por que NO alcanza con pegar el texto tal cual lo exporta Onvio).
// Devuelve { cuentas, duplicadas }.
//
// Un código puede aparecer en más de una fila, y no es teórico: en el Acumulado R$ la cuenta
// 4230400000 "Intereses resarcitorios" está en la fila 199 escrita con dos espacios después
// del código y en la 208 con uno solo. El BUSCARV busca el TEXTO, así que la 199 levanta su
// saldo y la 208 queda en cero para siempre, sin que nada avise.
//
// Se usa la PRIMERA fila para escribir (una sola de las dos puede levantar el importe) y se
// devuelven las repetidas para avisarlas: cuál de las dos sobra es una decisión contable.
function leerPlanDeCuentas(wb, layout) {
  const ws = wb.getWorksheet(layout.sheet);
  const cuentas = {};
  const duplicadas = [];
  for (let r = layout.planDeCuentas.desde; r <= layout.planDeCuentas.hasta; r++) {
    const texto = ctTextoDe(ws, r, layout.keyCol).trim();
    if (!texto) continue;
    const m = CT_RE_CUENTA.exec(texto);
    if (!m) continue;
    const codigo = m[1];
    // La columna del BUSCARV se guarda POR FILA y no se da por sentada: no todas las
    // secciones usan la misma (en RESULTADOS puede correrse), y los controles necesitan
    // leer el saldo en la columna que esa fila usa de verdad.
    const ficha = { fila: r, texto, colSaldo: ctColSaldoDeFila(ws, r, layout) };
    if (cuentas[codigo]) {
      duplicadas.push({ codigo, fila: r, texto, filaPrevia: cuentas[codigo].fila, textoPrevio: cuentas[codigo].texto });
      cuentas[codigo].otrasFilas = (cuentas[codigo].otrasFilas || []).concat([ficha]);
    } else {
      cuentas[codigo] = ficha;
    }
  }
  return { cuentas, duplicadas };
}

// La columna donde esta fila tiene su BUSCARV. Si no encuentra ninguna, cae en la del layout.
function ctColSaldoDeFila(ws, fila, layout) {
  for (let c = 1; c <= Math.min(ws.columnCount, 15); c++) {
    const f = ctFormulaDe(ws, fila, c);
    if (f && CT_RE_VLOOKUP.test(f)) return c;
  }
  return layout.saldoCol;
}

if (typeof module !== "undefined") {
  module.exports = {
    derivarLayoutSaldos, leerPlanDeCuentas,
    ctColLetraANumero, ctColNumeroALetra,
  };
}
