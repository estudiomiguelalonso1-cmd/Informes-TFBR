// Lee el "Balance de Sumas y Saldos" que TFBR exporta de su sistema contable (Onvio),
// hoja "Sheet1". Trae las dos monedas en la misma planilla (a diferencia del Informe A
// de SCA, que usa un export de una sola moneda): columnas "Debe ($)/Haber ($)/Saldo ($)"
// y "Debe (R$)/Haber (R$)/Saldo (R$)", agrupadas bajo tres capítulos (ACTIVO, PASIVO,
// RESULTADOS — TFBR no trae un capítulo de PATRIMONIO NETO en este export; los
// movimientos de PN se cargan aparte, a mano, en EEPN).
//
// Mismo criterio que informe-c/parser_balances.js: las columnas se ubican por el TEXTO
// del encabezado (nunca por posición fija) y se verifica con Debe-Haber=Saldo antes de
// confiar en lo que se encontró — un export mal formateado tiene que frenar el proceso,
// no generar un balance en cero en silencio.

const CAPITULOS_TFBR = ["ACTIVO", "PASIVO", "RESULTADOS"];
const CAPITULO_POR_DIGITO_TFBR = { 1: "ACTIVO", 2: "PASIVO", 4: "RESULTADOS" };

// "1110100330 - FONDO FIJO BS. AS."  ->  { codigo: "1110100330", nombre: "FONDO FIJO BS. AS." }
const RE_CUENTA_TFBR = /^\s*(\d{6,})\s*-\s*(.+?)\s*$/;

function ptNormTexto(s) {
  if (s === null || s === undefined) return "";
  return String(s).replace(/\s+/g, " ").trim().toUpperCase();
}

function ptNumero(v) {
  return typeof v === "number" ? v : 0;
}

// Igual algoritmo que ubicarColumna de informe-c: busca el encabezado por texto, y si el
// importe no cae en esa misma columna (por combinación de celdas u otro corrimiento),
// vota entre las filas de cuenta cuál columna cercana trae los números de verdad.
function ptUbicarColumna(filas, merges, filaEnc, filasDeCuenta, textoBuscado) {
  const enc = filas[filaEnc] || [];
  let col = null;
  for (let c = 0; c < enc.length; c++) {
    if (ptNormTexto(enc[c]) === ptNormTexto(textoBuscado)) { col = c; break; }
  }
  if (col === null) return null;

  const m = (merges || []).find(x => x.s.r === filaEnc && x.s.c <= col && col <= x.e.c);
  const desde = m ? m.s.c : col;
  const hasta = m ? m.e.c : col;

  const votos = {};
  for (const f of filasDeCuenta) {
    for (let c = desde; c <= hasta; c++) {
      if (typeof f[c] === "number") votos[c] = (votos[c] || 0) + 1;
    }
  }
  const ganador = Object.entries(votos).sort((a, b) => b[1] - a[1])[0];
  if (ganador) return Number(ganador[0]);

  for (let c = col + 1; c <= col + 4; c++) {
    let cuantos = 0;
    for (const f of filasDeCuenta) if (typeof f[c] === "number") cuantos++;
    if (cuantos >= Math.max(3, filasDeCuenta.length * 0.5)) return c;
  }
  return col;
}

// filas: XLSX.utils.sheet_to_json(ws, {header:1}). merges: ws['!merges'].
function parseSumasYSaldosTFBR(filas, merges) {
  let filaEnc = null;
  for (let r = 0; r < Math.min(filas.length, 40); r++) {
    const textos = (filas[r] || []).map(ptNormTexto);
    if (textos.some(t => t.startsWith("SALDO (")) && textos.some(t => t.startsWith("DEBE ("))) {
      filaEnc = r;
      break;
    }
  }
  if (filaEnc === null) {
    throw new Error(
      'No encontré la fila de encabezados del export (la que dice "Debe ($)", "Saldo (R$)", etc). ' +
      '¿Es el "Balance de Sumas y Saldos" de Onvio?'
    );
  }

  const filasDeCuenta = [];
  for (let r = filaEnc + 1; r < filas.length; r++) {
    const a = (filas[r] || [])[0];
    if (a !== null && a !== undefined && RE_CUENTA_TFBR.test(String(a))) filasDeCuenta.push(filas[r]);
  }
  if (!filasDeCuenta.length) throw new Error("El export no tiene ninguna línea de cuenta.");

  const cols = {
    debe_ars: ptUbicarColumna(filas, merges, filaEnc, filasDeCuenta, "Debe ($)"),
    haber_ars: ptUbicarColumna(filas, merges, filaEnc, filasDeCuenta, "Haber ($)"),
    saldo_ars: ptUbicarColumna(filas, merges, filaEnc, filasDeCuenta, "Saldo ($)"),
    debe_brl: ptUbicarColumna(filas, merges, filaEnc, filasDeCuenta, "Debe (R$)"),
    haber_brl: ptUbicarColumna(filas, merges, filaEnc, filasDeCuenta, "Haber (R$)"),
    saldo_brl: ptUbicarColumna(filas, merges, filaEnc, filasDeCuenta, "Saldo (R$)"),
  };

  for (const moneda of ["ars", "brl"]) {
    const cd = cols["debe_" + moneda], ch = cols["haber_" + moneda], cs = cols["saldo_" + moneda];
    if (cd === null || ch === null || cs === null) continue;
    let miradas = 0, cierran = 0;
    for (const f of filasDeCuenta) {
      const d = typeof f[cd] === "number" ? f[cd] : 0;
      const h = typeof f[ch] === "number" ? f[ch] : 0;
      const sa = typeof f[cs] === "number" ? f[cs] : 0;
      if (!d && !h && !sa) continue;
      miradas++;
      if (Math.abs((d - h) - sa) < 0.02) cierran++;
    }
    if (miradas >= 5 && cierran < miradas * 0.9) {
      throw new Error(
        `Las columnas de ${moneda === "ars" ? "pesos" : "reales"} del export no cierran: ` +
        `el saldo tendría que ser debe menos haber y sólo da en ${cierran} de ${miradas} cuentas. ` +
        `Puede que el reporte de Onvio haya cambiado de formato. NO se cargó nada.`
      );
    }
  }

  const faltan = Object.entries(cols).filter(([, v]) => v === null).map(([k]) => k);
  if (faltan.length) {
    throw new Error(`No pude ubicar en el export las columnas: ${faltan.join(", ")}.`);
  }

  const cuentas = [];
  const discrepancias = [];
  let capituloActual = null;

  for (let r = filaEnc + 1; r < filas.length; r++) {
    const fila = filas[r] || [];
    const a = fila[0];
    if (a === null || a === undefined) continue;
    const texto = String(a).trim();
    if (!texto) continue;

    const comoCapitulo = ptNormTexto(texto);
    if (CAPITULOS_TFBR.includes(comoCapitulo)) { capituloActual = comoCapitulo; continue; }
    if (comoCapitulo.startsWith("TOTALES GENERALES")) continue;

    const m = RE_CUENTA_TFBR.exec(texto);
    if (!m) continue;

    const codigo = m[1];
    const porDigito = CAPITULO_POR_DIGITO_TFBR[codigo[0]] || null;
    if (capituloActual && porDigito && capituloActual !== porDigito) {
      discrepancias.push({ codigo, nombre: m[2], seccion: capituloActual, porDigito });
    }

    cuentas.push({
      codigo,
      nombre: m[2],
      capitulo: capituloActual || porDigito,
      debe_ars: ptNumero(fila[cols.debe_ars]),
      haber_ars: ptNumero(fila[cols.haber_ars]),
      saldo_ars: ptNumero(fila[cols.saldo_ars]),
      debe_brl: ptNumero(fila[cols.debe_brl]),
      haber_brl: ptNumero(fila[cols.haber_brl]),
      saldo_brl: ptNumero(fila[cols.saldo_brl]),
    });
  }

  return {
    cuentas,
    columnas: cols,
    filaEncabezados: filaEnc,
    discrepanciasCapitulo: discrepancias,
    totales: {
      saldo_ars: cuentas.reduce((s, c) => s + c.saldo_ars, 0),
      saldo_brl: cuentas.reduce((s, c) => s + c.saldo_brl, 0),
    },
  };
}

function capituloDeCodigoTFBR(codigo) {
  return CAPITULO_POR_DIGITO_TFBR[String(codigo).trim()[0]] || null;
}

if (typeof module !== "undefined") {
  module.exports = {
    parseSumasYSaldosTFBR, capituloDeCodigoTFBR, ptUbicarColumna,
    CAPITULOS_TFBR, CAPITULO_POR_DIGITO_TFBR,
  };
}
