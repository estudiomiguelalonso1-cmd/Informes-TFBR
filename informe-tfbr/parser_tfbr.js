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

// El tipo de cambio de cierre, cuando el export lo trae.
//
// El de agosto de 2026 lo incluye en el encabezado ("TC 31/8" y al lado 291.3301). El de julio
// no lo traía y había que tipearlo. Leerlo de acá saca un dato manual del cierre, y además hace
// falta para verificar la columna de saldos en reales (ver más abajo).
function ptBuscarTcCierre(filas) {
  for (let r = 0; r < Math.min(filas.length, 14); r++) {
    const f = filas[r] || [];
    for (let c = 0; c < f.length; c++) {
      if (typeof f[c] !== "string") continue;
      if (!/^T\.?\s*C\.?\b/i.test(f[c].trim())) continue;
      for (let d = 1; d <= 3; d++) {
        if (typeof f[c + d] === "number" && f[c + d] > 0) {
          return { valor: f[c + d], etiqueta: f[c].trim(), fila: r, col: c + d };
        }
      }
    }
  }
  return null;
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

  const tc = ptBuscarTcCierre(filas);

  const cols = {
    debe_ars: ptUbicarColumna(filas, merges, filaEnc, filasDeCuenta, "Debe ($)"),
    haber_ars: ptUbicarColumna(filas, merges, filaEnc, filasDeCuenta, "Haber ($)"),
    saldo_ars: ptUbicarColumna(filas, merges, filaEnc, filasDeCuenta, "Saldo ($)"),
    debe_brl: ptUbicarColumna(filas, merges, filaEnc, filasDeCuenta, "Debe (R$)"),
    haber_brl: ptUbicarColumna(filas, merges, filaEnc, filasDeCuenta, "Haber (R$)"),
    saldo_brl: ptUbicarColumna(filas, merges, filaEnc, filasDeCuenta, "Saldo (R$)"),
  };

  // El control de que las columnas sean las que decimos que son.
  //
  // Se mira SÓLO en pesos: ahí el saldo siempre es debe menos haber, y si las columnas
  // estuvieran corridas los tres números no tendrían relación entre sí. Alcanza con eso para
  // atrapar un cambio de formato.
  //
  // En reales NO se controla la resta, y no es que el archivo esté mal: el saldo en reales de
  // las cuentas patrimoniales no se usa (ver convertirSaldosEnReales). El control anterior
  // exigía la resta en las dos monedas y rechazaba el export de agosto de 2026 entero, con un
  // mensaje que hacía pensar en un archivo roto.
  {
    const cd = cols.debe_ars, ch = cols.haber_ars, cs = cols.saldo_ars;
    if (cd !== null && ch !== null && cs !== null) {
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
          `Las columnas de pesos del export no cierran: el saldo tendría que ser debe menos ` +
          `haber y sólo da en ${cierran} de ${miradas} cuentas. Puede que el reporte de Onvio ` +
          `haya cambiado de formato. NO se cargó nada.`
        );
      }
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

  // Onvio deja pedir el reporte con las cuentas saldadas incluidas, y hay que pedirlo asi.
  //
  // El reporte corto omite las cuentas con saldo CERO, y "cero" lo mide en pesos. Hay cuentas
  // de resultado que cierran en cero en pesos y NO en reales, porque los movimientos del anio
  // se convirtieron a tipos de cambio distintos y no se cancelan: en agosto de 2026,
  // "4212100000 FLETES" tiene $ 0,00 y R$ 3.356,67. Con el reporte corto esa plata no llega al
  // informe en reales y el unico sintoma es la diferencia de cambio saliendo corta.
  //
  // Se reconoce por lo que le falta: en el reporte largo hay cientos de cuentas con el saldo en
  // pesos en cero; en el corto no hay ninguna, porque son justo las que saca.
  const enCeroEnPesos = cuentas.filter(c => !c.saldo_ars).length;
  const pareceReporteCorto = cuentas.length > 0 && enCeroEnPesos === 0;

  return {
    cuentas,
    columnas: cols,
    tcCierre: tc,
    pareceReporteCorto,
    cuentasEnCeroEnPesos: enCeroEnPesos,
    filaEncabezados: filaEnc,
    discrepanciasCapitulo: discrepancias,
    totales: {
      saldo_ars: cuentas.reduce((s, c) => s + c.saldo_ars, 0),
      saldo_brl: cuentas.reduce((s, c) => s + c.saldo_brl, 0),
    },
  };
}

// Un número escrito a mano, en cualquiera de las dos formas que se usan acá.
//
// El tipo de cambio se tipea como "291,3301" —coma decimal, que es como se escribe en
// Argentina— y Number() de eso da NaN. El sistema lo tomaba como que faltaba el dato y frenaba
// el cierre diciendo "falta el tipo de cambio" con el número a la vista en la pantalla.
//
// Cuál es el separador decimal se decide por el ÚLTIMO que aparece: "1.234,56" es mil
// doscientos treinta y cuatro con cincuenta y seis, y "1,234.56" también. Lo que queda antes
// son separadores de miles y se borran. Con un solo separador y exactamente tres dígitos
// detrás ("1.234") no hay forma de saberlo mirando el número, así que se toma como miles, que
// es lo que significa cuando alguien lo escribe así.
function ptNumeroTipeado(texto) {
  const t = String(texto == null ? "" : texto).trim().replace(/\s/g, "");
  if (!t) return null;
  if (!/[\d]/.test(t)) return null;

  const ultimaComa = t.lastIndexOf(",");
  const ultimoPunto = t.lastIndexOf(".");
  let limpio;

  if (ultimaComa >= 0 && ultimoPunto >= 0) {
    const dec = ultimaComa > ultimoPunto ? "," : ".";
    const miles = dec === "," ? "." : ",";
    limpio = t.split(miles).join("").replace(dec, ".");
  } else if (ultimaComa >= 0 || ultimoPunto >= 0) {
    const sep = ultimaComa >= 0 ? "," : ".";
    const partes = t.split(sep);
    const detras = partes[partes.length - 1].length;
    // Varios separadores, o uno con tres dígitos detrás: son miles.
    limpio = (partes.length > 2 || detras === 3)
      ? partes.join("")
      : partes.slice(0, -1).join("") + "." + partes[partes.length - 1];
  } else {
    limpio = t;
  }

  const n = Number(limpio);
  return isFinite(n) ? n : null;
}

// Redondeo a centavos, como lo hace la planilla: el medio centavo va para afuera del cero,
// no al par más cercano. Con Math.round a secas, -0.005 daría -0.00 y 0.005 daría 0.01, y los
// negativos quedarían un centavo corridos respecto del reporte.
function ptRedondearCentavos(x) {
  const signo = x < 0 ? -1 : 1;
  return signo * Math.round(Math.abs(x) * 100) / 100;
}

// El saldo en reales que usan los informes, a partir del export ORIGINAL del sistema.
//
// La regla es de contaduría, no del programa:
//
//   Cuentas 1, 2 y 3 (activo, pasivo, patrimonio): NO se usa el saldo en reales que trae el
//   export. Se toma el saldo en PESOS y se divide por el tipo de cambio de cierre. Son cuentas
//   patrimoniales: valen lo que valen a la fecha de cierre, no la suma de los movimientos
//   convertidos cada uno al cambio de su día.
//
//   Cuentas 4 (resultados): se toma el saldo en reales tal cual viene. Un resultado es la
//   acumulación de lo que pasó durante el período, cada movimiento al cambio de su momento, y
//   reexpresarlo al cambio de cierre lo cambiaría de sentido.
//
// Las columnas de debe y haber no se usan para nada de esto.
//
// Verificado contra el export de agosto de 2026 ya convertido por contaduría: las 144 cuentas
// de los dos archivos dan igual al centavo.
function convertirSaldosEnReales(cuentas, tipoDeCambio) {
  const tc = typeof tipoDeCambio === "number" ? tipoDeCambio : ptNumeroTipeado(tipoDeCambio);
  if (!tc || !isFinite(tc) || tc <= 0) {
    throw new Error(
      "Falta el tipo de cambio de cierre: sin él no se puede calcular el saldo en reales de " +
      "las cuentas de activo, pasivo y patrimonio. NO se procesó nada."
    );
  }
  return (cuentas || []).map(c => {
    const rubro = String(c.codigo).trim()[0];
    if (rubro === "4") return { ...c, saldo_brl_export: c.saldo_brl };
    return {
      ...c,
      saldo_brl_export: c.saldo_brl,
      saldo_brl: ptRedondearCentavos(c.saldo_ars / tc),
    };
  });
}

function capituloDeCodigoTFBR(codigo) {
  return CAPITULO_POR_DIGITO_TFBR[String(codigo).trim()[0]] || null;
}

if (typeof module !== "undefined") {
  module.exports = {
    parseSumasYSaldosTFBR, capituloDeCodigoTFBR, ptUbicarColumna, ptBuscarTcCierre,
    convertirSaldosEnReales, ptRedondearCentavos, ptNumeroTipeado,
    CAPITULOS_TFBR, CAPITULO_POR_DIGITO_TFBR,
  };
}
