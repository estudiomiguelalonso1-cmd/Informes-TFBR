// Renglones que un archivo no tiene y necesita, porque una cuenta no tiene dónde ir.
//
// No se agregan todos los renglones que le falten a un archivo respecto de los otros. Eso
// serían 75, y la mayoría no sirve: texto del pie del informe que el lector de rótulos toma
// por un renglón ("Las notas y anexos adjuntos forman parte…"), el mismo concepto escrito de
// dos formas ("- Anticipo de vacaciones" y "- Anticipos de Vacaciones"), y los renglones de
// préstamo que prestamos.js acaba de unificar — volver a crearlos desharía ese trabajo.
//
// Se agregan sólo los que hacen falta para que ninguna cuenta quede sin renglón propio. Hoy
// son cuatro, y están acá con nombre y apellido en vez de deducidos: crear una fila mueve
// todas las fórmulas de abajo, y eso no se hace por una regla general que algún mes podría
// disparar sola.
//
// Dónde se inserta, y cómo llega al total. Hay dos formas de totalizar en estos archivos, y
// cada una necesita algo distinto:
//
//  - Con un SUM de rango (Activo, Pasivo, Anexo II). Se inserta DENTRO del rango y Excel lo
//    estira solo: un SUM(D10:D18) se estira insertando en la 18, no en la 19. Una fila
//    agregada fuera del rango entra en SALDOS y no en el total, sin que nada lo muestre.
//
//  - Con una suma de celdas nombradas una por una (el EERR: "+C21+C22+C23+C25+C24"). Ahí no
//    hay rango que estirar: la fila nueva no la suma nadie hasta que se le agrega su término
//    al total. Se agrega con el MISMO SIGNO que tiene el renglón de referencia, que es el
//    vecino del que se copió la ubicación.

const RENGLONES_A_AGREGAR = [
  // El EERR del Acumulado R$ tenía un solo renglón "Otros ingresos y egresos" que juntaba
  // tres cuentas que los otros archivos abren por separado.
  { archivo: "balance_acumulado_brl", hoja: "EERR",   rotulo: "Otros ingresos",
    despuesDe: "Intereses" },
  { archivo: "balance_acumulado_brl", hoja: "EERR",   rotulo: "Recupero de gastos",
    despuesDe: "Otros ingresos" },
  { archivo: "balance_acumulado_brl", hoja: "Activo", rotulo: "- Intereses a Devengar Plazo Fijo",
    despuesDe: "- Fondo común de inversión BBVA" },

  // "- Adelanto de sueldos": lo tiene el Acumulado $ y le falta a los otros tres.
  { archivo: "balance_mensual_ars",    hoja: "Activo", rotulo: "- Adelanto de sueldos",
    despuesDe: ["- Adelanto SAC", "- Adelantos PDT"] },
  { archivo: "balance_mensual_brl",    hoja: "Activo", rotulo: "- Adelanto de sueldos",
    despuesDe: ["- Adelanto SAC", "- Adelantos PDT"] },
  { archivo: "balance_acumulado_brl",  hoja: "Activo", rotulo: "- Adelanto de sueldos",
    despuesDe: ["- Adelanto de sac", "- Adelantos PDT"] },

  // "- Anticipo de vacaciones": lo tienen los dos Mensuales y les falta a los Acumulados.
  { archivo: "balance_acumulado_ars",  hoja: "Activo", rotulo: "- Anticipo de vacaciones",
    despuesDe: ["- Adelanto de sueldos", "- Adelantos PDT"] },
  { archivo: "balance_acumulado_brl",  hoja: "Activo", rotulo: "- Anticipo de vacaciones",
    despuesDe: ["- Adelanto de sueldos", "- Adelantos PDT"] },
];

// El subtotal más chico que contiene a esa fila: el SUM de una columna que la abarca.
function rfSubtotalQueContiene(ws, fila) {
  let mejor = null;
  ws.eachRow({ includeEmpty: false }, (row, r) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      const v = cell.value;
      if (!v || typeof v !== "object" || typeof v.formula !== "string") return;
      const m = /^SUM\(\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)\)$/i.exec(v.formula);
      if (!m || m[1].toUpperCase() !== m[3].toUpperCase()) return;
      const desde = Math.min(+m[2], +m[4]), hasta = Math.max(+m[2], +m[4]);
      if (fila < desde || fila > hasta) return;
      if (!mejor || (hasta - desde) < (mejor.hasta - mejor.desde)) mejor = { desde, hasta, total: r };
    });
  });
  return mejor;
}

// El total que suma esa fila nombrándola: "+C21+C22+C23+C25+C24" suma la C24. Devuelve dónde
// está y con qué signo aparece la fila, para agregar el término nuevo igual.
function rfTotalQueSuma(ws, fila, col) {
  const letra = ctColNumeroALetra(col);
  const re = new RegExp("([+-]?)\\s*(?<![A-Z0-9_$!.])\\$?" + letra + "\\$?" + fila + "(?!\\d)");
  let encontrado = null;
  ws.eachRow({ includeEmpty: false }, (row, r) => {
    if (r === fila || encontrado) return;
    row.eachCell({ includeEmpty: false }, (cell, c) => {
      const v = cell.value;
      if (encontrado || !v || typeof v !== "object" || typeof v.formula !== "string") return;
      if (/^SUM\(/i.test(v.formula)) return;
      // Sólo el total de la misma columna: otra columna es otra cosa.
      if (c !== col) return;
      const m = re.exec(v.formula);
      if (!m) return;
      encontrado = { fila: r, col: c, formula: v.formula, signo: m[1] === "-" ? -1 : 1 };
    });
  });
  return encontrado;
}

// Crea un renglón con ese rótulo, después de `despuesDe`, y lo deja sumando en el total de
// su sección. Idempotente: si el rótulo ya está, no hace nada.
function crearRenglon(wb, layout, hoja, rotulo, despuesDe, log = () => {}) {
  const mapa = chMapaHoja(wb, layout, hoja);
  if (!mapa) return { hecho: false, motivo: "el archivo no tiene esa hoja" };

  const yaEsta = mapa.renglones.find(x => chNorm(x.rotulo) === chNorm(rotulo));
  if (yaEsta) return { hecho: false, yaEstaba: true, fila: yaEsta.fila };

  const anclas = Array.isArray(despuesDe) ? despuesDe : [despuesDe];
  const ancla = anclas.map(a => mapa.renglones.find(x => chNorm(x.rotulo) === chNorm(a))).find(Boolean);
  if (!ancla) return { hecho: false, motivo: `no encontré dónde ubicarlo (${anclas.join(" / ")})` };

  const sub = rfSubtotalQueContiene(mapa.ws, ancla.fila);
  const colImporte = ancla.cols.length ? ancla.cols[0].col : mapa.colsImporte[0];
  const total = sub ? null : rfTotalQueSuma(mapa.ws, ancla.fila, colImporte);
  if (!sub && !total) {
    return { hecho: false, motivo: `"${ancla.rotulo}" no está en ningún subtotal ni en ninguna suma` };
  }

  const colRotulo = chCeldaDelRotulo(mapa.ws, ancla.fila, colImporte);
  if (!colRotulo) return { hecho: false, motivo: "no ubiqué la columna del rótulo" };

  const fila = sub ? Math.min(ancla.fila + 1, sub.hasta) : ancla.fila + 1;
  const modificadas = insertRowEn(wb, hoja, fila);
  // El renglón nuevo tiene que verse igual que sus vecinos. Insertar deja la fila sin formato,
  // y en el informe impreso eso se nota: otra letra y el importe sin separador de miles. El
  // modelo es el ancla, que después de insertar quedó una fila más arriba o en su lugar.
  copiarFormatoDeFila(mapa.ws, ancla.fila >= fila ? ancla.fila + 1 : ancla.fila, fila);
  mapa.ws.getCell(fila, colRotulo.col).value = rotulo;
  mapa.ws.getCell(fila, colImporte).value = 0;
  mapa.ws.getRow(fila).hidden = false;

  let comoSuma;
  if (sub) {
    comoSuma = `dentro del subtotal ${sub.desde}-${sub.hasta}`;
  } else {
    // La inserción corrió el total: se lo vuelve a buscar en vez de confiar en dónde estaba.
    const t = rfTotalQueSuma(mapa.ws, ancla.fila, colImporte);
    if (!t) {
      log(`  ⚠ ${hoja}: agregué "${rotulo}" en la fila ${fila} pero NO pude sumarla al total.`);
      return { hecho: true, fila, sinSumar: true };
    }
    const letra = ctColNumeroALetra(colImporte);
    mapa.ws.getCell(t.fila, t.col).value = { formula: `${t.formula}${t.signo < 0 ? "-" : "+"}${letra}${fila}` };
    comoSuma = `sumada en el total de la fila ${t.fila}`;
  }

  log(`  ${hoja}: renglón "${rotulo}" agregado en la fila ${fila}, después de ` +
      `"${ancla.rotulo}" y ${comoSuma} (${modificadas} fórmula(s) reacomodadas).`);
  return { hecho: true, fila, comoSuma };
}

function agregarRenglonesFaltantes(wb, layout, archivoId, log = () => {}) {
  const mios = RENGLONES_A_AGREGAR.filter(r => r.archivo === archivoId);
  const agregados = [], salteados = [];
  for (const r of mios) {
    const res = crearRenglon(wb, layout, r.hoja, r.rotulo, r.despuesDe, log);
    if (res.yaEstaba) continue;
    if (!res.hecho) { salteados.push({ ...r, motivo: res.motivo }); log(`  ⚠ No agregué "${r.rotulo}" en ${r.hoja}: ${res.motivo}.`); continue; }
    agregados.push({ hoja: r.hoja, fila: res.fila, rotulo: r.rotulo, sinSumar: res.sinSumar });
  }
  return { agregados, salteados };
}

// Ninguna fila oculta, en ninguna hoja.
//
// Las plantillas venían con filas escondidas para no imprimir ceros, y eso escondía cosas que
// había que ver: el préstamo de Cuba restando en el Pasivo estuvo meses en una fila oculta, y
// el renglón donde se juntan los adelantos al personal tampoco se veía. Un informe que esconde
// renglones obliga a confiar en que lo escondido está bien.
//
// Se corre al final, después de todas las inserciones: insertar filas reacomoda las marcas de
// oculto, así que hacerlo antes no queda.
function mostrarRenglonesConImporte(wb, layout, planDeCuentas, escritas, log = () => {}) {
  const mostradas = [];
  for (const ws of wb.worksheets) {
    ws.eachRow({ includeEmpty: true }, (row, r) => {
      if (!row.hidden) return;
      row.hidden = false;
      mostradas.push({ hoja: ws.name, fila: r });
    });
  }
  if (mostradas.length) {
    const porHoja = {};
    mostradas.forEach(x => { porHoja[x.hoja] = (porHoja[x.hoja] || 0) + 1; });
    log(`  Filas ocultas: se mostraron ${mostradas.length} (` +
        Object.entries(porHoja).map(([h, n]) => `${h}: ${n}`).join(", ") + ").");
  }
  return mostradas;
}

// Que todo el bloque se vea parejo, no sólo las filas nuevas.
//
// Las plantillas originales ya traían filas sin formato —casi todas ocultas, en cero— y eso se
// nota en cuanto una de ellas recibe un importe y se muestra: queda con otra letra y el número
// sin separador de miles, en el medio de un bloque prolijo. Se le copia el formato al vecino
// con formato más cercano, mirando primero hacia arriba.
//
// Sólo se completa lo que está VACÍO: una celda que ya tiene fuente o formato de número se
// respeta tal cual. Los totales en negrita, los títulos de sección y las columnas de porcentaje
// del Anexo I están así a propósito, y emparejarlos sería romper el diseño en vez de arreglarlo.
function uniformarFormatoDeBloques(wb, layout, log = () => {}) {
  const emparejadas = [];

  for (const ws of wb.worksheets) {
    if (ws.name === layout.sheet) continue;
    const mapa = chMapaHoja(wb, layout, ws.name);
    if (!mapa || mapa.renglones.length < 2) continue;

    const filas = mapa.renglones.map(r => r.fila).sort((a, b) => a - b);
    const desde = filas[0];

    // El bloque se estira hacia abajo mientras haya filas con contenido. Un renglon que el
    // motor agrega DESPUES del ultimo del mapa —el caso de "Alquileres" al final del Anexo II
    // del Acumulado $— caia fuera del rango y se quedaba sin tipografia. Se corta en la
    // primera fila vacia, que es donde termina el cuadro.
    let hasta = filas[filas.length - 1];
    while (hasta < ws.rowCount) {
      let tieneAlgo = false;
      for (let c = 1; c <= ws.columnCount && !tieneAlgo; c++) {
        if (ws.getCell(hasta + 1, c).value != null) tieneAlgo = true;
      }
      if (!tieneAlgo) break;
      hasta++;
    }
    // Las columnas del bloque: la del rótulo y las de importe.
    const cols = new Set(mapa.colsImporte);
    for (const r of mapa.renglones) {
      const c = chCeldaDelRotulo(ws, r.fila, r.cols.length ? r.cols[0].col : mapa.colsImporte[0]);
      if (c) cols.add(c.col);
    }

    // Y cualquier otra columna con contenido dentro del bloque. La del total por fila del
    // Anexo II —"SUM(D9:F9)", la suma de los tres centros de costo— no es ni el rotulo ni un
    // centro de costo, asi que no entraba por ninguno de los dos caminos de arriba: las filas
    // que agrega el motor quedaban ahi sin tipografia, y Excel las dibuja en la fuente por
    // defecto del libro, que no es la del informe. Copiarle el formato a una celda que no
    // tiene ninguno no puede romper nada.
    for (let col = 1; col <= ws.columnCount; col++) {
      if (cols.has(col)) continue;
      for (let f = desde; f <= hasta; f++) {
        if (ws.getCell(f, col).value != null) { cols.add(col); break; }
      }
    }

    for (const col of cols) {
      // El formato de referencia va arrastrándose fila por fila: así una celda vacía toma el de
      // la última con formato que quedó arriba, que es la vecina de verdad y no una de otra
      // sección diez filas más abajo.
      let ultimoConFormato = null;
      const pendientes = [];
      for (let f = desde; f <= hasta; f++) {
        const st = ws.getCell(f, col).style;
        const tiene = st && (st.font || st.numFmt);
        if (tiene) {
          ultimoConFormato = st;
          // Lo que venía sin formato arriba toma el primero que aparece hacia abajo.
          for (const p of pendientes) {
            ws.getCell(p, col).style = JSON.parse(JSON.stringify(st));
            emparejadas.push({ hoja: ws.name, fila: p, col });
          }
          pendientes.length = 0;
          continue;
        }
        if (ultimoConFormato) {
          ws.getCell(f, col).style = JSON.parse(JSON.stringify(ultimoConFormato));
          emparejadas.push({ hoja: ws.name, fila: f, col });
        } else {
          pendientes.push(f);            // todavía no vimos ninguna con formato
        }
      }
    }
  }

  if (emparejadas.length) {
    const porHoja = {};
    emparejadas.forEach(x => { porHoja[x.hoja] = (porHoja[x.hoja] || 0) + 1; });
    log(`  Formato: ${emparejadas.length} celda(s) sin formato tomaron el de su vecina (` +
        Object.entries(porHoja).map(([h, n]) => `${h}: ${n}`).join(", ") + ").");
  }
  return emparejadas;
}

if (typeof module !== "undefined") {
  const ch = require("./config_hojas.js");
  global.chMapaHoja = ch.chMapaHoja;
  global.chNorm = ch.chNorm;
  global.chCeldaDelRotulo = ch.chCeldaDelRotulo;
  const fh = require("./formula_hojas.js");
  global.insertRowEn = fh.insertRowEn;
  global.copiarFormatoDeFila = fh.copiarFormatoDeFila;
  global.ctColNumeroALetra = require("./config_tfbr.js").ctColNumeroALetra;
  module.exports = {
    RENGLONES_A_AGREGAR, agregarRenglonesFaltantes, crearRenglon, mostrarRenglonesConImporte,
    uniformarFormatoDeBloques,
    rfSubtotalQueContiene, rfTotalQueSuma,
  };
}
