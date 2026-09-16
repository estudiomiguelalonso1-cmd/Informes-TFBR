// Qué cuenta alimenta cada renglón, en CUALQUIER hoja — no sólo en el Anexo II.
//
// El Anexo II era el único que se podía configurar, pero el problema es el mismo en todas:
// Activo, Pasivo, EERR y EEPN también nombran las cuentas una por una, y una cuenta que no
// está nombrada en ningún renglón no llega a ningún estado. Lo que cambia de una hoja a otra
// es sólo la geometría: en qué columna está el texto del renglón y en cuál el importe.
//
// Nada de eso se escribe a mano acá. Se deduce del archivo, como todo lo demás en este
// proyecto: los cuatro archivos tienen geometrías distintas —el Anexo II del Mensual $ pone
// el importe en D/E/F y el Activo en F— y una tabla escrita a mano se desactualizaría con el
// primer cambio de plantilla.
//
// Cómo se deduce:
//  - Renglón que engancha cuentas = celda con fórmula que nombra filas de SALDOS.
//  - Columna del importe = la columna de esa celda.
//  - Rótulo = el primer TEXTO a su izquierda. Tiene que ser texto de verdad: en el Anexo II
//    el importe va en E y la columna D suele tener un cero, y tomar ese cero como rótulo
//    dejaba la lista entera llamándose "0".

function chTexto(ws, r, c) {
  const v = ws.getCell(r, c).value;
  if (v == null) return "";
  if (typeof v === "object") return v.richText ? v.richText.map(t => t.text).join("") : "";
  return String(v);
}

function chNorm(t) {
  return String(t || "").normalize("NFD").replace(/[̀-ͯ]/g, "")
    .toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();
}

// El rótulo de un renglón: el primer texto a la izquierda de la columna del importe.
// Se saltean números y fórmulas — son importes de otra columna, no el nombre del renglón.
function chRotuloDe(ws, fila, colImporte) {
  for (let c = colImporte - 1; c >= 1; c--) {
    const v = ws.getCell(fila, c).value;
    if (typeof v === "number") continue;
    if (v && typeof v === "object" && typeof v.formula === "string") continue;
    const t = chTexto(ws, fila, c).trim();
    if (!t) continue;
    if (/^-?[\d.,]+$/.test(t)) continue;
    return t;
  }
  return "";
}

// El mapa de una hoja: sus renglones, qué filas de SALDOS lee cada uno, y qué filas del plan
// ya están tomadas.
//
// Un renglón es una FILA, no una celda. En el Anexo II una misma fila tiene hasta tres celdas
// de importe —administración, comercialización, financieros— y las tres son el mismo concepto
// con el mismo rótulo. Tratándolas como renglones distintos, dos entradas comparten nombre y
// no hay forma de decir a cuál se quiso mover una cuenta.
//
// Y entran también los renglones VACÍOS: los que tienen rótulo pero hoy no leen ninguna
// cuenta. Un renglón al que se le saca su última cuenta deja de tener fórmula, y si sólo se
// miraran las fórmulas desaparecería del mapa — con lo cual ya no se podría volver a usar
// como destino, ni siquiera para devolverle la cuenta que se le acaba de sacar.
function chMapaHoja(wb, layout, nombreHoja) {
  const ws = wb.getWorksheet(nombreHoja);
  if (!ws) return null;

  const porFila = {};
  const leidas = new Set();
  const colsImporte = new Set();
  let minFila = Infinity, maxFila = -Infinity;

  ws.eachRow({ includeEmpty: false }, (row, r) => {
    row.eachCell({ includeEmpty: false }, (cell, c) => {
      const v = cell.value;
      if (!v || typeof v !== "object" || typeof v.formula !== "string") return;
      if (!/SALDOS!/.test(v.formula)) return;

      // Sólo cuentan las referencias a las columnas donde vive el saldo de una cuenta: la
      // deudora, la acreedora y la cruda del VLOOKUP. Cualquier otra columna de SALDOS es otra
      // cosa — en el Mensual $ la celda E47 es SUM(D32:D47), un subtotal parado al final del
      // bloque, y el EESP lo lee. Tomándolo por una cuenta, el EESP entraba al editor como si
      // enganchara la fila 47, y mover algo ahí habría reemplazado un subtotal por una cuenta.
      const colsSaldo = new Set([layout.deudorCol, layout.acreedorCol, layout.saldoCol]
        .map(n => ctColNumeroALetra(n)));
      const filas = [];
      for (const m of v.formula.matchAll(/SALDOS!\$?([A-Z]{1,3})\$?(\d+)/g)) {
        if (!colsSaldo.has(m[1].toUpperCase())) continue;
        const f = +m[2];
        if (!filas.includes(f)) filas.push(f);   // la deudora y la acreedora son la misma cuenta
      }
      if (!filas.length) return;
      filas.forEach(f => leidas.add(f));
      colsImporte.add(c);
      minFila = Math.min(minFila, r); maxFila = Math.max(maxFila, r);

      const reng = porFila[r] || (porFila[r] = {
        fila: r, rotulo: chRotuloDe(ws, r, c), cols: [], filasSaldos: [], vacio: false,
      });
      if (!reng.rotulo) reng.rotulo = chRotuloDe(ws, r, c);
      reng.cols.push({ col: c, dir: cell.address });
      for (const f of filas) if (!reng.filasSaldos.includes(f)) reng.filasSaldos.push(f);
    });
  });

  if (!Object.keys(porFila).length) return null;

  // Renglones con rótulo pero sin cuentas, en el tramo donde viven los demás.
  //
  // El tramo se abre un poco a cada lado a propósito. Si se tomara justo de la primera a la
  // última fila CON fórmula, vaciar el primer renglón de la hoja lo dejaría fuera del tramo y
  // el renglón desaparecería para siempre: no se lo podría volver a elegir como destino ni
  // para devolverle la cuenta que se le acaba de sacar. Pasaba con "Amortizaciones", que es el
  // primer renglón del Anexo II del Acumulado R$.
  const colPrincipal = Math.min(...colsImporte);
  const MARGEN = 12;
  const desde = Math.max(1, minFila - MARGEN), hasta = maxFila + MARGEN;
  for (let r = desde; r <= hasta; r++) {
    if (porFila[r]) continue;
    const rot = chRotuloDe(ws, r, colPrincipal);
    if (!rot) continue;
    if (/^total|^conceptos|^subtotal/i.test(rot.trim())) continue;
    const v = ws.getCell(r, colPrincipal).value;
    // Sólo si la celda de importe está libre: un número escrito a mano no es un renglón
    // vacío, es un renglón que no sale de ninguna cuenta y no se toca desde acá.
    if (v != null && !(typeof v === "number" && Math.abs(v) < 0.005)) continue;
    porFila[r] = {
      fila: r, rotulo: rot, cols: [{ col: colPrincipal, dir: `${ctColNumeroALetra(colPrincipal)}${r}` }],
      filasSaldos: [], vacio: true,
    };
  }

  const renglones = Object.values(porFila).sort((a, b) => a.fila - b.fila);
  return { ws, hoja: nombreHoja, renglones, leidas, colsImporte: [...colsImporte].sort((a, b) => a - b) };
}

// Las hojas que enganchan cuentas, en el orden en que están en el archivo. EESP y Bienes no
// aparecen: sólo suman otras hojas, no leen cuentas, y no hay nada que configurar en ellas.
function chHojasConfigurables(wb, layout) {
  const hojas = [];
  for (const ws of wb.worksheets) {
    if (ws.name === layout.sheet) continue;
    const m = chMapaHoja(wb, layout, ws.name);
    if (!m) continue;
    // Sólo cuentan las filas que son del plan: una hoja que lee una fila cargada a mano
    // (una diferencia de cambio, por ejemplo) no tiene cuentas que configurar.
    const delPlan = [...m.leidas].filter(f =>
      f >= layout.planDeCuentas.desde && f <= layout.planDeCuentas.hasta);
    if (!delPlan.length) continue;
    hojas.push({ hoja: ws.name, renglones: m.renglones.length, cuentas: delPlan.length });
  }
  return hojas;
}

// ¿Se puede editar esta hoja moviendo cuentas de un rótulo a otro?
//
// Hace falta que el rótulo identifique el renglón. En el Anexo II sí: una fila es un concepto
// y sus columnas son los centros de costo del mismo concepto. En el Anexo I no: una fila es un
// bien —"Software"— y sus columnas son cosas distintas (valor de origen, altas, bajas,
// amortización del ejercicio), con signos opuestos y cuentas distintas. Mover "Software" ahí
// no quiere decir nada, y elegir una columna al azar mandaría el importe a la sección
// equivocada.
//
// Se detecta sin nombrar ninguna hoja: si dentro de una misma fila dos celdas de importe no
// coinciden en el signo, las columnas no son intercambiables y la hoja no se edita desde acá.
// También descalifica que el mismo rótulo aparezca en dos filas.
function chHojaEditable(wb, layout, nombreHoja) {
  const mapa = chMapaHoja(wb, layout, nombreHoja);
  if (!mapa) return { editable: false, motivo: "no engancha cuentas" };

  for (const r of mapa.renglones) {
    const signos = new Set();
    for (const c of r.cols) {
      const v = mapa.ws.getCell(r.fila, c.col).value;
      if (!v || typeof v !== "object" || typeof v.formula !== "string") continue;
      const s = chSignoDelRenglon(v.formula, layout);
      if (s) signos.add(s);
    }
    if (signos.size > 1) {
      return {
        editable: false,
        motivo: `en la fila ${r.fila} ("${r.rotulo}") las columnas no son el mismo concepto: ` +
                `unas suman el saldo y otras lo restan`,
      };
    }
  }

  // Un rótulo repetido NO descalifica la hoja entera: descalifica ese rótulo. El Anexo II del
  // Mensual $ tiene dos "Amortizaciones" y el Activo dos "- Plazo Fijo", y por dos renglones
  // ambiguos no tiene sentido dejar sin editar las otras noventa cuentas. De esos dos se
  // encarga engancharEnHoja, que se niega a mover una cuenta a un rótulo que aparece dos veces.
  return { editable: true };
}

// El rubro de una hoja: el primer dígito que más se repite entre las cuentas que ya lee.
// 1 activo, 2 pasivo, 3 patrimonio, 4 resultado. Sirve para saber qué cuentas ofrecerle:
// en el Activo no tiene sentido listar cuentas de gasto.
//
// Se deduce en vez de escribirse porque las hojas no se llaman igual en los cuatro archivos y
// porque una hoja puede leer alguna cuenta de otro rubro sin que eso cambie de qué es la hoja.
function chRubroDeHoja(wb, layout, nombreHoja, planPorFila) {
  const m = chMapaHoja(wb, layout, nombreHoja);
  if (!m) return null;
  const votos = {};
  for (const f of m.leidas) {
    const info = planPorFila[f];
    if (!info) continue;
    const d = String(info.cod)[0];
    votos[d] = (votos[d] || 0) + 1;
  }
  const ganador = Object.entries(votos).sort((a, b) => b[1] - a[1])[0];
  return ganador ? ganador[0] : null;
}

// Las cuentas del plan, por fila.
function chPlanPorFila(wb, layout) {
  const S = wb.getWorksheet(layout.sheet);
  const porFila = {};
  for (let r = layout.planDeCuentas.desde; r <= layout.planDeCuentas.hasta; r++) {
    const m = /^\s*([\d.]+)\s*-?\s*(.*)$/.exec(chTexto(S, r, layout.keyCol).trim());
    if (m) porFila[r] = { cod: m[1].replace(/\./g, ""), nom: m[2].trim(), fila: r };
  }
  return porFila;
}

// ¿El renglón suma el saldo tal cual, o invertido?
//
// Esto NO es un detalle de formato. El Activo muestra "deudora menos acreedora"
// (+SALDOS!B7-SALDOS!C7) porque un activo es deudor; el Pasivo y el EERR muestran lo
// contrario (+SALDOS!C83-SALDOS!B83) porque una deuda y un ingreso son acreedores, y ahí se
// muestran en positivo. Enganchar una cuenta nueva con el signo del Anexo II en un renglón
// del Pasivo le restaría el importe en vez de sumárselo: el renglón daría de menos justo el
// doble de la cuenta, y nada lo avisaría.
//
// El signo se lee del propio renglón — de cómo están escritos los términos que ya tiene — en
// vez de decidirse por el nombre de la hoja, que en los cuatro archivos no es igual.
function chSignoDelRenglon(formula, layout) {
  const f = chAplanarSignos(formula);
  const d = ctColNumeroALetra(layout.deudorCol);
  const a = ctColNumeroALetra(layout.acreedorCol);
  const cruda = ctColNumeroALetra(layout.saldoCol);

  // Un puntaje, no una cadena de casos: la deudora suma a favor, la acreedora en contra, y la
  // columna cruda (que YA es deudora menos acreedora) suma a favor con su propio signo.
  //
  //   +B-C  ->  (+1) - (-1) = +2   directo     (Activo: un activo es deudor)
  //   +C-B  ->  (-1) - (+1) = -2   invertido   (Pasivo, EERR: deuda e ingreso son acreedores)
  //   +C    ->   0   - (+1) = -1   invertido   (renglones que sólo miran la acreedora)
  //   -D    ->  -1                 invertido   (la cruda con el signo cambiado)
  //
  // Antes esto era una cadena de if y se le escapaba justo el caso "+C" solo, que es el más
  // común en el Pasivo: devolvía "no sé", caía en el signo mayoritario de la hoja, y ahí el
  // Pasivo del Acumulado R$ terminaba restando ocho renglones que tenían que sumar.
  let puntaje = 0;
  for (const m of f.matchAll(/([+-])\s*SALDOS!\$?([A-Z]{1,3})\$?\d+/g)) {
    const signo = m[1] === "-" ? -1 : 1;
    const col = m[2].toUpperCase();
    if (col === d) puntaje += signo;
    else if (col === a) puntaje -= signo;
    else if (col === cruda) puntaje += signo;
  }
  if (puntaje > 0) return 1;
  if (puntaje < 0) return -1;
  return 0;   // no pude decidirlo mirando este renglón
}

// Deja cada referencia a SALDOS con su signo verdadero, sin paréntesis de por medio.
//
// "- Proveedores" del Pasivo es -(SALDOS!F53+SALDOS!F54+…): adentro los términos están escritos
// con "+", pero todos restan, porque el paréntesis entero va restando. Contando los signos tal
// como están escritos, ese renglón parecía sumar 42 cuentas y restar 3, y el signo salía al
// revés — justo en el renglón con 44 proveedores adentro.
function chAplanarSignos(formula) {
  let f = String(formula || "").replace(/\s+/g, "");
  // Se resuelven los paréntesis de adentro hacia afuera, tantas veces como haga falta.
  for (let vuelta = 0; vuelta < 8 && /\(/.test(f); vuelta++) {
    const antes = f;
    // El (?<![A-Za-z]) deja afuera los paréntesis de una función: SUM(SALDOS!D53:D79) no es
    // un grupo con signo, es una llamada, y desarmarla dejaba "-SUM+SALDOS!D53:D79".
    f = f.replace(/([+-]?)(?<![A-Za-z])\(([^()]*)\)/g, (todo, signo, dentro) => {
      if (!/SALDOS!/.test(dentro)) return todo;
      const cuerpo = signo === "-"
        ? dentro.replace(/([+-]?)(SALDOS!)/g, (t, s, x) => (s === "-" ? "+" : "-") + x)
        : dentro.replace(/^([^+-])/, "+$1");
      return cuerpo;
    });
    if (f === antes) break;
  }
  return f.replace(/^([^+-])/, "+$1");
}

// El signo que usa una hoja, visto en todos sus renglones. Sirve cuando el renglón de destino
// está vacío y no tiene términos de dónde deducirlo.
function chSignoDeHoja(mapa, layout) {
  let directo = 0, invertido = 0;
  for (const r of mapa.renglones) {
    for (const c of r.cols) {
      const v = mapa.ws.getCell(r.fila, c.col).value;
      if (!v || typeof v !== "object" || typeof v.formula !== "string") continue;
      const s = chSignoDelRenglon(v.formula, layout);
      if (s === 1) directo++; else if (s === -1) invertido++;
    }
  }
  return invertido > directo ? -1 : 1;
}

// Engancha la cuenta al renglón `rotulo` de `nombreHoja`.
//
// Se la saca antes de los otros renglones DE ESA MISMA HOJA, no del libro entero: que el
// Activo y el EESP lean los dos la misma cuenta es correcto —cada hoja la muestra en su
// apertura— y borrarla de la otra hoja rompería un estado para arreglar el otro.
function engancharEnHoja(wb, layout, nombreHoja, cod, rotulo) {
  const plan = leerPlanDeCuentas(wb, layout).cuentas;
  const cuenta = plan[cod];
  if (!cuenta) return { hecho: false, motivo: `${cod} no está en el plan de cuentas de este archivo` };

  const mapa = chMapaHoja(wb, layout, nombreHoja);
  if (!mapa) return { hecho: false, motivo: `${nombreHoja} no engancha cuentas en este archivo` };

  const candidatos = mapa.renglones.filter(r => r.rotulo && chNorm(r.rotulo) === chNorm(rotulo));
  if (candidatos.length > 1) {
    // El Anexo I repite los mismos nombres en cada sección (valores de origen, amortizaciones):
    // con el rótulo solo no se sabe a qué renglón se quiso mover la cuenta, y elegir uno al
    // azar la mandaría a la sección equivocada. Mejor no hacer nada y decirlo.
    return {
      hecho: false,
      motivo: `"${rotulo}" aparece ${candidatos.length} veces en ${nombreHoja} (filas ` +
              `${candidatos.map(c => c.fila).join(", ")}): el rótulo solo no alcanza para saber a cuál va`,
    };
  }
  const destino = candidatos[0];
  if (!destino) return { hecho: false, motivo: `"${rotulo}" no existe en ${nombreHoja}` };

  // Dentro del renglón, la columna que ya usa. En el Anexo II eso decide el centro de costo:
  // un concepto que hoy suma sólo en comercialización sigue siendo de comercialización.
  const col = destino.cols.length ? destino.cols[0].col : mapa.colsImporte[0];

  // El signo del renglón se lee ACÁ, antes de sacarle nada. Si la cuenta que se está moviendo
  // era la única del renglón, sacarla lo deja vacío y el signo se perdería: había que caer en
  // el signo mayoritario de la hoja, y en el EERR del Acumulado $ eso daba vuelta
  // "Diferencia de cotización", que es el único renglón de resultados que suma en vez de restar.
  const antes = mapa.ws.getCell(destino.fila, col).value;
  let signo = chSignoDelRenglon(
    (antes && typeof antes === "object" && typeof antes.formula === "string") ? antes.formula : "",
    layout);

  const filas = [cuenta].concat(cuenta.otrasFilas || []).map(f => f.fila);
  for (const r of mapa.renglones) {
    for (const c of r.cols) {
      for (const f of filas) rtQuitarTermino(mapa.ws, c.dir, f);
    }
  }

  const celda = mapa.ws.getCell(destino.fila, col);
  const v = celda.value;
  const previo = (v && typeof v === "object" && typeof v.formula === "string") ? v.formula : "";

  // Si el renglón ya estaba vacío de entrada, no hay de dónde sacarlo: manda el mayoritario de
  // la hoja. Y eso se avisa, porque puede errarle: en el EERR del Acumulado $ la mayoría de los
  // renglones resta (son ingresos, que son acreedores) pero "Diferencia de cotización" suma, y
  // una cuenta nueva metida en un renglón vacío de esa hoja saldría con el signo de la mayoría.
  // El dato viaja en el resultado para que la pantalla lo diga en vez de que se descubra solo
  // cuando el número no cierre.
  const signoDeducido = !signo;
  if (!signo) signo = chSignoDeHoja(mapa, layout);

  const d = ctColNumeroALetra(layout.deudorCol), a = ctColNumeroALetra(layout.acreedorCol);
  const termino = signo < 0
    ? `+SALDOS!${a}${cuenta.fila}-SALDOS!${d}${cuenta.fila}`
    : `+SALDOS!${d}${cuenta.fila}-SALDOS!${a}${cuenta.fila}`;

  celda.value = { formula: previo ? `${previo}${termino}` : termino };
  return {
    hecho: true, hoja: nombreHoja, fila: destino.fila, col,
    rotulo: destino.rotulo, signo, signoDeducido,
  };
}

// ------------------------------------------------- el mismo renglón, con el mismo nombre
//
// El mismo concepto se escribía distinto en cada archivo: "Fondo común de inversión BBVA",
// "Fondo Común de inversión", "Fondo Comun de Inversión BBVA" y "Fondo común inversión BBVA"
// son los cuatro renglones de la cuenta 1160100000. Como los informes se comparan entre sí,
// cuatro nombres para lo mismo son cuatro renglones distintos.
//
// El renglón se busca POR LA CUENTA QUE LEE, no por su nombre anterior ni por su dirección:
// el nombre es justamente lo que varía, y las direcciones se mueven en cada corrida.
//
// Se renombra SÓLO si ese renglón lee esa cuenta y ninguna otra. Un renglón que agrupa varias
// cuentas no se puede renombrar con el nombre de una: en el Acumulado R$, "Otros ingresos y
// egresos" junta cuatro cuentas que los otros archivos reparten en tres renglones, y ponerle
// el nombre de una sola escondería a las otras tres.
const RENOMBRES_RENGLON = [
  { hoja: "EERR",   cuenta: "4120300000", a: "Resultado venta bienes de uso" },
  { hoja: "Activo", cuenta: "1110101050", a: "- Fondo Fijo Adelantos PDT" },
  { hoja: "Activo", cuenta: "1160100000", a: "- Fondo Común de Inversión BBVA" },
  { hoja: "Activo", cuenta: "1160300000", a: "- Intereses a Devengar Plazo Fijo" },
  { hoja: "Pasivo", cuenta: "2110407000", a: "- Plan mis Facilidades" },
  { hoja: "Pasivo", cuenta: "2110408000", a: "- Intereses a devengar" },
];

// Cuentas que van a un renglón determinado, decidido con contaduría, más allá de dónde las
// tenga hoy cada archivo.
const ASIGNACIONES_APROBADAS = [
  // El anticipo de vacaciones va junto con los adelantos al personal: es lo mismo, y el
  // Acumulado $ ya lo tenía así.
  { hoja: "Activo", cuenta: "1140500300", rotulo: "- Adelanto al personal" },
  // El resultado por tenencia de FCI es un interés. Tres archivos ya lo tenían en "Intereses";
  // el Mensual R$ y el Acumulado R$ lo tenían mezclado en "Otros ingresos".
  { hoja: "EERR",   cuenta: "4231100000", rotulo: "Intereses" },
];

function chCeldaDelRotulo(ws, fila, colImporte) {
  for (let c = colImporte - 1; c >= 1; c--) {
    const v = ws.getCell(fila, c).value;
    if (typeof v === "number") continue;
    if (v && typeof v === "object" && typeof v.formula === "string") continue;
    const t = chTexto(ws, fila, c).trim();
    if (!t || /^-?[\d.,]+$/.test(t)) continue;
    return { fila, col: c };
  }
  return null;
}

function aplicarRenombresRenglon(wb, layout, log = () => {}) {
  const plan = leerPlanDeCuentas(wb, layout).cuentas;
  const hechos = [];

  for (const r of RENOMBRES_RENGLON) {
    const cuenta = plan[r.cuenta];
    if (!cuenta) continue;                       // este archivo no tiene esa cuenta
    const mapa = chMapaHoja(wb, layout, r.hoja);
    if (!mapa) continue;

    const filas = [cuenta].concat(cuenta.otrasFilas || []).map(f => f.fila);
    const destino = mapa.renglones.find(x => x.filasSaldos.some(f => filas.includes(f)));
    if (!destino) continue;
    if (chNorm(destino.rotulo) === chNorm(r.a)) continue;          // ya se llama así

    // Un renglón que agrupa varias cuentas no lleva el nombre de una sola.
    const cuantas = destino.filasSaldos.filter(f =>
      Object.values(plan).some(c => c.fila === f || (c.otrasFilas || []).some(o => o.fila === f))).length;
    if (cuantas > 1) {
      log(`  ⚠ ${r.hoja}: no renombré el renglón de ${r.cuenta} a "${r.a}" porque agrupa ` +
          `${cuantas} cuentas; ponerle el nombre de una escondería a las otras.`);
      continue;
    }

    if (mapa.renglones.some(x => x !== destino && chNorm(x.rotulo) === chNorm(r.a))) {
      log(`  ⚠ ${r.hoja}: no renombré "${destino.rotulo}" porque "${r.a}" ya existe en esa hoja.`);
      continue;
    }

    const celda = chCeldaDelRotulo(mapa.ws, destino.fila, destino.cols[0].col);
    if (!celda) continue;
    const antes = destino.rotulo;
    mapa.ws.getCell(celda.fila, celda.col).value = r.a;
    hechos.push({ hoja: r.hoja, fila: destino.fila, de: antes, a: r.a, cuenta: r.cuenta });
    log(`  ${r.hoja} fila ${destino.fila}: "${antes}" pasa a llamarse "${r.a}" (igual que en los ` +
        `otros archivos).`);
  }
  return hechos;
}

// Manda cada cuenta al renglón que se decidió, esté donde esté hoy.
function aplicarAsignaciones(wb, layout, log = () => {}) {
  const plan = leerPlanDeCuentas(wb, layout).cuentas;
  const hechos = [], salteadas = [];
  for (const a of ASIGNACIONES_APROBADAS) {
    if (!plan[a.cuenta]) continue;
    const mapa = chMapaHoja(wb, layout, a.hoja);
    if (!mapa) continue;
    const filas = [plan[a.cuenta]].concat(plan[a.cuenta].otrasFilas || []).map(f => f.fila);
    const actual = mapa.renglones.find(x => x.filasSaldos.some(f => filas.includes(f)));
    if (actual && chNorm(actual.rotulo) === chNorm(a.rotulo)) continue;   // ya está donde va

    const r = engancharEnHoja(wb, layout, a.hoja, a.cuenta, a.rotulo);
    if (!r.hecho) { salteadas.push({ ...a, motivo: r.motivo }); log(`  ⚠ ${a.hoja}: ${a.cuenta} — ${r.motivo}.`); continue; }
    hechos.push({ ...a, de: actual ? actual.rotulo : null, fila: r.fila });
    log(`  ${a.hoja}: ${a.cuenta} pasa de ${actual ? `"${actual.rotulo}"` : "ningún renglón"} ` +
        `a "${a.rotulo}".`);
  }
  return { hechos, salteadas };
}

if (typeof module !== "undefined") {
  const cfg = require("./config_tfbr.js");
  global.derivarLayoutSaldos = cfg.derivarLayoutSaldos;
  global.leerPlanDeCuentas = cfg.leerPlanDeCuentas;
  global.ctFormulaNetaAnexo = cfg.ctFormulaNetaAnexo;
  global.ctColNumeroALetra = cfg.ctColNumeroALetra;
  global.rtQuitarTermino = require("./rotulos_anexo.js").rtQuitarTermino;
  module.exports = {
    chTexto, chNorm, chRotuloDe, chMapaHoja, chHojasConfigurables,
    chRubroDeHoja, chPlanPorFila, engancharEnHoja,
    chSignoDelRenglon, chSignoDeHoja, chAplanarSignos, chHojaEditable,
    RENOMBRES_RENGLON, aplicarRenombresRenglon, chCeldaDelRotulo,
    ASIGNACIONES_APROBADAS, aplicarAsignaciones,
  };
}
