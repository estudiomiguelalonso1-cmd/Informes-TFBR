// Alta de una cuenta que viene en el export y todavía no está en el plan de cuentas de
// SALDOS. Es un caso normal, no una excepción: entre junio y julio 2026 aparecieron 6 cuentas
// nuevas en el Mensual $, por 11,66 millones de pesos. Sin darlas de alta, ese importe no
// entra en ningún total y el balance cierra igual, sin avisar.
//
// La fila se inserta al lado de sus hermanas por código (prefijo común más largo), copiándoles
// la fórmula. Lo que NO se decide solo es a qué línea de la Nota 4 o del Anexo II va: eso lo
// tiene que elegir una persona, igual que hoy.

// Hasta dónde llega el prefijo compartido entre dos códigos. Ordena la cuenta nueva entre sus
// hermanas de la misma familia (4223xxxxx cerca de las otras 4223…).
const IC_MARCA = String.fromCharCode(1);
const IC_MARCA_RE = new RegExp(IC_MARCA + "(" + String.fromCharCode(92) + 'd+)' + IC_MARCA, "g");

function icPrefijoComun(a, b) {
  let i = 0;
  while (i < a.length && i < b.length && a[i] === b[i]) i++;
  return i;
}

function icTexto(ws, fila, col) {
  const v = ws.getCell(fila, col).value;
  if (v === null || v === undefined) return "";
  if (typeof v === "object") return v.richText ? v.richText.map(t => t.text).join("") : "";
  return String(v);
}

// ¿Esta fila es la ÚLTIMA de algún rango que la suma? Si lo es, insertar DEBAJO deja la cuenta
// nueva fuera del subtotal: Excel no estira un SUM(B7:B120) cuando se inserta en la 121. Hay
// que insertar EN la fila de la vecina para que el rango se expanda solo.
//
// Es el mismo error que en el sistema de Southern Copper dejó dos proveedores nuevos afuera
// del subtotal y descuadró el pasivo por 1.915.260 sin que nada lo delatara.
function icFilaCierraUnSubtotal(wb, hoja, fila) {
  const reRango = new RegExp(
    "SUM\\(\\s*(?:'?" + hoja + "'?!)?\\$?([A-Z]{1,3})\\$?(\\d+)\\s*:\\s*" +
    "(?:'?" + hoja + "'?!)?\\$?([A-Z]{1,3})\\$?(\\d+)\\s*\\)", "gi");
  const reNombra = new RegExp("'?" + hoja + "'?!");
  for (const ws of wb.worksheets) {
    const esLaHoja = ws.name === hoja;
    let cierra = false;
    ws.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (cierra) return;
        const v = cell.value;
        if (!(v && typeof v === "object" && typeof v.formula === "string")) return;
        if (!esLaHoja && !reNombra.test(v.formula)) return;
        reRango.lastIndex = 0;
        let m;
        while ((m = reRango.exec(v.formula))) {
          if (+m[2] <= fila && fila === +m[4]) { cierra = true; return; }
        }
      });
    });
    if (cierra) return true;
  }
  return false;
}

// El texto de la clave se arma con el MISMO formato que usa la vecina (dos espacios, un
// espacio, guión…). Si se escribe con otro formato, el BUSCARV que arma el motor sigue
// funcionando —porque escribe en la zona de pegado el texto exacto de la plantilla— pero el
// archivo queda con dos convenciones distintas, que es justamente de dónde salen los códigos
// repetidos que hoy hay que resolver a mano.
function icClaveComoLaVecina(textoVecina, codigo, nombre) {
  const m = /^\s*(\d{6,})(\s*-?\s*)/.exec(textoVecina);
  const separador = m ? m[2] : "  ";
  return `${codigo}${separador}${nombre}`;
}

// Copia la fórmula de la vecina cambiando su número de fila por el de la nueva.
function icCopiarFormula(ws, filaOrigen, filaDestino, col) {
  const origen = ws.getCell(filaOrigen, col);
  const destino = ws.getCell(filaDestino, col);
  destino.style = origen.style;
  const v = origen.value;
  if (v && typeof v === "object" && typeof v.formula === "string") {
    destino.value = {
      formula: v.formula.replace(
        new RegExp("\\b([A-Z]{1,3})" + filaOrigen + "\\b", "g"),
        (_, c) => c + filaDestino
      ),
    };
    return true;
  }
  return false;
}

// Inserta la cuenta en SALDOS y devuelve dónde quedó. NO la engancha a ningún estado: eso es
// una decisión aparte (ver el checklist que devuelve el motor).
function insertarCuentaEnSaldos(wb, layout, planDeCuentas, cuenta, log = () => {}) {
  const ws = wb.getWorksheet(layout.sheet);

  const hermanas = Object.entries(planDeCuentas)
    .map(([codigo, info]) => ({ codigo, ...info }))
    .filter(h => h.fila >= layout.planDeCuentas.desde && h.fila <= layout.planDeCuentas.hasta);
  if (!hermanas.length) {
    throw new Error(`No hay ninguna cuenta en el plan de ${layout.sheet} para usar de modelo.`);
  }

  const vecina = hermanas.reduce((a, b) => {
    const pa = icPrefijoComun(a.codigo, cuenta.codigo);
    const pb = icPrefijoComun(b.codigo, cuenta.codigo);
    if (pa !== pb) return pa > pb ? a : b;
    return a.fila >= b.fila ? a : b;
  });

  const cierra = icFilaCierraUnSubtotal(wb, layout.sheet, vecina.fila);
  const filaNueva = cierra ? vecina.fila : vecina.fila + 1;

  const modificadas = insertRowEn(wb, layout.sheet, filaNueva);

  // el plan en memoria queda corrido
  for (const info of Object.values(planDeCuentas)) {
    if (info.fila >= filaNueva) info.fila += 1;
    if (info.otrasFilas) for (const o of info.otrasFilas) if (o.fila >= filaNueva) o.fila += 1;
  }
  const filaVecina = vecina.fila >= filaNueva ? vecina.fila + 1 : vecina.fila;

  const clave = icClaveComoLaVecina(icTexto(ws, filaVecina, layout.keyCol), cuenta.codigo, cuenta.nombre);
  ws.getCell(filaNueva, layout.keyCol).value = clave;
  ws.getCell(filaNueva, layout.keyCol).style = ws.getCell(filaVecina, layout.keyCol).style;

  for (const col of [layout.deudorCol, layout.acreedorCol, layout.saldoCol]) {
    icCopiarFormula(ws, filaVecina, filaNueva, col);
  }

  planDeCuentas[cuenta.codigo] = {
    fila: filaNueva,
    texto: clave,
    colSaldo: layout.saldoCol,
  };

  log(`  Alta: "${clave}" en ${layout.sheet}!${filaNueva}, copiando la fórmula de ` +
      `"${vecina.texto.trim()}" (fila ${filaVecina})` +
      (cierra ? `, insertada DENTRO del subtotal que cerraba esa fila` : "") +
      `. ${modificadas} referencia(s) reacomodadas.`);

  return { fila: filaNueva, clave, vecina: vecina.texto.trim() };
}

// ---------------------------------------------------------------- cuentas gemelas
//
// Una "gemela" es la misma cuenta ya cargada con el código MAL: un dígito menos que el de
// Onvio (`421180000  REPARACIONES` contra `4211800000  REPARACIONES`). Como el BUSCARV busca
// por texto y el texto lleva el código, la fila vieja nunca levanta importe y queda en cero
// para siempre. Pero las líneas del Anexo II apuntan a ESA fila, así que dar de alta la
// cuenta con el código bueno sin tocar nada más mete el importe en los totales de SALDOS y
// lo deja afuera del Anexo II y del EERR: el balance y el estado de resultados dejan de
// contar lo mismo.
//
// Entre junio y julio 2026 esto pasó con 4 cuentas del Mensual $, por 18 millones de pesos, y
// el cierre de julio se resolvió repuntando esas líneas a mano.

function icNombreNormalizado(texto) {
  const m = /^\s*(\d{6,})\s*-?\s*(.*)$/.exec(String(texto).trim());
  return (m ? m[2] : String(texto))
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();
}

// ¿Los dos códigos son el mismo salvo un dígito de más o de menos? Se compara así y no por
// "uno empieza con el otro" para no emparejar cuentas de la misma familia que son distintas.
function icCodigosCasiIguales(a, b) {
  const [corto, largo] = a.length <= b.length ? [a, b] : [b, a];
  if (largo.length - corto.length !== 1) return false;
  for (let i = 0; i < largo.length; i++) {
    if (largo.slice(0, i) + largo.slice(i + 1) === corto) return true;
  }
  return false;
}

// Busca en el plan una cuenta con el mismo nombre y el código casi igual.
function buscarGemela(planDeCuentas, codigo, nombre) {
  const objetivo = icNombreNormalizado(codigo + " " + nombre);
  for (const [c, info] of Object.entries(planDeCuentas)) {
    if (c === codigo) continue;
    if (!icCodigosCasiIguales(c, codigo)) continue;
    if (icNombreNormalizado(info.texto) !== objetivo) continue;
    return { codigo: c, ...info };
  }
  return null;
}

// Mueve a la fila nueva todas las referencias sueltas que apuntaban a la gemela. Sólo toca la
// fila exacta (`SALDOS!C185` -> `SALDOS!C190`), nunca los rangos ni las otras filas que la
// misma fórmula referencie.
function repuntarGemela(wb, hoja, filaVieja, filaNueva, log = () => {}) {
  const reNombra = new RegExp("'?" + hoja + "'?!");
  const reRef = new RegExp("((?:'?" + hoja + "'?!)\\$?[A-Z]{1,3}\\$?)" + filaVieja + "(?!\\d)", "g");
  const reRango = /\$?[A-Z]{1,3}\$?\d+\s*:\s*\$?[A-Z]{1,3}\$?\d+/g;
  const movidas = [];

  for (const ws of wb.worksheets) {
    ws.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const v = cell.value;
        if (!(v && typeof v === "object" && typeof v.formula === "string")) return;
        if (!reNombra.test(v.formula)) return;

        // Los rangos se apartan para no tocarlos y se reponen al final. El marcador es un
        // caracter de control y no un numero: con numeros, al reponer los rangos se pisarian
        // los numeros de fila de las demas referencias de la misma formula.
        const rangos = [];
        const sinRangos = v.formula.replace(reRango, (m) => IC_MARCA + (rangos.push(m) - 1) + IC_MARCA);
        reRef.lastIndex = 0;
        if (!reRef.test(sinRangos)) return;
        reRef.lastIndex = 0;
        const nueva = sinRangos.replace(reRef, (_, pre) => pre + filaNueva)
          .replace(IC_MARCA_RE, (_, i) => rangos[+i]);
        cell.value = { formula: nueva };
        movidas.push(`${ws.name}!${cell.address}`);
      });
    });
  }

  if (movidas.length) {
    log(`     ${movidas.length} referencia(s) repuntadas de la fila ${filaVieja} a la ${filaNueva}: ${movidas.join(", ")}`);
  }
  return movidas;
}

if (typeof module !== "undefined") {
  global.insertRowEn = require("./formula_hojas.js").insertRowEn;
  module.exports = {
    insertarCuentaEnSaldos, icPrefijoComun, icClaveComoLaVecina, icFilaCierraUnSubtotal,
    buscarGemela, repuntarGemela, icCodigosCasiIguales, icNombreNormalizado,
  };
}
