// Renglones que leen media cuenta.
//
// En SALDOS cada cuenta ocupa dos columnas: la deudora y la acreedora. El saldo de verdad es la
// resta de las dos, y por eso el resto del motor arma las referencias con ctFormulaNetaAnexo.
// Pero los maestros vienen con decenas de renglones que nombran una sola de las dos columnas.
//
// Mientras la cuenta cae siempre del mismo lado, no se nota: la columna que falta vale cero y
// la cuenta llega completa. El día que la cuenta cambia de signo, su importe entero desaparece
// del informe y el balance no cierra, sin que nada avise.
//
// Pasó en agosto de 2026. "1140900900 SEGUROS A DEV. FEDERACIÓN" venía acreedora todos los
// meses y el renglón de "Seguros a devengar" del Activo la leía como "-SALDOS!C28". En agosto
// quedó deudora por 866.748,53 y ese renglón dio cero: el Estado de Situación Patrimonial del
// Mensual $ descuadró en 881.272,63, de los cuales 866.748,53 eran esta cuenta.
//
// Qué se hace. A cada referencia suelta se le agrega la columna que le falta, con el signo
// cambiado —que es lo que significa "el saldo es deudora menos acreedora"— y sin tocar nada
// más de la fórmula. Un renglón que hoy dice "-SALDOS!C28" pasa a decir "-SALDOS!C28+SALDOS!B28".
//
// Por qué se puede hacer sin preguntar renglón por renglón: el término que se agrega vale cero
// en todos los meses en los que el informe cerraba bien, así que no cambia ningún número que ya
// estuviera bien. Solo aparece cuando la cuenta se dio vuelta, que es el caso que hoy se pierde.
//
// Lo que NO se toca: las referencias que ya nombran las dos columnas, los rangos (SUM(D53:D79)
// y compañía, que toman la columna entera a propósito) y cualquier fila que no sea una cuenta
// del plan —subtotales, líneas manuales, la de diferencia de cambio—.

// Las referencias sueltas de una fórmula: [{ signo, col, fila, i, largo }].
function mcReferencias(formula, colDeudora, colAcreedora) {
  const refs = [];
  const re = /([+-]?)\s*SALDOS!(\$?)([A-Z]{1,3})(\$?)(\d+)(?!\d)/g;
  let m;
  while ((m = re.exec(formula)) !== null) {
    // Un rango (D53:D79) toma la columna entera a propósito: no es media cuenta.
    if (formula[m.index + m[0].length] === ":") continue;
    if (formula[m.index - 1] === ":") continue;
    const col = m[3];
    if (col !== colDeudora && col !== colAcreedora) continue;
    refs.push({ signo: m[1] === "-" ? "-" : "+", col, fila: Number(m[5]), i: m.index, largo: m[0].length });
  }
  return refs;
}

// Completa una fórmula. Devuelve la nueva, o null si no había nada que completar.
function mcCompletar(formula, colDeudora, colAcreedora, esCuenta) {
  const refs = mcReferencias(formula, colDeudora, colAcreedora);
  if (!refs.length) return null;

  // Por fila: qué columnas nombra.
  const porFila = {};
  for (const r of refs) {
    if (!porFila[r.fila]) porFila[r.fila] = [];
    porFila[r.fila].push(r);
  }

  const aAgregar = [];
  for (const fila of Object.keys(porFila)) {
    if (!esCuenta(Number(fila))) continue;
    const cols = new Set(porFila[fila].map(r => r.col));
    if (cols.size !== 1) continue;                 // ya nombra las dos
    const ref = porFila[fila][0];
    const falta = ref.col === colDeudora ? colAcreedora : colDeudora;
    const signo = ref.signo === "-" ? "+" : "-";   // el saldo es deudora MENOS acreedora
    aAgregar.push({ fila: Number(fila), texto: `${signo}SALDOS!${falta}${fila}` });
  }
  if (!aAgregar.length) return null;

  return formula + aAgregar.map(x => x.texto).join("");
}

// Recorre las hojas que se imprimen y completa lo que esté a medias.
function completarMediaColumna(wb, layout, log = () => {}) {
  const s = wb.getWorksheet(layout.sheet);
  const cd = ctColNumeroALetra(layout.deudorCol);
  const ca = ctColNumeroALetra(layout.acreedorCol);

  // Qué filas de SALDOS son una cuenta del plan. Una línea manual o un subtotal no tiene par de
  // columnas que completar.
  const esCuenta = (fila) => {
    if (fila < layout.planDeCuentas.desde || fila > layout.planDeCuentas.hasta) return false;
    let t = s.getCell(fila, layout.keyCol).value;
    if (t && typeof t === "object") t = t.richText ? t.richText.map(x => x.text).join("") : "";
    return /^\d{10}\b/.test(String(t || "").trim());
  };

  const completados = [];
  for (const ws of wb.worksheets) {
    if (ws.name === layout.sheet) continue;
    ws.eachRow({ includeEmpty: false }, (row, r) => {
      row.eachCell({ includeEmpty: false }, (cell, c) => {
        const v = cell.value;
        if (!v || typeof v !== "object" || typeof v.formula !== "string") return;
        const nueva = mcCompletar(v.formula, cd, ca, esCuenta);
        if (!nueva) return;
        completados.push({ hoja: ws.name, celda: cell.address, antes: v.formula, ahora: nueva });
        cell.value = { formula: nueva };
      });
    });
  }

  if (completados.length) {
    const porHoja = {};
    completados.forEach(x => { porHoja[x.hoja] = (porHoja[x.hoja] || 0) + 1; });
    log(`  ${completados.length} renglón(es) leían una sola columna de una cuenta y ahora leen ` +
        `las dos (` + Object.entries(porHoja).map(([h, n]) => `${h}: ${n}`).join(", ") + ").");
    for (const x of completados.slice(0, 5)) log(`    ${x.hoja}!${x.celda}: "${x.antes}" → "${x.ahora}"`);
    if (completados.length > 5) log(`    … y ${completados.length - 5} más.`);
  }
  return completados;
}

if (typeof module !== "undefined") {
  global.ctColNumeroALetra = require("./config_tfbr.js").ctColNumeroALetra;
  module.exports = { completarMediaColumna, mcCompletar, mcReferencias };
}
