// Rótulos que el Anexo II de un archivo no tiene y los otros sí.
//
// Los cuatro balances tienen que abrir el gasto con los mismos rótulos, o no se pueden
// comparar entre sí y una cuenta termina colgada de un renglón ajeno según el archivo. Los dos
// Mensuales tienen "Verificación Técnica Vehicular" y "Verificación Policial"; los dos
// Acumulados no, y por eso esas cuentas estaban sumadas dentro de "Deudores incobrables",
// "Gastos varios clientes" y "Mantenimiento y lav. de flota".
//
// El renglón se inserta EN LA ÚLTIMA FILA DEL RANGO que suma el total (la 84 de un
// SUM(D11:D84)), no en la fila del total ni debajo. Excel estira el rango cuando se inserta
// DENTRO, y la 84 todavía lo está; la 85 ya no, y el renglón nuevo quedaría fuera del total
// sin que nada lo muestre — el importe entraría en SALDOS y no llegaría al EERR.
//
// Y antes de engancharlo se le saca la cuenta al renglón que la venía sumando, o el importe se
// contaría dos veces.
const ROTULOS_ANEXO = [
  { archivo: "balance_acumulado_ars", rotulo: "Verificación Técnica Vehicular", cod: "4223500000", cc: "E" },
  { archivo: "balance_acumulado_ars", rotulo: "Verificación Policial",          cod: "4223700000", cc: "E" },
  { archivo: "balance_acumulado_brl", rotulo: "Verificación Técnica Vehicular", cod: "4223500000", cc: "E" },
  { archivo: "balance_acumulado_brl", rotulo: "Verificación Policial",          cod: "4223700000", cc: "E" },
];

const RT_CC = { D: 4, E: 5, F: 6 };

function rtNorm(t) {
  return String(t || "").normalize("NFD").replace(/[̀-ͯ]/g, "")
    .toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();
}
function rtTexto(ws, r, c) {
  const v = ws.getCell(r, c).value;
  if (v == null) return "";
  if (typeof v === "object") return v.richText ? v.richText.map(t => t.text).join("") : "";
  return String(v);
}

// El bloque de conceptos: desde el primer rótulo hasta la fila del total (la que lleva el SUM).
function rtUbicarBloque(ax) {
  let desde = null, hasta = null, total = null, rangoTotal = null, colImporte = null;
  for (let r = 1; r <= ax.rowCount; r++) {
    for (const c of [4, 5, 6]) {
      const v = ax.getCell(r, c).value;
      if (!v || typeof v !== "object" || typeof v.formula !== "string") continue;
      // El total es el SUM que abarca VARIAS FILAS de su misma columna — SUM(D11:D84).
      // El SUM(D11:F11) de cada renglón abarca columnas, no filas, y no es el total.
      const t = /^SUM\(\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)\)$/i.exec(v.formula);
      if (t && t[1].toUpperCase() === t[3].toUpperCase() && +t[4] > +t[2]) {
        if (total === null) { total = r; rangoTotal = { desde: +t[2], hasta: +t[4] }; }
        continue;
      }
      if (/^SUM\(/i.test(v.formula)) continue;
      const m = /SALDOS!\$?([A-Z]{1,3})\$?\d+/.exec(v.formula);
      if (!m) continue;
      if (desde === null) desde = r;
      hasta = r;
      if (!colImporte) colImporte = m[1];        // la columna de SALDOS que leen los renglones
    }
  }
  // El bloque llega hasta la última fila CON RÓTULO antes del total, no hasta la última que
  // lee una cuenta. Los renglones que todavía no tienen cuenta asignada están al final —
  // recién creados, esperando que se les asigne una— y si quedaran fuera del bloque nadie los
  // encontraría: la asignación crearía un duplicado en vez de usar el que ya está.
  if (rangoTotal) {
    for (let r = rangoTotal.hasta; r > (hasta || 0); r--) {
      const t = ax.getCell(r, 2).value;
      const texto = (t && typeof t === "object")
        ? (t.richText ? t.richText.map(x => x.text).join("") : "")
        : (t == null ? "" : String(t));
      if (texto.trim() && !/^total|^conceptos/i.test(texto.trim())) { hasta = r; break; }
    }
  }
  return { desde, hasta, total, rangoTotal, colImporte };
}

// Saca del renglón `celda` el término que lee la fila `fila` de SALDOS.
function rtQuitarTermino(ax, celda, fila) {
  const v = ax.getCell(celda).value;
  if (!v || typeof v !== "object" || typeof v.formula !== "string") return false;
  const re = new RegExp(`[+-]?\\s*SALDOS!\\$?[A-Z]{1,3}\\$?${fila}(?!\\d)`, "g");
  let nueva = v.formula.replace(re, "");
  if (nueva === v.formula) return false;
  nueva = nueva.replace(/^\s*\+/, "").trim();
  ax.getCell(celda).value = nueva ? { formula: nueva } : 0;
  return true;
}

// Agrega los rótulos que le falten a este archivo. Idempotente: si ya está, no hace nada.
function agregarRotulosAnexo(wb, archivoId, layout, planDeCuentas, log = () => {}) {
  const agregados = [], salteados = [];
  const mios = ROTULOS_ANEXO.filter(r => r.archivo === archivoId);
  if (!mios.length) return { agregados, salteados };

  const ax = wb.getWorksheet("Anexo II");
  if (!ax) return { agregados, salteados: mios.map(r => ({ ...r, motivo: "el archivo no tiene Anexo II" })) };

  for (const r of mios) {
    const bloque = rtUbicarBloque(ax);
    if (bloque.total === null || !bloque.rangoTotal) {
      salteados.push({ ...r, motivo: "no ubiqué la fila del total ni su rango" }); continue;
    }

    // ¿ya existe el rótulo?
    let existe = false;
    for (let x = bloque.desde; x <= bloque.hasta; x++) {
      if (rtNorm(rtTexto(ax, x, 2)) === rtNorm(r.rotulo)) { existe = true; break; }
    }
    if (existe) continue;

    const cuenta = planDeCuentas[r.cod];
    if (!cuenta) { salteados.push({ ...r, motivo: `la cuenta ${r.cod} no está en SALDOS` }); continue; }

    // quién viene sumando esa cuenta hoy: hay que sacársela antes, o se cuenta dos veces
    const quitados = [];
    for (let x = bloque.desde; x <= bloque.hasta; x++) {
      for (const c of [4, 5, 6]) {
        const dir = `${String.fromCharCode(64 + c)}${x}`;
        if (rtQuitarTermino(ax, dir, cuenta.fila)) quitados.push(`${dir} "${rtTexto(ax, x, 2).trim()}"`);
      }
    }

    // se inserta en la ÚLTIMA fila del rango del total, que todavía está dentro: así el SUM
    // se estira y el renglón nuevo queda contado
    const filaNueva = bloque.rangoTotal ? bloque.rangoTotal.hasta : bloque.hasta;
    const modificadas = insertRowEn(wb, "Anexo II", filaNueva);
    const colImp = bloque.colImporte || "C";
    const cc = RT_CC[r.cc] || 5;
    ax.getCell(filaNueva, 2).value = r.rotulo;
    ax.getCell(filaNueva, 3).value = { formula: `SUM(D${filaNueva}:F${filaNueva})` };
    for (const c of [4, 5, 6]) ax.getCell(filaNueva, c).value = 0;
    ax.getCell(filaNueva, cc).value = { formula: `+SALDOS!${colImp}${cuenta.fila}` };

    agregados.push({ rotulo: r.rotulo, fila: filaNueva, cod: r.cod, quitadoDe: quitados });
    log(`  Anexo II: renglón "${r.rotulo}" agregado en la fila ${filaNueva}, leyendo ${r.cod} ` +
        `(${modificadas} fórmula(s) reacomodadas)` +
        (quitados.length ? `. Se lo saqué a ${quitados.join(", ")}` : "") + ".");
  }
  return { agregados, salteados };
}

if (typeof module !== "undefined") {
  const fh = require("./formula_hojas.js");
  global.insertRowEn = fh.insertRowEn;
  module.exports = { ROTULOS_ANEXO, agregarRotulosAnexo, rtUbicarBloque, rtQuitarTermino };
}
