// Rangos del plan de cuentas convertidos a rastreo cuenta por cuenta.
//
// Qué pasaba. Algunas líneas no nombraban las cuentas: tomaban un tramo entero del plan.
// "- Proveedores" del Pasivo era -SUM(SALDOS!D53:D79)-SALDOS!D88-SALDOS!D101-SALDOS!D114-
// SALDOS!D115: un rango de 27 filas, más cuatro proveedores sueltos que alguien fue
// agregando a mano cuando quedaron fuera del tramo.
//
// Por qué molesta. El rango es una apuesta a que el plan no se mueva. Una cuenta nueva entra
// o no entra según dónde caiga su código, y nadie lo decide: lo decide el orden alfabético.
// Si cae adentro, se suma sin que nadie lo haya aprobado; si cae afuera, no se suma y no
// avisa. Los cuatro términos agregados a mano son la prueba de que ya pasó. Y al revés
// también: si el día de mañana se borra una cuenta del medio del tramo, el rango se encoge
// solo y el renglón cambia de valor sin que nadie haya tocado nada.
//
// Qué hace esto. Reemplaza cada SUM(SALDOS!D53:D79) por la suma explícita de las cuentas que
// hoy están adentro: (SALDOS!D53+SALDOS!D54+...). El número no cambia — es la misma suma
// escrita término a término — pero a partir de ahí cada cuenta está nombrada, y una cuenta
// nueva no entra sola: aparece en el control de cuentas sin destino y hay que decidir.
//
// Cuándo corre. DESPUÉS del alta de cuentas nuevas del mes, a propósito: las cuentas que
// entran este mes todavía las levanta el rango como siempre, y recién quedan fijadas después.
// Corriendo antes, una cuenta nueva que hoy el rango levanta se perdería en silencio, que es
// justo lo que se quiere evitar.
//
// Qué NO hace. Si dentro del rango hay una fila que no es una cuenta y tiene un número (un
// subtotal metido en el medio), no toca esa fórmula y lo informa: expandirla cambiaría el
// resultado. Hoy no pasa en ninguno de los cuatro archivos — los rangos son 100% filas de
// cuenta — pero el día que pase hay que mirarlo, no adivinar.

function rrTexto(ws, r, c) {
  const v = ws.getCell(r, c).value;
  if (v == null) return "";
  if (typeof v === "object") return v.richText ? v.richText.map(t => t.text).join("") : "";
  return String(v);
}

function rrNumero(ws, r, c) {
  const v = ws.getCell(r, c).value;
  if (typeof v === "number") return v;
  if (v && typeof v === "object" && typeof v.result === "number") return v.result;
  return 0;
}

function rrEsCuenta(texto) {
  return /^\s*[\d.]{6,}/.test(String(texto || ""));
}

// SUM(SALDOS!D53:D79) — un rango de una sola columna dentro de un SUM.
const RR_RE_SUM = /SUM\(\s*SALDOS!\$?([A-Z]{1,3})\$?(\d+)\s*:\s*\$?([A-Z]{1,3})\$?(\d+)\s*\)/gi;
// Un rango suelto, sin SUM alrededor: no se toca, pero se informa.
const RR_RE_RANGO = /SALDOS!\$?[A-Z]{1,3}\$?\d+\s*:\s*\$?[A-Z]{1,3}\$?\d+/i;

function expandirRangosSaldos(wb, layout, log = () => {}) {
  const S = wb.getWorksheet(layout.sheet);
  const expandidos = [], salteados = [];

  for (const ws of wb.worksheets) {
    if (ws.name === layout.sheet) continue;
    ws.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const v = cell.value;
        if (!v || typeof v !== "object" || typeof v.formula !== "string") return;
        if (!RR_RE_RANGO.test(v.formula)) return;

        const donde = `${ws.name}!${cell.address}`;
        let corto = false;
        const nueva = v.formula.replace(RR_RE_SUM, (entero, colA, desde, colB, hasta) => {
          if (colA.toUpperCase() !== colB.toUpperCase()) {   // rango que cruza columnas
            salteados.push({ donde, formula: entero, motivo: "el rango abarca más de una columna" });
            corto = true; return entero;
          }
          const d = Math.min(+desde, +hasta), h = Math.max(+desde, +hasta);
          const col = colA.toUpperCase();
          const filas = [];
          const intrusas = [];
          for (let r = d; r <= h; r++) {
            const texto = rrTexto(S, r, layout.keyCol).trim();
            if (rrEsCuenta(texto)) { filas.push(r); continue; }
            // Una fila que no es cuenta solo es inofensiva si además está en cero: en un SUM
            // aporta nada y sacarla no cambia el resultado.
            if (Math.abs(rrNumero(S, r, col.split("").reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0))) > 0.005) {
              intrusas.push(`fila ${r} "${texto.slice(0, 40)}"`);
            }
          }
          if (intrusas.length) {
            salteados.push({
              donde, formula: entero,
              motivo: `dentro del rango hay ${intrusas.length} fila(s) que no son cuentas y tienen ` +
                      `importe (${intrusas.join(", ")}); expandirlo cambiaría el número`,
            });
            corto = true; return entero;
          }
          if (!filas.length) {
            salteados.push({ donde, formula: entero, motivo: "el rango no contiene ninguna cuenta" });
            corto = true; return entero;
          }
          expandidos.push({ donde, rango: `${col}${d}:${col}${h}`, cuentas: filas.length });
          return `(${filas.map(r => `SALDOS!${col}${r}`).join("+")})`;
        });

        if (corto || nueva === v.formula) return;
        cell.value = { formula: nueva };
      });
    });
  }

  for (const e of expandidos) {
    log(`  ${e.donde}: el rango ${e.rango} pasó a nombrar sus ${e.cuentas} cuenta(s) una por una.`);
  }
  for (const s of salteados) {
    log(`  ⚠ ${s.donde}: NO expandí "${s.formula}" — ${s.motivo}.`);
  }
  return { expandidos, salteados };
}

// Qué hoja lee cada fila del plan. Con los rangos ya expandidos esto es exacto: cada
// referencia nombra una fila.
function quienLeeCadaFila(wb, layout) {
  const lee = {};
  for (const ws of wb.worksheets) {
    if (ws.name === layout.sheet) continue;
    ws.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const v = cell.value;
        if (!v || typeof v !== "object" || typeof v.formula !== "string") return;
        let f = v.formula;
        // Un rango que haya quedado sin expandir igual cuenta como lectura de todo el tramo:
        // si no, se denunciarían como perdidas cuentas que sí entran en el total.
        for (const m of f.matchAll(/SALDOS!\$?[A-Z]{1,3}\$?(\d+)\s*:\s*\$?[A-Z]{1,3}\$?(\d+)/g)) {
          for (let r = +m[1]; r <= +m[2]; r++) (lee[r] = lee[r] || new Set()).add(ws.name);
        }
        f = f.replace(/SALDOS!\$?[A-Z]{1,3}\$?\d+\s*:\s*\$?[A-Z]{1,3}\$?\d+/g, "");
        for (const m of f.matchAll(/SALDOS!\$?[A-Z]{1,3}\$?(\d+)/g)) {
          (lee[+m[1]] = lee[+m[1]] || new Set()).add(ws.name);
        }
      });
    });
  }
  return lee;
}

// Plata de ESTE mes que no llega a ninguna hoja.
//
// Se mira contra lo que el motor acaba de escribir, no contra los saldos que el archivo trae
// en caché: esos son los del mes pasado y dirían que sobra o falta lo que no corresponde.
//
// Este es el control que no existía. Una cuenta con importe que ninguna hoja lee entra en los
// totales de SALDOS —así que Debe sigue igual a Haber y el balance cierra— pero no aparece en
// ningún estado. No hay forma de notarlo mirando el resultado.
function cuentasSinDestino(wb, layout, planDeCuentas, escritas) {
  const lee = quienLeeCadaFila(wb, layout);
  const sueltas = [];
  for (const [cod, importe] of Object.entries(escritas || {})) {
    if (Math.abs(importe) < 0.005) continue;
    const info = planDeCuentas[cod];
    if (!info) continue;                       // sin fila en el plan: ya se avisa aparte
    const filas = [info].concat(info.otrasFilas || []);
    if (filas.some(f => lee[f.fila])) continue;
    sueltas.push({ cod, nom: (info.texto || "").replace(/^\s*[\d.]+\s*-?\s*/, "").trim(), importe });
  }
  return sueltas.sort((a, b) => Math.abs(b.importe) - Math.abs(a.importe));
}

if (typeof module !== "undefined") {
  module.exports = {
    expandirRangosSaldos, quienLeeCadaFila, cuentasSinDestino, RR_RE_SUM, RR_RE_RANGO,
  };
}
