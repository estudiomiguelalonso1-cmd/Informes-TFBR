// El saldo en reales, calculado a partir del export ORIGINAL del sistema.
//
// La regla es de contaduría:
//   cuentas 1, 2 y 3  ->  saldo en PESOS dividido el tipo de cambio de cierre
//   cuentas 4         ->  el saldo en reales que trae el export, tal cual
//
// Las columnas de debe y haber no se usan.
//
// Por qué hay un test dedicado. El error que esto previene no se ve: si la conversión queda mal
// —por el rubro, por el redondeo o porque se toma la columna equivocada— el balance en reales
// cierra igual y los números están todos corridos. Y el redondeo importa: medio centavo para el
// lado equivocado en 84 cuentas patrimoniales mueve el total.
//
// La referencia son los dos exports que contaduría ya convirtió a mano para agosto de 2026.
// Los cuatro archivos —los dos originales y los dos convertidos— están en inputs/2026-08/.
//
// Correr con: node informe-tfbr/test_reales.js

const path = require("path");
const fs = require("fs");
const XLSX = require("xlsx");
const P = require("./parser_tfbr.js");

const DIR = path.join(__dirname, "..", "inputs", "2026-08");
const TC = 291.3301;   // el que usó contaduría para convertir esos archivos

const CASOS = [
  { label: "mensual",   original: "sumas_y_saldos_mensual.xls",   convertido: "sumas_y_saldos_mensual_convertido.xls" },
  { label: "acumulado", original: "sumas_y_saldos_acumulado.xls", convertido: "sumas_y_saldos_acumulado_convertido.xls" },
];

function leer(archivo) {
  const buf = fs.readFileSync(path.join(DIR, archivo));
  const libro = XLSX.read(buf, { type: "buffer" });
  const ws = libro.Sheets[libro.SheetNames[0]];
  return P.parseSumasYSaldosTFBR(XLSX.utils.sheet_to_json(ws, { header: 1 }), ws["!merges"]);
}

(async () => {
  let fallas = 0;
  const fallo = (m) => { console.log(`   ✗ ${m}`); fallas++; };

  for (const c of CASOS) {
    const faltan = [c.original, c.convertido].filter(f => !fs.existsSync(path.join(DIR, f)));
    if (faltan.length) { fallo(`${c.label}: faltan los archivos ${faltan.join(", ")} en inputs/2026-08/`); continue; }

    const original = leer(c.original);
    const esperado = leer(c.convertido);
    const porCodigo = {};
    for (const x of esperado.cuentas) porCodigo[x.codigo] = x;

    const convertidas = P.convertirSaldosEnReales(original.cuentas, TC);

    let comparadas = 0, distintas = 0;
    const porRubro = {};
    for (const x of convertidas) {
      const e = porCodigo[x.codigo];
      if (!e) continue;
      comparadas++;
      const rubro = String(x.codigo)[0];
      porRubro[rubro] = (porRubro[rubro] || 0) + 1;
      if (Math.abs(x.saldo_brl - e.saldo_brl) > 0.005) {
        distintas++;
        if (distintas <= 5) {
          fallo(`${c.label} ${x.codigo} ${x.nombre.slice(0, 24)}: calculado ` +
                `${x.saldo_brl.toFixed(2)} y contaduría puso ${e.saldo_brl.toFixed(2)}`);
        }
      }
    }

    console.log(`\n== ${c.label}: ${comparadas} cuentas comparadas (` +
                Object.entries(porRubro).sort().map(([r, n]) => `rubro ${r}: ${n}`).join(", ") + ")");
    if (!distintas) console.log("   todas iguales al centavo");

    // Las de rubro 4 no se tocan: tienen que quedar exactamente como venían.
    const tocadas4 = convertidas.filter(x => String(x.codigo)[0] === "4" &&
      x.saldo_brl_export !== undefined && x.saldo_brl !== x.saldo_brl_export);
    if (tocadas4.length) fallo(`${c.label}: se tocaron ${tocadas4.length} cuenta(s) de resultados y no había que tocarlas`);

    // Y las patrimoniales sí: si alguna quedó igual al export, la conversión no se aplicó.
    const sinTocar = convertidas.filter(x => String(x.codigo)[0] !== "4" &&
      Math.abs(x.saldo_ars) > 1 && x.saldo_brl === x.saldo_brl_export);
    if (sinTocar.length) {
      fallo(`${c.label}: ${sinTocar.length} cuenta(s) patrimoniales quedaron con el saldo del ` +
            `export en vez del convertido (ej. ${sinTocar[0].codigo})`);
    }
  }

  // Sin tipo de cambio no se puede: tiene que frenar, no seguir con cualquier número.
  try {
    P.convertirSaldosEnReales([{ codigo: "1110100330", saldo_ars: 100, saldo_brl: 1 }], "");
    fallo("sin tipo de cambio la conversión siguió adelante en vez de frenar");
  } catch (e) { /* esperado */ }

  console.log(fallas ? `\n✗ ${fallas} falla(s).` : "\n✓ Pasa: el saldo en reales sale igual al que hace contaduría.");
  process.exit(fallas ? 1 : 0);
})();
