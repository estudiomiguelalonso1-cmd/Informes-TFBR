// Que la fecha de cierre quede bien en TODOS lados, no solo en los títulos.
//
// Cada informe tiene una celda con una FECHA de verdad —"EESP!F9" en los mensuales, "EESP!F13"
// en los acumulados— y la leen por fórmula el EERR, el EEPN, el Activo, el Pasivo, el Anexo I
// y el Anexo II. El actualizador de fechas solo tocaba textos, así que esa celda se quedaba con
// el mes anterior: los títulos decían "31 de Agosto de 2026" y seis recuadros 31/07/2026.
//
// Lo otro que cuida este test es lo contrario: que NO se toque el 31/12/2005 que hay en el EESP
// y en el Activo. No es del período —es un dato fijo— y pisarlo sería peor que el problema.
//
// Correr con: node informe-tfbr/test_fechas.js

const path = require("path");
const fs = require("fs");
const ExcelJS = require("exceljs");

global.ExcelJS = ExcelJS;
const FU = require("./formula_utils.js");
const fi = require("./fechas_informe.js");

const PERIODO = "2026-08";
const DIA = 31;
const ESPERADA = "2026-08-31";
const FIJA = "2005-12-31";        // no es del período: tiene que quedar como está

const MAESTROS = [
  { label: "Mensual $",    archivo: "base_bm_ars.xlsx" },
  { label: "Mensual R$",   archivo: "base_bm_brl.xlsx" },
  { label: "Acumulado $",  archivo: "base_ba_ars.xlsx" },
  { label: "Acumulado R$", archivo: "base_ba_brl.xlsx" },
];

const iso = (d) => d.toISOString().slice(0, 10);

// Toda celda con una fecha de verdad, y quién la lee por fórmula desde otra hoja.
function fechasReales(wb) {
  const celdas = [];
  for (const ws of wb.worksheets) {
    ws.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (cell.value instanceof Date) {
          celdas.push({ hoja: ws.name, dir: cell.address, fecha: iso(cell.value) });
        }
      });
    });
  }
  for (const x of celdas) {
    x.lectores = [];
    const re = new RegExp(`'?${x.hoja}'?!${x.dir}(?![0-9])`);
    for (const ws of wb.worksheets) {
      ws.eachRow({ includeEmpty: false }, (row) => {
        row.eachCell({ includeEmpty: false }, (cell) => {
          const v = cell.value;
          if (v && typeof v === "object" && typeof v.formula === "string" && re.test(v.formula)) {
            x.lectores.push(`${ws.name}!${cell.address}`);
          }
        });
      });
    }
  }
  return celdas;
}

(async () => {
  let fallas = 0;
  const fallo = (m) => { console.log(`   ✗ ${m}`); fallas++; };

  for (const m of MAESTROS) {
    const wb = await FU.abrirWorkbook(fs.readFileSync(path.join(__dirname, m.archivo)));

    const antes = fechasReales(wb);
    const deCierre = antes.filter(x => x.fecha !== FIJA);
    if (!deCierre.length) { fallo(`${m.label}: no encontré ninguna celda con la fecha de cierre`); continue; }

    fi.actualizarFechasDelInforme(wb, PERIODO, () => {}, DIA);
    const despues = fechasReales(wb);

    console.log(`\n== ${m.label}`);
    for (const x of despues) {
      const previa = antes.find(a => a.hoja === x.hoja && a.dir === x.dir);
      const lectores = x.lectores.length ? `la leen ${x.lectores.length}` : "no la lee nadie";
      console.log(`   ${(x.hoja + "!" + x.dir).padEnd(16)} ${previa.fecha} → ${x.fecha}  (${lectores})`);

      if (previa.fecha === FIJA) {
        // El dato fijo no se toca.
        if (x.fecha !== FIJA) fallo(`${m.label} ${x.hoja}!${x.dir}: pisó el ${FIJA}, que no es del período`);
      } else if (x.fecha !== ESPERADA) {
        fallo(`${m.label} ${x.hoja}!${x.dir}: quedó en ${x.fecha} y tendría que decir ${ESPERADA}` +
              (x.lectores.length ? ` — la leen ${x.lectores.join(", ")}` : ""));
      }
    }

    // La de cierre es la que arrastra a las otras hojas: si no la lee nadie, algo cambió de
    // lugar y este test dejaría de cuidar lo que importa.
    const cierre = despues.filter(x => x.fecha === ESPERADA);
    if (!cierre.some(x => x.lectores.length >= 3)) {
      fallo(`${m.label}: la fecha de cierre no la lee ninguna otra hoja — revisá si se movió`);
    }
  }

  console.log(fallas ? `\n✗ ${fallas} falla(s).` : "\n✓ La fecha de cierre queda bien en las celdas de fecha, y el dato fijo no se toca.");
  process.exit(fallas ? 1 : 0);
})();
