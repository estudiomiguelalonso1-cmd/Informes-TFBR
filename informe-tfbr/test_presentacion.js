// Que el informe generado SE VEA como el original: fechas al día y filas nuevas con el mismo
// formato que sus vecinas.
//
// Las dos cosas fallan en silencio. Una fecha vieja en el encabezado no rompe ninguna fórmula:
// el informe cierra perfecto y dice "31 de Julio" en septiembre. Y una fila insertada nace sin
// tipografía ni formato de número, así que el renglón sale con otra letra y el importe sin
// separador de miles — en pantalla se ve el número, en el papel se ve distinto al resto.
//
// Correr con: node informe-tfbr/test_presentacion.js

const path = require("path");
const fs = require("fs");
const XLSX = require("xlsx");
const ExcelJS = require("exceljs");

global.XLSX = XLSX;
const P = require("./parser_tfbr.js");
const cfg = require("./config_tfbr.js");
const ch = require("./config_hojas.js");
const fi = require("./fechas_informe.js");
const motor = require("./motor_tfbr.js");

const PERIODO_EXPORT = "2026-07";
const PERIODO_NUEVO = "2026-09";      // se cierra septiembre con el export de julio: sólo
                                      // interesa cómo queda la presentación, no los importes
const MAESTROS = [
  { label: "Mensual $",    archivo: "base_bm_ars.xlsx", periodo: "mensual",   campo: "saldo_ars", id: "balance_mensual_ars" },
  { label: "Mensual R$",   archivo: "base_bm_brl.xlsx", periodo: "mensual",   campo: "saldo_brl", id: "balance_mensual_brl" },
  { label: "Acumulado $",  archivo: "base_ba_ars.xlsx", periodo: "acumulado", campo: "saldo_ars", id: "balance_acumulado_ars" },
  { label: "Acumulado R$", archivo: "base_ba_brl.xlsx", periodo: "acumulado", campo: "saldo_brl", id: "balance_acumulado_brl" },
];

function leerExport(periodo) {
  const buf = fs.readFileSync(path.join(__dirname, "..", "inputs", PERIODO_EXPORT, `sumas_y_saldos_${periodo}.xls`));
  const libro = XLSX.read(buf, { type: "buffer" });
  const ws = libro.Sheets[libro.SheetNames[0]];
  return P.parseSumasYSaldosTFBR(XLSX.utils.sheet_to_json(ws, { header: 1 }), ws["!merges"]).cuentas;
}

// Con qué letra y qué formato de número se ve una celda.
function pinta(cell) {
  const st = cell.style || {};
  return JSON.stringify({
    fuente: st.font ? `${st.font.name || "?"}/${st.font.size || "?"}${st.font.bold ? "/b" : ""}` : null,
    numero: st.numFmt || null,
  });
}

(async () => {
  let fallas = 0;
  const fallo = (m) => { console.log(`   ✗ ${m}`); fallas++; };
  const ex = { mensual: leerExport("mensual"), acumulado: leerExport("acumulado") };

  for (const m of MAESTROS) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path.join(__dirname, m.archivo));

    const antes = fi.fiPeriodoActualDe(wb);
    const res = motor.procesarMaestroTFBR({
      wb, cuentasExport: ex[m.periodo], campoSaldo: m.campo, archivoId: m.id,
      cuentasAcumulado: ex.acumulado, log: () => {},
    });
    // Las fechas las escribe el paso del período, que el motor no corre: es lo que hace app.js
    // con lo que se carga en pantalla.
    const fechas = fi.actualizarFechasDelInforme(wb, PERIODO_NUEVO, () => {});

    console.log(`\n== ${m.label}`);

    // 1. No puede quedar ninguna fecha del período viejo.
    const despues = fi.fiPeriodoActualDe(wb);
    const [anio, mes] = PERIODO_NUEVO.split("-").map(Number);
    if (!despues || despues.anio !== anio || despues.mes !== mes) {
      fallo(`la fecha que domina el archivo quedó en ` +
            `${despues ? fi.FI_MESES[despues.mes - 1] + " " + despues.anio : "ninguna"} ` +
            `y el período es ${fi.FI_MESES[mes - 1]} ${anio}`);
    } else {
      console.log(`   fechas: ${antes ? fi.FI_MESES[antes.mes - 1] + " " + antes.anio : "?"} → ` +
                  `${fi.FI_MESES[mes - 1]} ${anio}, ${fechas.cambiadas.length} celda(s)`);
    }

    // 2. Aplicarlo de nuevo no cambia nada.
    const otra = fi.actualizarFechasDelInforme(wb, PERIODO_NUEVO, () => {});
    if (otra.cambiadas.length) fallo(`una segunda pasada volvió a tocar ${otra.cambiadas.length} celda(s)`);

    // 3. TODO el bloque parejo, no sólo las filas que creamos.
    //
    //    Las plantillas originales ya traían filas sin formato —casi todas ocultas, en cero— y
    //    eso se nota apenas una recibe importe y se muestra: otra letra y el número sin
    //    separador de miles en el medio de un bloque prolijo. Así que se revisa fila por fila
    //    de cada bloque: ninguna puede quedar en blanco si su vecina tiene formato.
    const layout = cfg.derivarLayoutSaldos(wb);
    let revisadas = 0, sinFormato = 0;
    for (const h of ch.chHojasConfigurables(wb, layout)) {
      const mapa = ch.chMapaHoja(wb, layout, h.hoja);
      if (!mapa || mapa.renglones.length < 2) continue;
      const filas = mapa.renglones.map(x => x.fila).sort((a, b) => a - b);
      const desde = filas[0], hasta = filas[filas.length - 1];
      for (const col of mapa.colsImporte) {
        let vistoConFormato = false;
        for (let f = desde; f <= hasta; f++) {
          const st = mapa.ws.getCell(f, col).style;
          const tiene = st && (st.font || st.numFmt);
          if (tiene) { vistoConFormato = true; continue; }
          if (!vistoConFormato) continue;      // arriba todavía no había ninguna con formato
          revisadas++; sinFormato++;
          if (sinFormato <= 4) {
            fallo(`${h.hoja} fila ${f} col ${col}: quedó sin formato y arriba hay filas con formato`);
          }
        }
        revisadas++;
      }
    }
    console.log(`   formato: ${res.resumen.celdasEmparejadas || 0} celda(s) emparejadas; ` +
                (sinFormato ? `quedan ${sinFormato} SIN FORMATO` : "no queda ninguna sin formato"));
  }

  console.log(fallas ? `\n✗ ${fallas} falla(s).` : "\n✓ Las 4 pasan: fechas al día y filas nuevas con el formato de sus vecinas.");
  process.exit(fallas ? 1 : 0);
})();
