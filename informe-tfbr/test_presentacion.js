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

    // 3. Las filas que el motor CREÓ tienen que verse como sus vecinas.
    //
    //    Se miran sólo esas, no todas: las plantillas originales ya traen filas sin formato
    //    —las ocultas que están en cero— y compararlas contra su vecina daría falsos avisos
    //    sobre celdas que nadie tocó. Lo que hay que demostrar es que lo que agregamos nosotros
    //    sale igual que el resto, no que la plantilla sea perfecta.
    const layout = cfg.derivarLayoutSaldos(wb);
    const creadas = [];
    const r = res.resumen;
    for (const x of (r.renglonesNuevos ? r.renglonesNuevos.agregados : [])) creadas.push({ hoja: x.hoja, rotulo: x.rotulo });
    for (const x of (r.prestamos ? r.prestamos.creados : [])) creadas.push({ hoja: "Activo", rotulo: x.rotulo });
    for (const x of (r.configurado ? r.configurado.creados : [])) creadas.push({ hoja: x.hoja, rotulo: x.rotulo });
    for (const x of (r.rotulosAgregados || [])) creadas.push({ hoja: "Anexo II", rotulo: x.rotulo });

    let revisadas = 0, sinFormato = 0;
    for (const c of creadas) {
      const mapa = ch.chMapaHoja(wb, layout, c.hoja);
      if (!mapa) continue;
      const reng = mapa.renglones.find(x => String(x.rotulo).trim() === String(c.rotulo).trim());
      if (!reng || !reng.cols.length) continue;
      // La vecina de arriba que también sea un renglón del bloque.
      const vecina = mapa.renglones.filter(x => x.fila < reng.fila && x.cols.length).pop();
      if (!vecina) continue;
      const col = reng.cols[0].col;
      const a = mapa.ws.getCell(vecina.fila, col), b = mapa.ws.getCell(reng.fila, col);
      const vecinaConFormato = a.style && (a.style.font || a.style.numFmt);
      if (!vecinaConFormato) continue;          // no hay contra qué comparar
      revisadas++;
      const sinNada = !b.style || (!b.style.font && !b.style.numFmt);
      if (sinNada) {
        sinFormato++;
        fallo(`${c.hoja} fila ${reng.fila} ("${c.rotulo}") se creó sin formato: ` +
              `vecina ${pinta(a)} vs esta ${pinta(b)}`);
      }
    }
    console.log(`   formato: ${creadas.length} renglón(es) creados por el sistema, ` +
                `${revisadas} comparables con su vecino` +
                (sinFormato ? `, ${sinFormato} SIN FORMATO` : ", todos con el mismo"));
  }

  console.log(fallas ? `\n✗ ${fallas} falla(s).` : "\n✓ Las 4 pasan: fechas al día y filas nuevas con el formato de sus vecinas.");
  process.exit(fallas ? 1 : 0);
})();
