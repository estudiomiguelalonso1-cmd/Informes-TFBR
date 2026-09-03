// Intervenciones explícitas sobre defectos ya diagnosticados (docs/defect_diagnosis.md).
// Se aplican como paso separado del motor, DESPUÉS de escribir el staging — nunca en
// silencio dentro de otra función — para que quede claro en el log qué se tocó y por qué.
// Mismo espíritu que config/fixes/*.yml del scaffold abandonado (portado como datos, no
// como código Python): dos primitivas nomás, freeze_static y replace_formula.

// Solo las aprobadas. Las pendientes de confirmar con contaduría (EESP!F13, la familia de
// #REF! de Pasivo!K / Anexo II!G / EEPN!K17) están documentadas en
// docs/defect_diagnosis.md pero NO acá — no se tocan hasta que el usuario confirme.
const FIXES_APROBADOS_TFBR = {
  balance_mensual_ars: [
    {
      id: "anexo_ii_freeze_2012",
      descripcion:
        "Anexo II!D122/D126/D129/D130/D132 tenían un vínculo externo muerto a un archivo de " +
        "mayo 2012. Nada los usa río abajo (confirmado). Se congelan como valores fijos, " +
        "iguales a los últimos números en caché, para romper el vínculo sin cambiar lo que " +
        "se ve en el reporte.",
      acciones: [
        { hoja: "Anexo II", celda: "D122", accion: "freeze_static", valor: 3310746.86 },
        { hoja: "Anexo II", celda: "D126", accion: "freeze_static", valor: 2583219.17 },
        { hoja: "Anexo II", celda: "D129", accion: "freeze_static", valor: 78529.45 },
        { hoja: "Anexo II", celda: "D130", accion: "freeze_static", valor: 89824.30 },
        { hoja: "Anexo II", celda: "D132", accion: "freeze_static", valor: 36354509.30 },
      ],
    },
  ],
  balance_mensual_brl: [],
  balance_acumulado_ars: [],
  balance_acumulado_brl: [],
};

function ftColLetraANumero(letras) {
  return letras.toUpperCase().split("").reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0);
}

function ftCeldaANumeros(celda) {
  const m = /^([A-Z]{1,3})(\d+)$/i.exec(celda);
  if (!m) throw new Error(`Dirección de celda inválida: ${celda}`);
  return { col: ftColLetraANumero(m[1]), fila: parseInt(m[2], 10) };
}

// Aplica los fixes aprobados para `archivoId` sobre un workbook ya abierto con ExcelJS.
// Devuelve la lista de acciones efectivamente aplicadas, para loguear.
function aplicarFixesAprobados(wb, archivoId, log = () => {}) {
  const grupos = FIXES_APROBADOS_TFBR[archivoId] || [];
  const aplicadas = [];

  for (const grupo of grupos) {
    for (const a of grupo.acciones) {
      const ws = wb.getWorksheet(a.hoja);
      if (!ws) throw new Error(`Fix '${grupo.id}': el archivo no tiene la hoja '${a.hoja}'.`);
      const { col, fila } = ftCeldaANumeros(a.celda);
      const cell = ws.getCell(fila, col);

      if (a.accion === "freeze_static") {
        cell.value = a.valor;
      } else if (a.accion === "replace_formula") {
        cell.value = { formula: a.valor.replace(/^=/, "") };
      } else {
        throw new Error(`Fix '${grupo.id}': acción desconocida '${a.accion}'.`);
      }
      aplicadas.push(`${a.hoja}!${a.celda}`);
    }
    log(`  Fix aplicado: ${grupo.descripcion}`);
  }

  return aplicadas;
}

if (typeof module !== "undefined") {
  module.exports = { FIXES_APROBADOS_TFBR, aplicarFixesAprobados };
}
