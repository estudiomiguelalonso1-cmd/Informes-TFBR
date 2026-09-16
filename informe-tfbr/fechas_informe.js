// Las fechas de cierre que el informe muestra, actualizadas al período que se está cerrando.
//
// El maestro de cada mes sale del mes anterior, así que arrastra la fecha vieja en todos los
// encabezados. Hasta acá sólo se actualizaba la etiqueta del tipo de cambio del Acumulado R$;
// el resto quedaba con el mes pasado y había que corregirlo a mano en cuatro archivos:
//
//   Mensual $     SALDOS!A2 "del 01/7/2026 al 31/7/2026", EESP fila 3, EEPN!B18, Activo!B3
//   Mensual R$    SALDOS!B2, SALDOS!F2, EESP fila 3, EEPN!B18
//   Acumulado $   EESP fila 7, EEPN!B17 "Total al 31.7.2026", Activo!B3
//   Acumulado R$  SALDOS!F1, EESP fila 7, EEPN!B17
//
// El EERR y el Anexo I no están en la lista porque toman la fecha del EESP por fórmula.
//
// Cómo se hace, y por qué así. No se escribe una fecha con un formato decidido acá: en estos
// archivos conviven "31/7/2026", "31/07/2026", "31.7.2026", "31/7/26", "31 de Julio de 2026" y
// "Al 31 de Julio del 2026". Se reemplaza el pedazo de texto que es la fecha y se deja el resto
// como está, respetando separador, ceros a la izquierda, año de dos o cuatro dígitos y cómo
// está escrito el mes.
//
// Y no se toca cualquier fecha: sólo LA FECHA DE CIERRE QUE EL ARCHIVO YA TRAE, que se deduce
// mirando cuál es la que más se repite. Así no se pisan fechas que no son del período — hay un
// 31/12/2005 en el Activo y un "2009" en la hoja Bienes que no tienen nada que ver.

const FI_MESES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
                  "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];

// "31/7/2026", "31/07/26", "31.7.2026", "31-07-2026"
const FI_RE_NUM = /\b(\d{1,2})([\/.\-])(\d{1,2})\2(\d{2}|\d{4})\b/g;
// "31 de Julio de 2026", "31 de Julio del 2026", "31 de Julio 2026"
const FI_RE_LARGA = new RegExp(
  "\\b(\\d{1,2})(\\s+de\\s+)(" + FI_MESES.join("|") + "|" +
  FI_MESES.map(m => m[0] + m.slice(1).toLowerCase()).join("|") + "|" +
  FI_MESES.map(m => m.toLowerCase()).join("|") + ")(\\s+(?:de[l]?\\s+)?)(\\d{4})\\b", "g");

function fiTexto(ws, r, c) {
  const v = ws.getCell(r, c).value;
  if (v == null) return null;
  if (typeof v === "object") {
    if (v.richText) return v.richText.map(t => t.text).join("");
    return null;                      // fórmulas y fechas de verdad no se tocan como texto
  }
  if (typeof v !== "string") return null;
  return v;
}

function fiUltimoDia(anio, mes) { return new Date(anio, mes, 0).getDate(); }

// Todas las fechas que aparecen en un texto, con dónde empiezan y de qué forma están escritas.
function fiFechasEn(texto) {
  const found = [];
  FI_RE_NUM.lastIndex = 0;
  for (const m of texto.matchAll(FI_RE_NUM)) {
    const anio = m[4].length === 2 ? 2000 + Number(m[4]) : Number(m[4]);
    found.push({ tipo: "num", i: m.index, largo: m[0].length, dia: +m[1], mes: +m[3], anio,
                 sep: m[2], diaPad: m[1].length === 2, mesPad: m[3].length === 2, anio2: m[4].length === 2 });
  }
  FI_RE_LARGA.lastIndex = 0;
  for (const m of texto.matchAll(FI_RE_LARGA)) {
    const mes = FI_MESES.findIndex(x => x === m[3].toUpperCase()) + 1;
    if (!mes) continue;
    found.push({ tipo: "larga", i: m.index, largo: m[0].length, dia: +m[1], mes, anio: +m[5],
                 sepDe: m[2], sepAnio: m[4], comoEscribeElMes: m[3] });
  }
  return found.sort((a, b) => a.i - b.i);
}

// El mes escrito como lo escribe el archivo: TODO MAYÚSCULAS, Capitalizado o minúsculas.
function fiMesComoEstaba(modelo, mes) {
  const nombre = FI_MESES[mes - 1];
  if (modelo === modelo.toUpperCase()) return nombre;
  if (modelo === modelo.toLowerCase()) return nombre.toLowerCase();
  return nombre[0] + nombre.slice(1).toLowerCase();
}

function fiEscribir(f, dia, mes, anio) {
  if (f.tipo === "num") {
    const d = f.diaPad ? String(dia).padStart(2, "0") : String(dia);
    const m = f.mesPad ? String(mes).padStart(2, "0") : String(mes);
    const a = f.anio2 ? String(anio).slice(-2) : String(anio);
    return `${d}${f.sep}${m}${f.sep}${a}`;
  }
  return `${f.dia === 1 ? "1" : String(dia)}${f.sepDe}${fiMesComoEstaba(f.comoEscribeElMes, mes)}${f.sepAnio}${anio}`;
}

// La fecha de cierre que el archivo trae hoy: la que más veces aparece. Se cuenta por mes y
// año, no por día, porque el mismo cierre se escribe con el día 1 (inicio del período) y con
// el último día (cierre) — "del 01/07/2026 al 31/07/2026" es un solo cierre, no dos fechas.
function fiPeriodoActualDe(wb) {
  const votos = {};
  for (const ws of wb.worksheets) {
    ws.eachRow({ includeEmpty: false }, (row, r) => {
      if (r > 60) return;
      row.eachCell({ includeEmpty: false }, (cell, c) => {
        const t = fiTexto(ws, r, c);
        if (!t) return;
        for (const f of fiFechasEn(t)) {
          const k = `${f.anio}-${f.mes}`;
          votos[k] = (votos[k] || 0) + 1;
        }
      });
    });
  }
  const ganador = Object.entries(votos).sort((a, b) => b[1] - a[1])[0];
  if (!ganador) return null;
  const [anio, mes] = ganador[0].split("-").map(Number);
  return { anio, mes, veces: ganador[1] };
}

// Actualiza las fechas del período al mes que se está cerrando.
//
// Sólo se tocan las fechas del mes que el archivo trae como cierre: una fecha de otro mes o de
// otro año es otra cosa (un dato histórico, una nota) y no se pisa.
//
// El día se conserva salvo que sea el último del mes: "del 01/07 al 31/07" tiene que pasar a
// "del 01/08 al 31/08", así que el 1 sigue siendo 1 y el 31 pasa a ser el último día de agosto.
function actualizarFechasDelInforme(wb, periodo, log = () => {}) {
  const [anioStr, mesStr] = String(periodo || "").split("-");
  const anio = parseInt(anioStr, 10), mes = parseInt(mesStr, 10);
  if (!anio || !mes || mes < 1 || mes > 12) {
    return { cambiadas: [], motivo: "el período no tiene la forma AAAA-MM" };
  }

  const actual = fiPeriodoActualDe(wb);
  if (!actual) return { cambiadas: [], motivo: "no encontré ninguna fecha de cierre en el archivo" };
  if (actual.anio === anio && actual.mes === mes) {
    return { cambiadas: [], yaEstaba: true, actual };
  }

  const ultimoNuevo = fiUltimoDia(anio, mes);
  const ultimoViejo = fiUltimoDia(actual.anio, actual.mes);
  const cambiadas = [];

  for (const ws of wb.worksheets) {
    ws.eachRow({ includeEmpty: false }, (row, r) => {
      if (r > 60) return;
      row.eachCell({ includeEmpty: false }, (cell, c) => {
        const t = fiTexto(ws, r, c);
        if (!t) return;
        const fechas = fiFechasEn(t).filter(f => f.anio === actual.anio && f.mes === actual.mes);
        if (!fechas.length) return;

        // De atrás para adelante, así los índices de las anteriores no se corren.
        let nuevo = t;
        for (const f of fechas.slice().reverse()) {
          const dia = f.dia === ultimoViejo ? ultimoNuevo : f.dia;
          nuevo = nuevo.slice(0, f.i) + fiEscribir(f, dia, mes, anio) + nuevo.slice(f.i + f.largo);
        }
        if (nuevo === t) return;
        cell.value = nuevo;
        cambiadas.push({ hoja: ws.name, celda: cell.address, de: t.trim(), a: nuevo.trim() });
      });
    });
  }

  if (cambiadas.length) {
    log(`  Fechas: ${cambiadas.length} celda(s) pasaron de ${FI_MESES[actual.mes - 1]} ` +
        `${actual.anio} a ${FI_MESES[mes - 1]} ${anio}.`);
    for (const c of cambiadas.slice(0, 6)) log(`    ${c.hoja}!${c.celda}: "${c.de}" → "${c.a}"`);
    if (cambiadas.length > 6) log(`    … y ${cambiadas.length - 6} más.`);
  }
  return { cambiadas, actual };
}

if (typeof module !== "undefined") {
  module.exports = {
    FI_MESES, actualizarFechasDelInforme, fiPeriodoActualDe, fiFechasEn, fiEscribir,
  };
}
