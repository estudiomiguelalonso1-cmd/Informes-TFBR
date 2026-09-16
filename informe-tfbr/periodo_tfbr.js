// Los datos del período que NO vienen del export de Onvio y hoy se tipean a mano:
// el tipo de cambio de cierre y la cifra mensual del cuadro "Explicación dif de cambio".
//
// Los dos viven solo en el Balance Acumulado R$ (los otros tres archivos no los tienen).
// Igual que el resto del motor, las celdas no están hardcodeadas: se ubican leyendo el
// archivo, porque una dirección anotada a mano se desactualiza sin avisar — de hecho, la
// primera investigación de estos archivos tenía mal la celda del TC por ese motivo.

const MESES_ES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
                  "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];

// "D" a partir del número de columna, para poder nombrar celdas concretas en los avisos.
// Local a este archivo (config_tfbr.js tiene ctColNumeroALetra, pero periodo_tfbr.js no
// depende de él en Node y no vale la pena acoplarlos por un aviso).
function pfColLetra(n) {
  let s = "";
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = (n - r - 1) / 26; }
  return s;
}

function pfTexto(ws, fila, col) {
  const v = ws.getCell(fila, col).value;
  if (v === null || v === undefined) return "";
  if (typeof v === "object") {
    if (v.richText) return v.richText.map(t => t.text).join("");
    if (v.result !== undefined && typeof v.result !== "object") return String(v.result);
    return "";
  }
  return String(v);
}

// El TC se reconoce por su etiqueta ("T.C. Reales al 31/07/2026"), y el valor es la celda
// de al lado. Devuelve null si este archivo no tiene TC — que es el caso de 3 de los 4.
function ubicarTcCierre(ws) {
  for (let r = 1; r <= Math.min(ws.rowCount, 6); r++) {
    for (let c = 1; c <= Math.min(ws.columnCount, 15); c++) {
      const t = pfTexto(ws, r, c).trim();
      if (!/^T\.?\s*C\.?\b/i.test(t)) continue;
      const valor = ws.getCell(r, c + 1).value;
      if (typeof valor !== "number") continue;
      return { filaEtiqueta: r, colEtiqueta: c, colValor: c + 1, etiqueta: t };
    }
  }
  return null;
}

// El cuadro arranca con el título "EXPLICACIÓN DIF DE CAMBIO" y sigue con una fila por mes,
// más una de acumulado, y cierra con el SUM. Cada fila se identifica por su texto, así que
// no hace falta saber de antemano qué fila es cada mes (ni suponer que agosto es la
// siguiente a julio: la fila de acumulado está justo en el medio).
function ubicarCuadroDifCambio(ws) {
  let filaTitulo = null, colTitulo = null;
  for (let r = 1; r <= ws.rowCount && filaTitulo === null; r++) {
    for (let c = 1; c <= Math.min(ws.columnCount, 15); c++) {
      const t = pfTexto(ws, r, c).trim().toUpperCase();
      if (t.startsWith("EXPLICACI") && t.includes("DIF") && t.includes("CAMBIO")) {
        filaTitulo = r; colTitulo = c; break;
      }
    }
  }
  if (filaTitulo === null) return null;

  const colValor = colTitulo + 1;
  const filas = [];
  for (let r = filaTitulo + 1; r <= Math.min(ws.rowCount, filaTitulo + 30); r++) {
    // La fila del total lleva el SUM y NO lleva etiqueta, así que se mira antes que el
    // texto: si se saltea por no tener etiqueta, el recorrido sigue de largo y termina
    // leyendo la zona de pegado como si fueran meses.
    const v = ws.getCell(r, colValor).value;
    if (v && typeof v === "object" && typeof v.formula === "string") {
      return pfCerrarCuadro({ filaTitulo, colEtiqueta: colTitulo, colValor, filas,
                              filaTotal: r, formulaTotal: v.formula });
    }
    const t = pfTexto(ws, r, colTitulo).trim();
    if (!t) continue;
    const mes = MESES_ES.findIndex(m => t.toUpperCase().startsWith(m)) + 1;
    filas.push({ fila: r, etiqueta: t, mes: mes || null });
  }
  return pfCerrarCuadro({ filaTitulo, colEtiqueta: colTitulo, colValor, filas,
                          filaTotal: null, formulaTotal: null });
}

// Dónde habría que insertar la fila de un mes que todavía no está: justo debajo del último
// mes cargado. NO al final del cuadro — después del último mes viene la línea de acumulado
// semestral, y el SUM del total la incluye. Una fila agregada debajo de esa línea cae fuera
// del rango del SUM (Excel solo lo estira cuando se inserta DENTRO), así que el importe no
// sumaría y el único síntoma sería el tie-out de más abajo dando distinto, sin decir por qué.
function pfCerrarCuadro(cuadro) {
  const meses = cuadro.filas.filter(f => f.mes !== null);
  cuadro.ultimoMes = meses.length ? meses[meses.length - 1] : null;
  // las filas sin mes que quedaron DESPUÉS del último mes (la de acumulado semestral)
  cuadro.filasNoMes = cuadro.ultimoMes
    ? cuadro.filas.filter(f => f.mes === null && f.fila > cuadro.ultimoMes.fila)
    : cuadro.filas.filter(f => f.mes === null);
  cuadro.filaAInsertar = cuadro.ultimoMes ? cuadro.ultimoMes.fila + 1 : null;
  return cuadro;
}

function pfUltimoDiaDelMes(anio, mes) {
  return new Date(anio, mes, 0).getDate();
}

// La etiqueta del mes nuevo, calcada del formato que ya usan las filas del cuadro
// ("JULIO 01/07/2026-31/07/2026"). Se copia el espaciado alrededor del guión del último mes
// en vez de fijarlo acá: la fila de acumulado usa " - " con espacios y los meses no, así que
// el formato del cuadro es algo que se lee del archivo, no que se decide en el código.
function pfEtiquetaDelMes(etiquetaModelo, anio, mes, diaCierre) {
  const ultimo = pfUltimoDiaDelMes(anio, mes);
  const dd = String(diaCierre ? Math.min(diaCierre, ultimo) : ultimo).padStart(2, "0");
  const mm = String(mes).padStart(2, "0");
  const desde = `01/${mm}/${anio}`;
  const hasta = `${dd}/${mm}/${anio}`;
  // separador tal cual lo escribe el archivo, si se puede leer del modelo
  const m = /\d{2}\/\d{2}\/\d{4}(\s*-\s*)\d{2}\/\d{2}\/\d{4}/.exec(etiquetaModelo || "");
  const sep = m ? m[1] : "-";
  return `${MESES_ES[mes - 1]} ${desde}${sep}${hasta}`;
}

// Agrega la fila del mes al cuadro. Devuelve {ok:false, motivo} sin tocar nada cuando no se
// puede hacer con seguridad — insertar una fila corre todo lo que está debajo, así que es
// preferible dejarlo pendiente y avisar antes que mover algo que rompa los controles.
function pfInsertarFilaDelMes(wb, ws, cuadro, anio, mes, valor, log, diaCierre) {
  if (!cuadro.filaAInsertar || !cuadro.ultimoMes) {
    return { ok: false, motivo: "no se pudo ubicar el último mes del cuadro" };
  }
  // Un hueco (cerrar septiembre con agosto sin cargar) casi siempre es un período mal
  // elegido: el maestro de cada mes sale del anterior. Insertar acá taparía el hueco en
  // lugar de mostrarlo, así que se frena y se avisa.
  if (cuadro.ultimoMes.mes && mes > cuadro.ultimoMes.mes + 1) {
    return { ok: false, motivo: "hay meses sin cargar en el medio" };
  }
  if (cuadro.ultimoMes.mes && mes <= cuadro.ultimoMes.mes) {
    return { ok: false, motivo: "el mes es anterior al último cargado" };
  }
  // El plan de cuentas guarda el número de fila de cada cuenta, y los controles del cierre lo
  // usan tal cual. Si la fila nueva cayera dentro del plan, esos números quedarían corridos y
  // el control de "cada cuenta se levantó" fallaría por todos lados sin motivo real. Hoy el
  // cuadro está debajo del plan, pero se verifica en vez de darlo por sentado.
  const layout = derivarLayoutSaldos(wb);
  if (cuadro.filaAInsertar <= layout.planDeCuentas.hasta) {
    return { ok: false, motivo: "la fila caería dentro del plan de cuentas" };
  }

  const modificadas = insertRowEn(wb, layout.sheet, cuadro.filaAInsertar);
  // El mes nuevo del cuadro, con el formato del mes anterior (que la inserción corrió uno).
  copiarFormatoDeFila(ws, cuadro.filaAInsertar + 1, cuadro.filaAInsertar);
  const etiqueta = pfEtiquetaDelMes(cuadro.ultimoMes.etiqueta, anio, mes, diaCierre);
  ws.getCell(cuadro.filaAInsertar, cuadro.colEtiqueta).value = etiqueta;
  ws.getCell(cuadro.filaAInsertar, cuadro.colValor).value = valor;
  log(`  Explicación dif de cambio: fila nueva ${cuadro.filaAInsertar} "${etiqueta}" = ${valor} ` +
      `(${modificadas} fórmula(s) reacomodadas; el total la toma solo).`);
  return { ok: true, fila: cuadro.filaAInsertar, etiqueta, modificadas };
}

// El texto que se lee en el checklist, con el Excel abierto al lado. Dice la fila exacta y
// contra qué etiquetas verificar que cayó en el lugar correcto, porque una instrucción del
// tipo "agregala al cuadro" lleva justo al error que se quiere evitar.
function pfAvisoFaltaFilaMes(cuadro, mes, valor, motivo) {
  const nombreMes = MESES_ES[mes - 1];
  const col = pfColLetra(cuadro.colValor);
  let msg = `El cuadro "Explicación dif de cambio" no tiene fila para ${nombreMes}` +
            (motivo ? ` y no la agregué sola (${motivo})` : "") +
            `: hay que agregarla a mano en el Excel`;

  if (cuadro.filaAInsertar && cuadro.filasNoMes.length) {
    // el caso de hoy: después del último mes viene la línea de acumulado, y el total la suma
    msg += `. Insertá una fila en la ${cuadro.filaAInsertar}, debajo de ` +
           `"${cuadro.ultimoMes.etiqueta}" y ARRIBA de "${cuadro.filasNoMes[0].etiqueta}"` +
           (cuadro.filaTotal
             ? `, así el total (${col}${cuadro.filaTotal}${cuadro.formulaTotal ? ` = ${cuadro.formulaTotal}` : ""}) la toma solo`
             : "") +
           `. Si la agregás debajo de "${cuadro.filasNoMes[0].etiqueta}" queda FUERA del total ` +
           `y el importe no suma`;
  } else if (cuadro.filaAInsertar) {
    msg += `. Va en la fila ${cuadro.filaAInsertar}, debajo de "${cuadro.ultimoMes.etiqueta}"` +
           (cuadro.filaTotal ? `, dentro del rango que suma ${col}${cuadro.filaTotal}` : "");
  }

  if (valor !== null && valor !== undefined) {
    msg += `. El importe que cargaste (${valor}) NO se escribió: ponelo en esa fila`;
  }

  // Si entre el último mes del cuadro y el que se está cerrando hay un hueco, lo más probable
  // es que se haya elegido mal el período (el maestro de cada mes sale del anterior, así que
  // saltear uno no debería poder pasar). Sin este aviso, el texto de arriba manda a insertar
  // el mes nuevo pegado al último cargado, tapando el hueco en vez de mostrarlo.
  if (cuadro.ultimoMes && cuadro.ultimoMes.mes && mes > cuadro.ultimoMes.mes + 1) {
    const faltan = [];
    for (let m = cuadro.ultimoMes.mes + 1; m < mes; m++) faltan.push(MESES_ES[m - 1]);
    msg += ` ⚠ OJO: el cuadro llega hasta ${MESES_ES[cuadro.ultimoMes.mes - 1]} y falta(n) ` +
           `${faltan.join(", ")} en el medio. Revisá que el período sea el correcto antes de ` +
           `insertar nada`;
  }
  return msg + ".";
}

// La línea "DIFERENCIA DE CAMBIO" del plan: la que hace cerrar el balance en reales.
//
// Existe en los dos archivos en R$, no sólo en el Acumulado. Está justo después de la última
// cuenta y entra en los subtotales de la hoja: lleva el importe en la columna acreedora y el
// mismo número con el signo cambiado en la columna del saldo.
//
// Por qué hace falta. Las cuentas patrimoniales se convierten al tipo de cambio de cierre y las
// de resultado vienen acumuladas a sus propios cambios: la diferencia entre los dos criterios es
// exactamente la diferencia de cambio del período, y sin esta línea el balance en reales no
// cierra. Antes cerraba solo porque todas las cuentas salían de debe menos haber en reales.
function pfUbicarLineaDifCambio(ws, layout) {
  for (let r = layout.planDeCuentas.hasta - 2; r < layout.stagingRange.filaDesde; r++) {
    for (let c = 1; c <= 4; c++) {
      const t = pfTexto(ws, r, c).trim();
      if (/^diferencia\s+de\s+cambio$/i.test(t)) return { fila: r, colEtiqueta: c };
    }
  }
  return null;
}

function pfEscribirLineaDifCambio(ws, layout, linea, valor, log) {
  // Acreedora el importe, y la columna del saldo el mismo número con el signo cambiado, que es
  // como está escrita la línea en los dos archivos.
  ws.getCell(linea.fila, layout.acreedorCol).value = valor;
  ws.getCell(linea.fila, layout.saldoCol).value = -valor;
  log(`  Línea "DIFERENCIA DE CAMBIO" (fila ${linea.fila}): ${valor}.`);
}

// Escribe en el maestro lo que corresponda de este período. Devuelve qué escribió y qué no,
// para que la pantalla lo muestre: un dato que no se pudo escribir tiene que quedar a la
// vista como paso manual pendiente, no desaparecer.
// La suma del cuadro de meses, leyendo las filas que tienen importe.
function pfTotalDelCuadro(ws, cuadro) {
  let total = 0, hubo = false;
  for (const f of cuadro.filas) {
    const v = ws.getCell(f.fila, cuadro.colValor).value;
    const n = typeof v === "number" ? v
      : (v && typeof v === "object" && typeof v.result === "number") ? v.result : null;
    if (n === null) continue;
    total += n; hubo = true;
  }
  return hubo ? total : null;
}

function escribirDatosDelPeriodo(wb, { periodo, tcCierre, difCambioMes, diaCierre }, log = () => {}) {
  const ws = wb.getWorksheet("SALDOS");
  const layout0 = derivarLayoutSaldos(wb);
  const [anioStr, mesStr] = String(periodo).split("-");
  const anio = parseInt(anioStr, 10);
  const mes = parseInt(mesStr, 10);
  const hecho = [];
  const pendiente = [];

  // Las fechas de todos los encabezados, en todas las hojas. Va primero porque la etiqueta del
  // tipo de cambio es una de ellas: si el TC se escribiera antes con su propia fecha, después
  // vendría el actualizador y la volvería a tocar, y el formato de cada archivo se perdería.
  const fechas = actualizarFechasDelInforme(wb, periodo, log, diaCierre || null);
  if (fechas.cambiadas.length) {
    hecho.push(`${fechas.cambiadas.length} fecha(s) de encabezado actualizadas al cierre del período`);
  } else if (fechas.motivo) {
    pendiente.push(`No pude actualizar las fechas del informe: ${fechas.motivo}. Revisá los ` +
                   `encabezados a mano antes de entregarlo.`);
  }

  const tc = ubicarTcCierre(ws);
  if (tc && tcCierre !== null && tcCierre !== undefined && tcCierre !== "") {
    // Sólo el número. La etiqueta con la fecha ya la actualizó actualizarFechasDelInforme,
    // respetando el formato que usa cada archivo ("31/7/26" en uno, "31/07/2026" en otro).
    // El TC se tipea con coma decimal ("291,3301"): Number() de eso da NaN y escribía basura.
    ws.getCell(tc.filaEtiqueta, tc.colValor).value = ptNumeroTipeado(tcCierre);
    hecho.push(`TC de cierre ${tcCierre} escrito en SALDOS`);
    log(`  TC de cierre: ${tcCierre}.`);
  }

  const cuadro = ubicarCuadroDifCambio(ws);
  if (cuadro) {
    const fila = cuadro.filas.find(f => f.mes === mes);
    const valor = ptNumeroTipeado(difCambioMes);
    if (fila && valor !== null) {
      ws.getCell(fila.fila, cuadro.colValor).value = valor;
      hecho.push(`Diferencia de cambio de ${fila.etiqueta}: ${valor}`);
      log(`  Explicación dif de cambio: ${valor} en "${fila.etiqueta}".`);
    } else if (!fila && valor !== null) {
      // La fila del mes nuevo se inserta sola, en el único lugar donde el total la toma:
      // debajo del último mes y arriba del acumulado semestral (ver pfCerrarCuadro). Si por
      // lo que sea no se puede hacer con seguridad, no se fuerza: queda el aviso con la
      // ubicación exacta para hacerlo a mano.
      const ins = pfInsertarFilaDelMes(wb, ws, cuadro, anio, mes, valor, log, diaCierre || null);
      if (ins.ok) {
        hecho.push(`Diferencia de cambio de ${ins.etiqueta}: ${valor} (fila ${ins.fila}, agregada al cuadro)`);
      } else {
        pendiente.push(pfAvisoFaltaFilaMes(cuadro, mes, valor, ins.motivo));
        log(`  ⚠ No agregué la fila de ${MESES_ES[mes - 1]} (${ins.motivo}): queda pendiente a mano.`);
      }
    } else if (!fila) {
      // Sin importe cargado no hay nada que escribir, así que tampoco se inserta la fila:
      // una fila vacía en el cuadro no ayuda y habría que sacarla después.
      pendiente.push(pfAvisoFaltaFilaMes(cuadro, mes, valor));
      log(`  ⚠ Sin fila para ${MESES_ES[mes - 1]} en el cuadro de dif de cambio: queda pendiente a mano.`);
    } else if (valor === null) {
      pendiente.push(
        `Falta la cifra de diferencia de cambio de ${MESES_ES[mes - 1]} (fila "${fila.etiqueta}").`
      );
      log(`  ⚠ No se cargó la diferencia de cambio de ${MESES_ES[mes - 1]}: queda pendiente a mano.`);
    }
  }

  // Y la línea del plan que hace cerrar el balance.
  //
  // Donde hay cuadro de meses, la línea lleva el ACUMULADO del año (la suma del cuadro), no el
  // importe del mes: es lo que estaba escrito a mano y lo que el control de la propia planilla
  // verifica. Donde no hay cuadro —el Mensual R$— lleva el importe del mes, que es lo que ese
  // informe muestra.
  const linea = pfUbicarLineaDifCambio(ws, layout0);
  if (linea) {
    const valorMes = ptNumeroTipeado(difCambioMes);
    if (cuadro) {
      const total = pfTotalDelCuadro(ws, cuadro);
      if (total !== null) {
        pfEscribirLineaDifCambio(ws, layout0, linea, total, log);
        hecho.push(`Línea "DIFERENCIA DE CAMBIO" actualizada al acumulado del cuadro: ${total.toFixed(2)}`);
      }
    } else if (valorMes !== null) {
      pfEscribirLineaDifCambio(ws, layout0, linea, valorMes, log);
      hecho.push(`Línea "DIFERENCIA DE CAMBIO": ${valorMes}`);
    } else {
      pendiente.push(
        `Falta la diferencia de cambio del mes. Sin ella la línea "DIFERENCIA DE CAMBIO" ` +
        `(SALDOS fila ${linea.fila}) queda con la del mes pasado y el balance en reales no cierra.`
      );
    }
  }

  return { hecho, pendiente, tieneTc: !!tc, tieneCuadro: !!cuadro, tieneLinea: !!linea };
}

if (typeof module !== "undefined") {
  // En el navegador estos vienen de los <script> que index.html carga antes que este archivo;
  // en Node hay que traerlos a mano, igual que hacen motor_tfbr.js y validar_tfbr.js.
  const fh = require("./formula_hojas.js");
  const cfg = require("./config_tfbr.js");
  global.insertRowEn = fh.insertRowEn;
  global.copiarFormatoDeFila = fh.copiarFormatoDeFila;
  global.derivarLayoutSaldos = cfg.derivarLayoutSaldos;
  global.actualizarFechasDelInforme = require("./fechas_informe.js").actualizarFechasDelInforme;
  global.ptNumeroTipeado = require("./parser_tfbr.js").ptNumeroTipeado;
  module.exports = {
    MESES_ES, ubicarTcCierre, ubicarCuadroDifCambio, escribirDatosDelPeriodo,
    pfUbicarLineaDifCambio, pfTotalDelCuadro,
    pfAvisoFaltaFilaMes, pfColLetra, pfEtiquetaDelMes, pfInsertarFilaDelMes,
  };
}
