// Líneas del Anexo II que leen una cuenta que no es la de su rótulo.
//
// En el Balance Acumulado $, 25 líneas leen la cuenta de UNA FILA MÁS ARRIBA de la que
// corresponde a su rótulo: la línea "Catering" lee la de arriba de CATERING, "Adelanto Viaje"
// lee la de arriba de ADELANTO VIAJE, y así. Cada línea lee a su vecina en cadena, y al final
// de cada cadena queda una cuenta que no lee nadie. Los otros tres balances no lo tienen.
//
// Mientras las cuentas del final de cada cadena estén en cero no se nota, porque el Anexo II
// cierra igual. El mes que alguna tenga saldo, ese importe no entra en el Anexo II y el
// renglón que debería mostrarlo exhibe el de otra cuenta.
//
// Acá están SOLO las 12 que se pueden corregir sin crear un problema nuevo: aquellas cuya
// cuenta de destino no la está leyendo ninguna otra línea que se quede donde está. Las 13
// restantes necesitan una definición de contaduría (a qué línea le corresponde la cuenta que
// hoy comparten) y están en docs/anexo_ii_corrimiento.md.
//
// La cuenta se identifica por CÓDIGO + NOMBRE, nunca por número de fila: las filas cambian
// entre meses y cada vez que se da de alta una cuenta.
const REPUNTES_ANEXO = [
  { archivo: "balance_acumulado_ars", hoja: "Anexo II", celda: "D23", rotulo: "Cesión de Derechos",   cod: "4225000000", nom: "CESIÓN DE DERECHOS" },
  { archivo: "balance_acumulado_ars", hoja: "Anexo II", celda: "D43", rotulo: "Gastos Hotelería",     cod: "4218000000", nom: "GASTOS HOTELERÍA" },
  { archivo: "balance_acumulado_ars", hoja: "Anexo II", celda: "D44", rotulo: "Gastos Gestoría",      cod: "4219000000", nom: "GASTOS GESTORÍA" },
  { archivo: "balance_acumulado_ars", hoja: "Anexo II", celda: "D46", rotulo: "Habilitaciones",       cod: "4216000000", nom: "HABILITACIONES" },
  { archivo: "balance_acumulado_ars", hoja: "Anexo II", celda: "D79", rotulo: "Siniestros",           cod: "4222300000", nom: "SINIESTROS" },
  { archivo: "balance_acumulado_ars", hoja: "Anexo II", celda: "E12", rotulo: "Accidentes",           cod: "4222200000", nom: "ACCIDENTES" },
  { archivo: "balance_acumulado_ars", hoja: "Anexo II", celda: "E14", rotulo: "Adelanto Viaje",       cod: "4226000000", nom: "ADELANTO VIAJE" },
  { archivo: "balance_acumulado_ars", hoja: "Anexo II", celda: "E22", rotulo: "Catering",             cod: "4211900000", nom: "CATERING" },
  { archivo: "balance_acumulado_ars", hoja: "Anexo II", celda: "E28", rotulo: "Deudores incobrables", cod: "4223600000", nom: "DEUDORES INCOBRABLES" },
  // Fila 79 tiene las dos columnas de centro de costo cargadas. D79 va a SINIESTROS; a E79 le
  // corresponde SINIESTROS CONVENIO HSBC, que es la que queda sin lector al correr la cadena.
  { archivo: "balance_acumulado_ars", hoja: "Anexo II", celda: "E79", rotulo: "Siniestros",           cod: "4230600000", nom: "SINIESTROS CONVENIO HSBC" },
  { archivo: "balance_acumulado_ars", hoja: "Anexo II", celda: "F52", rotulo: "Impuestos varios",     cod: "4230800000", nom: "IMPUESTOS VARIOS" },
  { archivo: "balance_acumulado_ars", hoja: "Anexo II", celda: "F56", rotulo: "imp. Al cheque",       cod: "4230900000", nom: "IMP. AL CHEQUE" },
];

const RA_RE_REF = /SALDOS!\$?[A-Z]{1,3}\$?(\d+)/g;

function raNormaliza(t) {
  return String(t || "").normalize("NFD").replace(/[̀-ͯ]/g, "")
    .toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();
}

// Todas las filas de SALDOS que lee cada celda del Anexo II, para poder comprobar antes de
// escribir que la cuenta de destino no la esté levantando ya otro renglón.
function raCenso(wb, hoja) {
  const ws = wb.getWorksheet(hoja);
  const porFila = {};
  if (!ws) return porFila;
  ws.eachRow((row, r) => row.eachCell((cell, c) => {
    const v = cell.value;
    if (!v || typeof v !== "object" || typeof v.formula !== "string") return;
    if (/^SUM\(/i.test(v.formula)) return;
    const vistas = new Set();
    let m;
    RA_RE_REF.lastIndex = 0;
    while ((m = RA_RE_REF.exec(v.formula))) vistas.add(Number(m[1]));
    vistas.forEach(f => (porFila[f] = porFila[f] || []).push(cell.address));
  }));
  return porFila;
}

// Corrige las líneas de este archivo. Idempotente: si ya apuntan bien, no hay nada que hacer.
function aplicarRepuntesAnexo(wb, archivoId, planDeCuentas, log = () => {}) {
  const hechos = [];
  let salteados = [];
  const mios = REPUNTES_ANEXO.filter(r => r.archivo === archivoId);
  if (!mios.length) return { hechos, salteados };

  const censo = raCenso(wb, mios[0].hoja);

  // Las correcciones se encadenan: la cuenta que le toca a una línea suele estar ocupada por
  // otra que también se mueve, y hasta que esa no se corre la primera no puede entrar. Se dan
  // pasadas hasta que una no cambie nada, en vez de depender del orden de la lista.
  let pendientes = mios.slice();
  for (let pasada = 0; pasada < mios.length + 1 && pendientes.length; pasada++) {
    const antes = hechos.length;
    salteados = [];
    pendientes = raUnaPasada(wb, pendientes, planDeCuentas, censo, hechos, salteados, log);
    if (hechos.length === antes) break;   // ninguna avanzó: lo que queda no se destraba solo
  }
  return { hechos, salteados };
}

// Una pasada sobre las que faltan. Devuelve las que no se pudieron hacer todavía.
function raUnaPasada(wb, lista, planDeCuentas, censo, hechos, salteados, log) {
  const faltan = [];
  for (const r of lista) {
    const ws = wb.getWorksheet(r.hoja);
    if (!ws) { salteados.push({ ...r, motivo: `el archivo no tiene la hoja '${r.hoja}'` }); continue; }

    // la fila donde vive hoy la cuenta, por código + nombre
    const info = planDeCuentas[r.cod];
    const candidatas = info ? [info].concat(info.otrasFilas || []) : [];
    const destino = candidatas.find(f => raNormaliza(f.texto).includes(raNormaliza(r.nom)));
    if (!destino) {
      salteados.push({ ...r, motivo: `no encontré la cuenta ${r.cod} "${r.nom}" en SALDOS` });
      log(`  ⚠ ${r.celda}: no encontré ${r.cod} "${r.nom}", lo dejo como está.`);
      continue;
    }

    const celda = ws.getCell(r.celda);
    const v = celda.value;
    if (!v || typeof v !== "object" || typeof v.formula !== "string") {
      salteados.push({ ...r, motivo: `${r.celda} ya no tiene una fórmula` });
      continue;
    }
    const refs = [];
    let m; RA_RE_REF.lastIndex = 0;
    while ((m = RA_RE_REF.exec(v.formula))) refs.push(Number(m[1]));
    // Solo se tocan las que leen UNA sola cuenta: si alguien le agregó otra, la línea ya no es
    // la que se analizó y hay que volver a mirarla, no reescribirla a ciegas.
    if (refs.length !== 1) {
      salteados.push({ ...r, motivo: `${r.celda} lee ${refs.length} cuentas, no una` });
      log(`  ⚠ ${r.celda} ("${r.rotulo}") lee ${refs.length} cuentas: no la toco.`);
      continue;
    }
    if (refs[0] === destino.fila) continue;   // ya estaba corregida (idempotente)

    // Nadie más puede estar leyendo la cuenta de destino, o el importe se contaría dos veces.
    const otros = (censo[destino.fila] || []).filter(a => a !== r.celda);
    if (otros.length) {
      // puede ser que la otra línea también se mueva: se reintenta en la pasada siguiente
      salteados.push({ ...r, motivo: `${r.nom} ya la lee ${otros.join(", ")}` });
      faltan.push(r);
      continue;
    }

    celda.value = { formula: v.formula.replace(
      new RegExp(`(SALDOS!\\$?[A-Z]{1,3}\\$?)${refs[0]}(?!\\d)`), `$1${destino.fila}`) };
    censo[destino.fila] = [r.celda];
    censo[refs[0]] = (censo[refs[0]] || []).filter(a => a !== r.celda);
    hechos.push({ ...r, filaVieja: refs[0], filaNueva: destino.fila });
    log(`  ${r.hoja}!${r.celda} ("${r.rotulo}") ahora lee ${r.cod} ${r.nom} (fila ${destino.fila}, antes ${refs[0]}).`);
  }
  return faltan;
}

if (typeof module !== "undefined") {
  module.exports = { REPUNTES_ANEXO, aplicarRepuntesAnexo, raCenso, raNormaliza };
}
