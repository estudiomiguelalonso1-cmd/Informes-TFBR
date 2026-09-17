// Motor de informe-tfbr: escribe SOLO la zona de pegado dentro de SALDOS (ver
// config_tfbr.js) y aplica los fixes aprobados. Todo lo demás (EESP, EERR, EEPN, Activo,
// Pasivo, Anexo I, Anexo II) se resuelve con las fórmulas que el maestro ya tiene, cuando
// se abre en Excel real — el motor nunca intenta recalcular nada (ver docs/formula_analysis.md
// y el plan de arquitectura: TFBR sigue el patrón de round-trip por Excel real, no el de
// calcular en JS).

// derivarLayoutSaldos / leerPlanDeCuentas / ctColNumeroALetra vienen de config_tfbr.js,
// cargado antes que este archivo (ver index.html) — se usan acá como globales, igual que
// motor_balances.js usa insertRowEn de formula_hojas.js. El require() de más abajo es solo
// para que los tests de Node (que no cargan <script> en orden) tengan lo mismo disponible.

// Empareja las cuentas del export contra el plan de cuentas del maestro POR CÓDIGO, nunca
// por el texto tal cual lo trae Onvio ("1120100100 - FORD", con guión) porque el VLOOKUP
// del maestro busca el texto EXACTO que la plantilla ya tiene ("1120100100  FORD", sin
// guión, doble espacio). Ver docs/formula_analysis.md, sección de riesgo de matching.
function emparejarConPlan(cuentasExport, planDeCuentas, campoSaldo) {
  const matcheadas = []; // [{texto, saldo}] listas para escribir en el staging
  const sinMapear = [];  // cuentas del export sin fila correspondiente en el maestro
  const escritas = {};   // codigo -> saldo escrito, para el control cuenta por cuenta
  for (const c of cuentasExport) {
    const enPlan = planDeCuentas[c.codigo];
    if (!enPlan) { sinMapear.push(c); continue; }
    matcheadas.push({ codigo: c.codigo, texto: enPlan.texto, saldo: c[campoSaldo] });
    escritas[c.codigo] = c[campoSaldo];
  }
  return { matcheadas, sinMapear, escritas };
}

// Limpia el rango de staging completo (evita que queden filas de un mes anterior con más
// cuentas que el actual) y escribe las cuentas emparejadas, una por fila, empezando en la
// esquina superior izquierda del rango.
function escribirStaging(wb, layout, matcheadas, log = () => {}) {
  const ws = wb.getWorksheet(layout.sheet);
  const { colDesde, filaDesde, colHasta, filaHasta } = layout.stagingRange;
  const capacidad = filaHasta - filaDesde + 1;

  if (matcheadas.length > capacidad) {
    throw new Error(
      `El rango de staging (${ctColNumeroALetra(colDesde)}${filaDesde}:${ctColNumeroALetra(colHasta)}${filaHasta}, ` +
      `${capacidad} filas) no alcanza para las ${matcheadas.length} cuentas emparejadas. ` +
      "NO se escribió nada. Hay que ampliar el rango en la plantilla antes de seguir."
    );
  }

  for (let r = filaDesde; r <= filaHasta; r++) {
    ws.getCell(r, colDesde).value = null;
    ws.getCell(r, colDesde + 1).value = null;
  }
  // Se escribe el NÚMERO DE CUENTA, no el texto: es lo que busca la fórmula de cada fila
  // (ver migrarClaveANumero). Con el texto completo, renombrar una cuenta dejaba la fila en
  // cero hasta que el motor volviera a escribir el staging.
  matcheadas.forEach((m, i) => {
    ws.getCell(filaDesde + i, colDesde).value = m.codigo;
    ws.getCell(filaDesde + i, colDesde + 1).value = m.saldo;
  });

  log(`  SALDOS: ${matcheadas.length} cuenta(s) escritas en el staging ` +
      `(${ctColNumeroALetra(colDesde)}${filaDesde}:${ctColNumeroALetra(colHasta)}${filaHasta}).`);
}

// Cuántas filas de colchón se dejan al ampliar la zona de pegado. El plan de cuentas crece de
// a poco —entre julio y agosto de 2026 el acumulado pasó de 87 a 91 cuentas— y ampliar de a una
// fila por mes significa insertar filas en cada cierre. Con margen se hace una vez y aguanta.
const MOTOR_MARGEN_STAGING = 15;

// La capacidad que tienen que tener los cuatro archivos, tengan las cuentas que tengan hoy.
//
// Venían con capacidades distintas y por casualidad: 118, 58, 123 y 106 filas. Eso significa que
// el mismo plan de cuentas entra en unos y en otros no, y que el primero en quedarse corto es
// siempre el mismo sin que haya una razón. Igualarlos hace que los cuatro aguanten lo mismo y
// que, cuando uno tenga que crecer, crezcan todos a la vez y por el mismo motivo.
//
// 150 es cómodo sobre las 91 cuentas que trae hoy el acumulado, y a razón de las cuatro cuentas
// nuevas que aparecieron entre julio y agosto de 2026 da para varios años. Si algún mes hiciera
// falta más, se amplía sola igual: esto es un piso, no un techo.
const MOTOR_CAPACIDAD_MINIMA = 150;

// La zona de pegado se quedó corta: se le agregan filas.
//
// El Acumulado R$ tenía 88 filas y agosto de 2026 trajo 91 cuentas. Sin esto el motor frena el
// cierre entero con "hay que ampliar el rango en la plantilla", que es correcto pero deja el
// trabajo a mano justo en el paso que el sistema existe para automatizar.
//
// Las filas se insertan EN LA ÚLTIMA DEL RANGO, no debajo: ahí todavía están dentro, así que el
// rango del BUSCARV se estira solo y lo que haya más abajo se corre sin romperse. Insertar
// debajo dejaría las filas nuevas fuera del rango y las cuentas que cayeran ahí no se
// levantarían — el balance cerraría igual y el importe no estaría.
function ampliarZonaDePegado(wb, layout, necesarias, log = () => {}) {
  const sr = layout.stagingRange;
  const capacidad = sr.filaHasta - sr.filaDesde + 1;
  // El piso es el mismo para los cuatro; si un mes trajera más cuentas que eso, manda lo que
  // haga falta más el colchón.
  const objetivo = Math.max(MOTOR_CAPACIDAD_MINIMA, necesarias + MOTOR_MARGEN_STAGING);
  if (capacidad >= objetivo) return { ampliada: false, capacidad };

  const faltan = objetivo - capacidad;
  const modificadas = [];
  for (let i = 0; i < faltan; i++) {
    modificadas.push(insertRowEn(wb, layout.sheet, layout.stagingRange.filaHasta));
    layout = derivarLayoutSaldos(wb);
  }
  const nueva = layout.stagingRange;
  log(`  Zona de pegado: tenía ${capacidad} filas, este mes hacían falta ${necesarias} y el ` +
      `mínimo de los cuatro archivos es ${MOTOR_CAPACIDAD_MINIMA}. Se agregaron ${faltan}: ` +
      `ahora va de ${nueva.filaDesde} a ${nueva.filaHasta} (${objetivo} filas).`);
  return { ampliada: true, faltan, capacidad, objetivo, layout };
}

// Corre el proceso completo para UN archivo/moneda. No guarda el archivo (eso lo decide
// quien llama, según si va a pedir más pasos antes de bajar el .xlsx).
function procesarMaestroTFBR({ wb, cuentasExport, campoSaldo, archivoId = null, altaAutomatica = true,
                              rotulosGuardados = null, cuentasAcumulado = null,
                              configuracion = null, log = () => {} }) {
  // Las cuentas sin saldo se sacan ACA, antes que nada.
  //
  // El export de Onvio hay que pedirlo con las cuentas saldadas incluidas —hay cuentas de
  // resultado que cierran en cero en pesos y no en reales, porque los movimientos del anio se
  // convirtieron a tipos de cambio distintos y no se cancelan— pero eso trae las 940 cuentas
  // del plan oficial. De esas, unas 700 no estan en el plan del maestro, y el alta automatica
  // de mas abajo les insertaba una fila a cada una: cada insercion recorre las formulas de
  // todas las hojas para reacomodarlas, asi que la pagina se colgaba.
  //
  // Un renglon en cero no aporta nada: el VLOOKUP de una cuenta que no se pego ya devuelve
  // cero. Se filtra por el saldo de ESTE archivo, que es lo correcto en las dos monedas —
  // "4212100000 FLETES" tiene $ 0,00 y R$ 3.356,67: no entra en los informes en pesos y si en
  // los de reales.
  const cuentasDelExport = cuentasExport;
  cuentasExport = cuentasExport.filter(c => c[campoSaldo]);
  const sinSaldo = cuentasDelExport.length - cuentasExport.length;
  if (sinSaldo) {
    log(`  ${sinSaldo} cuenta(s) del export vienen sin saldo en este archivo y no se pegan ` +
        `(de ${cuentasDelExport.length}, quedan ${cuentasExport.length}).`);
  }


  // Primero de todo: poner el plan de cuentas de SALDOS de acuerdo con el plan oficial. Corrige
  // los códigos tipeados con un dígito de menos y junta las filas que quedan repetidas. Va
  // antes que cualquier otra cosa porque cambia a qué fila resuelve cada código, que es lo que
  // usan el alta de cuentas nuevas y el emparejamiento de más abajo.
  const limpieza = limpiarPlanDeCuentas(wb, log);
  if (limpieza.corregidos.length || limpieza.reasignados.length || limpieza.fusionadas.length) {
    log(`  Plan de cuentas: ${limpieza.corregidos.length} código(s) corregidos, ` +
        `${limpieza.reasignados.length} reasignado(s), ${limpieza.fusionadas.length} fila(s) fusionadas.`);
  }

  let layout = derivarLayoutSaldos(wb);
  let { cuentas: planDeCuentas, duplicadas } = leerPlanDeCuentas(wb, layout);

  // Cuentas que vienen en el export y no estan en el plan. Es un caso normal: entre junio y
  // julio 2026 aparecieron 6 en el Mensual $, por 11,66 millones de pesos. Si no se dan de
  // alta, ese importe no entra en ningun total y el balance cierra igual, sin avisar.
  const altas = [];
  if (altaAutomatica) {
    const faltantes = cuentasExport.filter(c => !planDeCuentas[c.codigo]);
    for (const c of faltantes) {
      // Si la cuenta ya esta con el codigo mal (un digito menos), las lineas de los estados
      // apuntan a ESA fila, que nunca levanta importe. Hay que moverlas a la fila nueva o el
      // importe entra en SALDOS pero no llega al Anexo II ni al EERR.
      const gemela = buscarGemela(planDeCuentas, c.codigo, c.nombre);
      const { fila, clave } = insertarCuentaEnSaldos(wb, layout, planDeCuentas, c, log);
      let repuntadas = [];
      if (gemela) {
        const filaGemela = planDeCuentas[gemela.codigo].fila;
        repuntadas = repuntarGemela(wb, layout.sheet, filaGemela, fila, log);
      }
      altas.push({
        codigo: c.codigo, nombre: c.nombre, fila, clave,
        gemela: gemela ? gemela.codigo : null, repuntadas,
        saldo: c[campoSaldo],
      });
    }
    if (faltantes.length) {
      // Las filas se movieron: el layout y el plan que teniamos quedaron viejos.
      layout = derivarLayoutSaldos(wb);
      ({ cuentas: planDeCuentas, duplicadas } = leerPlanDeCuentas(wb, layout));
    }
  }

  // Líneas del Anexo II que leen la cuenta de al lado en vez de la suya. Va después de las
  // decisiones de duplicados porque esas mueven filas, y el repunte apunta a la fila donde la
  // cuenta quedó, no donde estaba.
  const repuntes = aplicarRepuntesAnexo(wb, archivoId, planDeCuentas, log);

  // Rótulos que este archivo no tiene y los otros sí. Va después del repunte: el repunte libera
  // la cuenta del renglón ajeno, y esto le da el suyo.
  const rotulos = agregarRotulosAnexo(wb, archivoId, layout, planDeCuentas, log);

  // Los cuatro archivos con los mismos rótulos, tomando el Mensual $ como modelo, y cada
  // cuenta leída por un solo renglón.
  const unificacion = unificarRotulosAnexo(wb, planDeCuentas, log);

  // Cuentas de gasto que ningún renglón del Anexo II lee. Primero se enganchan solas las que
  // ya tienen decisión tomada en un mes anterior; las que quedan se devuelven para que la
  // pantalla pregunte a qué rótulo van. Va acá, después de la unificación, porque hasta este
  // punto los renglones todavía se están moviendo.
  const rotulosAuto = aplicarRotulosGuardados(wb, rotulosGuardados, log);
  const decididas = rotulosGuardados || {};
  const sinRotulo = cuentasSinRotulo(wb).filter(c => !(c.cod in decididas));
  for (const c of sinRotulo) {
    log(`  ⚠ ${c.cod} ${c.nom} está en SALDOS pero ningún renglón del Anexo II la lee: ` +
        `hay que decir a qué rótulo va, o su importe no llega al estado de resultados.`);
  }

  const { matcheadas, sinMapear, escritas } = emparejarConPlan(cuentasExport, planDeCuentas, campoSaldo);

  // Antes de escribir: que haya lugar. El plan de cuentas crece y la zona de pegado no.
  const ampliacion = ampliarZonaDePegado(wb, layout, matcheadas.length, log);
  if (ampliacion.ampliada) layout = ampliacion.layout;

  escribirStaging(wb, layout, matcheadas, log);

  // Los rangos del plan pasan a nombrar sus cuentas una por una. Va DESPUÉS del alta: las
  // cuentas que entran este mes todavía las levanta el rango como siempre, y recién después
  // quedan fijadas. Al revés, una cuenta nueva que hoy el rango levanta se perdería sin avisar.
  const rastreo = expandirRangosSaldos(wb, layout, log);

  // Anexo I: el valor de origen de los bienes de uso sale de las cuentas, no de un número
  // tipeado. Va después de escribir el staging porque necesita saber qué se pegó: con las
  // cuentas de bienes de uso adentro arma la fórmula, y si no vinieron escribe el importe.
  // Renglones que se llaman de una cuenta y suman otra. Va después de expandir los rangos,
  // que es cuando cada renglón nombra de verdad las cuentas que lo alimentan.
  const renombres = aplicarRenombresRenglon(wb, layout, log);

  // Los préstamos al personal van juntos en el Activo, en los cuatro archivos, y fuera del
  // Pasivo — son cuentas 114, de activo, y en el Pasivo entraban restando.
  const prestamos = consolidarPrestamos(wb, layout, log);

  // Los renglones que le faltan a este archivo y alguna cuenta necesita.
  const renglonesNuevos = agregarRenglonesFaltantes(wb, layout, archivoId, log);

  // Cuentas que van a un renglón decidido con contaduría, esté donde estén hoy. Va DESPUÉS de
  // consolidar los préstamos: en dos archivos el renglón "- Adelanto al personal" no existía
  // con ese nombre y lo crea ese paso, así que corriendo antes no había dónde mandarlas.
  const asignaciones = aplicarAsignaciones(wb, layout, log);

  // Y encima de todo eso, lo que la persona configuró desde el panel. Va último de los pasos
  // de cableado, a propósito: una decisión tomada a mano manda sobre cualquier regla del
  // código. Ese es el punto de que el sistema sea configurable sin tocar el programa.
  const configurado = aplicarConfiguracion(wb, layout, configuracion, log);
  layout = configurado.layout || layout;

  const anexoI = completarAnexoI(wb, layout,
    { escritas, cuentasAcumulado: cuentasAcumulado || cuentasExport, campoSaldo }, log);

  // Y con todo ya cableado, el control que faltaba: plata de este mes que no llega a ninguna
  // hoja. Entra en los totales de SALDOS —el balance cierra igual— pero no está en ningún
  // estado, y mirando el resultado no hay forma de darse cuenta.
  // La configuración pudo insertar filas: el plan hay que releerlo antes de los controles.
  ({ cuentas: planDeCuentas, duplicadas } = leerPlanDeCuentas(wb, layout));

  // Lo último: que no quede oculto ningún renglón con importe. Va al final porque insertar
  // filas reacomoda las marcas de oculto.
  const mostrados = mostrarRenglonesConImporte(wb, layout, planDeCuentas, escritas, log);

  // Y que todo el bloque se vea parejo: las plantillas traen filas sin formato que, apenas
  // reciben un importe y se muestran, salen con otra letra y el número sin separador de miles.
  const formato = uniformarFormatoDeBloques(wb, layout, log);

  // Las que la persona marcó como que no van a ningún renglón no se denuncian: quedaron así
  // a propósito, y avisar todos los meses enseña a ignorar el aviso.
  const excluidas = cuentasExcluidas(configuracion);
  const sinDestino = cuentasSinDestino(wb, layout, planDeCuentas, escritas)
    .filter(c => !excluidas.has(c.cod));
  for (const c of sinDestino) {
    log(`  ⚠ ${c.cod} ${c.nom} tiene ${c.importe.toFixed(2)} y ninguna hoja la lee: ` +
        `su importe no llega a ningún estado.`);
  }

  // Una cuenta de resultado recien dada de alta que no quedo referenciada por ninguna linea
  // entra en los totales de SALDOS pero no en el Anexo II, y entonces el EERR no la cuenta.
  // No se inventa a que concepto va: se avisa.
  const sinEnganchar = altas.filter(a =>
    String(a.codigo)[0] === "4" && !a.repuntadas.length);
  for (const a of sinEnganchar) {
    log(`  ⚠ ${a.clave} se dio de alta en SALDOS pero no quedo enganchada a ninguna linea del ` +
        `Anexo II: hay que agregarla a mano, o su importe no llega al estado de resultados.`);
  }

  // Los códigos repetidos se clasifican en vez de avisarlos en bloque: no son todos el mismo
  // problema (ver duplicados_tfbr.js) y mezclarlos hace que el aviso no sirva para actuar.
  const casosDuplicados = duplicadas.length
    ? clasificarDuplicados(wb, layout, duplicadas)
    : [];
  for (const c of casosDuplicados) {
    log(`  ⚠ Código repetido ${c.codigo} (filas ${c.filas.map(f => f.fila).join(" y ")}): ${c.motivo}`);
  }

  if (sinMapear.length) {
    log(`  ⚠ ${sinMapear.length} cuenta(s) del export NO están en el plan de cuentas de este ` +
        `maestro y quedaron SIN incluir en ningún total: ` +
        sinMapear.map(c => `${c.codigo} (${c.nombre})`).join(", ") + ".");
  }

  const totalEscrito = matcheadas.reduce((s, m) => s + m.saldo, 0);
  const totalExport = cuentasExport.reduce((s, c) => s + c[campoSaldo], 0);
  log(`  Total escrito: ${totalEscrito.toFixed(2)} | Total del export: ${totalExport.toFixed(2)}` +
      (Math.abs(totalEscrito - totalExport) > 0.02
        ? " ⚠ NO COINCIDEN (revisar cuentas sin mapear arriba)."
        : " (coinciden)."));

  return {
    layout,
    planDeCuentas,
    escritas,
    resumen: {
      cuentasExport: cuentasExport.length,
      cuentasEscritas: matcheadas.length,
      sinMapear: sinMapear.map(c => ({ codigo: c.codigo, nombre: c.nombre, saldo: c[campoSaldo] })),
      duplicadas: casosDuplicados,
      limpiezaPlan: {
        corregidos: limpieza.corregidos, reasignados: limpieza.reasignados,
        fusionadas: limpieza.fusionadas, trabadas: limpieza.trabadas,
        sinExplicar: limpieza.sinExplicar, ambiguos: limpieza.ambiguos,
      },
      repuntesAnexo: repuntes.hechos,
      rotulosAgregados: rotulos.agregados,
      rotulosUnificados: unificacion,
      rotulosSalteados: rotulos.salteados,
      repuntesSalteados: repuntes.salteados,
      altas,
      sinEnganchar: sinEnganchar.map(a => ({ codigo: a.codigo, clave: a.clave })),
      rotulosAuto,
      sinRotulo,
      anexoI,
      zonaDePegado: ampliacion,
      renglonesRenombrados: renombres,
      asignaciones,
      configurado,
      renglonesNuevos,
      renglonesMostrados: mostrados,
      celdasEmparejadas: formato.length,
      prestamos,
      rangosExpandidos: rastreo.expandidos,
      rangosSalteados: rastreo.salteados,
      sinDestino,
      totalEscrito,
      totalExport,
    },
  };
}

if (typeof module !== "undefined") {
  const cfg = require("./config_tfbr.js");
  global.derivarLayoutSaldos = cfg.derivarLayoutSaldos;
  global.leerPlanDeCuentas = cfg.leerPlanDeCuentas;
  global.ctColNumeroALetra = cfg.ctColNumeroALetra;
  global.clasificarDuplicados = require("./duplicados_tfbr.js").clasificarDuplicados;
  const ic = require("./insertar_cuenta.js");
  global.insertarCuentaEnSaldos = ic.insertarCuentaEnSaldos;
  global.buscarGemela = ic.buscarGemela;
  global.repuntarGemela = ic.repuntarGemela;
  global.aplicarRepuntesAnexo = require("./repuntes_anexo.js").aplicarRepuntesAnexo;
  global.limpiarPlanDeCuentas = require("./limpieza_plan.js").limpiarPlanDeCuentas;
  global.insertRowEn = require("./formula_hojas.js").insertRowEn;
  global.agregarRotulosAnexo = require("./rotulos_anexo.js").agregarRotulosAnexo;
  global.unificarRotulosAnexo = require("./rotulos_unificados.js").unificarRotulosAnexo;
  const cn = require("./cuentas_nuevas.js");
  global.aplicarRotulosGuardados = cn.aplicarRotulosGuardados;
  global.cuentasSinRotulo = cn.cuentasSinRotulo;
  const rr = require("./rastreo_rangos.js");
  global.expandirRangosSaldos = rr.expandirRangosSaldos;
  global.completarAnexoI = require("./anexo_i.js").completarAnexoI;
  global.aplicarRenombresRenglon = require("./config_hojas.js").aplicarRenombresRenglon;
  global.aplicarAsignaciones = require("./config_hojas.js").aplicarAsignaciones;
  const cg = require("./config_guardada.js");
  global.aplicarConfiguracion = cg.aplicarConfiguracion;
  global.cuentasExcluidas = cg.cuentasExcluidas;
  global.agregarRenglonesFaltantes = require("./renglones_faltantes.js").agregarRenglonesFaltantes;
  global.mostrarRenglonesConImporte = require("./renglones_faltantes.js").mostrarRenglonesConImporte;
  global.uniformarFormatoDeBloques = require("./renglones_faltantes.js").uniformarFormatoDeBloques;
  global.consolidarPrestamos = require("./prestamos.js").consolidarPrestamos;
  global.cuentasSinDestino = rr.cuentasSinDestino;
  module.exports = { emparejarConPlan, escribirStaging, procesarMaestroTFBR };
}
