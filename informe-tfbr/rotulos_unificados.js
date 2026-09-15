// Los cuatro Anexo II con los mismos rótulos, tomando el MENSUAL $ como modelo.
//
// CUENTA_DE_ROTULO no está escrita a mano: se derivó del cableado real del Mensual $, que es
// el archivo que contaduría mantiene y el más completo de los cuatro. De ahí salieron 79
// cuentas con un rótulo único, que se copian tal cual.
//
// Las 9 que el modelo leía desde DOS renglones (doble conteo dentro del propio modelo) se
// resolvieron una por una: queda el rótulo que coincide con el nombre de la cuenta. Los dos
// renglones de impuestos a los débitos y a los créditos se unificaron en uno solo, porque las
// dos cuentas resultaron ser la misma (ver UNIFICACIONES en limpieza_plan.js).
//
// La regla que sostiene todo: cada cuenta la lee UN SOLO renglón. Al engancharla en el suyo se
// la saca de cualquier otro, o el importe se cuenta dos veces — que es lo que venía pasando con
// GASTOS DE LOGÍSTICA (tres renglones del Mensual $) y con TASAS AFIP (dos).
//
// Ningún rótulo se borra. Los que quedan sin cuenta quedan disponibles para asignarles una.
//
// Estas cuatro cuentas de gasto quedan A PROPÓSITO sin rótulo, como en el modelo, para que se
// les asigne uno desde la configuración cuando haga falta:
//   4211200000 GASTOS EN EQ. TELEFÓNICOS   4211100000 REDONDEO
//   4213100000 GRATIFICACIONES             4220800000 LAVADO DE FLOTA

const RENOMBRAR_ROTULOS = [
  { a: "Acuerdo Seclo",                             de: ["Acuerdo SECLO"] },
  { a: "Gastos médicos",                            de: ["Gastos Medicos", "Gastos Médicos"] },
  { a: "Gastos de representación",                  de: ["Gastos de Representación"] },
  { a: "Gastos obra social",                        de: ["Gastos Obra Social"] },
  { a: "Gastos gestoría",                           de: ["Gastos Gestoria", "Gastos Gestoría"] },
  { a: "Honorarios directores",                     de: ["Honorarios director"] },
  { a: "Honorarios legales",                        de: ["Honorarios Legales"] },
  { a: "Gastos de logística",                       de: ["Gastos de logistica"] },
  { a: "Servicios de limpieza",                     de: ["Servicio de limpieza"] },
  { a: "Deudores incobrables",                      de: ["Deudores Incobrables"] },
  { a: "Tasas Senasa",                              de: ["Tasas SENASA"] },
  { a: "Servicio de transporte y almacenamiento",   de: ["Serv de transporte y almacenamiento", "Serv de Transporte y Almacenamiento"] },
  { a: "Salario Complementario del Decreto 332/20", de: ["Salario Complementario del decreto 332/20"] },
  { a: "Ropa de trabajo",                           de: ["Gs Ropa de Trabajo"] },
  { a: "Embargo radicación vehicular",              de: ["Embargo Rdicación Vehicular"] },
  { a: "Impuestos al cheque",                       de: ["Impuestos al Cheuque"] },
  { a: "Reparaciones de flota",                     de: ["Reparaciones flota", "Repuestos y reparaciones"] },
  // Las dos cuentas resultaron ser una sola, así que los dos renglones se unifican.
  { a: "IMP. A LOS DÉBITOS Y CRÉDITOS",             de: ["Imp. A los créditos", "Imp. A los débitos", "Imp. A los débitos y créditos"] },
  // Los Acumulados las tenían juntas; el modelo las tiene separadas.
  { a: "Tasas AFIP",                                de: ["Tasas IGJ, AFIP"] },

  // Sinónimos entre archivos. En cada par queda el nombre del Mensual $, que es el modelo.
  { a: "Mensajería",                                de: ["Mensajería y trámites"] },
  { a: "Seguridad",                                 de: ["Seguridad y vigilancia"] },
  { a: "Alquiler cochera",                          de: ["Alquiler de cochera"] },
  { a: "Gastos Capacitación",                       de: ["Capacitación"] },
  { a: "Gastos PC",                                 de: ["Gastos Computación"] },
  { a: "Impuestos al cheque",                       de: ["imp. Al cheque"] },
  { a: "Mantenimiento de flota",                    de: ["Mantenimiento y lav. de flota"] },
  { a: "Gastos de logística",                       de: ["Logística"] },
  { a: "Telefonos",                                 de: ["Gastos Telefonico", "Gastos Telefonía"] },
  { a: "Bs Personales Acc. Y Participacion",        de: ["Imp. Bienes Acc. Y Part. Soc.", "Imp. Bienes Sociedades"] },
];

// cuenta -> rótulo + centro de costo (D administración, E comercialización, F financieros).
// Derivada del Mensual $. Vale para los cuatro: la cuenta es la misma en todos.
const CUENTA_DE_ROTULO = [
  { cod: "4210100000", cc: "E", rotulo: "Gastos Generales"                          },
  { cod: "4210200000", cc: "D", rotulo: "Molilidad y viáticos"                      },
  { cod: "4210300000", cc: "E", rotulo: "Sueldos"                                   },
  { cod: "4210400000", cc: "E", rotulo: "Honorarios profesionales"                  },
  { cod: "4210500000", cc: "D", rotulo: "Librería"                                  },
  { cod: "4210600000", cc: "E", rotulo: "Cargas Sociales"                           },
  { cod: "4210700000", cc: "D", rotulo: "Mantenimiento de maquinas"                 },
  { cod: "4210800000", cc: "E", rotulo: "Cargas Sociales"                           },
  { cod: "4210900000", cc: "D", rotulo: "Gastos PC"                                 },
  { cod: "4211300000", cc: "D", rotulo: "Gastos obra social"                        },
  { cod: "4211400000", cc: "D", rotulo: "Gastos trámites"                           },
  { cod: "4211500000", cc: "E", rotulo: "Honorarios directores"                     },
  { cod: "4211600000", cc: "D", rotulo: "Servicios de limpieza"                     },
  { cod: "4211700000", cc: "D", rotulo: "Gastos Capacitación"                       },
  { cod: "4211800000", cc: "E", rotulo: "Reparaciones"                              },
  { cod: "4212000000", cc: "E", rotulo: "Mantenimiento"                             },
  { cod: "4212100000", cc: "E", rotulo: "Fletes"                                    },
  { cod: "4212200000", cc: "D", rotulo: "Legalizaciones"                            },
  { cod: "4212300000", cc: "E", rotulo: "Gastos varios personal"                    },
  { cod: "4212400000", cc: "D", rotulo: "Autónomos Directores"                      },
  { cod: "4212500000", cc: "E", rotulo: "Amortizaciones"                            },
  { cod: "4212600000", cc: "D", rotulo: "Agua"                                      },
  { cod: "4212700000", cc: "D", rotulo: "Alquiler fotocopiadora"                    },
  { cod: "4212800000", cc: "D", rotulo: "Tasas AFIP"                                },
  { cod: "4212900000", cc: "D", rotulo: "Tasas IGJ"                                 },
  { cod: "4213000000", cc: "D", rotulo: "Alquiler Oficina Móvil"                    },
  { cod: "4213200000", cc: "E", rotulo: "Acuerdo Seclo"                             },
  { cod: "4213300000", cc: "E", rotulo: "Salario Complementario del Decreto 332/20" },
  { cod: "4214000000", cc: "D", rotulo: "Gastos médicos"                            },
  { cod: "4215000000", cc: "E", rotulo: "Honorarios legales"                        },
  { cod: "4216000000", cc: "D", rotulo: "Habilitaciones"                            },
  { cod: "4217000000", cc: "D", rotulo: "Sellados"                                  },
  { cod: "4218000000", cc: "D", rotulo: "Gastos Hotelería"                          },
  { cod: "4219000000", cc: "D", rotulo: "Gastos gestoría"                           },
  { cod: "4220100000", cc: "E", rotulo: "Seguros"                                   },
  { cod: "4220200000", cc: "E", rotulo: "Comunicaciones"                            },
  { cod: "4220300000", cc: "E", rotulo: "Locomoción"                                },
  { cod: "4220400000", cc: "D", rotulo: "Custodia"                                  },
  { cod: "4220500000", cc: "E", rotulo: "Despachantes"                              },
  { cod: "4220600000", cc: "E", rotulo: "Mantenimiento de flota"                    },
  { cod: "4220700000", cc: "E", rotulo: "Combustible"                               },
  { cod: "4220900000", cc: "D", rotulo: "Mensajería"                                },
  { cod: "4221000000", cc: "E", rotulo: "Peajes"                                    },
  { cod: "4221100000", cc: "E", rotulo: "Estacionamiento"                           },
  { cod: "4221200000", cc: "E", rotulo: "CNRT"                                      },
  { cod: "4221300000", cc: "E", rotulo: "Gastos varios"                             },
  { cod: "4221400000", cc: "E", rotulo: "Seguridad"                                 },
  { cod: "4221500000", cc: "D", rotulo: "Refrigerios"                               },
  { cod: "4221600000", cc: "E", rotulo: "Reparaciones de flota"                     },
  { cod: "4221700100", cc: "E", rotulo: "Fletes"                                    },
  { cod: "4221700200", cc: "E", rotulo: "Estadía"                                   },
  { cod: "4221800000", cc: "E", rotulo: "Multas"                                    },
  { cod: "4221900000", cc: "E", rotulo: "Gastos de representación"                  },
  { cod: "4222000000", cc: "E", rotulo: "Gastos de logística"                       },
  { cod: "4222100000", cc: "E", rotulo: "Alquiler de autos"                         },
  { cod: "4222200000", cc: "E", rotulo: "Accidentes"                                },
  { cod: "4222300000", cc: "E", rotulo: "Siniestros"                                },
  { cod: "4222400000", cc: "E", rotulo: "Donaciones"                                },
  { cod: "4222500000", cc: "E", rotulo: "Estibajes"                                 },
  { cod: "4222600000", cc: "E", rotulo: "Alquiler dpto."                            },
  { cod: "4222700000", cc: "E", rotulo: "Comisiones"                                },
  { cod: "4222800000", cc: "E", rotulo: "Alquiler cochera"                          },
  { cod: "4222900000", cc: "D", rotulo: "Ropa de trabajo"                           },
  { cod: "4223000000", cc: "D", rotulo: "Bs Personales Acc. Y Participacion"        },
  { cod: "4223100000", cc: "E", rotulo: "Patentes"                                  },
  { cod: "4223200000", cc: "E", rotulo: "Tasas Senasa"                              },
  { cod: "4223300000", cc: "E", rotulo: "Juicios"                                   },
  { cod: "4223400000", cc: "E", rotulo: "Gastos varios clientes"                    },
  { cod: "4223500000", cc: "E", rotulo: "Verificación Técnica Vehicular"            },
  { cod: "4223600000", cc: "E", rotulo: "Deudores incobrables"                      },
  { cod: "4223700000", cc: "E", rotulo: "Verificación Policial"                     },
  { cod: "4223800000", cc: "E", rotulo: "Gastos varios clientes"                    },
  { cod: "4223900000", cc: "D", rotulo: "Ropa de trabajo"                           },
  { cod: "4224100000", cc: "D", rotulo: "Indemnizaciones"                           },
  { cod: "4224200000", cc: "E", rotulo: "Publicidad"                                },
  { cod: "4225000000", cc: "D", rotulo: "Cesión de Derechos"                        },
  { cod: "4226000000", cc: "E", rotulo: "Adelanto Viaje"                            },
  { cod: "4228000000", cc: "E", rotulo: "Servicio de transporte y almacenamiento"   },
  { cod: "4230100000", cc: "F", rotulo: "Gastos y comisiones bancarias"             },
  { cod: "4230200000", cc: "F", rotulo: "IMP. A LOS DÉBITOS Y CRÉDITOS"             },
  { cod: "4230400000", cc: "D", rotulo: "Intereses resarcitorios"                   },
  { cod: "4230500000", cc: "E", rotulo: "Leasing"                                   },
  { cod: "4230700000", cc: "E", rotulo: "Embargo radicación vehicular"              },
  { cod: "4230800000", cc: "F", rotulo: "Impuestos varios"                          },
  { cod: "4230900000", cc: "F", rotulo: "Impuestos al cheque"                       },
  { cod: "4231000000", cc: "F", rotulo: "R.E.C.P.A.M."                              },
  { cod: "4231200000", cc: "F", rotulo: "Gastos y comisiones bancarias"             },
  { cod: "4240100000", cc: "E", rotulo: "Combustible"                               },

  // Las seis que quedaban sueltas. En los Mensuales no las leía ningún renglón, y en los
  // Acumulados colgaban de uno ajeno — REDONDEO de "Refrigerios", GRATIFICACIONES de "Acuerdo
  // Seclo" —, así que cada archivo reportaba una cosa distinta. Cada una pasa al rótulo de su
  // propio nombre.
  { cod: "4211100000", cc: "E", rotulo: "Redondeo"                                  },
  { cod: "4211200000", cc: "D", rotulo: "Telefonos"                                 },
  { cod: "4211900000", cc: "E", rotulo: "Catering"                                  },
  { cod: "4213100000", cc: "E", rotulo: "Gratificaciones"                           },
  // LAVADO DE FLOTA va con el mantenimiento: el rótulo que los Acumulados traían se llamaba
  // "Mantenimiento y lav. de flota", así que agruparlas es lo que ya venía haciéndose.
  { cod: "4220800000", cc: "E", rotulo: "Mantenimiento de flota"                    },
  // SINIESTROS CONVENIO HSBC comparte renglón con SINIESTROS, como en el Mensual $.
  { cod: "4230600000", cc: "E", rotulo: "Siniestros"                                },
];

// Los 99 rótulos que tienen que estar en los CUATRO archivos. Es la unión de lo que cada uno
// traía, ya con los sinónimos unificados. Los que un archivo no usa quedan en cero: no molestan
// y están listos para que se les asigne una cuenta desde la configuración.
const ROTULOS_COMUNES = [
  // GRATIFICACIONES no tenía rótulo en ningún archivo: en los Acumulados colgaba de "Acuerdo
  // Seclo" y de "Sueldos". Es una cuenta con centro de costo propio, así que lleva su renglón.
  "Gratificaciones",
  "Acuerdo Seclo",
  "Adelanto Viaje",
  "Alquiler de autos",
  "Alquiler cochera",
  "Alquiler dpto.",
  "Alquiler fotocopiadora",
  "Alquiler Oficina Móvil",
  "Amortizaciones",
  "Almuerzos",
  "Autónomos Directores",
  "Agua",
  "Ajuste saldo proveedores",
  "Bs Personales Acc. Y Participacion",
  "Sueldos",
  "Cargas Sociales",
  "Salario Complementario del Decreto 332/20",
  "Cesión de Derechos",
  "Comisiones",
  "Gastos varios clientes",
  "Verificación Técnica Vehicular",
  "Accidentes",
  "Telefonos",
  "Librería",
  "Leasing",
  "Viaticos",
  "Reparaciones de flota",
  "Catering",
  "CNRT",
  "Gastos bancarios",
  "Intereses",
  "Combustible",
  "Comunicaciones",
  "Correo",
  "Custodia",
  "Despachantes",
  "Donaciones",
  "Publicidad",
  "Estadía",
  "Estacionamiento",
  "Estibajes",
  "Embargo radicación vehicular",
  "Fletes",
  "Gastos Capacitación",
  "Gastos Medicos",
  "Gastos de Representación",
  "Gastos Generales",
  "Gastos de instalación",
  "Ropa de trabajo",
  "Gastos trámites",
  "Gastos PC",
  "Gastos varios",
  "Gastos varios personal",
  "Gastos y comisiones bancarias",
  "Gastos Obra Social",
  "Gastos Hotelería",
  "Gastos Gestoria",
  "Habilitaciones",
  "Honorarios directores",
  "Honorarios profesionales",
  "Honorarios Legales",
  "IMP. A LOS DÉBITOS Y CRÉDITOS",
  "Intereses resarcitorios",
  "Impuestos varios",
  "Impuestos al cheque",
  "Juicios",
  "Indemnizaciones",
  "Legalizaciones",
  "Locomoción",
  "Gastos de logística",
  "Mantenimiento",
  "Mantenimiento de flota",
  "Mantenimiento de maquinas",
  "Mensajería",
  "Medicina Laboral",
  "Molilidad y viáticos",
  "Multas",
  "Patentes",
  "Peajes",
  "Reparaciones",
  "Refrigerios",
  "Seguros",
  "Seguridad",
  "Servicios de limpieza",
  "Deudores Incobrables",
  "Sellados",
  "Siniestros",
  "Verificación Policial",
  "Viajes y Estadías",
  "Servicio de transporte y almacenamiento",
  "Tasas AFIP",
  "Tasas IGJ",
  "R.E.C.P.A.M.",
  "Tasas SENASA",
  "DIFERENCIA",
  "ADELANTOS CASA MATRIZ",
  "Redondeo",
  "Adicional obra social",
  "Indumentaria",
  "Alquileres",
];

const RU_CC = { D: 4, E: 5, F: 6 };
const ru = {
  norm: t => String(t || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim(),
  texto: (ws, r, c) => {
    const v = ws.getCell(r, c).value;
    if (v == null) return "";
    if (typeof v === "object") return v.richText ? v.richText.map(t => t.text).join("") : "";
    return String(v);
  },
};

// Unifica cómo se escribe cada rótulo. Solo texto: no mueve ningún importe.
function renombrarRotulos(ax, log = () => {}) {
  const bloque = rtUbicarBloque(ax);
  const hechos = [];
  if (bloque.desde === null) return hechos;
  for (let r = bloque.desde; r <= bloque.hasta; r++) {
    const t = ru.texto(ax, r, 2).trim();
    if (!t) continue;
    for (const g of RENOMBRAR_ROTULOS) {
      if (!g.de.some(v => ru.norm(v) === ru.norm(t))) continue;
      if (ru.norm(t) === ru.norm(g.a)) break;
      ax.getCell(r, 2).value = g.a;
      hechos.push({ fila: r, de: t, a: g.a });
      log(`  Anexo II!B${r}: "${t}" -> "${g.a}"`);
      break;
    }
  }
  return hechos;
}

// Deja cada cuenta leída por UN solo renglón: el de su rótulo. Crea el rótulo si falta.
function asignarCuentasARotulos(wb, planDeCuentas, log = () => {}) {
  const ax = wb.getWorksheet("Anexo II");
  const enganchadas = [], creados = [], quitadas = [], salteadas = [];
  if (!ax) return { enganchadas, creados, quitadas, salteadas };

  for (const m of CUENTA_DE_ROTULO) {
    const cuenta = planDeCuentas[m.cod];
    if (!cuenta) { salteadas.push({ ...m, motivo: "la cuenta no está en SALDOS" }); continue; }

    let bloque = rtUbicarBloque(ax);
    if (!bloque.rangoTotal) { salteadas.push({ ...m, motivo: "no ubiqué el total" }); continue; }

    let filaRotulo = null;
    for (let r = bloque.desde; r <= bloque.hasta; r++) {
      if (ru.norm(ru.texto(ax, r, 2)) === ru.norm(m.rotulo)) { filaRotulo = r; break; }
    }

    // Sacarla de todo lo que no sea SU celda. No alcanza con saltear su renglón: la cuenta
    // puede estar además en otra columna de centro de costo del MISMO renglón, y entonces el
    // importe se cuenta dos veces dentro de la misma línea. Se limpia todo y queda solo la
    // columna que dice el modelo.
    const suCelda = filaRotulo === null ? null
      : `${String.fromCharCode(64 + (RU_CC[m.cc] || 5))}${filaRotulo}`;
    for (let r = bloque.desde; r <= bloque.hasta; r++) {
      for (const c of [4, 5, 6]) {
        const dir = `${String.fromCharCode(64 + c)}${r}`;
        if (dir === suCelda) continue;
        if (rtQuitarTermino(ax, dir, cuenta.fila)) {
          quitadas.push({ cod: m.cod, de: `${dir} "${ru.texto(ax, r, 2).trim()}"` });
          log(`  Anexo II: ${m.cod} sacada de ${dir} "${ru.texto(ax, r, 2).trim()}"`);
        }
      }
    }

    if (filaRotulo === null) {
      // dentro del rango del total, para que el SUM lo tome (ver rotulos_anexo.js)
      bloque = rtUbicarBloque(ax);
      filaRotulo = bloque.rangoTotal.hasta;
      insertRowEn(wb, "Anexo II", filaRotulo);
      ax.getCell(filaRotulo, 2).value = m.rotulo;
      ax.getCell(filaRotulo, 3).value = { formula: `SUM(D${filaRotulo}:F${filaRotulo})` };
      for (const c of [4, 5, 6]) ax.getCell(filaRotulo, c).value = 0;
      creados.push({ rotulo: m.rotulo, fila: filaRotulo });
      log(`  Anexo II: renglón "${m.rotulo}" creado en la fila ${filaRotulo}`);
    }

    const colImp = bloque.colImporte || "C";
    const celda = ax.getCell(filaRotulo, RU_CC[m.cc] || 5);
    const v = celda.value;
    const ya = v && typeof v === "object" && typeof v.formula === "string" &&
      new RegExp("SALDOS!\\$?[A-Z]{1,3}\\$?" + cuenta.fila + "(?!\\d)").test(v.formula);
    if (!ya) {
      const previo = (v && typeof v === "object" && typeof v.formula === "string") ? v.formula : "";
      celda.value = { formula: previo ? `${previo}+SALDOS!${colImp}${cuenta.fila}`
                                      : `+SALDOS!${colImp}${cuenta.fila}` };
      enganchadas.push({ cod: m.cod, rotulo: m.rotulo, fila: filaRotulo });
    }
  }
  return { enganchadas, creados, quitadas, salteadas };
}

// Agrega los rótulos que este archivo no tenga, vacíos. Así los cuatro abren el gasto con la
// misma lista y se pueden comparar renglón por renglón.
function completarRotulos(wb, log = () => {}) {
  const ax = wb.getWorksheet("Anexo II");
  const agregados = [];
  if (!ax) return agregados;
  for (const rotulo of ROTULOS_COMUNES) {
    const bloque = rtUbicarBloque(ax);
    if (!bloque.rangoTotal) break;
    let existe = false;
    for (let r = bloque.desde; r <= bloque.hasta; r++) {
      if (ru.norm(ru.texto(ax, r, 2)) === ru.norm(rotulo)) { existe = true; break; }
    }
    if (existe) continue;
    const fila = bloque.rangoTotal.hasta;      // dentro del rango, para que el total lo tome
    insertRowEn(wb, "Anexo II", fila);
    ax.getCell(fila, 2).value = rotulo;
    ax.getCell(fila, 3).value = { formula: `SUM(D${fila}:F${fila})` };
    for (const c of [4, 5, 6]) ax.getCell(fila, c).value = 0;
    agregados.push({ rotulo, fila });
    log(`  Anexo II: renglón "${rotulo}" agregado vacío en la fila ${fila}`);
  }
  return agregados;
}

function unificarRotulosAnexo(wb, planDeCuentas, log = () => {}) {
  const ax = wb.getWorksheet("Anexo II");
  if (!ax) return { renombrados: [], enganchadas: [], creados: [], quitadas: [], salteadas: [], completados: [] };
  const renombrados = renombrarRotulos(ax, log);
  const asignadas = asignarCuentasARotulos(wb, planDeCuentas, log);
  const completados = completarRotulos(wb, log);
  return { renombrados, ...asignadas, completados };
}

if (typeof module !== "undefined") {
  const fh = require("./formula_hojas.js");
  const ra = require("./rotulos_anexo.js");
  global.insertRowEn = fh.insertRowEn;
  global.rtUbicarBloque = ra.rtUbicarBloque;
  global.rtQuitarTermino = ra.rtQuitarTermino;
  module.exports = { RENOMBRAR_ROTULOS, CUENTA_DE_ROTULO, ROTULOS_COMUNES,
    unificarRotulosAnexo, renombrarRotulos, asignarCuentasARotulos, completarRotulos };
}
