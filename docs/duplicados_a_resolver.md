# Códigos repetidos — lo que sigue sin resolverse

Actualizado después de aplicar las definiciones del documento *Cuentas TFBR*
(las filas marcadas en rojo). De los 17 casos originales ya se resolvieron 10:
8 borrando la fila que sobraba, y 2 —los de `ADELANTO VIAJE`— corrigiéndole el código,
que resultaron no ser duplicados sino una cuenta mal codificada (sección 0).

Ninguna de las cuentas que quedan viene con movimiento en el Sumas y Saldos de julio,
así que el cierre de julio no está afectado. El problema se activa el mes que alguna
tenga importe.

## 0. RESUELTO — ADELANTO VIAJE no estaba repetida: tenía el código mal (2)

Los dos casos de `ADELANTO VIAJE` no eran un código compartido por dos cuentas distintas,
sino **una cuenta cargada con el código de la vecina**. Los dos archivos en pesos coinciden
en cuál es cuál, y ahí no hay ningún duplicado:

| código | Mensual $ | Acumulado $ |
|---|---|---|
| `4223600000` | DEUDORES INCOBRABLES | DEUDORES INCOBRABLES |
| `4225000000` | CESIÓN DE DERECHOS | CESIÓN DE DERECHOS |
| `4226000000` | ADELANTO DE VIAJE | ADELANTO VIAJE |

En los dos archivos en reales, en cambio, `ADELANTO VIAJE` estaba cargada con `4223600000`
(Mensual R$) y con `4225000000` (Acumulado R$), y en ninguno de los dos existía una fila
para `4226000000`.

Por eso **no se borra la fila: se le corrige el código** a `4226000000`. Borrarla habría
dejado a los dos archivos en reales sin el concepto "Adelanto Viaje" del Anexo II y sin
ninguna fila para esa cuenta, así que el mes que Onvio mandara movimiento habría entrado
como cuenta sin mapear.

De paso arregla un error que ya estaba activo: `4223600000` resolvía a la fila de
`ADELANTO VIAJE` (la primera de las dos), así que un importe de `DEUDORES INCOBRABLES` se
habría reportado en la línea "Adelanto Viaje" del Anexo II.

Está en `decisiones_duplicados.js` como `recodificarNombre` + `codigoNuevo`. Después de
aplicarlo, `Anexo II!E10` (Mensual R$) y `Anexo II!E14` (Acumulado R$) —las dos rotuladas
"Adelanto Viaje"— leen la fila que ahora sí dice `4226000000  ADELANTO VIAJE`.

## 1. Ya definidos, pero falta un dato para poder borrarlos (2)

En estos dos, la fila que se decidió eliminar **está siendo leída por una línea del
Anexo II**. Si se borra sin más, esa línea queda en `#REF!`. Hace falta definir qué
cuenta pasa a leer cada una de esas líneas (o si la línea queda en cero).

Ojo: los dos son del Acumulado $, y las dos líneas que los referencian **ya estaban leyendo
la cuenta equivocada** antes de todo esto. No son un problema de duplicados: son dos casos
del corrimiento que se describe en la sección 3.

### Balance Acumulado $ — 4211100000

- Fila 130: `4211100000  REDONDEO`
  - la lee `Anexo II!D74`, que es la línea **"Refrigerios"**
- Fila 131: `4211100000  SERVICIOS DE LIMPIEZA` **(la que se eliminaría)**
  - la lee `Anexo II!D45`, que es la línea **"Gastos Telefonico"**

**Falta definir:** al eliminar `SERVICIOS DE LIMPIEZA`, ¿qué cuenta pasa a leer esa línea del Anexo II?

### Balance Acumulado $ — 4211200000

- Fila 127: `4211200000  GASTOS EN EQ. TELEFÓNICOS`
  - la lee `Anexo II!E17`, que es la línea **"Adicional obra social"**
- Fila 132: `4211200000  GASTOS TELEFÓNICOS` **(la que se eliminaría)**
  - la lee `Anexo II!D42`, que es la línea **"Gastos obra social"**

**Falta definir:** al eliminar `GASTOS TELEFONICOS`, ¿qué cuenta pasa a leer esa línea del Anexo II?

## 2. Sin definición todavía: el mismo importe se cuenta dos veces (5)

Estos no estaban en el documento. Las dos filas tienen el mismo código **y el mismo
nombre**, así que las dos levantan el mismo importe del Sumas y Saldos, y las dos
alimentan líneas de los estados: el importe entra dos veces.

### Balance Mensual R$ — 4212800000

- Fila 161: `4212800000  TASAS AFIP`
  - la lee `Anexo II!D105` — línea **"Tasas AFIP"** — `+SALDOS!C161`
- Fila 217: `4212800000  TASAS AFIP`
  - la lee `Anexo II!E21` — línea **"Ajuste saldo proveedores"** — `+SALDOS!C217`

**Falta definir:** cuál de las dos filas queda, y en qué línea tiene que reportarse el importe.

### Balance Acumulado $ — 4120500000

- Fila 114: `4120500000  INTERESES GANADOS`
  - la lee `EERR!C25` — línea **"Intereses"** — `+SALDOS!D115+SALDOS!D120+SALDOS!D114`
- Fila 120: `4120500000  INTERESES GANADOS`
  - la lee `EERR!C25` — línea **"Intereses"** — `+SALDOS!D115+SALDOS!D120+SALDOS!D114`

**Falta definir:** cuál de las dos filas queda, y en qué línea tiene que reportarse el importe.

### Balance Acumulado $ — 4222900000

- Fila 188: `4222900000  ROPA DE TRABAJO`
  - la lee `Anexo II!E32` — línea **"Estibajes"** — `+SALDOS!C188`
- Fila 193: `4222900000  ROPA DE TRABAJO`
  - la lee `Anexo II!E53` — línea **"Imp. Bienes Acc. Y Part. Soc."** — `+SALDOS!C193`

**Falta definir:** cuál de las dos filas queda, y en qué línea tiene que reportarse el importe.

### Balance Acumulado R$ — 4120500000

- Fila 115: `4120500000  INTERESES GANADOS`
  - la lee `EERR!C24` — línea **"Intereses"** — `+SALDOS!E115`
- Fila 121: `4120500000  INTERESES GANADOS`
  - la lee `EERR!C22` — línea **"Otros ingresos y egresos"** — `+SALDOS!E116+SALDOS!E121+SALDOS!E187+SALDOS!E117`

**Falta definir:** cuál de las dos filas queda, y en qué línea tiene que reportarse el importe.

### Balance Acumulado R$ — 4230400000

- Fila 196: `4230400000  INTERESES RESARCITORIOS`
  - la lee `Anexo II!E58` — línea **"Intereses resarcitorios"** — `+SALDOS!D196`
- Fila 205: `4230400000 INTERESES RESARCITORIOS`
  - la lee `Anexo II!F58` — línea **"Intereses resarcitorios"** — `+SALDOS!D205`

**Falta definir:** cuál de las dos filas queda, y en qué línea tiene que reportarse el importe.

## 3. Aparte: líneas del Anexo II que leen una cuenta que no es la suya

Esto salió al revisar los casos de arriba y es un problema distinto, del **Acumulado $**:
varias líneas del Anexo II leen una cuenta que no corresponde a su propio rótulo, como si en
algún momento se hubieran corrido una fila.

| Línea del Anexo II | Rótulo de la línea | Cuenta que lee |
|---|---|---|
| `D42` | Gastos obra social | `4211200000 GASTOS TELEFÓNICOS` |
| `D45` | Gastos Telefonico | `4211100000 SERVICIOS DE LIMPIEZA` |
| `D74` | Refrigerios | `4211100000 REDONDEO` |
| `E17` | Adicional obra social | `4211200000 GASTOS EN EQ. TELEFÓNICOS` |

Hoy no cambia ningún número porque esas cuentas están en cero, pero conviene revisarlo junto
con los casos del punto 1: es la misma definición (qué cuenta tiene que leer cada línea).
