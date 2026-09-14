# Códigos repetidos — lo que sigue sin resolverse

Actualizado después de aplicar las definiciones del documento *Cuentas TFBR*
(las filas marcadas en rojo). De los 17 casos originales ya se resolvieron 8.

Ninguna de las cuentas que quedan viene con movimiento en el Sumas y Saldos de julio,
así que el cierre de julio no está afectado. El problema se activa el mes que alguna
tenga importe.

## 1. Ya definidos, pero falta un dato para poder borrarlos (4)

En estos cuatro, la fila que se decidió eliminar **está siendo leída por una línea del
Anexo II**. Si se borra sin más, esa línea queda en `#REF!`. Hace falta definir qué
cuenta pasa a leer cada una de esas líneas (o si la línea queda en cero).

### Balance Mensual R$ — 4223600000

- Fila 202: `4223600000  ADELANTO VIAJE` **(la que se eliminaría)**
  - la lee `Anexo II!E10`, que es la línea **"Adelanto Viaje"**
- Fila 203: `4223600000  DEUDORES INCOBRABLES`
  - la lee `EERR!C21`, que es la línea **"Otros Ingresos"**
  - la lee `Anexo II!E97`, que es la línea **"Deudores Incobrables"**

**Falta definir:** al eliminar `ADELANTO VIAJE`, ¿qué cuenta pasa a leer esa línea del Anexo II?

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

### Balance Acumulado R$ — 4225000000

- Fila 186: `4225000000  CESIÓN DE DERECHOS`
  - la lee `Anexo II!D22`, que es la línea **"Cesión de Derechos"**
- Fila 192: `4225000000  ADELANTO VIAJE` **(la que se eliminaría)**
  - la lee `Anexo II!E14`, que es la línea **"Adelanto Viaje"**

**Falta definir:** al eliminar `ADELANTO VIAJE`, ¿qué cuenta pasa a leer esa línea del Anexo II?

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
