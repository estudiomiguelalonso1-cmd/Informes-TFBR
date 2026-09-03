# Códigos repetidos que quedan sin resolver

Los que se podían resolver sin cambiar ningún número ya se borraron de los maestros
corregidos (la fila huérfana de cada par, dejando la que está enganchada a los estados).
Los que siguen acá **cambian algo del balance según cómo se resuelvan**, así que
necesitan una decisión contable. La app los avisa en cada corrida y no los toca.

Hoy casi todas estas cuentas están en cero, así que el balance de julio no está
distorsionado — el problema se activa el mes que alguna tenga movimiento.

## 1. El mismo código usado por dos cuentas diferentes (8)

Hay que decidir qué código le corresponde a cada cuenta. Borrar una fila acá elimina
una cuenta real del balance.

| Archivo | Código | Fila | Cuenta | La referencia |
|---|---|---|---|---|
| Balance Mensual R$ | 4223600000 | 207 | 4223600000  ADELANTO VIAJE | `Anexo II!E10` |
| Balance Mensual R$ | 4223600000 | 208 | 4223600000  DEUDORES INCOBRABLES | `EERR!C21`, `Anexo II!E97` |
| Balance Acumulado $ | 4211100000 | 132 | 4211100000  REDONDEO | `Anexo II!D74` |
| Balance Acumulado $ | 4211100000 | 133 | 4211100000  SERVICIOS DE LIMPIEZA | `Anexo II!D45` |
| Balance Acumulado $ | 4211200000 | 129 | 4211200000  GASTOS EN EQ. TELEFÓNICOS | `Anexo II!E17` |
| Balance Acumulado $ | 4211200000 | 134 | 4211200000  GASTOS TELEFÓNICOS | `Anexo II!D42` |
| Balance Acumulado $ | 4211700000 | 136 | 4211700000  GASTOS DE CAPACITACION | `Anexo II!D71` |
| Balance Acumulado $ | 4211700000 | 142 | 4211700000  GASTOS CAPACITACIÓN | — |
| Balance Acumulado $ | 4230200000 | 211 | 4230200000  IMP. A LOS DÉBITOS Y CRÉDITOS LEY 25,413 | — |
| Balance Acumulado $ | 4230200000 | 212 | 4230200000  IMP. A LOS DEBITOS | `Anexo II!F54` |
| Balance Acumulado $ | 4212800000 | 152 | 4212800000  TASA AFIP | `Anexo II!D82` |
| Balance Acumulado $ | 4212800000 | 221 | 4212800000  TASAS AFIP | — |
| Balance Acumulado R$ | 4225000000 | 188 | 4225000000  CESIÓN DE DERECHOS | `Anexo II!D22` |
| Balance Acumulado R$ | 4225000000 | 194 | 4225000000  ADELANTO VIAJE | `Anexo II!E14` |
| Balance Acumulado R$ | 4230200000 | 196 | 4230200000  IMP. A LOS DÉBITOS Y CRÉDITOS LEY 25,413 | `Anexo II!F56` |
| Balance Acumulado R$ | 4230200000 | 197 | 4230200000  IMP. A LOS DEBITOS | — |

## 2. Doble conteo: el mismo importe entra dos veces (5)

Las dos filas levantan el mismo importe y las dos están enganchadas a los estados.
Hay que decidir en qué línea tiene que quedar.

### Balance Mensual R$ — 4212800000

- **Fila 166**: `4212800000  TASAS AFIP` — `Anexo II!D105` = `+SALDOS!C166`
- **Fila 222**: `4212800000  TASAS AFIP` — `Anexo II!E21` = `+SALDOS!C222`

las dos filas están referenciadas por líneas distintas (Anexo II!D105 y Anexo II!E21): cuál sobrevive cambia en qué línea del estado se reporta el importe.

### Balance Acumulado $ — 4120500000

- **Fila 116**: `4120500000  INTERESES GANADOS` — `EERR!C25` = `+SALDOS!D117+SALDOS!D122+SALDOS!D116`
- **Fila 122**: `4120500000  INTERESES GANADOS` — `EERR!C25` = `+SALDOS!D117+SALDOS!D122+SALDOS!D116`

la misma fórmula (EERR!C25) suma las dos filas, así que el importe entra dos veces en esa línea. Hay que sacar uno de los dos términos de la fórmula antes de borrar la fila.

### Balance Acumulado $ — 4222900000

- **Fila 190**: `4222900000  ROPA DE TRABAJO` — `Anexo II!E32` = `+SALDOS!C190`
- **Fila 195**: `4222900000  ROPA DE TRABAJO` — `Anexo II!E53` = `+SALDOS!C195`

las dos filas están referenciadas por líneas distintas (Anexo II!E32 y Anexo II!E53): cuál sobrevive cambia en qué línea del estado se reporta el importe.

### Balance Acumulado R$ — 4120500000

- **Fila 117**: `4120500000  INTERESES GANADOS` — `EERR!C24` = `+SALDOS!E117`
- **Fila 123**: `4120500000  INTERESES GANADOS` — `EERR!C22` = `+SALDOS!E118+SALDOS!E123+SALDOS!E189+SALDOS!E119`

las dos filas están referenciadas por líneas distintas (EERR!C24 y EERR!C22): cuál sobrevive cambia en qué línea del estado se reporta el importe.

### Balance Acumulado R$ — 4230400000

- **Fila 199**: `4230400000  INTERESES RESARCITORIOS` — `Anexo II!E58` = `+SALDOS!D199`
- **Fila 208**: `4230400000 INTERESES RESARCITORIOS` — `Anexo II!F58` = `+SALDOS!D208`

las dos filas están referenciadas por líneas distintas (Anexo II!E58 y Anexo II!F58): cuál sobrevive cambia en qué línea del estado se reporta el importe.
