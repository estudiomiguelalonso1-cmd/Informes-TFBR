# Códigos repetidos — resuelto

Los cuatro balances quedaron **sin códigos repetidos**. Este documento queda como registro de
qué eran y cómo se resolvieron.

## Qué eran

Los códigos del plan de cuentas de la hoja `SALDOS` se tipean a mano, y decenas quedaron con
**un dígito de menos** — nueve en vez de diez, siempre un cero:

```
422040000  CUSTODIA        ← lo que tenía el balance
4220400000 CUSTODIA        ← lo que dice el plan oficial
```

Una cuenta con el código mal **nunca levanta importe**: el motor empareja por código contra lo
que manda Onvio, no la encuentra, y el importe entra en cero sin que nada avise. Cuando más
adelante alguien daba de alta la cuenta con el código correcto, la fila vieja se quedaba ahí.
De ahí salían los códigos repetidos.

No eran 17 casos sueltos: era un solo problema con muchas manifestaciones.

## Cómo se resuelven

Con `informe-tfbr/limpieza_plan.js`, que corre al principio de cada corrida y compara contra
`plan_oficial.js` (el plan exportado del sistema contable). Cada cambio se valida por
**código + nombre** contra esa referencia; nada se decide por criterio.

Ver el encabezado de `limpieza_plan.js` para los pasos.

### Resultado sobre los cuatro maestros

| | Códigos corregidos | Reasignados | Filas fusionadas |
|---|---|---|---|
| Mensual $ | 38 | 0 | 6 |
| Mensual R$ | 37 | 1 | 5 |
| Acumulado $ | 2 | 1 | 8 |
| Acumulado R$ | 0 | 1 | 4 |

**Códigos repetidos: 13 → 0.** Los cuatro archivos siguen emparejando el 100% del export,
sin cuentas sin mapear, y sin errores nuevos.

## Las tres que necesitaron definición de contaduría

Están en la tabla `UNIFICACIONES` de `limpieza_plan.js`, cada una con el motivo al lado:

- **`IMP. A LOS CREDITOS`** — el código `4230300000` no existe en el plan oficial. El plan
  tiene una sola cuenta que cubre débitos y créditos (`4230200000`); esta se había desprendido
  de aquella. Se unifican.
- **`IMP. A LOS DEBITOS`** — la misma cuenta, con el nombre cortado a la mitad.
- **`GASTOS TELEFÓNICOS`** — el plan tiene una sola cuenta de teléfono,
  `4211200000 GASTOS EN EQ. TELEFÓNICOS`. No eran dos cuentas: era la misma cargada dos veces
  con el nombre escrito distinto. Queda el nombre oficial.

## Lo que sigue sin resolverse

Una fila por archivo, y ninguna es un duplicado:

- **Mensual $ y Mensual R$** — `422300000 IMP BIENES SOCIEDADES`. No hay nada con ese nombre
  en el plan oficial. Lo más parecido es `4223000000 BS. PERSONALES ACCIO. Y PARTICIPACIONES`,
  que es a lo que apunta el nombre, pero no coinciden lo suficiente como para decidirlo solo.
- **Acumulado $ y Acumulado R$** — una fila con el "código" `412` y el texto `OTROS INGRESOS`.
  No es una cuenta: es un título de agrupación que quedó dentro del rango del plan.

## Aparte: el corrimiento del Anexo II

Distinto de los duplicados y todavía abierto: en el Acumulado $, 25 líneas del Anexo II leen
la cuenta de una fila más arriba de la que corresponde a su rótulo. Doce ya se corrigieron
(ver `repuntes_anexo.js`); las trece restantes comparten su cuenta de destino con otra línea
y necesitan definición. El detalle está en esa lista y en el historial del repositorio.
