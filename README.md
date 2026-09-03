# Informes TFBR — cierre mensual

Automatiza la parte mecánica del cierre mensual de TRANSPORTES FURLONG DO BRASIL S.A.:
toma el Sumas y Saldos exportado de Onvio y lo vuelca en los 4 balances
(Mensual $, Mensual R$, Acumulado $, Acumulado R$).

Corre entero en el navegador: no hay que instalar nada, se entra por un link.

**App:** https://estudiomiguelalonso1-cmd.github.io/Informes-TFBR/

## Cómo funcionan los archivos (importante)

Cada balance **se calcula solo**. La hoja `SALDOS` tiene el plan de cuentas completo, y cada
fila trae su saldo con un `BUSCARV` contra la zona de pegado que está más abajo en esa misma
hoja. De `SALDOS` salen EESP, EERR, EEPN, Activo, Pasivo, Anexo I y Anexo II.

Por eso el motor escribe **únicamente en la zona de pegado**. Todo lo demás lo resuelve
Excel con las fórmulas que el archivo ya tiene, al abrirlo. Escribir importes a mano en los
estados pisaría esas fórmulas.

Los 4 archivos tienen la zona de pegado en lugares distintos (columnas y filas diferentes en
cada uno). La app no los tiene anotados: los **deduce leyendo las fórmulas de cada archivo**,
así que si mañana cambian de lugar, sigue funcionando.

## Uso mensual

1. Exportás de Onvio los dos Sumas y Saldos (mensual y acumulado), como siempre.
2. Los subís a la app, junto con el período y el tipo de cambio de cierre.
3. Revisás los controles en pantalla (cuentas emparejadas, totales, cuentas sin mapear).
4. Descargás los 4 borradores, los abrís en Excel — ahí es donde se recalcula todo —,
   revisás y guardás.
5. Subís los 4 revisados y cerrás el mes: quedan guardados en este repositorio como los
   maestros del mes siguiente.

El paso 4 no es un capricho: quien calcula esas fórmulas es Excel. Reproducir en el navegador
la cadena `SALDOS → 7 hojas` sería frágil y podría dar totales mal sin avisar. Usando el
archivo que Excel ya calculó, lo que se guarda es exactamente lo que se vio en pantalla.

## Primera vez (una sola vez)

1. **Convertir los 4 balances a .xlsx**: abrilos en Excel → Archivo → Guardar como → "Libro
   de Excel (.xlsx)". El formato viejo `.xls` no conserva las fórmulas al procesarlo en el
   navegador, por eso se convierte una vez y de ahí en más se trabaja siempre en `.xlsx`.
2. **Configurar GitHub** en la app (⚙): un token personal con permiso `repo`, y el
   repositorio. El token queda guardado solo en tu navegador.
3. **Subir los 4 maestros** convertidos. Quedan guardados acá y no hay que volver a subirlos.

## Lo que la app NO toca (sigue siendo manual)

- **Anexo II, columnas J a M** (reparto de gastos por centro de costo): son sumas de
  comprobantes cargadas a mano, sin fuente de datos que las respalde.
- **EEPN**: movimientos de patrimonio neto.
- **Anexo I**: altas y bajas de bienes de uso.
- **Impuesto a las ganancias** en los acumulados.

## Arreglos aplicados y pendientes

La app aplica **un** arreglo automático, ya aprobado: congela `Anexo II!D122/D126/D129/D130/D132`
del Mensual $, que tenían un vínculo a un archivo de mayo 2012 que ya no existe (los números
que se ven no cambian; lo que se rompe es el vínculo muerto).

Quedan **pendientes de confirmar con contaduría**, y por eso la app no los toca:

- `EESP!F13` del Mensual $ — vínculo muerto a un balance 2004 de otra empresa del grupo. Hoy
  vale 0, así que corregirlo no cambiaría ningún número.
- La columna de comparación contra ejercicios 2002-2004 (`Pasivo!K`, `Anexo II!G`, `EEPN!K17`
  según el archivo), rota con `#REF!` desde hace años. No tiene datos vivos abajo: reconstruir
  la fórmula daría 0 igual.

El detalle completo de cada celda está en [`docs/defect_diagnosis.md`](docs/defect_diagnosis.md),
y el mapa de todas las fórmulas de los 4 archivos en
[`docs/formula_analysis.md`](docs/formula_analysis.md).

## Archivos

```
index.html          portada
estilos.css         estilos compartidos
informe-tfbr/
  index.html        la app
  app.js            conecta la pantalla con el motor
  parser_tfbr.js    lee el Sumas y Saldos de Onvio (columnas por encabezado, no por posición)
  config_tfbr.js    deduce el layout de SALDOS leyendo las fórmulas del archivo
  motor_tfbr.js     empareja por código y escribe la zona de pegado
  fixes_tfbr.js     los arreglos explícitos aprobados
  github_tfbr.js    lee y guarda los maestros y el historial en este repositorio
  formula_hojas.js  reacomoda fórmulas al insertar filas (ExcelJS no lo hace solo)
  formula_utils.js  materializa fórmulas compartidas al abrir un archivo
  trazabilidad.js   (todavía sin usar: adaptado de otro informe, para cuando se automatice
                     el alta de cuentas nuevas)
docs/               la investigación de los 4 archivos base
config/             direcciones de celda confirmadas, como referencia de test
```
