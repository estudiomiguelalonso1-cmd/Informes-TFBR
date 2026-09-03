# Automatización del cierre mensual TFBR

Automatiza la parte mecánica del cierre mensual de TRANSPORTES FURLONG DO BRASIL S.A.: tomar el
export de "Sumas y Saldos" de Onvio y volcarlo en los 4 balances Excel (Mensual $, Mensual R$,
Acumulado $, Acumulado R$), recalculando y validando antes de entregar el archivo.

**Lo que NO automatiza a propósito** (sigue siendo manual, como hoy): movimientos de patrimonio
neto, altas/bajas de bienes de uso, y el reparto de gastos por centro de costo del Anexo II — el
archivo generado incluye un checklist con las celdas exactas que faltan completar.

Contexto completo del proceso original y de cada decisión de diseño: `docs/formula_analysis.md` y
`docs/defect_diagnosis.md`.

## Estado actual

🚧 **Fase 0 (spike) en curso.** El pipeline (`src/`, `macros/uno_pipeline.py`) está escrito pero
**todavía no se probó contra un archivo real** — ver "Cómo seguir" abajo. No usar el resultado de
un run para un cierre real hasta confirmar la Fase 1 (ver plan).

## Flujo mensual (una vez pasada la Fase 1)

1. Exportar de Onvio los dos "Sumas y Saldos" (mensual y acumulado), como siempre.
2. Guardarlos en `inputs/<YYYY-MM>/sumas_y_saldos_mensual.xls` y `..._acumulado.xls` (nombres
   fijos), `git add` + commit + push.
3. En GitHub → Actions → **Monthly Close** → *Run workflow*, completar:
   - `period`: `2026-07`
   - `tc_cierre`: el tipo de cambio de cierre del período
   - `dif_cambio_mes`: solo si hace falta cargar la fila del mes en el cuadro de diferencia de
     cambio del Balance Acumulado R$ (ver checklist del run anterior)
   - `target_files`: vacío o `balance_mensual_ars,balance_mensual_brl,balance_acumulado_ars,balance_acumulado_brl`
4. Revisar el resumen que aparece en la pantalla del run (validaciones + checklist de pasos
   manuales pendientes).
5. Descargar el artifact, completar en Excel real lo que falta, y guardar la copia final donde
   se guarde hoy.

## Cómo seguir (Fase 0)

1. **Instalar LibreOffice** (local, para probar rápido antes de tocar GitHub Actions) o confirmar
   que el workflow de Actions arranca bien (`apt install libreoffice-calc python3-uno`).
2. **Confirmar el layout de `Hoja1`** del Sumas y Saldos **acumulado** (el mensual ya está
   confirmado: `Hoja1!C7` = TC CIERRE, ver `docs/formula_analysis.md`) — ¿es igual?
3. **Probar `macros/uno_pipeline.SheetWriter`** contra una copia de
   `templates/balance_mensual_ars.xls` (que tiene el vínculo externo muerto real): abrir con
   `UpdateDocMode=NO_UPDATE`, escribir un par de celdas de prueba en `SALDOS`, `recalc()`, y
   comparar visualmente contra lo que Excel real muestra hoy para las mismas celdas.
4. Recién ahí correr `python -m src.pipeline --period 2026-07 --tc-cierre 293.1757 --target
   balance_mensual_ars` de punta a punta y revisar `output/2026-07/`.

## Estructura

```
templates/    los 4 .xls "base" (se copian, nunca se editan in place)
config/       layouts por archivo, fixes aprobados/pendientes, checks de validación
src/          orquestador Python (sin dependencia de UNO salvo pipeline.py/uno_pipeline.py)
macros/       cliente UNO (LibreOffice headless)
inputs/       exports de Onvio por período (committeados, dan auditoría/historial)
output/       generado, gitignored — solo vive como artifact del workflow run
docs/         investigación original de los archivos base (referencia viva)
```
