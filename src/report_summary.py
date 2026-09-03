"""Arma el resumen en markdown que se escribe en $GITHUB_STEP_SUMMARY: una tabla de
validaciones por archivo + un checklist de lo que sigue siendo manual cada mes.
No se imprime en un log que nadie lee - esto es lo primero que el usuario ve al abrir el run.
"""
from __future__ import annotations

from src.validate import ValidationReport

# Checklist de pasos manuales fijos (direcciones confirmadas en docs/formula_analysis.md
# y docs/defect_diagnosis.md). No cambia de mes a mes, salvo la fila de BAR que depende del mes.
MANUAL_STEPS = {
    "balance_mensual_ars": [
        "EEPN!C11:H11 y C19:H19 - movimientos de patrimonio neto del mes",
        "Anexo I!B11:B17, C11:C17, E11:E17, F11:F17, J11:J17 - altas/bajas de bienes de uso",
        "Anexo II!J:M (filas ~40-105) - reparto de gastos por centro de costo (ADM CENTRAL/PACHECO/CORDOBA/TRANSPORTES)",
    ],
    "balance_mensual_brl": [
        "EEPN!C11:H11 y C19:H19 - movimientos de patrimonio neto del mes",
        "Anexo I - altas/bajas de bienes de uso (mismo criterio que en $, ver Anexo I!D17 con nota aparte)",
        "Anexo II!J:M (filas ~38-100) - reparto de gastos por centro de costo (revisar conversion a R$, ver defect_diagnosis.md)",
    ],
    "balance_acumulado_ars": [
        "EEPN - movimientos de patrimonio neto acumulados",
        "EERR!C28 (Impuesto a las Ganancias) - hoy es un 0 tipeado, confirmar si corresponde cargar un valor real",
    ],
    "balance_acumulado_brl": [
        "EEPN - movimientos de patrimonio neto acumulados",
        "EERR!C27 (Impuesto a las Ganancias) - hoy esta en blanco, confirmar si corresponde cargar un valor",
        "Bienes!A6 - revisar el calculo de 0.5% sobre PN (rotulo interno dice '2009', formulas si estan vivas)",
    ],
}

PENDING_BACKLOG_NOTE = (
    "**Pendiente de confirmar con contaduria (no se toca todavia):** "
    "`EESP!F13` en Balance Mensual $ (vinculo muerto 2004, valor en cache = 0), y la familia de "
    "`#REF!` de la columna de comparacion 2002-2004 muerta en `Pasivo!K` / `Anexo II!G` / "
    "`EEPN!K17` de varios archivos. Ver `docs/defect_diagnosis.md` para el detalle completo."
)


def render_checks_table(report: ValidationReport) -> str:
    lines = [f"### {report.file_id}", "", "| Check | Celda | Esperado | Real | Resultado |", "|---|---|---|---|---|"]
    for c in report.checks:
        actual_str = "N/D" if c.actual is None else f"{c.actual:,.2f}"
        status = "OK" if c.passed else "**FALLO**"
        lines.append(f"| {c.description} | `{c.cell}` | {c.expected:,.2f} | {actual_str} | {status} |")

    if report.unmapped_accounts:
        lines.append("")
        lines.append(f"**Cuentas sin mapear al template (NO incluidas en ningun total):** {', '.join(report.unmapped_accounts)}")

    if report.new_errors:
        lines.append("")
        lines.append(f"**Errores NUEVOS detectados (no estaban documentados):** {', '.join(report.new_errors)}")

    if report.known_errors:
        lines.append("")
        lines.append(f"_Errores preexistentes conocidos (no bloquean, ver docs/defect_diagnosis.md): {', '.join(report.known_errors)}_")

    lines.append("")
    lines.append(f"**Resultado general: {'PASA' if report.passed else 'FALLA - revisar antes de usar el archivo'}**")
    return "\n".join(lines)


def render_manual_checklist(file_id: str) -> str:
    steps = MANUAL_STEPS.get(file_id, [])
    lines = [f"#### Pasos manuales pendientes - {file_id}"]
    lines += [f"- [ ] {s}" for s in steps]
    return "\n".join(lines)


def render_full_summary(reports: list[ValidationReport]) -> str:
    parts = ["## Resultado del cierre automatizado", ""]
    for r in reports:
        parts.append(render_checks_table(r))
        parts.append("")
        parts.append(render_manual_checklist(r.file_id))
        parts.append("")
    parts.append("---")
    parts.append(PENDING_BACKLOG_NOTE)
    return "\n".join(parts)
