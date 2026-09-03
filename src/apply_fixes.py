"""Motor generico de 2 primitivas para aplicar los fixes de config/fixes/*.yml.

No sabe nada de contabilidad ni de layouts - solo sabe escribir un valor fijo o una formula
en una celda dada. La logica de "que corregir y por que" vive enteramente en la config,
para poder activar los fixes pendientes (config/fixes/pending_backlog.yml) sin tocar codigo
el dia que contaduria los confirme.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

FixAction = Literal["freeze_static", "replace_formula"]


@dataclass(frozen=True)
class CellFix:
    sheet: str
    address: str  # ej "D122"
    action: FixAction
    value: float | str


def _split_ref(ref: str) -> tuple[str, str]:
    """'Anexo II!D122' -> ('Anexo II', 'D122')"""
    sheet, address = ref.split("!", 1)
    return sheet, address


def load_fixes_for_file(fix_config: dict, file_id: str) -> list[CellFix]:
    """Aplana una entrada de approved.yml/pending_backlog.yml (con enabled=true) a CellFix."""
    fixes: list[CellFix] = []

    for _key, entry in fix_config.items():
        if not entry.get("enabled", True):
            continue

        # forma "approved.yml": un dict {file, action, cells:{ref: valor}} o {file, action, cell, new_value}
        if "entries" in entry:
            for sub in entry["entries"]:
                if sub.get("file") != file_id:
                    continue
                sheet, addr = _split_ref(sub["cell"])
                fixes.append(CellFix(sheet, addr, sub["action"], sub["new_value"]))
            continue

        if entry.get("file") != file_id:
            continue

        if "cells" in entry:  # freeze_static con multiples celdas
            for ref, val in entry["cells"].items():
                sheet, addr = _split_ref(ref)
                fixes.append(CellFix(sheet, addr, entry["action"], val))
        elif "cell" in entry:  # replace_formula de una sola celda
            sheet, addr = _split_ref(entry["cell"])
            fixes.append(CellFix(sheet, addr, entry["action"], entry["new_value"]))

    return fixes


def apply_fix(sheet_writer, fix: CellFix) -> None:
    """sheet_writer: objeto con .set_value(sheet, addr, value) y .set_formula(sheet, addr, formula)
    provisto por macros/uno_pipeline.py (o un doble de prueba). Mantiene este modulo libre de
    dependencias de UNO para poder testearlo sin LibreOffice instalado.
    """
    if fix.action == "freeze_static":
        sheet_writer.set_value(fix.sheet, fix.address, fix.value)
    elif fix.action == "replace_formula":
        if isinstance(fix.value, str) and fix.value.startswith("="):
            sheet_writer.set_formula(fix.sheet, fix.address, fix.value)
        else:
            sheet_writer.set_value(fix.sheet, fix.address, fix.value)
    else:
        raise ValueError(f"accion de fix desconocida: {fix.action!r}")
