"""Orquestador principal. Uso:

    python -m src.pipeline --period 2026-07 --tc-cierre 293.1757 \
        --target balance_mensual_ars [--dif-cambio-mes 3231.78]

Por cada archivo objetivo: copia el template, mapea las cuentas de Onvio contra el
plan de cuentas del template (por codigo, ver ingest_onvio.match_against_template),
pega en el staging, escribe TC CIERRE donde corresponda, aplica los fixes aprobados,
recalcula con LibreOffice, valida, y guarda el resultado + un resumen en markdown.
"""
from __future__ import annotations

import argparse
import calendar
import shutil
import sys
from pathlib import Path

import yaml

from src.apply_fixes import apply_fix, load_fixes_for_file
from src.ingest_onvio import match_against_template, parse_sumas_y_saldos
from src.report_summary import render_full_summary
from src.validate import build_report, run_checks, scan_new_errors

REPO_ROOT = Path(__file__).resolve().parent.parent
LAYOUTS_DIR = REPO_ROOT / "config" / "layouts"
FIXES_DIR = REPO_ROOT / "config" / "fixes"
VALIDATIONS_DIR = REPO_ROOT / "config" / "validations"

MONTH_NAMES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}


def _load_yaml(path: Path) -> dict:
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f)


def _load_all_layouts() -> dict[str, dict]:
    return {p.stem: _load_yaml(p) for p in LAYOUTS_DIR.glob("*.yml")}


def _read_template_keys(sheet_writer, layout: dict) -> dict[str, str]:
    """Lee la columna clave del template (ej SALDOS!A7:A241) y arma {codigo: texto_exacto}."""
    import re

    from macros.uno_pipeline import CELL_REF_RE

    saldos = layout["saldos"]
    sheet = saldos["sheet"]
    key_range = saldos["template_key_rows"]  # ej "A7:A241"
    start, end = key_range.split(":")
    m_start, m_end = CELL_REF_RE.match(start), CELL_REF_RE.match(end)
    col = m_start.group(1)
    row_start, row_end = int(m_start.group(2)), int(m_end.group(2))

    code_re = re.compile(r"^\s*([0-9]{6,})")
    out: dict[str, str] = {}
    for row in range(row_start, row_end + 1):
        text = sheet_writer.get_display_text(sheet, f"{col}{row}")
        if not text.strip():
            continue
        mm = code_re.match(text)
        if mm:
            out[mm.group(1)] = text
    return out


def process_file(
    file_id: str,
    layout: dict,
    period: str,
    tc_cierre: float,
    dif_cambio_mes: float | None,
    fixes_approved: dict,
    checks_config: dict,
    inputs_dir: Path,
    output_dir: Path,
    port: int,
) -> "ValidationReportLike":
    from macros.uno_pipeline import SheetWriter  # import tardio: no hace falta UNO para dry-run/tests

    year, month = period.split("-")
    month_name = MONTH_NAMES_ES[int(month)]

    template_path = REPO_ROOT / layout["template"]
    work_path = output_dir / f"{file_id}.xls"
    output_dir.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(template_path, work_path)

    onvio_filename = "sumas_y_saldos_mensual.xls" if layout["source_period"] == "mensual" else "sumas_y_saldos_acumulado.xls"
    onvio_path = inputs_dir / period / onvio_filename
    if not onvio_path.exists():
        raise FileNotFoundError(f"No se encontro el export de Onvio: {onvio_path}")

    accounts = parse_sumas_y_saldos(str(onvio_path), layout["source_currency"])

    with SheetWriter(str(work_path), port=port) as wr:
        template_keys = _read_template_keys(wr, layout)
        matched, unmapped = match_against_template(accounts, template_keys)

        saldos = layout["saldos"]
        wr.clear_range(saldos["sheet"], saldos["staging_range"])
        rows = [(text, value) for text, value in matched.items()]
        top_left = saldos["staging_range"].split(":")[0]
        wr.write_table(saldos["sheet"], top_left, rows)

        for target in layout.get("tc_cierre_targets", []):
            sheet, addr = target["cell"].split("!", 1)
            wr.set_value(sheet, addr, tc_cierre)

        dif_schedule = layout.get("dif_cambio_schedule")
        if dif_schedule and dif_cambio_mes is not None:
            cell = dif_schedule["cells_by_month"].get(int(month))
            if cell is None:
                raise ValueError(
                    f"{file_id}: no hay celda mapeada para el mes {month} en dif_cambio_schedule "
                    f"(ver la nota en config/layouts/{file_id}.yml antes de forzar esto)"
                )
            wr.set_value(dif_schedule["sheet"], cell, dif_cambio_mes)

        for fix in load_fixes_for_file(fixes_approved, file_id):
            apply_fix(wr, fix)

        wr.recalc()

        checks = run_checks(wr, file_id, checks_config)
        new_errors, known_errors = scan_new_errors(wr, file_id, checks_config)
        report = build_report(file_id, checks, unmapped, new_errors, known_errors)

        out_name = layout["output_name_pattern"].format(month_name=month_name, year=year)
        wr.save_as(str(output_dir / out_name.replace(".xls", ".xlsx")), fmt="xlsx")

    return report


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Automatizacion del cierre mensual TFBR")
    ap.add_argument("--period", required=True, help="YYYY-MM, ej 2026-07")
    ap.add_argument("--tc-cierre", type=float, required=True)
    ap.add_argument("--dif-cambio-mes", type=float, default=None)
    ap.add_argument("--target", action="append", dest="targets", help="id de layout (repetible). Default: todos")
    ap.add_argument("--inputs-dir", default=str(REPO_ROOT / "inputs"))
    ap.add_argument("--output-dir", default=str(REPO_ROOT / "output"))
    ap.add_argument("--port", type=int, default=2002)
    args = ap.parse_args(argv)

    layouts = _load_all_layouts()
    targets = args.targets or list(layouts.keys())

    fixes_approved = _load_yaml(FIXES_DIR / "approved.yml")
    checks_config = _load_yaml(VALIDATIONS_DIR / "checks.yml")

    reports = []
    any_failed = False
    for file_id in targets:
        if file_id not in layouts:
            print(f"ERROR: no existe config/layouts/{file_id}.yml", file=sys.stderr)
            return 2
        report = process_file(
            file_id=file_id,
            layout=layouts[file_id],
            period=args.period,
            tc_cierre=args.tc_cierre,
            dif_cambio_mes=args.dif_cambio_mes,
            fixes_approved=fixes_approved,
            checks_config=checks_config,
            inputs_dir=Path(args.inputs_dir),
            output_dir=Path(args.output_dir) / args.period,
            port=args.port,
        )
        reports.append(report)
        any_failed = any_failed or not report.passed

    summary_md = render_full_summary(reports)
    summary_path = Path(args.output_dir) / args.period / "SUMMARY.md"
    summary_path.parent.mkdir(parents=True, exist_ok=True)
    summary_path.write_text(summary_md, encoding="utf-8")
    print(summary_md)

    return 1 if any_failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
