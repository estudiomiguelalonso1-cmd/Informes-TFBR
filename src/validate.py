"""Corre los checks de config/validations/checks.yml contra un archivo ya recalculado,
mas la deteccion de cuentas sin mapear (ver ingest_onvio.match_against_template) y el
escaneo de errores nuevos vs conocidos. No decide que hacer con el resultado (eso es
pipeline.py) - solo junta hechos.
"""
from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class CheckResult:
    description: str
    cell: str
    expected: float
    actual: float | None
    passed: bool


@dataclass
class ValidationReport:
    file_id: str
    checks: list[CheckResult] = field(default_factory=list)
    unmapped_accounts: list[str] = field(default_factory=list)
    new_errors: list[str] = field(default_factory=list)
    known_errors: list[str] = field(default_factory=list)

    @property
    def passed(self) -> bool:
        return (
            all(c.passed for c in self.checks)
            and not self.unmapped_accounts
            and not self.new_errors
        )


def run_checks(sheet_writer, file_id: str, checks_config: dict, tolerance: float = 0.01) -> list[CheckResult]:
    results: list[CheckResult] = []
    for check in checks_config.get(file_id, []):
        cell = check["cell"]
        if cell == "TODO":
            # layout todavia no confirmado para este archivo - no se puede chequear, se avisa
            results.append(CheckResult(check["description"], cell, check["expect"], None, False))
            continue
        sheet, addr = cell.split("!", 1)
        actual = sheet_writer.get_value(sheet, addr)
        expected = check["expect"]
        results.append(
            CheckResult(check["description"], cell, expected, actual, abs(actual - expected) <= tolerance)
        )
    return results


def scan_new_errors(sheet_writer, file_id: str, checks_config: dict) -> tuple[list[str], list[str]]:
    """Devuelve (errores_nuevos, errores_conocidos_preexistentes)."""
    baseline = set(checks_config.get("known_baseline_errors", {}).get(file_id, []))
    found = set(sheet_writer.scan_errors())
    new = sorted(found - baseline)
    known = sorted(found & baseline)
    return new, known


def build_report(
    file_id: str,
    checks: list[CheckResult],
    unmapped_accounts: list[str],
    new_errors: list[str],
    known_errors: list[str],
) -> ValidationReport:
    return ValidationReport(
        file_id=file_id,
        checks=checks,
        unmapped_accounts=unmapped_accounts,
        new_errors=new_errors,
        known_errors=known_errors,
    )
