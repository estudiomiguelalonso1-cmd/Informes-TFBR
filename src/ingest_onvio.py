"""Parsea los exports "Sumas y Saldos" de Onvio a un diccionario {codigo: saldo}.

Los exports vienen en dos hojas historicamente vistas en los archivos reales:
  - "Sheet1": la exportacion cruda, sin formulas. Fila de headers en la fila 10 (indice 9,
    0-based), columna A = "Cuenta - Denominacion" con formato "<codigo> - <descripcion>",
    columnas D/G = "Saldo ($)" / "Saldo (R$)".
  - "Hoja1": hoja auxiliar con el TC CIERRE tipeado a mano en C7 (NO C6 - ver docs/formula_analysis.md,
    esa celda solo tiene la etiqueta de texto). Mismo formato de cuenta que Sheet1.

Extraemos el codigo de cuenta con una regex al inicio del string ("^\\s*(\\d+)") en vez de
confiar en el formato completo del texto - ver docs/formula_analysis.md, seccion de riesgo de
matching, sobre por que el texto completo de Onvio NO se debe pegar tal cual en el staging de
SALDOS (el VLOOKUP del template espera el texto EXACTO que el template ya tiene, no el de Onvio).
"""
from __future__ import annotations

import re
from dataclasses import dataclass

import pandas as pd

ACCOUNT_CODE_RE = re.compile(r"^\s*([0-9]{6,})")

# Ubicacion confirmada en Sheet1 (ver docs/formula_analysis.md): fila de headers en la fila 10
# (1-based Excel), datos desde la fila 11. Columnas: A=Cuenta, B=Debe($), C=Haber($), D=Saldo($),
# E=Debe(R$), F=Haber(R$), G=Saldo(R$).
SHEET1_HEADER_ROW_0BASED = 9
SHEET1_ACCOUNT_COL = 0
SHEET1_SALDO_COL = {"ARS": 3, "BRL": 6}


@dataclass(frozen=True)
class AccountBalance:
    code: str
    raw_label: str
    balance: float


def _extract_code(raw_label: str) -> str | None:
    m = ACCOUNT_CODE_RE.match(str(raw_label))
    return m.group(1) if m else None


def parse_sumas_y_saldos(xls_path: str, currency: str) -> dict[str, AccountBalance]:
    """Lee Sheet1 del export de Onvio y devuelve {codigo: AccountBalance}.

    currency: "ARS" o "BRL" - selecciona la columna Saldo($) o Saldo(R$).
    Cuentas con saldo 0 se incluyen igual (el template las necesita para no dejar
    huecos si un mes tienen movimiento y el siguiente no).
    """
    if currency not in SHEET1_SALDO_COL:
        raise ValueError(f"currency debe ser ARS o BRL, recibido: {currency!r}")

    df = pd.read_excel(xls_path, sheet_name="Sheet1", header=None, engine="xlrd")
    saldo_col = SHEET1_SALDO_COL[currency]

    out: dict[str, AccountBalance] = {}
    skipped_no_code: list[str] = []

    for i in range(SHEET1_HEADER_ROW_0BASED + 1, len(df)):
        raw_label = df.iat[i, SHEET1_ACCOUNT_COL] if SHEET1_ACCOUNT_COL < df.shape[1] else None
        if raw_label is None or (isinstance(raw_label, float) and pd.isna(raw_label)):
            continue
        raw_label = str(raw_label).strip()
        if not raw_label or raw_label.upper() in {"ACTIVO", "PASIVO", "RESULTADOS", "TOTALES GENERALES:"}:
            continue

        code = _extract_code(raw_label)
        if code is None:
            # fila de subtotal / seccion / texto libre - no es una cuenta, se ignora
            skipped_no_code.append(raw_label)
            continue

        raw_balance = df.iat[i, saldo_col] if saldo_col < df.shape[1] else None
        try:
            balance = float(raw_balance)
        except (TypeError, ValueError):
            continue

        if code in out:
            raise ValueError(
                f"Codigo de cuenta duplicado en el export: {code!r} "
                f"(filas con '{out[code].raw_label}' y '{raw_label}')"
            )
        out[code] = AccountBalance(code=code, raw_label=raw_label, balance=balance)

    return out


def match_against_template(
    accounts: dict[str, AccountBalance],
    template_keys: dict[str, str],
) -> tuple[dict[str, float], list[str]]:
    """Cruza las cuentas del export contra las claves del template de SALDOS.

    template_keys: {codigo: texto_exacto_de_la_fila_del_template} (ver ingest del layout,
    se arma leyendo SALDOS!<template_key_rows> del archivo destino).

    Devuelve: (matched, unmatched_codes)
      matched: {texto_exacto_del_template: saldo} - listo para pegar en el staging_range,
        garantizando que el VLOOKUP del template va a encontrar la fila.
      unmatched_codes: codigos del export sin fila correspondiente en el template - deben
        reportarse en el resumen de validacion como "cuenta sin mapear", nunca descartarse
        en silencio.
    """
    matched: dict[str, float] = {}
    unmatched: list[str] = []

    for code, acc in accounts.items():
        template_text = template_keys.get(code)
        if template_text is None:
            unmatched.append(code)
            continue
        matched[template_text] = acc.balance

    return matched, unmatched
