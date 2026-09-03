"""Cliente UNO para LibreOffice headless: abre un .xls, permite leer/escribir celdas,
forzar un recalculo real y guardar. Es la pieza mas nueva y riesgosa del diseno (ver
Fase 0 del plan) - probar esto primero, aislado, antes de conectarlo al resto del pipeline.

Requiere una instancia de soffice ya escuchando en un socket, por ejemplo:
    soffice --headless --invisible --nologo --nofirststartwizard \
            --accept="socket,host=localhost,port=2002;urp;"

Uso tipico:
    with SheetWriter("templates/balance_mensual_ars.xls") as wr:
        wr.clear_range("SALDOS", "A243:B360")
        wr.write_table("SALDOS", "A243", [("1110100330  FONDO FIJO BS. AS.", -265310.63), ...])
        wr.set_value("Hoja1", "C7", 293.1757)
        wr.recalc()
        errors = wr.scan_errors()
        wr.save_as("output/2026-07/balance_mensual_ars.xlsx", fmt="xlsx")
"""
from __future__ import annotations

import re
import subprocess
import time
from pathlib import Path

import uno
from com.sun.star.beans import PropertyValue

# ERROR.TYPE de Calc: numeros de error habituales que nos importan detectar.
# https://wiki.documentfoundation.org/Documentation/Calc_functions/ERROR.TYPE
KNOWN_ERROR_TEXTS = {"#REF!", "#VALUE!", "#NAME?", "#DIV/0!", "#N/A", "#NULL!", "#NUM!"}

CELL_REF_RE = re.compile(r"^([A-Z]+)(\d+)$")


def _prop(name: str, value) -> PropertyValue:
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def start_soffice(port: int = 2002, timeout: float = 30.0) -> subprocess.Popen:
    """Lanza soffice --headless escuchando en `port`. Devuelve el proceso (para poder matarlo).

    NOTA Fase 0: en el runner de GitHub Actions esto corre en Linux (apt install libreoffice-calc);
    en Windows local el ejecutable es distinto (ver LIBREOFFICE_BIN de entorno) - se resuelve el
    binario en orden: $LIBREOFFICE_BIN, luego 'soffice' en PATH, luego la ruta tipica de instalacion
    de Windows.
    """
    import os
    import shutil

    binary = os.environ.get("LIBREOFFICE_BIN")
    if not binary:
        binary = shutil.which("soffice") or shutil.which("soffice.exe")
    if not binary:
        for candidate in (
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ):
            if Path(candidate).exists():
                binary = candidate
                break
    if not binary:
        raise RuntimeError(
            "No se encontro el ejecutable de soffice. Fijar LIBREOFFICE_BIN o instalar LibreOffice."
        )

    proc = subprocess.Popen(
        [
            binary,
            "--headless",
            "--invisible",
            "--nologo",
            "--nofirststartwizard",
            "--norestore",
            f"--accept=socket,host=localhost,port={port};urp;",
        ]
    )

    deadline = time.time() + timeout
    last_err = None
    while time.time() < deadline:
        try:
            _connect(port)
            return proc
        except Exception as e:  # noqa: BLE001 - reintentamos hasta timeout
            last_err = e
            time.sleep(0.5)
    proc.kill()
    raise TimeoutError(f"soffice no respondio en {timeout}s en el puerto {port}: {last_err}")


def _connect(port: int):
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx
    )
    ctx = resolver.resolve(
        f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext"
    )
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    return desktop, ctx


class SheetWriter:
    """Envuelve un documento Calc abierto via UNO. No hace falta usar `with` si se prefiere
    llamar a .close() explicitamente (util para probar en una consola interactiva).
    """

    def __init__(self, path: str, port: int = 2002):
        self.path = Path(path).resolve()
        self.port = port
        self._desktop = None
        self._doc = None

    def __enter__(self) -> "SheetWriter":
        self.open()
        return self

    def __exit__(self, exc_type, exc, tb):
        self.close()

    def open(self) -> None:
        desktop, _ctx = _connect(self.port)
        self._desktop = desktop
        url = uno.systemPathToFileUrl(str(self.path))
        props = [
            _prop("Hidden", True),
            # CRITICO (ver plan, Fase 0): sin esto LibreOffice puede intentar resolver los
            # vinculos externos muertos (2004, 2012) al abrir y colgarse en un runner sin cabeza.
            _prop("UpdateDocMode", 0),  # com.sun.star.document.UpdateDocMode.NO_UPDATE
        ]
        self._doc = self._desktop.loadComponentFromURL(url, "_blank", 0, tuple(props))
        if self._doc is None:
            raise RuntimeError(f"No se pudo abrir {self.path}")

    def close(self) -> None:
        if self._doc is not None:
            self._doc.close(False)
            self._doc = None

    # -- acceso a celdas -----------------------------------------------------------------

    def _sheet(self, name: str):
        sheets = self._doc.getSheets()
        if not sheets.hasByName(name):
            raise KeyError(f"La hoja {name!r} no existe en {self.path.name}")
        return sheets.getByName(name)

    def _cell(self, sheet: str, address: str):
        m = CELL_REF_RE.match(address)
        if not m:
            raise ValueError(f"direccion de celda invalida: {address!r}")
        col_letters, row_str = m.groups()
        col = 0
        for ch in col_letters:
            col = col * 26 + (ord(ch) - ord("A") + 1)
        col -= 1
        row = int(row_str) - 1
        return self._sheet(sheet).getCellByPosition(col, row)

    def get_value(self, sheet: str, address: str) -> float:
        return self._cell(sheet, address).getValue()

    def get_formula(self, sheet: str, address: str) -> str:
        return self._cell(sheet, address).getFormula()

    def get_display_text(self, sheet: str, address: str) -> str:
        return self._cell(sheet, address).getString()

    def set_value(self, sheet: str, address: str, value: float) -> None:
        self._cell(sheet, address).setValue(float(value))

    def set_string(self, sheet: str, address: str, text: str) -> None:
        self._cell(sheet, address).setString(text)

    def set_formula(self, sheet: str, address: str, formula: str) -> None:
        self._cell(sheet, address).setFormula(formula)

    def clear_range(self, sheet: str, range_addr: str) -> None:
        """range_addr ej 'A243:B360' - borra contenido y formulas, no formato."""
        cell_range = self._sheet(sheet).getCellRangeByName(range_addr)
        # CONTENTS = VALUE | DATETIME | STRING | ANNOTATION | FORMULA (ver com.sun.star.sheet.CellFlags)
        cell_range.clearContents(1 | 2 | 4 | 16 | 32)

    def write_table(self, sheet: str, top_left: str, rows: list[tuple[str, float]]) -> None:
        """Escribe pares (texto, numero) empezando en top_left, una fila por par,
        texto en la columna de top_left y numero en la columna siguiente.
        """
        m = CELL_REF_RE.match(top_left)
        if not m:
            raise ValueError(f"direccion invalida: {top_left!r}")
        col_letters, row_str = m.groups()
        start_row = int(row_str)
        sh = self._sheet(sheet)
        for i, (text, value) in enumerate(rows):
            key_cell = self._cell(sheet, f"{col_letters}{start_row + i}")
            key_cell.setString(text)
            # columna siguiente (asume una sola letra de columna, valido para los 4 layouts confirmados)
            next_col = chr(ord(col_letters) + 1)
            val_cell = self._cell(sheet, f"{next_col}{start_row + i}")
            val_cell.setValue(float(value))

    # -- recalculo / guardado --------------------------------------------------------------

    def recalc(self) -> None:
        self._doc.calculateAll()

    def scan_errors(self) -> list[str]:
        """Recorre todas las hojas usadas buscando celdas con texto de error (#REF!, etc).
        Devuelve una lista de 'Hoja!Celda' - simple y suficiente para diffear contra
        known_baseline_errors en config/validations/checks.yml.
        """
        found: list[str] = []
        sheets = self._doc.getSheets()
        for i in range(sheets.getCount()):
            sh = sheets.getByIndex(i)
            cursor = sh.createCursor()
            cursor.gotoEndOfUsedArea(False)
            end = cursor.getRangeAddress()
            used = sh.getCellRangeByPosition(0, 0, end.EndColumn, end.EndRow)
            data = used.getDataArray()
            for r, row in enumerate(data):
                for c, val in enumerate(row):
                    if isinstance(val, str) and val.strip() in KNOWN_ERROR_TEXTS:
                        cell = used.getCellByPosition(c, r)
                        found.append(f"{sh.getName()}!{cell.AbsoluteName.split('.')[-1]}")
        return found

    def save_as(self, out_path: str, fmt: str = "xlsx") -> None:
        filter_name = {
            "xlsx": "Calc MS Excel 2007 XML",
            "xls": "MS Excel 97",
        }[fmt]
        out = Path(out_path).resolve()
        out.parent.mkdir(parents=True, exist_ok=True)
        url = uno.systemPathToFileUrl(str(out))
        self._doc.storeToURL(url, (_prop("FilterName", filter_name),))
