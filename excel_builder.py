"""Construcción del archivo Excel SIRE LyJ respetando la plantilla histórica.

Estructura reproducida:
- Hoja única "Hoja1", 80 columnas (RUC ... CLU39), fuente Calibri 8.
- Fila 1 de cabeceras en negrita, centradas, con ajuste de texto y alto 54.75.
- Filas de datos desde la fila 2. Fórmulas:
    B  = CONCATENATE(G,"-",H,"-",J)
    P  = O*0.18            (sólo cuando el usuario marca "IGV automático")
    Y  = SUM(O:X)          (sólo cuando el usuario marca "Total CP automático")
- Fila en blanco y luego fila de totales con =SUM(col2:col{última+1}) en O..Y,
  negrita, borde superior fino e inferior doble.
- Sección opcional de "mes anterior": fila en blanco, sus filas, fila en blanco
  y su propia fila de totales.
"""
from __future__ import annotations

import datetime as dt
import io
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import column_index_from_string, get_column_letter

HEADERS = [
    "RUC", "CONCATENADO", "Periodo", "REGISTRO", "Fecha de emisión", "Fecha Vcto/Pago",
    "Tipo CP/Doc.", "Serie del CDP", "Año", "Nro CP o Doc. Nro Inicial (Rango)",
    "Nro Final (Rango)", "Tipo Doc Identidad", "Nro Doc Identidad",
    "Apellidos Nombres/ Razón  Social", "BI Gravado DG", "IGV / IPM DG",
    "BI Gravado DGNG", "IGV / IPM DGNG", "BI Gravado DNG", "IGV / IPM DNG",
    "Valor Adq. NG", "ISC", "ICBPER", "Otros Trib/ Cargos", "Total CP", "Moneda",
    "Tipo de Cambio", "Fecha Emisión Doc Modificado", "Tipo CP Modificado",
    "Serie CP Modificado", "COD. DAM O DSI", "Nro CP Modificado", "CTA  DEBE",
    "CTA HABER", "DETALLE", "IMB", "CAR Orig/ Ind E o I", "Detracción", "Tipo de Nota",
    "Est. Comp.", "Incal",
] + [f"CLU{i}" for i in range(1, 40)]

assert len(HEADERS) == 80

COL_LETTERS = [get_column_letter(i) for i in range(1, len(HEADERS) + 1)]
HEADER_BY_COL = dict(zip(COL_LETTERS, HEADERS))

# Columnas numéricas con formato contable (O..Y)
NUM_COLS = ["O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y"]
# Columnas de importes editables del formulario (sin Total CP)
IMPORTE_COLS = ["O", "P", "Q", "R", "S", "T", "U", "V", "W", "X"]

# Columnas "extra" que casi nunca se llenan pero se mantienen en la plantilla.
EXTRA_COLS = ["I", "K", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AJ", "AK", "AL", "AM", "AN", "AO"]
CLU_COLS = COL_LETTERS[COL_LETTERS.index("AP"):]

ACC_FMT = '_ * #,##0.00_ ;_ * \\-#,##0.00_ ;_ * "-"??_ ;_ @_ '
DATE_FMT = "mm-dd-yy"  # formato de fecha corta del sistema (como en los originales)

WIDTHS = {
    "A": 15.3, "B": 17.7, "C": 7.7, "D": 9.9, "E": 10.7, "G": 7.6, "H": 5.3, "I": 5.0,
    "K": 7.1, "L": 8.7, "M": 15.0, "N": 42.9, "O": 12.0, "P": 12.0, "Q": 8.0, "U": 12.0,
    "Z": 4.7, "AA": 8.3, "AB": 9.6, "AD": 6.1, "AG": 9.7, "AI": 45.4, "AJ": 14.3,
}

FONT = Font(name="Calibri", size=8)
FONT_BOLD = Font(name="Calibri", size=8, bold=True)
THIN = Side(style="thin")
DOUBLE = Side(style="double")
TOTAL_BORDER = Border(top=THIN, bottom=DOUBLE)


# ---------------------------------------------------------------------------
# Utilidades de conversión
# ---------------------------------------------------------------------------
def _as_number_if_digits(value: Any):
    """Convierte '20519342961' en int; deja texto si no es sólo dígitos."""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return value
    s = str(value).strip()
    if s == "":
        return None
    if s.isdigit() and not (len(s) > 1 and s.startswith("0")):
        return int(s)
    return s


def _as_float(value: Any):
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _as_date(value: Any):
    if isinstance(value, dt.datetime):
        return value
    if isinstance(value, dt.date):
        return dt.datetime(value.year, value.month, value.day)
    return None


def igv_calculado(bi: float | None, tasa: float = 0.18) -> float:
    return round((bi or 0.0) * tasa, 2)


def total_calculado(fila: dict) -> float:
    return round(sum(_as_float(fila.get(c)) or 0.0 for c in IMPORTE_COLS), 2)


def valores_fila(fila: dict) -> dict:
    """Devuelve la fila con todos los valores resueltos (como los vería Excel).

    Se usa para la vista previa. ``fila`` es un dict con llaves = letras de columna
    más los flags ``_igv_auto`` y ``_total_auto``.
    """
    out = dict(fila)
    out["B"] = f'{fila.get("G") or ""}-{fila.get("H") or ""}-{fila.get("J") or ""}'
    if fila.get("_igv_auto"):
        out["P"] = igv_calculado(_as_float(fila.get("O")))
    if fila.get("_total_auto"):
        tmp = dict(out)
        out["Y"] = total_calculado(tmp)
    return out


def totales_seccion(filas: list[dict]) -> dict:
    tot = {}
    for c in NUM_COLS:
        tot[c] = round(sum(_as_float(valores_fila(f).get(c)) or 0.0 for f in filas), 2)
    return tot


# ---------------------------------------------------------------------------
# Construcción del libro
# ---------------------------------------------------------------------------
def _write_header(ws):
    for idx, title in enumerate(HEADERS, start=1):
        c = ws.cell(row=1, column=idx, value=title)
        c.font = FONT_BOLD
        c.alignment = Alignment(horizontal="center", vertical="bottom", wrap_text=True)
        letter = get_column_letter(idx)
        if letter in NUM_COLS:
            c.number_format = ACC_FMT
        elif letter == "G":
            c.number_format = "@"
    ws.row_dimensions[1].height = 54.75
    for letter, w in WIDTHS.items():
        ws.column_dimensions[letter].width = w


def _write_row(ws, r: int, fila: dict, tipo: str):
    """Escribe una fila de datos en la fila ``r`` de la hoja."""
    def put(col: str, value, fmt: str | None = None, center: bool = False):
        cell = ws[f"{col}{r}"]
        cell.value = value
        cell.font = FONT
        if fmt:
            cell.number_format = fmt
        if center:
            cell.alignment = Alignment(horizontal="center")
        return cell

    put("A", _as_number_if_digits(fila.get("A")), center=(tipo == "Venta"))
    put("B", f'=CONCATENATE(G{r},"-",H{r},"-",J{r})')
    put("C", fila.get("C") or None)
    put("D", fila.get("D") or None)
    put("E", _as_date(fila.get("E")), DATE_FMT)
    put("F", _as_date(fila.get("F")), DATE_FMT)
    put("G", str(fila.get("G") or ""), "@")
    put("H", (fila.get("H") or "").strip() or None)
    put("I", _as_number_if_digits(fila.get("I")))
    put("J", _as_number_if_digits(fila.get("J")), center=True)
    put("K", _as_number_if_digits(fila.get("K")))
    put("L", _as_number_if_digits(fila.get("L")))
    put("M", _as_number_if_digits(fila.get("M")), center=True)
    put("N", (fila.get("N") or "").strip() or None)

    # Importes. En Compras los ceros se dejan en blanco (como en los archivos
    # recientes); en Ventas se escriben como 0.
    for col in IMPORTE_COLS:
        if col == "P" and fila.get("_igv_auto"):
            put("P", f"=O{r}*0.18", ACC_FMT)
            continue
        v = _as_float(fila.get(col))
        if v == 0 and tipo != "Venta" and col not in ("O", "P"):
            v = None
        put(col, v, ACC_FMT)

    if fila.get("_total_auto"):
        put("Y", f"=SUM(O{r}:X{r})", ACC_FMT)
    else:
        put("Y", _as_float(fila.get("Y")), ACC_FMT)

    put("Z", (fila.get("Z") or "").strip() or None)
    put("AA", _as_float(fila.get("AA")))
    put("AI", (fila.get("AI") or "").strip() or None)

    for col in EXTRA_COLS + CLU_COLS:
        if col in ("I", "K"):
            continue
        v = fila.get(col)
        if v is None or str(v).strip() == "":
            continue
        put(col, _as_number_if_digits(v))


def _write_total(ws, r: int, first: int, last: int, bold: bool = True):
    """Fila de totales: =SUM(col{first}:col{last}) en O..Y."""
    for col in NUM_COLS:
        cell = ws[f"{col}{r}"]
        cell.value = f"=SUM({col}{first}:{col}{last})"
        cell.number_format = ACC_FMT
        cell.font = FONT_BOLD if bold else FONT
        cell.border = TOTAL_BORDER
    ws.row_dimensions[r].height = 10.8
    ws.row_dimensions[r + 1].height = 10.8


def build_workbook(documento: dict) -> bytes:
    """Genera el .xlsx y devuelve sus bytes.

    ``documento`` = {
        "tipo": "Compra" | "Venta",
        "secciones": [ [fila, fila, ...], [fila, ...] ]   # 1 o 2 secciones
    }
    Cada fila es un dict con llaves = letra de columna (A, C, D, E, ...) y los
    flags ``_igv_auto`` / ``_total_auto``.
    """
    tipo = documento.get("tipo", "Compra")
    secciones = [s for s in documento.get("secciones", []) if s]

    wb = Workbook()
    ws = wb.active
    ws.title = "Hoja1"
    _write_header(ws)

    r = 2
    for i, filas in enumerate(secciones):
        first = r
        for fila in filas:
            _write_row(ws, r, fila, tipo)
            r += 1
        last = r          # la fila en blanco también entra al rango SUM (como en los originales)
        r += 1            # fila en blanco
        _write_total(ws, r, first, last, bold=True)
        r += 2            # fila de total + fila en blanco antes de la siguiente sección

    ws.freeze_panes = "AI2" if tipo == "Venta" else "O2"
    ws.sheet_view.zoomScale = 120

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
