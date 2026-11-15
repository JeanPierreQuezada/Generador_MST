"""Generador de Archivo Contable - SIRE LyJ (Compras / Ventas).

Aplicación Streamlit que guía al usuario en 3 pasos:
  1. Tipo de registro, periodo y datos generales.
  2. Carga de comprobantes uno por uno (con valores por defecto editables).
  3. Vista previa, confirmación y descarga del Excel.
"""
from __future__ import annotations

import datetime as dt
import re

import pandas as pd
import streamlit as st

from catalogo import (
    MESES, MONEDAS, RUC_PROPIO, SERIES_CONOCIDAS, TIPOS_CP, TIPOS_CP_DESC,
    TIPOS_DOC, TIPOS_DOC_DESC, buscar_tercero, catalogo_por_tipo,
)
from excel_builder import (
    CLU_COLS, EXTRA_COLS, HEADER_BY_COL, HEADERS, COL_LETTERS, IMPORTE_COLS,
    NUM_COLS, build_workbook, igv_calculado, total_calculado, totales_seccion,
    valores_fila,
)

# ---------------------------------------------------------------------------
# Configuración y estilos
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Generador de Archivo Contable",
    page_icon="📗",
    layout="wide",
    initial_sidebar_state="collapsed",
)

CSS = """
<style>
:root {
  --verde-900: #1B5E20; --verde-700: #2E7D32; --verde-500: #43A047;
  --verde-100: #E8F5E9; --verde-50: #F2F8F3; --gris: #5F6B64; --borde: #D9E7DC;
}
.block-container { padding-top: 1.6rem; padding-bottom: 3rem; max-width: 1200px; }
h1, h2, h3 { color: var(--verde-900); }
.app-title { display:flex; align-items:center; gap:.7rem; margin-bottom:.2rem; }
.app-title h1 { margin:0; font-size:1.75rem; }
.app-sub { color: var(--gris); margin: 0 0 1.2rem 0; }
.stepper { display:flex; gap:.5rem; margin: .4rem 0 1.4rem 0; flex-wrap: wrap; }
.step { flex:1; min-width: 180px; padding:.55rem .9rem; border-radius: 10px;
        border:1px solid var(--borde); background:#fff; color:var(--gris); font-size:.9rem; }
.step b { display:block; font-size:.78rem; letter-spacing:.04em; text-transform:uppercase; }
.step.active { background: var(--verde-700); color:#fff; border-color: var(--verde-700); }
.step.done { background: var(--verde-100); color: var(--verde-900); border-color: var(--verde-100); }
.card { border:1px solid var(--borde); border-radius: 12px; padding: 1rem 1.1rem; background:#fff; }
.hint { color: var(--gris); font-size:.85rem; }
.badge { display:inline-block; background:var(--verde-100); color:var(--verde-900);
         border-radius:999px; padding:.1rem .6rem; font-size:.78rem; font-weight:600; }
.fila-item { font-size:.9rem; line-height:1.35; margin-bottom:.35rem; word-break: break-word; }
.fila-item small { color: var(--gris); }
.fila-item.editando { color: var(--verde-900); }
.total-box { background: var(--verde-50); border:1px solid var(--borde); border-radius: 10px;
             padding:.6rem .9rem; margin-top:.4rem; word-break: break-word; }
div[data-testid="stToggle"] label p, div[data-testid="stCheckbox"] label p { font-size:.9rem; }

/* Botones primarios: verde; deshabilitados: verde claro pero legible */
button[data-testid="stBaseButton-primary"], button[kind="primary"] {
    background: var(--verde-700) !important; border-color: var(--verde-700) !important; color:#fff !important; }
button[data-testid="stBaseButton-primary"]:disabled, button[kind="primary"]:disabled {
    background: #C8E6C9 !important; border-color: #C8E6C9 !important; color: var(--verde-900) !important;
    opacity: 1 !important; cursor: not-allowed; }
button[data-testid="stBaseButton-primary"]:disabled p, button[kind="primary"]:disabled p {
    color: var(--verde-900) !important; }
div[data-testid="stDownloadButton"] button { background: var(--verde-900) !important; color:#fff !important;
    border-color: var(--verde-900) !important; font-weight:600; }
button p { white-space: normal; }

/* Móvil */
@media (max-width: 768px) {
  .block-container { padding-left: .9rem; padding-right: .9rem; padding-top: 1rem; }
  .app-title h1 { font-size:1.3rem; }
  .app-sub { font-size:.88rem; }
  .stepper { gap:.35rem; }
  .step { min-width: 0; padding:.4rem .5rem; font-size:.78rem; }
  .step b { font-size:.66rem; }
  h2 { font-size:1.25rem !important; }
  h3 { font-size:1.05rem !important; }
  div[data-testid="stNumberInput"] button { min-width: 2rem; }
}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

HOY = dt.date.today()


# ---------------------------------------------------------------------------
# Estado
# ---------------------------------------------------------------------------
def mes_anterior(mes_idx: int, anio: int) -> tuple[int, int]:
    """mes_idx 1..12 -> (mes, año) del mes previo."""
    return (12, anio - 1) if mes_idx == 1 else (mes_idx - 1, anio)


def init_state():
    ss = st.session_state
    if "inicializado" in ss:
        return
    ss.inicializado = True
    ss.paso = 1
    ss.tipo = "Compra"
    ss.mes = MESES[HOY.month - 1]
    ss.anio = HOY.year
    ss.ruc_propio = RUC_PROPIO
    ss.editar_ruc = False
    ss.prefijo_auto = True
    ss.prefijo = "01"
    ss.incluir_anterior = False
    pm, pa = mes_anterior(HOY.month, HOY.year)
    ss.mes_ant = MESES[pm - 1]
    ss.anio_ant = pa
    ss.filas = []              # lista de dicts (letras de columna + flags + "_seccion")
    ss.editando = None         # índice de la fila en edición
    ss.descargado = False
    ss.nombre_archivo = ""
    reset_form(conservar_tercero=False)


def mes_idx(nombre: str) -> int:
    return MESES.index(nombre) + 1


def periodo_de(mes_nombre: str, anio: int) -> str:
    return f"{int(anio)}-{mes_idx(mes_nombre):02d}"


def prefijo_actual() -> str:
    ss = st.session_state
    if ss.prefijo_auto:
        return "01" if ss.tipo == "Compra" else f"{mes_idx(ss.mes):02d}"
    return (ss.prefijo or "01").strip()


def siguiente_registro(seccion: str, excluir_idx: int | None = None) -> str:
    ss = st.session_state
    n = sum(1 for i, f in enumerate(ss.filas) if f["_seccion"] == seccion and i != excluir_idx)
    return f"{prefijo_actual()}-{n + 1:04d}"


def periodo_seccion(seccion: str) -> str:
    ss = st.session_state
    if seccion == "anterior":
        return periodo_de(ss.mes_ant, ss.anio_ant)
    return periodo_de(ss.mes, ss.anio)


def fecha_por_defecto(seccion: str) -> dt.date:
    ss = st.session_state
    if seccion == "anterior":
        return dt.date(int(ss.anio_ant), mes_idx(ss.mes_ant), 1)
    m, a = mes_idx(ss.mes), int(ss.anio)
    if HOY.year == a and HOY.month == m:
        return HOY
    return dt.date(a, m, 1)


def reset_form(conservar_tercero: bool = True, seccion: str | None = None):
    """Vuelve a poner los valores por defecto en el formulario de comprobante."""
    ss = st.session_state
    es_venta = ss.tipo == "Venta"
    seccion = seccion or ("actual" if not ss.get("f_seccion") else ss.f_seccion)
    if seccion == "anterior" and not ss.incluir_anterior:
        seccion = "actual"
    terceros, detalles = catalogo_por_tipo(ss.tipo)
    primero = terceros[0] if terceros else {"nombre": "", "ruc": "", "series": [""], "detalle": ""}

    tercero_prev = ss.get("f_razon") if conservar_tercero else None
    ruc_prev = ss.get("f_ruc") if conservar_tercero else None
    serie_prev = ss.get("f_serie") if conservar_tercero else None

    ss.f_seccion = seccion
    ss.f_periodo = periodo_seccion(seccion)
    ss.f_registro = siguiente_registro(seccion)
    ss.f_fecha = fecha_por_defecto(seccion)
    ss.f_vcto_distinta = False
    ss.f_vcto = ss.f_fecha
    ss.f_tipo_cp = "01"
    ss.f_serie = serie_prev if serie_prev is not None else (primero["series"][0] if primero["series"] else "")
    ss.f_numero = ""
    ss.f_tipo_doc = "6"
    ss.f_razon = tercero_prev if tercero_prev is not None else primero["nombre"]
    ss.f_ruc = ruc_prev if ruc_prev is not None else primero["ruc"]
    ss.f_bi = 0.0 if es_venta else 84.75
    ss.f_igv_auto = es_venta
    ss.f_igv = 0.0 if es_venta else 15.25
    for c in ["Q", "R", "S", "T", "U", "V", "W", "X"]:
        ss[f"f_{c}"] = 0.0
    ss.f_total_auto = es_venta          # Compras: 100 fijo editable; Ventas: fórmula =SUM(O:X)
    ss.f_total = 0.0 if es_venta else 100.0
    ss.f_moneda = "PEN"
    ss.f_tc = 1.0
    ss.f_detalle = "" if es_venta else (primero.get("detalle") or "Combustible")
    ss.f_detalle_sug = None
    ss.f_usar_comision = es_venta
    ss.f_com_ini = ss.f_fecha - dt.timedelta(days=180)
    ss.f_com_fin = ss.f_fecha
    for c in EXTRA_COLS + CLU_COLS:
        ss[f"f_ex_{c}"] = ""
    ss.editando = None


def cargar_fila_en_form(idx: int):
    """Carga una fila ya registrada en el formulario para editarla."""
    ss = st.session_state
    f = ss.filas[idx]
    ss.f_seccion = f["_seccion"]
    ss.f_periodo = f.get("C", "")
    ss.f_registro = f.get("D", "")
    ss.f_fecha = f.get("E") or HOY
    ss.f_vcto = f.get("F") or ss.f_fecha
    ss.f_vcto_distinta = ss.f_vcto != ss.f_fecha
    ss.f_tipo_cp = f.get("G", "01")
    ss.f_serie = f.get("H", "")
    ss.f_numero = str(f.get("J", "") or "")
    ss.f_tipo_doc = str(f.get("L", "6"))
    ss.f_razon = f.get("N", "")
    ss.f_ruc = str(f.get("M", "") or "")
    ss.f_bi = float(f.get("O") or 0.0)
    ss.f_igv_auto = bool(f.get("_igv_auto"))
    ss.f_igv = float(f.get("P") or 0.0)
    for c in ["Q", "R", "S", "T", "U", "V", "W", "X"]:
        ss[f"f_{c}"] = float(f.get(c) or 0.0)
    ss.f_total_auto = bool(f.get("_total_auto"))
    ss.f_total = float(f.get("Y") or 0.0)
    ss.f_moneda = f.get("Z", "PEN")
    ss.f_tc = float(f.get("AA") or 1.0)
    ss.f_detalle = f.get("AI", "")
    ss.f_detalle_sug = None
    ss.f_usar_comision = False
    for c in EXTRA_COLS + CLU_COLS:
        ss[f"f_ex_{c}"] = str(f.get(c, "") or "")
    ss.editando = idx


def fila_desde_form() -> dict:
    ss = st.session_state
    fecha = ss.f_fecha
    vcto = ss.f_vcto if ss.f_vcto_distinta else fecha
    fila = {
        "_seccion": ss.f_seccion,
        "_igv_auto": bool(ss.f_igv_auto),
        "_total_auto": bool(ss.f_total_auto),
        "A": ss.ruc_propio.strip(),
        "C": ss.f_periodo.strip(),
        "D": ss.f_registro.strip(),
        "E": fecha,
        "F": vcto,
        "G": (ss.f_tipo_cp or "01").strip(),
        "H": (ss.f_serie or "").strip().upper(),
        "J": (ss.f_numero or "").strip(),
        "L": (ss.f_tipo_doc or "6").strip(),
        "M": (ss.f_ruc or "").strip(),
        "N": (ss.f_razon or "").strip(),
        "O": float(ss.f_bi or 0.0),
        "P": float(ss.f_igv or 0.0),
        "Y": float(ss.f_total or 0.0),
        "Z": (ss.f_moneda or "PEN").strip(),
        "AA": float(ss.f_tc or 1.0),
        "AI": detalle_final().strip(),
    }
    for c in ["Q", "R", "S", "T", "U", "V", "W", "X"]:
        fila[c] = float(ss[f"f_{c}"] or 0.0)
    for c in EXTRA_COLS + CLU_COLS:
        v = (ss.get(f"f_ex_{c}") or "").strip()
        if v:
            fila[c] = v
    if fila["_igv_auto"]:
        fila["P"] = igv_calculado(fila["O"])
    if fila["_total_auto"]:
        fila["Y"] = total_calculado(fila)
    return fila


def detalle_final() -> str:
    ss = st.session_state
    if ss.tipo == "Venta" and ss.get("f_usar_comision"):
        return f"Comisión por ventas del {ss.f_com_ini:%d/%m/%y} al {ss.f_com_fin:%d/%m/%y}"
    return ss.get("f_detalle") or ""


def validar_form() -> list[str]:
    ss = st.session_state
    errores = []
    if not (ss.f_serie or "").strip():
        errores.append("Falta la **Serie del CDP**.")
    if not (ss.f_numero or "").strip():
        errores.append("Falta el **Nro CP** (número de comprobante).")
    if not (ss.f_razon or "").strip():
        errores.append("Falta el **Apellidos Nombres / Razón Social**.")
    ruc = (ss.f_ruc or "").strip()
    if not ruc:
        errores.append("Falta el **Nro Doc Identidad** (RUC del tercero).")
    elif ss.f_tipo_doc == "6" and not re.fullmatch(r"\d{11}", ruc):
        errores.append("El **RUC del tercero** debe tener 11 dígitos.")
    if not re.fullmatch(r"\d{4}-\d{2}", (ss.f_periodo or "").strip()):
        errores.append("El **Periodo** debe tener el formato AAAA-MM.")
    if not (ss.f_registro or "").strip():
        errores.append("Falta el **REGISTRO**.")
    return errores


# --- callbacks -------------------------------------------------------------
def cb_guardar_fila():
    ss = st.session_state
    errores = validar_form()
    if errores:
        ss.form_errores = errores
        return
    fila = fila_desde_form()
    if ss.editando is None:
        ss.filas.append(fila)
        ss.flash = f"Comprobante {fila['D']} agregado."
    else:
        ss.filas[ss.editando] = fila
        ss.flash = f"Comprobante {fila['D']} actualizado."
    ss.form_errores = []
    ss.descargado = False
    reset_form(conservar_tercero=True, seccion=fila["_seccion"])


def cb_editar(idx: int):
    st.session_state.form_errores = []
    cargar_fila_en_form(idx)


def cb_eliminar(idx: int):
    ss = st.session_state
    ss.filas.pop(idx)
    ss.descargado = False
    ss.form_errores = []
    reset_form(conservar_tercero=True)


def cb_cancelar_edicion():
    st.session_state.form_errores = []
    reset_form(conservar_tercero=True)


def cb_cambio_tercero():
    """Al elegir un tercero conocido, autocompleta RUC, serie y detalle."""
    ss = st.session_state
    t = buscar_tercero(ss.f_razon, ss.tipo)
    if t:
        ss.f_ruc = t["ruc"]
        if t["series"]:
            ss.f_serie = t["series"][0]
        if t.get("detalle") and ss.tipo == "Compra":
            ss.f_detalle = t["detalle"]


def cb_cambio_seccion():
    ss = st.session_state
    ss.f_periodo = periodo_seccion(ss.f_seccion)
    ss.f_registro = siguiente_registro(ss.f_seccion, excluir_idx=ss.editando)
    ss.f_fecha = fecha_por_defecto(ss.f_seccion)
    if not ss.f_vcto_distinta:
        ss.f_vcto = ss.f_fecha


def cb_sugerencia_detalle():
    """Copia el detalle frecuente elegido al campo de texto."""
    ss = st.session_state
    if ss.get("f_detalle_sug"):
        ss.f_detalle = ss.f_detalle_sug
    ss.f_detalle_sug = None


def cb_descargado():
    st.session_state.descargado = True


def cb_reiniciar():
    """Borra todo el estado y vuelve al paso 1 con los valores por defecto."""
    for k in list(st.session_state.keys()):
        del st.session_state[k]


def cb_cambio_fecha():
    ss = st.session_state
    if not ss.f_vcto_distinta:
        ss.f_vcto = ss.f_fecha
    if ss.tipo == "Venta":
        ss.f_periodo = f"{ss.f_fecha.year}-{ss.f_fecha.month:02d}"


def cb_ir_paso2():
    ss = st.session_state
    ss.form_errores = []
    cambio_tipo = ss.get("tipo_en_filas") not in (None, ss.tipo)
    if cambio_tipo and ss.filas:
        ss.filas = []
    ss.tipo_en_filas = ss.tipo
    if not ss.incluir_anterior:
        ss.filas = [f for f in ss.filas if f["_seccion"] == "actual"]
    # Actualiza el periodo/registro de las filas existentes si cambió el periodo
    for i, f in enumerate(ss.filas):
        if f["_seccion"] == "actual" and ss.tipo == "Compra":
            f["C"] = periodo_seccion("actual")
        elif f["_seccion"] == "anterior":
            f["C"] = periodo_seccion("anterior")
        f["A"] = ss.ruc_propio.strip()
    reset_form(conservar_tercero=bool(ss.filas))
    ss.paso = 2


def cb_ir_paso(n: int):
    st.session_state.paso = n


def nombre_archivo_defecto() -> str:
    ss = st.session_state
    base = "SIRE_LyJ" if ss.tipo == "Compra" else "SIRE_LyJ_Ventas"
    return f"{base}_{ss.mes}_{int(ss.anio)}"


# ---------------------------------------------------------------------------
# Componentes de interfaz
# ---------------------------------------------------------------------------
def stepper(paso: int):
    pasos = [("1", "Tipo y periodo"), ("2", "Comprobantes"), ("3", "Vista previa y descarga")]
    html = '<div class="stepper">'
    for i, (n, t) in enumerate(pasos, start=1):
        cls = "active" if i == paso else ("done" if i < paso else "")
        html += f'<div class="step {cls}"><b>Paso {n}</b>{t}</div>'
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)


def encabezado():
    st.markdown(
        '<div class="app-title"><span style="font-size:1.9rem">📗</span>'
        "<h1>Generador de Archivo Contable</h1></div>"
        '<p class="app-sub">Registro SIRE LyJ de Compras y Ventas: completa los comprobantes, '
        "revisa la vista previa y descarga el Excel con la plantilla habitual.</p>",
        unsafe_allow_html=True,
    )


def fmt_money(v) -> str:
    try:
        return f"{float(v):,.2f}"
    except (TypeError, ValueError):
        return "-"


# ---------------------------------------------------------------------------
# PASO 1
# ---------------------------------------------------------------------------
def paso_1():
    ss = st.session_state
    st.subheader("1. ¿Qué registro vas a generar?")
    ss.tipo = st.radio(
        "Tipo de registro", ["Compra", "Venta"], horizontal=True,
        index=0 if ss.tipo == "Compra" else 1, label_visibility="collapsed",
        captions=["Registro de compras (proveedores)", "Registro de ventas (clientes)"],
    )

    st.markdown("#### Periodo")
    c1, c2, c3 = st.columns([2, 1, 3])
    with c1:
        ss.mes = st.selectbox("Mes", MESES, index=MESES.index(ss.mes))
    with c2:
        ss.anio = st.number_input("Año", min_value=2000, max_value=2100, value=int(ss.anio), step=1, format="%d")
    with c3:
        st.markdown(
            f'<div class="total-box"><span class="hint">Archivo a generar</span><br>'
            f"<b>{nombre_archivo_defecto()}.xlsx</b><br>"
            f'<span class="hint">Periodo: <b>{periodo_de(ss.mes, ss.anio)}</b></span></div>',
            unsafe_allow_html=True,
        )

    st.markdown("#### Datos generales")
    c1, c2 = st.columns(2)
    with c1:
        ss.editar_ruc = st.toggle("Editar RUC propio", value=ss.editar_ruc,
                                  help="El RUC de la columna 'RUC' no suele cambiar. Actívalo sólo si necesitas otro.")
        ss.ruc_propio = st.text_input("RUC propio (columna RUC)", value=ss.ruc_propio,
                                      disabled=not ss.editar_ruc, max_chars=11)
    with c2:
        ss.prefijo_auto = st.toggle(
            "Prefijo de REGISTRO automático", value=ss.prefijo_auto,
            help="Compras: 01-0001, 01-0002… · Ventas: MM-0001, MM-0002… (MM = mes). "
                 "Desactívalo para escribir otro prefijo.",
        )
        ss.prefijo = st.text_input("Prefijo de REGISTRO", value=prefijo_actual(),
                                   disabled=ss.prefijo_auto, max_chars=4)
        st.caption(f"Ejemplo: **{prefijo_actual()}-0001**, **{prefijo_actual()}-0002**…")

    st.markdown("#### Comprobantes de un mes anterior")
    ss.incluir_anterior = st.toggle(
        "Agregar también comprobantes de un mes anterior (sección aparte después del total)",
        value=ss.incluir_anterior,
    )
    if ss.incluir_anterior:
        c1, c2, _ = st.columns([2, 1, 3])
        with c1:
            ss.mes_ant = st.selectbox("Mes anterior", MESES, index=MESES.index(ss.mes_ant))
        with c2:
            ss.anio_ant = st.number_input("Año anterior", min_value=2000, max_value=2100,
                                          value=int(ss.anio_ant), step=1, format="%d")
        st.caption("Esas filas se escriben debajo del total del mes con su propia fila de totales. "
                   "El nombre del archivo no cambia.")

    st.write("")
    st.button("Continuar a comprobantes ➜", type="primary", on_click=cb_ir_paso2, width="stretch")


# ---------------------------------------------------------------------------
# PASO 2
# ---------------------------------------------------------------------------
def formulario_comprobante():
    ss = st.session_state
    es_venta = ss.tipo == "Venta"
    terceros, detalles = catalogo_por_tipo(ss.tipo)
    editando = ss.editando is not None

    titulo = f"✏️ Editando comprobante {ss.filas[ss.editando]['D']}" if editando else "➕ Nuevo comprobante"
    st.markdown(f"#### {titulo}")

    if ss.incluir_anterior:
        st.radio(
            "Sección", ["actual", "anterior"], key="f_seccion", horizontal=True,
            format_func=lambda s: f"Mes actual ({periodo_seccion('actual')})" if s == "actual"
            else f"Mes anterior ({periodo_seccion('anterior')})",
            on_change=cb_cambio_seccion,
        )
    else:
        ss.f_seccion = "actual"

    # --- Fechas y registro
    c1, c2, c3 = st.columns([1.3, 1, 1])
    with c1:
        st.date_input("Fecha de emisión", key="f_fecha", format="DD/MM/YYYY", on_change=cb_cambio_fecha)
    with c2:
        st.text_input("Periodo", key="f_periodo", help="Formato AAAA-MM. Se completa con el mes y año elegidos.")
    with c3:
        st.text_input("REGISTRO", key="f_registro", help="Correlativo. Se propone el siguiente número automáticamente.")
    c1, c2 = st.columns([1.3, 2])
    with c1:
        st.toggle("Fecha Vcto/Pago distinta", key="f_vcto_distinta",
                  help="Por defecto la fecha de vencimiento/pago es igual a la de emisión.")
    with c2:
        if ss.f_vcto_distinta:
            st.date_input("Fecha Vcto/Pago", key="f_vcto", format="DD/MM/YYYY", label_visibility="collapsed")
        else:
            st.caption(f"Fecha Vcto/Pago = {ss.f_fecha:%d/%m/%Y} (igual a la emisión)")

    # --- Comprobante
    c1, c2, c3, c4 = st.columns([1.3, 1, 1, 1])
    with c1:
        opciones_cp = TIPOS_CP if ss.f_tipo_cp in TIPOS_CP else TIPOS_CP + [ss.f_tipo_cp]
        st.selectbox("Tipo CP/Doc.", opciones_cp, key="f_tipo_cp",
                     format_func=lambda c: TIPOS_CP_DESC.get(c, c), accept_new_options=True)
    with c2:
        opciones_serie = sorted(set(SERIES_CONOCIDAS + ([ss.f_serie] if ss.f_serie else [])))
        st.selectbox("Serie del CDP", opciones_serie, key="f_serie", accept_new_options=True,
                     placeholder="Elige o escribe la serie",
                     help="Series conocidas de los archivos históricos. Escribe una nueva y pulsa Enter para usarla.")
    with c3:
        st.text_input("Nro CP (Nro inicial)", key="f_numero", placeholder="p. ej. 7606")
    with c4:
        opciones_doc = TIPOS_DOC if ss.f_tipo_doc in TIPOS_DOC else TIPOS_DOC + [ss.f_tipo_doc]
        st.selectbox("Tipo Doc Identidad", opciones_doc, key="f_tipo_doc",
                     format_func=lambda c: TIPOS_DOC_DESC.get(c, c), accept_new_options=True)

    # --- Tercero
    c1, c2 = st.columns([2.2, 1])
    with c1:
        nombres = [t["nombre"] for t in terceros]
        if ss.f_razon and ss.f_razon not in nombres:
            nombres = nombres + [ss.f_razon]
        st.selectbox(
            "Apellidos Nombres / Razón Social", nombres, key="f_razon", accept_new_options=True,
            on_change=cb_cambio_tercero,
            placeholder="Elige un tercero conocido o escribe uno nuevo",
            help="Al elegir uno conocido se autocompleta el RUC (y la serie habitual).",
        )
    with c2:
        st.text_input("Nro Doc Identidad (RUC)", key="f_ruc", max_chars=11)

    # --- Importes
    st.markdown("##### Importes")
    st.toggle("IGV automático (18 % de la BI)", key="f_igv_auto",
              help="En el Excel se escribe la fórmula =BI*0.18. Desactívalo para escribir el IGV a mano.")
    ss.f_total_auto = es_venta   # Ventas: fórmula =SUM(O:X). Compras: valor fijo editable (100 por defecto).
    c1, c2, c3 = st.columns(3)
    with c1:
        st.number_input("BI Gravado DG", key="f_bi", min_value=0.0, step=0.01, format="%.2f")
    with c2:
        if ss.f_igv_auto:
            st.number_input("IGV / IPM DG (18 %)", value=igv_calculado(ss.f_bi), disabled=True,
                            format="%.2f", key="f_igv_show")
        else:
            st.number_input("IGV / IPM DG", key="f_igv", min_value=0.0, step=0.01, format="%.2f")
    with c3:
        if es_venta:
            prev = fila_desde_form()
            st.number_input("Total CP (fórmula)", value=prev["Y"], disabled=True, format="%.2f", key="f_total_show",
                            help="En Ventas se escribe la fórmula =SUM(BI … Otros) en el Excel.")
        else:
            st.number_input("Total CP", key="f_total", min_value=0.0, step=0.01, format="%.2f",
                            help="En Compras es un valor (100 por defecto). Edítalo si el comprobante es distinto.")
    c1, c2, _ = st.columns(3)
    with c1:
        opciones_mon = MONEDAS if ss.f_moneda in MONEDAS else MONEDAS + [ss.f_moneda]
        st.selectbox("Moneda", opciones_mon, key="f_moneda", accept_new_options=True)
    with c2:
        st.number_input("Tipo de Cambio", key="f_tc", min_value=0.0, step=0.001, format="%.3f")

    with st.expander("Otros importes (DGNG, DNG, NG, ISC, ICBPER, Otros)", expanded=False):
        cols = st.columns(4)
        etiquetas = [("Q", "BI Gravado DGNG"), ("R", "IGV / IPM DGNG"), ("S", "BI Gravado DNG"),
                     ("T", "IGV / IPM DNG"), ("U", "Valor Adq. NG"), ("V", "ISC"), ("W", "ICBPER"),
                     ("X", "Otros Trib/ Cargos")]
        for i, (col, et) in enumerate(etiquetas):
            with cols[i % 4]:
                st.number_input(et, key=f"f_{col}", min_value=0.0, step=0.01, format="%.2f")

    # --- Detalle
    st.markdown("##### Detalle")
    if es_venta:
        st.toggle("Generar detalle «Comisión por ventas del … al …»", key="f_usar_comision")
    if es_venta and ss.f_usar_comision:
        c1, c2, c3 = st.columns([1, 1, 2])
        with c1:
            st.date_input("Desde", key="f_com_ini", format="DD/MM/YYYY")
        with c2:
            st.date_input("Hasta", key="f_com_fin", format="DD/MM/YYYY")
        with c3:
            st.text_input("DETALLE (generado)", value=detalle_final(), disabled=True)
    else:
        c1, c2 = st.columns([2, 1.2])
        with c1:
            st.text_input("DETALLE", key="f_detalle", placeholder="Escribe el detalle libremente")
        with c2:
            st.selectbox("Detalles frecuentes", list(detalles), key="f_detalle_sug", index=None,
                         placeholder="Elegir uno…", on_change=cb_sugerencia_detalle,
                         help="Al elegir uno se copia al campo DETALLE, donde puedes modificarlo.")

    with st.expander("Más columnas de la plantilla (opcional, casi nunca se usan)", expanded=False):
        st.caption("Año, Nro Final, documento modificado, cuentas, IMB, detracción, etc. Déjalas vacías si no aplican.")
        cols = st.columns(3)
        for i, col in enumerate(EXTRA_COLS):
            with cols[i % 3]:
                st.text_input(f"{HEADER_BY_COL[col]} ({col})", key=f"f_ex_{col}")
        with st.expander("Columnas CLU1 … CLU39"):
            cols = st.columns(4)
            for i, col in enumerate(CLU_COLS):
                with cols[i % 4]:
                    st.text_input(HEADER_BY_COL[col], key=f"f_ex_{col}")

    for e in ss.get("form_errores", []):
        st.error(e)

    b1, b2 = st.columns([1.4, 1])
    with b1:
        st.button("💾 Guardar cambios" if editando else "➕ Agregar comprobante",
                  type="primary", on_click=cb_guardar_fila, width="stretch")
    with b2:
        if editando:
            st.button("Cancelar edición", on_click=cb_cancelar_edicion, width="stretch")


def lista_comprobantes():
    ss = st.session_state
    st.markdown("#### Comprobantes registrados")
    if not ss.filas:
        st.info("Aún no hay comprobantes. Completa el formulario y pulsa **Agregar comprobante**.")
        return

    for seccion, titulo in [("actual", f"Mes actual · {periodo_seccion('actual')}"),
                            ("anterior", f"Mes anterior · {periodo_seccion('anterior')}")]:
        idxs = [i for i, f in enumerate(ss.filas) if f["_seccion"] == seccion]
        if not idxs:
            continue
        st.markdown(f'<span class="badge">{titulo}</span>', unsafe_allow_html=True)
        for i in idxs:
            f = valores_fila(ss.filas[i])
            en_edicion = ss.editando == i
            with st.container(border=True):
                st.markdown(
                    f'<div class="fila-item{" editando" if en_edicion else ""}">'
                    f'<b>{f["D"]}</b> · {f["E"]:%d/%m/%Y} · <b>{f["B"]}</b><br>'
                    f'{f["N"]} <small>({f["M"]})</small><br>'
                    f'<small>BI {fmt_money(f["O"])} · IGV {fmt_money(f["P"])} · '
                    f'Total {fmt_money(f["Y"])} {f["Z"]} · {f["AI"] or "sin detalle"}</small></div>',
                    unsafe_allow_html=True,
                )
                with st.container(horizontal=True, horizontal_alignment="left"):
                    st.button("✏️ Editar", key=f"ed_{i}", on_click=cb_editar, args=(i,),
                              disabled=en_edicion, width="content")
                    st.button("🗑️ Eliminar", key=f"del_{i}", on_click=cb_eliminar, args=(i,),
                              width="content")
        tot = totales_seccion([ss.filas[i] for i in idxs])
        st.markdown(
            f'<div class="total-box">Total sección · BI <b>{fmt_money(tot["O"])}</b> · '
            f'IGV <b>{fmt_money(tot["P"])}</b> · Total CP <b>{fmt_money(tot["Y"])}</b> · '
            f"{len(idxs)} comprobante(s)</div>",
            unsafe_allow_html=True,
        )
        st.write("")


def paso_2():
    ss = st.session_state
    st.subheader(f"2. Comprobantes de {ss.tipo.lower()}s · {ss.mes} {int(ss.anio)}")
    if ss.get("flash"):
        st.success(ss.flash)
        ss.flash = ""

    col_form, col_lista = st.columns([1.45, 1], gap="large")
    with col_form:
        with st.container(border=True):
            formulario_comprobante()
    with col_lista:
        with st.container(border=True):
            lista_comprobantes()

    st.write("")
    b1, b2, _ = st.columns([1, 1.5, 3])
    with b1:
        st.button("⬅ Volver", on_click=cb_ir_paso, args=(1,), width="stretch")
    with b2:
        st.button("Ver vista previa ➜", type="primary", on_click=cb_ir_paso, args=(3,),
                  disabled=not ss.filas, width="stretch",
                  help=None if ss.filas else "Agrega al menos un comprobante para continuar.")


# ---------------------------------------------------------------------------
# PASO 3
# ---------------------------------------------------------------------------
def documento_actual() -> dict:
    ss = st.session_state
    secciones = [[f for f in ss.filas if f["_seccion"] == "actual"]]
    ant = [f for f in ss.filas if f["_seccion"] == "anterior"]
    if ant:
        secciones.append(ant)
    return {"tipo": ss.tipo, "secciones": secciones}


def dataframe_previa(doc: dict, todas_las_columnas: bool) -> pd.DataFrame:
    registros = []
    for filas in doc["secciones"]:
        for f in filas:
            v = valores_fila(f)
            registros.append({HEADER_BY_COL[c]: v.get(c) for c in COL_LETTERS})
        registros.append({h: None for h in HEADERS})
        tot = totales_seccion(filas)
        fila_tot = {h: None for h in HEADERS}
        fila_tot[HEADER_BY_COL["N"]] = "TOTAL"
        for c in NUM_COLS:
            fila_tot[HEADER_BY_COL[c]] = tot[c]
        registros.append(fila_tot)
        registros.append({h: None for h in HEADERS})
    if registros:
        registros.pop()  # sin fila en blanco al final
    df = pd.DataFrame(registros, columns=HEADERS)
    num_headers = {HEADER_BY_COL[c] for c in NUM_COLS + ["AA"]}
    date_headers = {HEADER_BY_COL["E"], HEADER_BY_COL["F"]}
    for h in HEADERS:
        if h in num_headers:
            df[h] = pd.to_numeric(df[h], errors="coerce")
        elif h in date_headers:
            df[h] = df[h].apply(lambda v: v.strftime("%d/%m/%Y") if isinstance(v, (dt.date, dt.datetime)) else "")
        else:
            df[h] = df[h].apply(lambda v: "" if v is None else str(v))
    if not todas_las_columnas:
        visibles = [HEADER_BY_COL[c] for c in COL_LETTERS[:27]] + [HEADER_BY_COL["AI"]]
        df = df[visibles]
    df.index = range(2, len(df) + 2)  # misma numeración de filas que Excel
    return df


def paso_3():
    ss = st.session_state
    st.subheader("3. Vista previa del Excel")
    doc = documento_actual()

    todas = st.toggle("Mostrar las 80 columnas de la plantilla", value=False)
    df = dataframe_previa(doc, todas)
    col_cfg = {HEADER_BY_COL[c]: st.column_config.NumberColumn(format="%.2f") for c in NUM_COLS}
    col_cfg[HEADER_BY_COL["AA"]] = st.column_config.NumberColumn(format="%.3f")
    st.dataframe(df, width="stretch", column_config=col_cfg, height=min(600, 60 + 35 * len(df)))
    st.caption("Los valores de CONCATENADO, IGV automático, Total CP automático y la fila TOTAL se "
               "escriben como fórmulas en el Excel; aquí se muestran ya calculados. "
               "El índice de la izquierda coincide con la fila de Excel.")

    n = len(ss.filas)
    tot = totales_seccion(doc["secciones"][0])
    st.markdown(
        f'<div class="total-box">{n} comprobante(s) · Mes actual: BI <b>{fmt_money(tot["O"])}</b> · '
        f'IGV <b>{fmt_money(tot["P"])}</b> · Total CP <b>{fmt_money(tot["Y"])}</b></div>',
        unsafe_allow_html=True,
    )

    st.markdown("#### Descargar")
    c1, c2 = st.columns([2, 1], vertical_alignment="bottom")
    with c1:
        nombre = st.text_input("Nombre del archivo", value=ss.nombre_archivo or nombre_archivo_defecto())
        ss.nombre_archivo = nombre
    nombre_limpio = re.sub(r"[^\w\-. ]+", "_", nombre).strip() or nombre_archivo_defecto()
    if not nombre_limpio.lower().endswith(".xlsx"):
        nombre_limpio += ".xlsx"
    datos = build_workbook(doc)
    with c2:
        st.download_button(
            "⬇️ Confirmar y descargar Excel", data=datos, file_name=nombre_limpio,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch", type="primary", on_click=cb_descargado,
        )
    st.caption(f"Se descargará **{nombre_limpio}** con la hoja «Hoja1», las 80 columnas y las fórmulas de la plantilla. "
               "Si algo no cuadra, vuelve a editar: no se pierde nada.")

    st.write("")
    b1, b2, _ = st.columns([1.2, 1.5, 3])
    with b1:
        st.button("⬅ Volver a editar", on_click=cb_ir_paso, args=(2,), width="stretch")
    with b2:
        if ss.descargado:
            with st.popover("🔄 Empezar de nuevo", width="stretch"):
                st.markdown("Se borrarán los comprobantes cargados y volverás al paso 1.")
                st.button("Sí, reiniciar todo", type="primary", on_click=cb_reiniciar, width="stretch")
    if ss.descargado:
        st.success("Archivo descargado. Puedes volver a editar y descargar de nuevo, o empezar otro archivo.")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    init_state()
    ss = st.session_state
    encabezado()
    stepper(ss.paso)
    if ss.paso == 1:
        paso_1()
    elif ss.paso == 2:
        paso_2()
    else:
        paso_3()


main()
