"""Catálogo de valores conocidos extraídos de los archivos históricos SIRE LyJ.

Todo lo que aparece aquí son valores por defecto: la aplicación siempre permite
editarlos o escribir valores nuevos.
"""

RUC_PROPIO = "20519342961"

MESES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre",
]

# Proveedores (Compras). "series" = series de comprobante vistas en los archivos.
PROVEEDORES = [
    {"ruc": "20615283011", "nombre": "Petronex Perú S.A.C.", "series": ["F111", "F112", "F113"], "detalle": "Combustible"},
    {"ruc": "20514636843", "nombre": "Estación de Servicios Petro World S.A.C.", "series": ["F091", "F092", "F093"], "detalle": "Combustible"},
    {"ruc": "20611410256", "nombre": "Estación de Servicios Santo Tomás de Lima SAC.", "series": ["F142"], "detalle": "Combustible"},
    {"ruc": "20503840121", "nombre": "REPSOL COMERCIAL S.A.C.", "series": ["F560"], "detalle": "Combustible"},
    {"ruc": "20514691607", "nombre": "Reencauchadora Palacios SAC.", "series": ["F001"], "detalle": ""},
    {"ruc": "20609572303", "nombre": "Importadora & Exportadora Alison S.A.C.", "series": ["F001"], "detalle": ""},
    {"ruc": "20606759348", "nombre": "Keyra Import E.I.R.L.", "series": ["E001"], "detalle": "Batería"},
    {"ruc": "10106851609", "nombre": "Julca Jatabe Juan Carlos", "series": ["E001"], "detalle": ""},
    {"ruc": "20508443669", "nombre": "CANDAO ELECTRIC SERVICES S.A.C", "series": ["F001"], "detalle": ""},
    {"ruc": "20505253681", "nombre": "DIESEL SILVER E.I.R.L.", "series": ["FPP1"], "detalle": ""},
    {"ruc": "20612555673", "nombre": "Turbo Lima S.A.C.", "series": ["E001"], "detalle": ""},
    {"ruc": "10073773445", "nombre": "Lubricantes Claudia de Pimentel Cruz Juan Santos.", "series": ["FF04"], "detalle": ""},
]

# Clientes (Ventas).
CLIENTES = [
    {"ruc": "20352810615", "nombre": "REENCAUCHADORA AMERICANA SAC.", "series": ["E001"], "detalle": ""},
    {"ruc": "20505337515", "nombre": "REENCAUCHADORA RELINO SAC.", "series": ["E001"], "detalle": ""},
]

DETALLES_COMPRAS = [
    "Combustible",
    "Batería",
    "Aceite Castrol, Filtro de aceite, Filtro de Petróleo",
    "Kit de embrague, bomba de agua, retenes",
    "Juego de Pastillas, jebes y ponchos, líquido, jebes bomba",
    "Empaque de culata",
    "Arrancador mit",
    "01 llanta 155R12C Roadwing",
]

DETALLES_VENTAS = [
    "Comisión por ventas del dd/mm/aa al dd/mm/aa",
]

# Tipo de comprobante (tabla 10 SUNAT). El habitual es 01 = Factura.
TIPOS_CP = ["01", "03", "07", "08", "12", "14"]
TIPOS_CP_DESC = {
    "01": "01 - Factura",
    "03": "03 - Boleta de venta",
    "07": "07 - Nota de crédito",
    "08": "08 - Nota de débito",
    "12": "12 - Ticket",
    "14": "14 - Recibo de servicios públicos",
}

# Tipo de documento de identidad (tabla 2 SUNAT). El habitual es 6 = RUC.
TIPOS_DOC = ["6", "1", "4", "7", "0"]
TIPOS_DOC_DESC = {
    "6": "6 - RUC",
    "1": "1 - DNI",
    "4": "4 - Carnet de extranjería",
    "7": "7 - Pasaporte",
    "0": "0 - Otros",
}

MONEDAS = ["PEN", "USD"]

SERIES_CONOCIDAS = sorted({s for p in PROVEEDORES + CLIENTES for s in p["series"]})


def catalogo_por_tipo(tipo: str):
    """Devuelve (lista de terceros, lista de detalles) según 'Compra' o 'Venta'."""
    if tipo == "Venta":
        return CLIENTES, DETALLES_VENTAS
    return PROVEEDORES, DETALLES_COMPRAS


def buscar_tercero(nombre: str, tipo: str):
    terceros, _ = catalogo_por_tipo(tipo)
    nombre_n = (nombre or "").strip().lower()
    for t in terceros:
        if t["nombre"].strip().lower() == nombre_n:
            return t
    return None
