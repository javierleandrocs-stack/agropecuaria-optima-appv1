"""
================================================================================
SISTEMA INTEGRAL FINCA COMPLETO (VERSIÓN WEB Y MÓVIL PARA IPHONE, ANDROID Y PC)
Empresa: AGROPECUARIA OPTIMA C.A
Archivo: app_movil_finca.py
Contiene los 7 Módulos Completos en Pestañas Visibles (Sin textos None):
 1. Listas de Precios (Checklist ☑/☐, Ajuste Masivo % y Editor de Catálogo).
 2. Órdenes de Despacho (Clientes, Kilos x Precio Unitario, Descuento de Stock).
 3. Notas de Crédito (Centros de Distribución y Afectación de Saldos).
 4. Notas de Débito y Movimientos (Preselector de Choferes y Rotación en Vivo).
 5. Inventario por Sede (Existencias en KG, Carga Manual, Alertas y PDF/Excel).
 6. Bases de Datos Maestras (Clientes, Centros de Distribución y Choferes).
 7. Historial General de Documentos (Reimpresión PDF/Excel y Reversión de Stock).
================================================================================
"""

import io
import sqlite3
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Tuple

import urllib3
import requests
from bs4 import BeautifulSoup
import streamlit as st
from PIL import Image, ImageDraw
from fpdf import FPDF
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


# ==============================================================================
# PARÁMETROS AJUSTABLES DEL SISTEMA
# ==============================================================================
NOMBRE_EMPRESA = "AGROPECUARIA OPTIMA C.A"
SEDE_PLANTA_PRINCIPAL = "PLANTA / FINCA PRINCIPAL"
OPCION_CHOFER_MANUAL = "-- Escribir Manualmente --"

UMBRAL_STOCK_BAJO_KG = 20.0
ALTO_MAX_LOGO_PDF_MM = 24.0
ANCHO_MAX_LOGO_PDF_MM = 88.0

PREFIJOS_DOCUMENTO = {
    "DESPACHO": "OD",
    "NOTA_CREDITO": "NC",
    "NOTA_DEBITO": "ND",
    "MOVIMIENTO": "MOV",
}

MODALIDADES_LISTA_PRECIOS = [
    "Lista Completa (Las 3 Escalas)",
    "Solo Precio Al Mayor (Mas de 50kg)",
    "Solo Precio Vendedor 10%",
    "Solo Precio Cliente Final",
    "Vendedor 10% + Cliente Final",
]

RUBROS_CATALOGO = [
    "Cortes Frescos",
    "Ahumados y Especiales",
    "Charcutería y Embutidos",
]

ESCALAS_PRECIO = {
    "Cliente Final": "precio_cliente",
    "Vendedor 10%": "precio_vendedor",
    "Al Mayor": "precio_mayor",
}

DIAS_SEMANA_ES = {
    0: "Lunes",
    1: "Martes",
    2: "Miércoles",
    3: "Jueves",
    4: "Viernes",
    5: "Sábado",
    6: "Domingo",
}

RUTA_BASE: Path = Path(__file__).resolve().parent
RUTA_DB: Path = RUTA_BASE / "sistema_finca_optima.db"
RUTA_LOGO: Path = RUTA_BASE / "logo_finca.png"


def obtener_pin_seguro() -> str:
    """
    Lee el PIN desde la bóveda privada 'Secrets' de Streamlit Cloud.
    Si no está configurado en Secrets, acepta 'a1b2c3' por defecto.
    """
    try:
        if "PIN_ACCESO" in st.secrets:
            return str(st.secrets["PIN_ACCESO"]).strip()
    except Exception:
        pass
    return "a1b2c3"


PIN_ACCESO_MOVIL = obtener_pin_seguro()


# ==============================================================================
# PASO 1: INICIALIZACIÓN DE BASE DE DATOS Y LOGOTIPO
# ==============================================================================
def asegurar_logo() -> None:
    if RUTA_LOGO.exists():
        return
    img = Image.new("RGBA", (440, 170), (63, 90, 42, 255))
    dib = ImageDraw.Draw(img)
    dib.rounded_rectangle((8, 8, 432, 162), radius=20, outline=(250, 204, 21, 255), width=4)
    img.save(RUTA_LOGO, format="PNG")


def obtener_marca_tiempo() -> Tuple[str, str, str]:
    ahora = datetime.now()
    return (
        DIAS_SEMANA_ES[ahora.weekday()],
        ahora.strftime("%d/%m/%Y"),
        ahora.strftime("%I:%M:%S %p"),
    )


def conectar_db() -> sqlite3.Connection:
    conn = sqlite3.connect(RUTA_DB, timeout=20, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys=ON;")
    return conn


def ejecutar_sql(query: str, params: tuple = ()) -> int:
    with conectar_db() as conn:
        cur = conn.execute(query, params)
        conn.commit()
        return cur.lastrowid


def consultar_sql(query: str, params: tuple = ()) -> List[sqlite3.Row]:
    with conectar_db() as conn:
        return conn.execute(query, params).fetchall()


CATALOGO_BASE = [
    ("CHULETA FRESCA", "Cortes Frescos", 5.85, 6.50, 6.99, 1),
    ("CHULETA AHUMADA", "Ahumados y Especiales", 4.95, 5.45, 5.99, 1),
    ("CHULETA DE PERNIL", "Cortes Frescos", 5.85, 6.50, 6.99, 1),
    ("LOMO DE CERDO", "Cortes Frescos", 8.10, 8.99, 9.99, 1),
    ("COSTILLITAS CHINAS", "Cortes Frescos", 3.60, 3.99, 4.39, 1),
    ("COSTILLAS FRESCAS", "Cortes Frescos", 6.50, 6.99, 7.69, 1),
    ("CERDO PICADO", "Cortes Frescos", 4.50, 4.99, 5.49, 1),
    ("HUESITOS DE CERDO", "Cortes Frescos", 3.60, 3.99, 4.40, 1),
    ("HUESITOS AHUMADOS", "Ahumados y Especiales", 2.50, 2.75, 2.99, 1),
    ("PERNIL CON CUERO", "Cortes Frescos", 4.50, 4.95, 5.50, 1),
    ("PERNIL SIN CUERO", "Cortes Frescos", 4.95, 5.50, 5.99, 1),
    ("SOLOMILLO", "Cortes Frescos", 6.30, 6.99, 7.69, 1),
    ("COSTILLA BBQ", "Ahumados y Especiales", 5.85, 6.50, 6.99, 1),
    ("PORKBELLY", "Cortes Frescos", 7.20, 7.99, 8.79, 1),
    ("MORTADELA ESPECIAL", "Charcutería y Embutidos", 3.30, 3.63, 3.99, 1),
    ("MORTADELA PREMIUM", "Charcutería y Embutidos", 4.20, 4.62, 4.99, 1),
    ("JAMON DE PIERNA", "Charcutería y Embutidos", 4.60, 4.99, 5.49, 1),
    ("JAMON DE ESPALDA", "Charcutería y Embutidos", 4.20, 4.60, 4.99, 1),
    ("JAMON AREPERO", "Charcutería y Embutidos", 4.50, 4.99, 5.49, 1),
    ("TOCINETA AHUMADA", "Ahumados y Especiales", 6.00, 6.60, 7.26, 1),
    ("CHORIZO (AJO, AHUMADO, PICANTE, OTROS)", "Charcutería y Embutidos", 5.50, 5.99, 6.59, 1),
    ("PATAS DE CERDO", "Cortes Frescos", 2.70, 2.99, 3.29, 1),
]


def inicializar_base_de_datos() -> None:
    with conectar_db() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS tasas_cambio (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha TEXT NOT NULL, hora TEXT NOT NULL,
                bcv_usd REAL NOT NULL, bcv_eur REAL NOT NULL, fuente TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS productos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL UNIQUE, rubro TEXT NOT NULL,
                precio_mayor REAL NOT NULL, precio_vendedor REAL NOT NULL,
                precio_cliente REAL NOT NULL, disponible INTEGER NOT NULL DEFAULT 1
            );
            CREATE TABLE IF NOT EXISTS clientes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL, documento_id TEXT NOT NULL,
                escala_preferida TEXT NOT NULL DEFAULT 'Cliente Final',
                telefono TEXT DEFAULT '', direccion TEXT DEFAULT ''
            );
            CREATE TABLE IF NOT EXISTS centros_distribucion (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo TEXT NOT NULL UNIQUE, nombre TEXT NOT NULL,
                encargado TEXT DEFAULT '', telefono TEXT DEFAULT '',
                ubicacion TEXT DEFAULT '', saldo_pendiente_ref REAL DEFAULT 0.0
            );
            CREATE TABLE IF NOT EXISTS choferes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL, cedula TEXT NOT NULL,
                placa TEXT NOT NULL, telefono TEXT DEFAULT '', vehiculo TEXT DEFAULT ''
            );
            CREATE TABLE IF NOT EXISTS inventario_sedes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sede TEXT NOT NULL, producto TEXT NOT NULL,
                kilos_disponibles REAL NOT NULL DEFAULT 0.0,
                ultima_actualizacion TEXT NOT NULL,
                UNIQUE(sede, producto)
            );
            CREATE TABLE IF NOT EXISTS documentos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                correlativo TEXT NOT NULL UNIQUE, tipo_doc TEXT NOT NULL,
                subtipo TEXT NOT NULL, entidad_origen TEXT NOT NULL,
                entidad_destino TEXT NOT NULL, doc_identidad TEXT DEFAULT '',
                referencia TEXT DEFAULT '', motivo TEXT DEFAULT '',
                chofer TEXT DEFAULT '', cedula_chofer TEXT DEFAULT '',
                placa_vehiculo TEXT DEFAULT '', afecto_saldo INTEGER DEFAULT 0,
                afecto_inventario INTEGER DEFAULT 1, observaciones TEXT DEFAULT '',
                modo_moneda TEXT NOT NULL, tasa_usd REAL NOT NULL,
                tasa_eur REAL NOT NULL, total_kilos REAL NOT NULL,
                total_ref REAL NOT NULL, total_bs REAL NOT NULL,
                dia_semana TEXT NOT NULL, fecha TEXT NOT NULL, hora TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS documento_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                documento_id INTEGER NOT NULL, producto TEXT NOT NULL,
                escala TEXT NOT NULL, kilos REAL NOT NULL,
                precio_unitario REAL NOT NULL, subtotal_ref REAL NOT NULL,
                FOREIGN KEY (documento_id) REFERENCES documentos(id) ON DELETE CASCADE
            );
            """
        )
        cols_prod = {f["name"] for f in conn.execute("PRAGMA table_info(productos)").fetchall()}
        if "disponible" not in cols_prod:
            conn.execute("ALTER TABLE productos ADD COLUMN disponible INTEGER NOT NULL DEFAULT 1;")

        if conn.execute("SELECT COUNT(*) AS c FROM productos").fetchone()["c"] == 0:
            conn.executemany(
                "INSERT INTO productos (nombre, rubro, precio_mayor, precio_vendedor, precio_cliente, disponible) VALUES (?, ?, ?, ?, ?, ?)",
                CATALOGO_BASE,
            )
        if conn.execute("SELECT COUNT(*) AS c FROM clientes").fetchone()["c"] == 0:
            conn.execute(
                "INSERT INTO clientes (nombre, documento_id, escala_preferida, telefono, direccion) VALUES (?, ?, ?, ?, ?)",
                ("CLIENTE MOSTRADOR GENERAL", "V-00000000", "Cliente Final", "", "Venta directa"),
            )
        if conn.execute("SELECT COUNT(*) AS c FROM centros_distribucion").fetchone()["c"] == 0:
            conn.executemany(
                "INSERT INTO centros_distribucion (codigo, nombre, encargado, telefono, ubicacion, saldo_pendiente_ref) VALUES (?, ?, ?, ?, ?, ?)",
                [
                    ("CD-01", "CENTRO DE DISTRIBUCION PRINCIPAL", "Encargado Sede 1", "", "Sede Central", 0.0),
                    ("CD-02", "CENTRO DE DISTRIBUCION SUCURSAL 2", "Encargado Sede 2", "", "Zona Comercial", 0.0),
                ],
            )
        if conn.execute("SELECT COUNT(*) AS c FROM choferes").fetchone()["c"] == 0:
            conn.execute(
                "INSERT INTO choferes (nombre, cedula, placa, telefono, vehiculo) VALUES (?, ?, ?, ?, ?)",
                ("CHOFER TRANSPORTE PLANTA", "V-12345678", "A00AA00", "0414-0000000", "Camión Cava Fría"),
            )
        if conn.execute("SELECT COUNT(*) AS c FROM tasas_cambio").fetchone()["c"] == 0:
            ahora = datetime.now()
            conn.execute(
                "INSERT INTO tasas_cambio (fecha, hora, bcv_usd, bcv_eur, fuente) VALUES (?, ?, ?, ?, ?)",
                (ahora.strftime("%Y-%m-%d"), ahora.strftime("%H:%M:%S"), 901.62, 1007.96, "Inicial"),
            )
        conn.commit()


# ==============================================================================
# PASO 2: LÓGICA DE TASAS BCV, STOCK POR SEDE Y CORRELATIVOS
# ==============================================================================
def obtener_tasas_db() -> Dict[str, float]:
    filas = consultar_sql("SELECT bcv_usd, bcv_eur FROM tasas_cambio ORDER BY id DESC LIMIT 1")
    if filas:
        return {"bcv_usd": float(filas[0]["bcv_usd"]), "bcv_eur": float(filas[0]["bcv_eur"])}
    return {"bcv_usd": 901.62, "bcv_eur": 1007.96}


def guardar_tasas_db(usd: float, eur: float, fuente: str = "Manual") -> None:
    ahora = datetime.now()
    ejecutar_sql(
        "INSERT INTO tasas_cambio (fecha, hora, bcv_usd, bcv_eur, fuente) VALUES (?, ?, ?, ?, ?)",
        (ahora.strftime("%Y-%m-%d"), ahora.strftime("%H:%M:%S"), round(usd, 2), round(eur, 2), fuente),
    )


def sincronizar_tasas_bcv_web() -> Tuple[bool, float, float]:
    actuales = obtener_tasas_db()
    usd, eur = actuales["bcv_usd"], actuales["bcv_eur"]
    exito = False
    try:
        r = requests.get("https://ve.dolarapi.com/v1/dolares/oficial", timeout=8)
        if r.status_code == 200 and r.json().get("promedio"):
            usd = round(float(r.json()["promedio"]), 2)
            exito = True
    except Exception:
        pass
    try:
        r_bcv = requests.get("https://www.bcv.org.ve/", headers={"User-Agent": "Mozilla/5.0"}, verify=False, timeout=10)
        if r_bcv.status_code == 200:
            sopa = BeautifulSoup(r_bcv.text, "html.parser")
            for div_id, es_euro in (("euro", True), ("dolar", False)):
                bloque = sopa.find("div", id=div_id)
                if bloque and bloque.find("strong"):
                    val = float(bloque.find("strong").text.strip().replace(".", "").replace(",", "."))
                    if es_euro:
                        eur = round(val, 2)
                    else:
                        usd = round(val, 2)
                    exito = True
    except Exception:
        pass
    if exito:
        guardar_tasas_db(usd, eur, fuente="Auto-BCV")
    return exito, usd, eur


def normalizar_nombre_sede(nombre_sede: str) -> str:
    limpio = str(nombre_sede).strip()
    if limpio.startswith(SEDE_PLANTA_PRINCIPAL):
        return SEDE_PLANTA_PRINCIPAL
    return limpio


def ajustar_stock_sede(conn: sqlite3.Connection, sede: str, producto: str, delta_kilos: float) -> None:
    sede_norm = normalizar_nombre_sede(sede)
    if not sede_norm or not producto:
        return
    _, f_str, h_str = obtener_marca_tiempo()
    conn.execute(
        """INSERT INTO inventario_sedes (sede, producto, kilos_disponibles, ultima_actualizacion)
           VALUES (?, ?, ?, ?)
           ON CONFLICT(sede, producto) DO UPDATE SET
               kilos_disponibles = ROUND(inventario_sedes.kilos_disponibles + excluded.kilos_disponibles, 2),
               ultima_actualizacion = excluded.ultima_actualizacion""",
        (sede_norm, producto.strip().upper(), round(delta_kilos, 2), f"{f_str} {h_str}"),
    )


def fijar_stock_exacto_sede(sede: str, producto: str, kilos_exactos: float) -> None:
    sede_norm = normalizar_nombre_sede(sede)
    _, f_str, h_str = obtener_marca_tiempo()
    ejecutar_sql(
        """INSERT INTO inventario_sedes (sede, producto, kilos_disponibles, ultima_actualizacion)
           VALUES (?, ?, ?, ?)
           ON CONFLICT(sede, producto) DO UPDATE SET
               kilos_disponibles = excluded.kilos_disponibles,
               ultima_actualizacion = excluded.ultima_actualizacion""",
        (sede_norm, producto.strip().upper(), round(kilos_exactos, 2), f"{f_str} {h_str}"),
    )


def generar_siguiente_correlativo(tipo_doc: str) -> str:
    prefijo = PREFIJOS_DOCUMENTO.get(tipo_doc, "DOC")
    fila = consultar_sql("SELECT COUNT(*) AS total FROM documentos WHERE tipo_doc = ?", (tipo_doc,))
    consecutivo = (fila[0]["total"] if fila else 0) + 1
    while True:
        candidato = f"{prefijo}-{consecutivo:04d}"
        if not consultar_sql("SELECT 1 FROM documentos WHERE correlativo = ?", (candidato,)):
            return candidato
        consecutivo += 1


# ==============================================================================
# PASO 3: GENERADORES DE PDF Y EXCEL EN MEMORIA (SIN IMPRIMIR 'NONE')
# ==============================================================================
def limpiar_pdf(texto: str) -> str:
    return str(texto).replace("€", "EUR").encode("latin-1", "replace").decode("latin-1")


def obtener_tasa_activa(modo_moneda: str, tasa_usd: float, tasa_eur: float) -> Tuple[float, str, str]:
    if ("Euro" in modo_moneda) or ("€" in modo_moneda):
        return tasa_eur, "EUR", "Euro BCV"
    return tasa_usd, "$", "Dólar BCV"


def formatear_monto_celda(valor_ref: float, modo_moneda: str, tasa_usd: float, tasa_eur: float) -> str:
    tasa, simb, _ = obtener_tasa_activa(modo_moneda, tasa_usd, tasa_eur)
    if modo_moneda.startswith("REF"):
        return f"{simb} {valor_ref:,.2f}"
    if modo_moneda.startswith("Bolívares"):
        return f"Bs {valor_ref * tasa:,.2f}"
    return f"{simb} {valor_ref:,.2f} / Bs {valor_ref * tasa:,.2f}"


def colocar_encabezado_pdf(pdf: FPDF, subtitulo: str, cintillo: str) -> None:
    if RUTA_LOGO.exists():
        try:
            with Image.open(RUTA_LOGO) as im:
                w_px, h_px = im.size
            prop = w_px / max(h_px, 1)
            alto = ALTO_MAX_LOGO_PDF_MM
            ancho = min(alto * prop, ANCHO_MAX_LOGO_PDF_MM)
            pdf.image(str(RUTA_LOGO), x=(210.0 - ancho) / 2.0, y=6, w=ancho, h=alto)
            pdf.set_y(6 + alto + 3)
        except Exception:
            pdf.set_y(10)
    pdf.set_fill_color(63, 90, 42)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 12.5)
    pdf.cell(0, 9, limpiar_pdf(f"{NOMBRE_EMPRESA}  -  {subtitulo}"), ln=True, align="C", fill=True)
    pdf.set_fill_color(217, 119, 6)
    pdf.set_font("Helvetica", "B", 9.2)
    pdf.cell(0, 6.5, limpiar_pdf(cintillo), ln=True, align="C", fill=True)
    pdf.ln(3)


def obtener_columnas_lista_precios(modalidad_lista: str) -> List[Tuple[str, str]]:
    if "Solo Precio Al Mayor" in modalidad_lista:
        return [("AL MAYOR (+50KG)", "precio_mayor")]
    if "Solo Precio Vendedor" in modalidad_lista:
        return [("VENDEDOR (10%)", "precio_vendedor")]
    if "Solo Precio Cliente" in modalidad_lista:
        return [("CLIENTE FINAL", "precio_cliente")]
    if "Vendedor 10% + Cliente" in modalidad_lista:
        return [("VENDEDOR (10%)", "precio_vendedor"), ("CLIENTE FINAL", "precio_cliente")]
    return [
        ("AL MAYOR (+50KG)", "precio_mayor"),
        ("VENDEDOR (10%)", "precio_vendedor"),
        ("CLIENTE FINAL", "precio_cliente"),
    ]


def generar_bytes_lista_precios(
    en_excel: bool, modalidad_lista: str, rubro_filtro: str,
    modo_moneda: str, tasa_usd: float, tasa_eur: float, solo_disponibles: bool
) -> bytes:
    if rubro_filtro == "Todos los Rubros":
        if solo_disponibles:
            productos = consultar_sql("SELECT * FROM productos WHERE disponible=1 ORDER BY rubro, nombre")
        else:
            productos = consultar_sql("SELECT * FROM productos ORDER BY rubro, nombre")
    else:
        if solo_disponibles:
            productos = consultar_sql("SELECT * FROM productos WHERE rubro=? AND disponible=1 ORDER BY nombre", (rubro_filtro,))
        else:
            productos = consultar_sql("SELECT * FROM productos WHERE rubro=? ORDER BY nombre", (rubro_filtro,))

    dia_s, fecha_s, hora_s = obtener_marca_tiempo()
    tasa_apl, simb, nom_tasa = obtener_tasa_activa(modo_moneda, tasa_usd, tasa_eur)
    cols_precio = obtener_columnas_lista_precios(modalidad_lista)

    if en_excel:
        wb = Workbook()
        ws = wb.active
        ws.title = "Lista_Precios"
        ws["A1"] = f"{NOMBRE_EMPRESA} - LISTA DE PRECIOS ({modalidad_lista})"
        ws["A1"].font = Font(size=13, bold=True, color="3F5A2A")
        ws.append([f"Fecha: {dia_s} {fecha_s} ({hora_s}) | Moneda: {modo_moneda} | Tasa {nom_tasa}: Bs {tasa_apl:,.2f}"])
        ws.append([])
        enc = ["Producto / Corte", "Categoría"]
        for t_col, _ in cols_precio:
            enc.extend([f"{t_col} ({simb})", f"{t_col} (Bs)"])
        ws.append(enc)
        for p in productos:
            fila = [p["nombre"], p["rubro"]]
            for _, c_db in cols_precio:
                v = float(p[c_db])
                fila.extend([round(v, 2), round(v * tasa_apl, 2)])
            ws.append(fila)
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.add_page()
    cintillo = f"MODALIDAD: {modalidad_lista.upper()}  |  TASA {nom_tasa.upper()}: Bs {tasa_apl:,.2f}  |  {dia_s.upper()} {fecha_s}"
    colocar_encabezado_pdf(pdf, "LISTA OFICIAL DE PRECIOS DISPONIBLES", cintillo)

    n_cols = len(cols_precio)
    if n_cols == 3:
        ancho_p = 36
    elif n_cols == 2:
        ancho_p = 46
    else:
        ancho_p = 60
    ancho_r = 38
    ancho_nom = 190 - ancho_r - (ancho_p * n_cols)

    pdf.set_fill_color(234, 228, 215)
    pdf.set_text_color(45, 66, 30)
    pdf.set_font("Helvetica", "B", 8.5)
    pdf.cell(ancho_nom, 8, "  PRODUCTO / CORTE", border=1, fill=True)
    pdf.cell(ancho_r, 8, "CATEGORIA", border=1, align="C", fill=True)
    for i, (t_col, _) in enumerate(cols_precio):
        pdf.cell(ancho_p, 8, limpiar_pdf(t_col), border=1, align="R", fill=True, ln=(i == n_cols - 1))

    for idx, p in enumerate(productos):
        if idx % 2 == 1:
            pdf.set_fill_color(246, 244, 238)
        else:
            pdf.set_fill_color(255, 255, 255)
        pdf.set_text_color(20, 20, 20)
        pdf.set_font("Helvetica", "B", 8.2)
        pdf.cell(ancho_nom, 7, limpiar_pdf(f"  {p['nombre']}")[:42], border=1, fill=True)
        pdf.set_font("Helvetica", "", 8)
        pdf.cell(ancho_r, 7, limpiar_pdf(p["rubro"])[:22], border=1, align="C", fill=True)
        pdf.set_font("Helvetica", "B", 8.2)
        for i, (_, c_db) in enumerate(cols_precio):
            txt = formatear_monto_celda(float(p[c_db]), modo_moneda, tasa_usd, tasa_eur)
            pdf.cell(ancho_p, 7, limpiar_pdf(f"{txt} "), border=1, align="R", fill=True, ln=(i == n_cols - 1))

    return bytes(pdf.output())


def generar_bytes_documento(documento_id: int, en_excel: bool = False) -> Tuple[bytes, str]:
    doc = consultar_sql("SELECT * FROM documentos WHERE id = ?", (documento_id,))[0]
    items = consultar_sql("SELECT * FROM documento_items WHERE documento_id = ?", (documento_id,))
    tasa_apl, simb, nom_tasa = obtener_tasa_activa(doc["modo_moneda"], doc["tasa_usd"], doc["tasa_eur"])
    base_nom = f"{doc['correlativo']}_{doc['fecha'].replace('/', '')}"

    if en_excel:
        wb = Workbook()
        ws = wb.active
        ws.title = doc["correlativo"]
        ws["A1"] = f"{NOMBRE_EMPRESA} - {doc['subtipo']} ({doc['correlativo']})"
        ws["A1"].font = Font(size=13, bold=True, color="3F5A2A")
        ws.append([f"Origen: {doc['entidad_origen']} | Destino: {doc['entidad_destino']} | Fecha: {doc['fecha']} {doc['hora']}"])
        ws.append([f"Detalle: {doc['motivo']} | Chofer: {doc['chofer'] or 'N/A'} | CI: {doc['cedula_chofer'] or 'N/A'} | Placa: {doc['placa_vehiculo'] or 'N/A'}"])
        ws.append([])
        ws.append(["Producto", "Escala", "Peso (KG)", f"Precio Unit ({simb})", f"Subtotal ({simb})", f"Subtotal Bs ({nom_tasa})"])
        for it in items:
            sr = float(it["subtotal_ref"])
            ws.append([it["producto"], it["escala"], round(float(it["kilos"]), 2), round(float(it["precio_unitario"]), 2), round(sr, 2), round(sr * tasa_apl, 2)])
        ws.append([])
        ws.append(["TOTALES", "", round(doc["total_kilos"], 2), "", round(doc["total_ref"], 2), round(doc["total_bs"], 2)])
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue(), f"{base_nom}.xlsx"

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=22)
    pdf.add_page()
    cintillo = f"DOCUMENTO: {doc['correlativo']}  |  FECHA: {doc['fecha']} {doc['hora']}  |  TASA {nom_tasa.upper()}: Bs {tasa_apl:,.2f}"
    colocar_encabezado_pdf(pdf, doc["subtipo"].upper(), cintillo)

    pdf.set_fill_color(246, 244, 238)
    pdf.set_draw_color(63, 90, 42)
    pdf.set_text_color(25, 25, 25)
    pdf.set_font("Helvetica", "B", 9)

    if doc["tipo_doc"] == "DESPACHO":
        pdf.cell(125, 6.5, limpiar_pdf(f"  CLIENTE: {doc['entidad_destino']}"), border="LTR", fill=True)
        pdf.cell(65, 6.5, limpiar_pdf(f"  RIF/CI: {doc['doc_identidad']}"), border="LTR", fill=True, ln=True)
        pdf.set_font("Helvetica", "", 8.5)
        pdf.cell(125, 6, limpiar_pdf(f"  DESPACHA: {doc['entidad_origen']}"), border="LBR", fill=True)
        pdf.cell(65, 6, limpiar_pdf(f"  DETALLE: {doc['motivo']}"), border="LBR", fill=True, ln=True)
    else:
        pdf.cell(95, 6.5, limpiar_pdf(f"  ORIGEN: {doc['entidad_origen']}"), border="LTR", fill=True)
        pdf.cell(95, 6.5, limpiar_pdf(f"  DESTINO: {doc['entidad_destino']}"), border="LTR", fill=True, ln=True)
        pdf.set_font("Helvetica", "", 8.5)
        pdf.cell(95, 6, limpiar_pdf(f"  MOTIVO: {doc['motivo']}"), border="LBR", fill=True)
        pdf.cell(95, 6, limpiar_pdf(f"  REFERENCIA: {doc['referencia'] or 'N/A'}"), border="LBR", fill=True, ln=True)
        if doc["chofer"] or doc["cedula_chofer"] or doc["placa_vehiculo"]:
            pdf.set_font("Helvetica", "B", 8.5)
            tr = f"  TRANSPORTE -> Chofer: {doc['chofer'] or 'N/A'} | CI: {doc['cedula_chofer'] or 'N/A'} | Placa: {doc['placa_vehiculo'] or 'N/A'}"
            pdf.cell(190, 6.5, limpiar_pdf(tr), border=1, fill=True, ln=True)

    if doc["observaciones"]:
        pdf.set_font("Helvetica", "I", 8.5)
        pdf.cell(190, 6, limpiar_pdf(f"  Observaciones: {doc['observaciones']}"), border=1, fill=True, ln=True)

    pdf.ln(3)
    pdf.set_fill_color(234, 228, 215)
    pdf.set_text_color(45, 66, 30)
    pdf.set_font("Helvetica", "B", 8.5)
    pdf.cell(78, 8, "  PRODUCTO / CORTE", border=1, fill=True)
    pdf.cell(28, 8, "ESCALA", border=1, align="C", fill=True)
    pdf.cell(24, 8, "PESO (KG)", border=1, align="R", fill=True)
    pdf.cell(30, 8, "PRECIO UNIT.", border=1, align="R", fill=True)
    pdf.cell(30, 8, "SUBTOTAL", border=1, align="R", fill=True, ln=True)

    for it in items:
        pdf.set_fill_color(255, 255, 255)
        pdf.set_text_color(20, 20, 20)
        pdf.set_font("Helvetica", "B", 8.2)
        pdf.cell(78, 7, limpiar_pdf(f"  {it['producto']}")[:44], border=1, fill=True)
        pdf.set_font("Helvetica", "", 8)
        pdf.cell(28, 7, limpiar_pdf(it["escala"]), border=1, align="C", fill=True)
        pdf.cell(24, 7, f"{float(it['kilos']):,.2f} kg ", border=1, align="R", fill=True)
        pu = formatear_monto_celda(float(it["precio_unitario"]), doc["modo_moneda"], doc["tasa_usd"], doc["tasa_eur"])
        stx = formatear_monto_celda(float(it["subtotal_ref"]), doc["modo_moneda"], doc["tasa_usd"], doc["tasa_eur"])
        pdf.cell(30, 7, limpiar_pdf(pu), border=1, align="R", fill=True)
        pdf.set_font("Helvetica", "B", 8.2)
        pdf.cell(30, 7, limpiar_pdf(stx), border=1, align="R", fill=True, ln=True)

    pdf.ln(3)
    pdf.set_fill_color(63, 90, 42)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 9.5)
    pdf.cell(60, 8, f"  TOTAL PESO: {doc['total_kilos']:,.2f} KG", border=1, fill=True)
    pdf.cell(62, 8, limpiar_pdf(f"  TOTAL REF: {simb} {doc['total_ref']:,.2f}"), border=1, align="C", fill=True)
    pdf.cell(68, 8, limpiar_pdf(f"TOTAL {nom_tasa.upper()}: Bs {doc['total_bs']:,.2f}  "), border=1, align="R", fill=True, ln=True)

    return bytes(pdf.output()), f"{base_nom}.pdf"


def generar_bytes_inventario(
    en_excel: bool, sede_filtro: str, filas_inv: List[Dict[str, object]],
    modo_moneda: str, t_usd: float, t_eur: float
) -> bytes:
    dia_s, fecha_s, hora_s = obtener_marca_tiempo()
    tasa_apl, simb, nom_tasa = obtener_tasa_activa(modo_moneda, t_usd, t_eur)

    if en_excel:
        wb = Workbook()
        ws = wb.active
        ws.title = "Inventario_Sedes"
        ws["A1"] = f"{NOMBRE_EMPRESA} - INVENTARIO POR SEDE ({sede_filtro})"
        ws.append([f"Emitido: {fecha_s} {hora_s} | Tasa {nom_tasa}: Bs {tasa_apl:,.2f}"])
        ws.append([])
        ws.append(["Sede", "Producto", "Stock (KG)", "Estado", f"Valor Mayor ({simb})", "Valor en Bs", "Actualizado"])
        for r in filas_inv:
            ws.append([r["sede"], r["producto"], r["kilos"], r["estado"], r["valor_ref"], round(float(r["valor_ref"]) * tasa_apl, 2), r["actualizado"]])
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.add_page()
    cintillo = f"SEDE: {sede_filtro.upper()}  |  FECHA: {dia_s} {fecha_s}  |  TASA {nom_tasa.upper()}: Bs {tasa_apl:,.2f}"
    colocar_encabezado_pdf(pdf, "REPORTE DE INVENTARIO POR SEDE", cintillo)

    pdf.set_fill_color(234, 228, 215)
    pdf.set_text_color(45, 66, 30)
    pdf.set_font("Helvetica", "B", 8.5)
    pdf.cell(62, 8, "  SEDE / CENTRO DIST.", border=1, fill=True)
    pdf.cell(54, 8, "  PRODUCTO / CORTE", border=1, fill=True)
    pdf.cell(24, 8, "STOCK (KG)", border=1, align="R", fill=True)
    pdf.cell(25, 8, f"VALOR ({simb})", border=1, align="R", fill=True)
    pdf.cell(25, 8, "ACTUALIZADO", border=1, align="C", fill=True, ln=True)

    tot_kg, tot_ref = 0.0, 0.0
    for r in filas_inv:
        kg = float(r["kilos"])
        vr = float(r["valor_ref"])
        tot_kg += kg
        tot_ref += vr
        pdf.set_fill_color(255, 255, 255)
        pdf.set_text_color(20, 20, 20)
        pdf.set_font("Helvetica", "", 8)
        pdf.cell(62, 7, limpiar_pdf(f"  {r['sede']}")[:34], border=1, fill=True)
        pdf.set_font("Helvetica", "B", 8)
        pdf.cell(54, 7, limpiar_pdf(f"  {r['producto']}")[:30], border=1, fill=True)
        pdf.cell(24, 7, f"{kg:,.2f} kg ", border=1, align="R", fill=True)
        pdf.set_font("Helvetica", "", 8)
        pdf.cell(25, 7, limpiar_pdf(f"{simb} {vr:,.2f} "), border=1, align="R", fill=True)
        pdf.cell(25, 7, limpiar_pdf(str(r["actualizado"])[:10]), border=1, align="C", fill=True, ln=True)

    pdf.ln(3)
    pdf.set_fill_color(63, 90, 42)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 9.5)
    pdf.cell(60, 8, f"  TOTAL STOCK: {tot_kg:,.2f} KG", border=1, fill=True)
    pdf.cell(62, 8, limpiar_pdf(f"  VALOR MAYOR: {simb} {tot_ref:,.2f}"), border=1, align="C", fill=True)
    pdf.cell(68, 8, limpiar_pdf(f"EQUIVALENTE: Bs {tot_ref * tasa_apl:,.2f}  "), border=1, align="R", fill=True, ln=True)
    return bytes(pdf.output())


def guardar_operacion_db(
    tipo_doc: str, subtipo: str, origen: str, destino: str, doc_identidad: str,
    referencia: str, motivo: str, chofer: str, cedula_chofer: str, placa_vehiculo: str,
    afecto_saldo: int, afecto_inventario: int, observaciones: str,
    modo_moneda: str, tasa_usd: float, tasa_eur: float, lineas_carrito: List[Dict[str, object]]
) -> int:
    correlativo = generar_siguiente_correlativo(tipo_doc)
    tasa_apl, _, _ = obtener_tasa_activa(modo_moneda, tasa_usd, tasa_eur)
    tot_kg = round(sum(float(x["kilos"]) for x in lineas_carrito), 2)
    tot_ref = round(sum(float(x["subtotal_ref"]) for x in lineas_carrito), 2)
    tot_bs = round(tot_ref * tasa_apl, 2)
    dia_sem, fecha_str, hora_str = obtener_marca_tiempo()

    origen_norm = normalizar_nombre_sede(origen)
    if tipo_doc != "DESPACHO":
        destino_norm = normalizar_nombre_sede(destino)
    else:
        destino_norm = destino.strip()

    columnas = (
        "correlativo", "tipo_doc", "subtipo", "entidad_origen", "entidad_destino",
        "doc_identidad", "referencia", "motivo", "chofer", "cedula_chofer",
        "placa_vehiculo", "afecto_saldo", "afecto_inventario", "observaciones",
        "modo_moneda", "tasa_usd", "tasa_eur", "total_kilos", "total_ref",
        "total_bs", "dia_semana", "fecha", "hora"
    )
    valores = (
        correlativo, tipo_doc, subtipo, origen_norm, destino_norm,
        doc_identidad, referencia, motivo, chofer, cedula_chofer,
        placa_vehiculo, afecto_saldo, afecto_inventario, observaciones,
        modo_moneda, tasa_usd, tasa_eur, tot_kg, tot_ref,
        tot_bs, dia_sem, fecha_str, hora_str
    )
    comodines = ", ".join(["?"] * len(valores))

    with conectar_db() as conn:
        cur = conn.execute(f"INSERT INTO documentos ({', '.join(columnas)}) VALUES ({comodines})", valores)
        doc_id = cur.lastrowid
        for lin in lineas_carrito:
            p_nom = str(lin["producto"])
            kg_l = float(lin["kilos"])
            conn.execute(
                "INSERT INTO documento_items (documento_id, producto, escala, kilos, precio_unitario, subtotal_ref) VALUES (?, ?, ?, ?, ?, ?)",
                (doc_id, p_nom, lin["escala"], kg_l, lin["precio_unitario"], lin["subtotal_ref"]),
            )
            if afecto_inventario == 1:
                if tipo_doc == "DESPACHO":
                    ajustar_stock_sede(conn, origen_norm, p_nom, -kg_l)
                elif tipo_doc in ("MOVIMIENTO", "NOTA_DEBITO"):
                    ajustar_stock_sede(conn, origen_norm, p_nom, -kg_l)
                    ajustar_stock_sede(conn, destino_norm, p_nom, +kg_l)
        conn.commit()
    return doc_id


# ==============================================================================
# PASO 4: INTERFAZ COMPLETA DE LOS 7 MÓDULOS EN PESTAÑAS SUPERIORES
# ==============================================================================
def ejecutar_interfaz_principal() -> None:
    asegurar_logo()
    inicializar_base_de_datos()

    icono_pag = str(RUTA_LOGO) if RUTA_LOGO.exists() else "🚜"
    st.set_page_config(
        page_title="Agropecuaria Óptima | Sistema Integral",
        page_icon=icono_pag,
        layout="wide",
        initial_sidebar_state="collapsed",
    )

    st.markdown(
        """
        <meta name="apple-mobile-web-app-capable" content="yes">
        <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
        <meta name="apple-mobile-web-app-title" content="AgroÓptima">
        <meta name="mobile-web-app-capable" content="yes">
        <style>
        .stApp { background-color: #0D2618; color: #ECFDF5; }
        .stButton>button, .stDownloadButton>button {
            width: 100%; border-radius: 8px; font-weight: bold; height: 2.9em;
        }
        button[data-baseweb="tab"] {
            font-size: 15px !important;
            font-weight: bold !important;
            padding: 10px 14px !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    for k_ses in ("carrito_desp", "carrito_nc", "carrito_nd", "ultimo_doc_id"):
        if k_ses not in st.session_state:
            if k_ses.startswith("carrito"):
                st.session_state[k_ses] = []
            else:
                st.session_state[k_ses] = None

    # Pantalla de bloqueo por PIN (a1b2c3)
    if PIN_ACCESO_MOVIL:
        if "autenticado" not in st.session_state:
            st.session_state["autenticado"] = False
        if not st.session_state["autenticado"]:
            if RUTA_LOGO.exists():
                st.image(str(RUTA_LOGO), width=200)
            st.subheader("🔒 Sistema Integral | Agropecuaria Óptima C.A")
            pin_in = st.text_input("Ingresa tu PIN de seguridad:", type="password")
            if st.button("🔓 Desbloquear Sistema Integral"):
                if pin_in.strip() == PIN_ACCESO_MOVIL:
                    st.session_state["autenticado"] = True
                    st.rerun()
                else:
                    st.error("PIN incorrecto. Verifica e intenta de nuevo.")
            st.stop()

    # Encabezado y Tasas BCV / Euro
    tasas_act = obtener_tasas_db()
    c_lg, c_ti = st.columns([1, 4])
    with c_lg:
        if RUTA_LOGO.exists():
            st.image(str(RUTA_LOGO), width=130)
    with c_ti:
        st.markdown("### 🚜 SISTEMA INTEGRAL: PRECIOS, DESPACHOS E INVENTARIO")
        st.caption("Agropecuaria Óptima C.A  •  7 Módulos Unificados en Vivo (iPhone • Android • Laptop)")

    with st.expander("💱 Tasas BCV / Euro y Respaldo de Base de Datos (Toca aquí para abrir)", expanded=False):
        cm, cu, ce = st.columns(3)
        modo_moneda = cm.selectbox(
            "Moneda para PDFs / Excels",
            [
                "REF Tasa BCV ($)", "REF Tasa Euro (€)",
                "Bolívares (Tasa BCV $)", "Bolívares (Tasa Euro €)",
                "Doble: REF ($) + Bs BCV", "Doble: REF (€) + Bs Euro",
            ],
        )
        tasa_usd = cu.number_input("💵 Tasa BCV ($)", value=float(tasas_act["bcv_usd"]), step=0.5, format="%.2f")
        tasa_eur = ce.number_input("💶 Tasa Euro (€)", value=float(tasas_act["bcv_eur"]), step=0.5, format="%.2f")

        bf, bb = st.columns(2)
        if bf.button("💾 Fijar Tasas Manuales"):
            guardar_tasas_db(tasa_usd, tasa_eur, "Manual")
            st.success("Tasas guardadas correctamente.")
            st.rerun()
        if bb.button("🔄 Sincronizar con BCV en Vivo"):
            ok, uv, ev = sincronizar_tasas_bcv_web()
            if ok:
                st.success(f"BCV actualizado: $ {uv:.2f} | € {ev:.2f}")
                st.rerun()
            else:
                st.warning("Sin respuesta del BCV; puedes fijarla manualmente.")

        st.markdown("---")
        cs1, cs2 = st.columns(2)
        if RUTA_DB.exists():
            cs1.download_button(
                "📥 Descargar Copia de Base de Datos (.DB)",
                data=RUTA_DB.read_bytes(),
                file_name="sistema_finca_optima.db",
                mime="application/octet-stream",
            )
        db_subida = cs2.file_uploader("📤 Subir Base de Datos (.db) desde Laptop", type=["db"])
        if db_subida is not None and st.button("✅ Aplicar Base de Datos Subida"):
            RUTA_DB.write_bytes(db_subida.getvalue())
            st.success("¡Base de datos actualizada en todos tus dispositivos!")
            st.rerun()

    tasa_apl, simb_mon, nom_tasa = obtener_tasa_activa(modo_moneda, tasa_usd, tasa_eur)
    st.info(f"📌 **Moneda Activa:** {modo_moneda}  |  **💵 BCV ($):** Bs {tasa_usd:,.2f}  |  **💶 Euro (€):** Bs {tasa_eur:,.2f}")

    prods_all = consultar_sql("SELECT * FROM productos ORDER BY nombre ASC")
    mapa_prods = {p["nombre"]: p for p in prods_all}
    centros_all = consultar_sql("SELECT * FROM centros_distribucion ORDER BY codigo ASC")
    mapa_centros = {f"{d['codigo']} - {d['nombre']}": d for d in centros_all}
    todas_sedes = [SEDE_PLANTA_PRINCIPAL] + list(mapa_centros.keys())

    def renderizar_carrito(clave_sesion: str, escala_defecto: str = "Cliente Final") -> None:
        st.markdown("#### 🛒 Carrito de Productos (Kilos x Precio Unitario)")
        nombres_p = list(mapa_prods.keys())
        if not nombres_p:
            st.warning("No hay productos registrados en el catálogo.")
            return
        c1, c2 = st.columns(2)
        prod_sel = c1.selectbox("Producto o Corte", nombres_p, key=f"p_{clave_sesion}")
        lista_esc = list(ESCALAS_PRECIO.keys())
        idx_e = lista_esc.index(escala_defecto) if escala_defecto in ESCALAS_PRECIO else 0
        esc_sel = c2.selectbox("Escala Manual", lista_esc, index=idx_e, key=f"e_{clave_sesion}")

        precio_sug = float(mapa_prods[prod_sel][ESCALAS_PRECIO[esc_sel]])
        c3, c4 = st.columns(2)
        kilos_in = c3.number_input("Cantidad en Kilos (KG)", min_value=0.0, value=10.0, step=1.0, format="%.2f", key=f"k_{clave_sesion}")
        precio_in = c4.number_input(f"Precio Unit. ({simb_mon})", min_value=0.0, value=precio_sug, step=0.1, format="%.2f", key=f"u_{clave_sesion}")

        b_add, b_rem, b_clr = st.columns(3)
        if b_add.button("➕ Agregar Rubro", key=f"add_{clave_sesion}") and kilos_in > 0:
            st.session_state[clave_sesion].append({
                "producto": prod_sel,
                "escala": esc_sel,
                "kilos": round(kilos_in, 2),
                "precio_unitario": round(precio_in, 2),
                "subtotal_ref": round(kilos_in * precio_in, 2),
            })
            st.rerun()
        if b_rem.button("🗑️ Quitar Último", key=f"rem_{clave_sesion}") and st.session_state[clave_sesion]:
            st.session_state[clave_sesion].pop()
            st.rerun()
        if b_clr.button("🧹 Vaciar Lista", key=f"clr_{clave_sesion}"):
            st.session_state[clave_sesion] = []
            st.rerun()

        lineas = st.session_state[clave_sesion]
        if lineas:
            st.dataframe(lineas, use_container_width=True)
            t_kg = sum(float(x["kilos"]) for x in lineas)
            t_rf = sum(float(x["subtotal_ref"]) for x in lineas)
            st.success(f"**PESO TOTAL:** {t_kg:,.2f} KG  |  **TOTAL REF:** {simb_mon} {t_rf:,.2f}  |  **TOTAL Bs:** Bs {t_rf * tasa_apl:,.2f}")

    def mostrar_descargas_doc(doc_id: int, prefijo_key: str) -> None:
        pdf_b, nom_pdf = generar_bytes_documento(doc_id, en_excel=False)
        xls_b, nom_xls = generar_bytes_documento(doc_id, en_excel=True)
        d1, d2 = st.columns(2)
        with d1:
            st.download_button(f"📄 Descargar PDF ({nom_pdf})", data=pdf_b, file_name=nom_pdf, mime="application/pdf", key=f"dl_pdf_{prefijo_key}_{doc_id}")
        with d2:
            st.download_button(f"📊 Descargar Excel ({nom_xls})", data=xls_b, file_name=nom_xls, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"dl_xls_{prefijo_key}_{doc_id}")

    # ==========================================================================
    # LAS 7 PESTAÑAS VISIBLES EN LA PARTE SUPERIOR (IGUAL QUE EN EL ESCRITORIO)
    # ==========================================================================
    (
        tab_listas,
        tab_despachos,
        tab_credito,
        tab_debito_mov,
        tab_inventario,
        tab_bases_datos,
        tab_historial,
    ) = st.tabs([
        "📋 1. Listas de Precios",
        "📦 2. Órdenes de Despacho",
        "💳 3. Notas de Crédito",
        "🔄 4. Débitos y Movimientos",
        "📊 5. Inventario por Sede",
        "👥 6. Clientes, Centros y Choferes",
        "📜 7. Historial General",
    ])

    # --------------------------------------------------------------------------
    # MÓDULO 1: LISTAS DE PRECIOS Y CHECKLIST MANUAL (SIN 'NONE')
    # --------------------------------------------------------------------------
    with tab_listas:
        st.subheader("📋 Generador de Listas de Precios y Checklist Manual")
        c1, c2 = st.columns(2)
        mod_lista = c1.selectbox("Precios Elegibles para el PDF / Excel", MODALIDADES_LISTA_PRECIOS)
        rub_lista = c2.selectbox("Filtrar por Categoría", ["Todos los Rubros"] + RUBROS_CATALOGO)
        solo_disp = st.checkbox("Incluir solo productos marcados como ☑ Disponibles", value=True)

        bytes_lista_pdf = generar_bytes_lista_precios(False, mod_lista, rub_lista, modo_moneda, tasa_usd, tasa_eur, solo_disp)
        bytes_lista_xls = generar_bytes_lista_precios(True, mod_lista, rub_lista, modo_moneda, tasa_usd, tasa_eur, solo_disp)

        dl1, dl2 = st.columns(2)
        with dl1:
            st.download_button(
                "📄 Descargar Lista de Precios (PDF)",
                data=bytes_lista_pdf,
                file_name=f"LISTA_PRECIOS_{datetime.now().strftime('%d%m%Y')}.pdf",
                mime="application/pdf",
            )
        with dl2:
            st.download_button(
                "📊 Descargar Lista de Precios (Excel)",
                data=bytes_lista_xls,
                file_name=f"LISTA_PRECIOS_{datetime.now().strftime('%d%m%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        st.divider()
        st.markdown("#### ☑️ Checklist de Productos Disponibles (Toca para marcar o quitar sin borrar)")

        bm1, bm2 = st.columns(2)
        if bm1.button("✅ Marcar Todos Disponibles"):
            ejecutar_sql("UPDATE productos SET disponible = 1")
            st.rerun()
        if bm2.button("⬜ Desmarcar Todos"):
            ejecutar_sql("UPDATE productos SET disponible = 0")
            st.rerun()

        if rub_lista == "Todos los Rubros":
            productos_vista = consultar_sql("SELECT * FROM productos ORDER BY id ASC")
        else:
            productos_vista = consultar_sql("SELECT * FROM productos WHERE rubro=? ORDER BY id ASC", (rub_lista,))

        for p in productos_vista:
            est = bool(int(p["disponible"]) == 1)
            p_cli_bs = float(p["precio_cliente"]) * tasa_apl
            lbl = (
                f"{p['nombre']} ({p['rubro']})  —  "
                f"Mayor: {simb_mon}{p['precio_mayor']:.2f} | "
                f"Vend 10%: {simb_mon}{p['precio_vendedor']:.2f} | "
                f"Cliente: {simb_mon}{p['precio_cliente']:.2f} (Bs {p_cli_bs:,.2f})"
            )
            n_est = st.checkbox(lbl, value=est, key=f"chk_p_{p['id']}")
            if n_est != est:
                nuevo_val = 1 if n_est else 0
                ejecutar_sql("UPDATE productos SET disponible = ? WHERE id = ?", (nuevo_val, p["id"]))
                st.rerun()

        st.divider()
        col_ed, col_mas = st.columns(2)
        with col_ed:
            st.markdown("#### 🥩 Agregar, Editar o Eliminar Producto")
            p_ed = st.selectbox("Seleccionar corte del catálogo:", ["-- Nuevo Producto --"] + list(mapa_prods.keys()))
            po = mapa_prods.get(p_ed)
            np = st.text_input("Nombre del Producto", value=po["nombre"] if po else "").strip().upper()
            idx_r = RUBROS_CATALOGO.index(po["rubro"]) if (po and po["rubro"] in RUBROS_CATALOGO) else 0
            rp = st.selectbox("Categoría / Rubro", RUBROS_CATALOGO, index=idx_r)
            cp1, cp2, cp3 = st.columns(3)
            pm = cp1.number_input("Mayor ($)", value=float(po["precio_mayor"]) if po else 5.0, format="%.2f")
            pv = cp2.number_input("Vendedor ($)", value=float(po["precio_vendedor"]) if po else 5.5, format="%.2f")
            pc = cp3.number_input("Cliente ($)", value=float(po["precio_cliente"]) if po else 6.0, format="%.2f")
            bg, be = st.columns(2)
            if bg.button("💾 Guardar / Actualizar Producto") and np:
                ejecutar_sql(
                    """INSERT INTO productos (nombre, rubro, precio_mayor, precio_vendedor, precio_cliente, disponible)
                       VALUES (?, ?, ?, ?, ?, 1)
                       ON CONFLICT(nombre) DO UPDATE SET rubro=excluded.rubro, precio_mayor=excluded.precio_mayor,
                       precio_vendedor=excluded.precio_vendedor, precio_cliente=excluded.precio_cliente""",
                    (np, rp, round(pm, 2), round(pv, 2), round(pc, 2)),
                )
                st.success(f"Producto '{np}' actualizado.")
                st.rerun()
            if po and be.button("🗑️ Eliminar Producto del Catálogo"):
                ejecutar_sql("DELETE FROM productos WHERE id=?", (po["id"],))
                st.rerun()

        with col_mas:
            st.markdown("#### ⚡ Ajuste Porcentual Masivo (%)")
            porc_in = st.number_input("Porcentaje a subir o bajar (Ej. 5 o -5)", value=0.0, step=1.0, format="%.2f")
            if st.button("📈 Aplicar Porcentaje a Todo el Catálogo") and porc_in != 0.0:
                factor = 1.0 + (porc_in / 100.0)
                ejecutar_sql(
                    """UPDATE productos SET
                       precio_mayor = ROUND(precio_mayor * ?, 2),
                       precio_vendedor = ROUND(precio_vendedor * ?, 2),
                       precio_cliente = ROUND(precio_cliente * ?, 2)""",
                    (factor, factor, factor),
                )
                st.success(f"Se aplicó {porc_in:+.2f}% a todos los precios.")
                st.rerun()

    # --------------------------------------------------------------------------
    # MÓDULO 2: ÓRDENES DE DESPACHO
    # --------------------------------------------------------------------------
    with tab_despachos:
        st.subheader("📦 Emitir Orden de Despacho a Cliente")
        clientes_db = consultar_sql("SELECT * FROM clientes ORDER BY nombre ASC")
        mapa_cli = {f"{c['nombre']} ({c['documento_id']})": c for c in clientes_db}
        cli_sel = st.selectbox("Cliente Registrado", list(mapa_cli.keys()))
        cli_obj = mapa_cli[cli_sel]

        with st.expander("⚡ Registro Rápido de Cliente Nuevo"):
            rn = st.text_input("Nombre o Razón Social del Cliente").strip().upper()
            rr = st.text_input("Cédula o RIF del Cliente").strip().upper()
            re = st.selectbox("Escala Preferida", list(ESCALAS_PRECIO.keys()))
            if st.button("💾 Guardar Cliente Nuevo") and rn:
                ejecutar_sql("INSERT INTO clientes (nombre, documento_id, escala_preferida) VALUES (?, ?, ?)", (rn, rr or "S/D", re))
                st.success(f"Cliente '{rn}' guardado.")
                st.rerun()

        sede_od = st.selectbox("Sede que Despacha (Descuenta Stock)", todas_sedes)
        obs_od = st.text_input("Observaciones de Entrega")
        desc_inv = st.checkbox("Descontar kilogramos del inventario de la sede", value=True)

        st.divider()
        renderizar_carrito("carrito_desp", cli_obj["escala_preferida"])

        if st.button("🚀 EMITIR ORDEN DE DESPACHO"):
            if not st.session_state["carrito_desp"]:
                st.warning("Agrega al menos un producto al carrito.")
            else:
                did = guardar_operacion_db(
                    "DESPACHO", "ORDEN DE DESPACHO", sede_od, cli_obj["nombre"], cli_obj["documento_id"],
                    "", f"Escala: {cli_obj['escala_preferida']}", "", "", "", 0,
                    1 if desc_inv else 0, obs_od, modo_moneda, tasa_usd, tasa_eur, st.session_state["carrito_desp"]
                )
                st.session_state["carrito_desp"] = []
                st.session_state["ultimo_doc_id"] = did
                st.success("¡Orden de Despacho emitida con éxito!")

        if st.session_state["ultimo_doc_id"]:
            mostrar_descargas_doc(st.session_state["ultimo_doc_id"], "desp")

    # --------------------------------------------------------------------------
    # MÓDULO 3: NOTAS DE CRÉDITO (CD)
    # --------------------------------------------------------------------------
    with tab_credito:
        st.subheader("💳 Nota de Crédito a Centro de Distribución")
        cd_sel = st.selectbox("Centro de Distribución Beneficiario", list(mapa_centros.keys()))
        mot_nc = st.selectbox("Motivo de Acreditación", ["Ajuste de Precio / Descuento Comercial", "Merma de Peso en Recepción", "Diferencia en Guía", "Bonificación Operativa"])
        ref_nc = st.text_input("Documento o Guía de Referencia")
        obs_nc = st.text_input("Observaciones de la Nota de Crédito")
        af_s = st.checkbox("Descontar monto del saldo del Centro", value=True)

        st.divider()
        renderizar_carrito("carrito_nc", "Al Mayor")

        if st.button("🚀 GENERAR NOTA DE CRÉDITO"):
            if not st.session_state["carrito_nc"]:
                st.warning("Agrega al menos un producto a acreditar.")
            else:
                cd_obj = mapa_centros[cd_sel]
                t_ref = round(sum(float(x["subtotal_ref"]) for x in st.session_state["carrito_nc"]), 2)
                if af_s:
                    ejecutar_sql("UPDATE centros_distribucion SET saldo_pendiente_ref = saldo_pendiente_ref - ? WHERE id=?", (t_ref, cd_obj["id"]))
                did = guardar_operacion_db(
                    "NOTA_CREDITO", "NOTA DE CREDITO - CENTRO DE DISTRIBUCION", SEDE_PLANTA_PRINCIPAL,
                    cd_sel, cd_obj["codigo"], ref_nc, mot_nc, "", "", "",
                    1 if af_s else 0, 0, obs_nc, modo_moneda, tasa_usd, tasa_eur, st.session_state["carrito_nc"]
                )
                st.session_state["carrito_nc"] = []
                st.session_state["ultimo_doc_id"] = did
                st.success("¡Nota de Crédito registrada correctamente!")

        if st.session_state["ultimo_doc_id"]:
            mostrar_descargas_doc(st.session_state["ultimo_doc_id"], "nc")

    # --------------------------------------------------------------------------
    # MÓDULO 4: DÉBITOS Y MOVIMIENTOS (ROTACIÓN Y TRANSPORTE)
    # --------------------------------------------------------------------------
    with tab_debito_mov:
        st.subheader("🔄 Notas de Débito y Rotación de Inventario entre Sedes")
        modo_op = st.selectbox("Modalidad de Operación", ["Movimiento entre Centros de Distribución", "Nota de Débito (Devolución NO Vendida)"])
        c_or, c_de = st.columns(2)
        orig_nd = c_or.selectbox("Sede Origen (Resta Stock -)", todas_sedes)
        dest_nd = c_de.selectbox("Sede Destino (Suma Stock +)", todas_sedes, index=min(1, len(todas_sedes) - 1))
        mot_nd = st.selectbox("Motivo del Registro", ["Rotación de Inventario Frío entre Sedes", "Abastecimiento desde Planta a Centro", "Retorno de Mercancía NO Vendida", "Cargo por Devolución"])

        st.markdown("#### 🚛 Datos de Transporte en Guía (Preselección de Choferes)")
        choferes_db = consultar_sql("SELECT * FROM choferes ORDER BY nombre ASC")
        mapa_ch = {f"{ch['nombre']} | Placa: {ch['placa']}": ch for ch in choferes_db}
        pre_ch = st.selectbox("Preseleccionar Chofer Registrado", [OPCION_CHOFER_MANUAL] + list(mapa_ch.keys()))
        ch_obj = mapa_ch.get(pre_ch)

        ct1, ct2, ct3 = st.columns(3)
        nom_ch = ct1.text_input("Nombre del Chofer", value=ch_obj["nombre"] if ch_obj else "").strip().upper()
        ced_ch = ct2.text_input("Cédula del Chofer", value=ch_obj["cedula"] if ch_obj else "").strip().upper()
        pla_ch = ct3.text_input("Placa del Vehículo", value=ch_obj["placa"] if ch_obj else "").strip().upper()
        obs_nd = st.text_input("Observaciones de Traslado / Cadena de Frío")
        af_nd = st.checkbox("Afectar saldo contable del Centro (en Nota de Débito)", value=True)

        st.divider()
        renderizar_carrito("carrito_nd", "Al Mayor")

        if st.button("🚀 EMITIR GUÍA Y ROTAR INVENTARIO"):
            if not st.session_state["carrito_nd"]:
                st.warning("Agrega productos a la lista para trasladar o devolver.")
            elif normalizar_nombre_sede(orig_nd) == normalizar_nombre_sede(dest_nd):
                st.error("La Sede Origen y la Sede Destino deben ser diferentes.")
            else:
                es_dev = "Devolución" in modo_op
                t_ref = round(sum(float(x["subtotal_ref"]) for x in st.session_state["carrito_nd"]), 2)
                af_s = 1 if (es_dev and af_nd) else 0
                if af_s and orig_nd in mapa_centros:
                    ejecutar_sql("UPDATE centros_distribucion SET saldo_pendiente_ref = saldo_pendiente_ref + ? WHERE id=?", (t_ref, mapa_centros[orig_nd]["id"]))
                did = guardar_operacion_db(
                    "NOTA_DEBITO" if es_dev else "MOVIMIENTO",
                    "NOTA DE DEBITO - DEVOLUCION MERCANCIA" if es_dev else "GUIA DE MOVIMIENTO ENTRE CENTROS",
                    orig_nd, dest_nd, "", pla_ch, mot_nd, nom_ch, ced_ch, pla_ch,
                    af_s, 1, obs_nd, modo_moneda, tasa_usd, tasa_eur, st.session_state["carrito_nd"]
                )
                st.session_state["carrito_nd"] = []
                st.session_state["ultimo_doc_id"] = did
                st.success(f"¡Inventario rotado exitosamente de '{orig_nd}' a '{dest_nd}'!")

        if st.session_state["ultimo_doc_id"]:
            mostrar_descargas_doc(st.session_state["ultimo_doc_id"], "nd")

    # --------------------------------------------------------------------------
    # MÓDULO 5: INVENTARIO POR SEDE
    # --------------------------------------------------------------------------
    with tab_inventario:
        st.subheader("📊 Control de Inventario en Vivo por Sede")
        c_f1, c_f2 = st.columns(2)
        sede_f_inv = c_f1.selectbox("Filtrar por Sede", ["Todas las Sedes"] + todas_sedes)
        est_f_inv = c_f2.selectbox("Mostrar Productos", ["Con Movimiento o Stock Registrado", "Catálogo Completo (22 Productos)", "Solo Stock Positivo"])

        with st.expander("⚖️ Carga Inicial / Ajuste Manual o Eliminación de Stock", expanded=False):
            s_aj = st.selectbox("Sede o Centro de Distribución", todas_sedes)
            p_aj = st.selectbox("Producto / Corte", list(mapa_prods.keys()))
            t_aj = st.selectbox("Tipo de Operación", ["Sumar Entrada (+ KG)", "Restar Salida (- KG)", "Fijar Stock Exacto (= KG)"])
            k_aj = st.number_input("Kilos (KG)", min_value=0.0, value=50.0, step=5.0, format="%.2f")

            b1, b2, b3 = st.columns(3)
            if b1.button("💾 Aplicar Stock"):
                if "Fijar" in t_aj:
                    fijar_stock_exacto_sede(s_aj, p_aj, k_aj)
                elif "Restar" in t_aj:
                    with conectar_db() as c:
                        ajustar_stock_sede(c, s_aj, p_aj, -abs(k_aj))
                        c.commit()
                else:
                    with conectar_db() as c:
                        ajustar_stock_sede(c, s_aj, p_aj, +abs(k_aj))
                        c.commit()
                st.success("Stock actualizado.")
                st.rerun()

            if b2.button("🧹 Poner en 0 KG"):
                fijar_stock_exacto_sede(s_aj, p_aj, 0.0)
                st.rerun()

            if b3.button("🗑️ Eliminar Fila"):
                ejecutar_sql(
                    "DELETE FROM inventario_sedes WHERE sede = ? AND producto = ?",
                    (normalizar_nombre_sede(s_aj), p_aj.strip().upper()),
                )
                st.success("Registro eliminado del inventario.")
                st.rerun()

        if sede_f_inv == "Todas las Sedes":
            regs_inv = consultar_sql("SELECT * FROM inventario_sedes ORDER BY sede, producto")
        else:
            regs_inv = consultar_sql("SELECT * FROM inventario_sedes WHERE sede=? ORDER BY producto", (sede_f_inv,))

        filas_vista: List[Dict[str, object]] = []
        if "Catálogo Completo" in est_f_inv:
            sedes_obj = todas_sedes if sede_f_inv == "Todas las Sedes" else [sede_f_inv]
            mapa_ex = {(r["sede"], r["producto"]): r for r in regs_inv}
            for s_n in sedes_obj:
                for p_n, p_row in mapa_prods.items():
                    rg = mapa_ex.get((s_n, p_n))
                    kg = float(rg["kilos_disponibles"]) if rg else 0.0
                    act = rg["ultima_actualizacion"] if rg else "Sin movimientos"
                    if kg <= 0:
                        est_txt = "🔴 Sin Stock"
                    elif kg < UMBRAL_STOCK_BAJO_KG:
                        est_txt = "🟡 Bajo Stock"
                    else:
                        est_txt = "🟢 En Stock"
                    vr = round(kg * float(p_row["precio_mayor"]), 2)
                    filas_vista.append({"sede": s_n, "producto": p_n, "kilos": round(kg, 2), "estado": est_txt, "valor_ref": vr, "actualizado": act})
        else:
            for r in regs_inv:
                kg = round(float(r["kilos_disponibles"]), 2)
                if "Positivo" in est_f_inv and kg <= 0:
                    continue
                p_row = mapa_prods.get(r["producto"])
                pm = float(p_row["precio_mayor"]) if p_row else 0.0
                if kg <= 0:
                    est_txt = "🔴 Sin Stock"
                elif kg < UMBRAL_STOCK_BAJO_KG:
                    est_txt = "🟡 Bajo Stock"
                else:
                    est_txt = "🟢 En Stock"
                filas_vista.append({"sede": r["sede"], "producto": r["producto"], "kilos": kg, "estado": est_txt, "valor_ref": round(kg * pm, 2), "actualizado": r["ultima_actualizacion"]})

        if filas_vista:
            st.dataframe(filas_vista, use_container_width=True)
            tk_inv = sum(float(x["kilos"]) for x in filas_vista)
            tr_inv = sum(float(x["valor_ref"]) for x in filas_vista)
            st.success(f"**STOCK TOTAL:** {tk_inv:,.2f} KG  |  **VALOR MAYOR:** {simb_mon} {tr_inv:,.2f}  |  **TOTAL Bs:** Bs {tr_inv * tasa_apl:,.2f}")

            inv_pdf = generar_bytes_inventario(False, sede_f_inv, filas_vista, modo_moneda, tasa_usd, tasa_eur)
            inv_xls = generar_bytes_inventario(True, sede_f_inv, filas_vista, modo_moneda, tasa_usd, tasa_eur)
            di1, di2 = st.columns(2)
            with di1:
                st.download_button(
                    "📄 Descargar Inventario (PDF)",
                    data=inv_pdf,
                    file_name=f"INVENTARIO_{datetime.now().strftime('%d%m%Y')}.pdf",
                    mime="application/pdf",
                )
            with di2:
                st.download_button(
                    "📊 Descargar Inventario (Excel)",
                    data=inv_xls,
                    file_name=f"INVENTARIO_{datetime.now().strftime('%d%m%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        else:
            st.info("Aún no hay kilos cargados en esta vista. Abre el cuadro de 'Carga Inicial / Ajuste Manual' arriba o elige 'Catálogo Completo'.")

    # --------------------------------------------------------------------------
    # MÓDULO 6: BASES DE DATOS (CLIENTES, CENTROS Y CHOFERES)
    # --------------------------------------------------------------------------
    with tab_bases_datos:
        st.subheader("👥 Bases de Datos Maestras: Clientes, Centros y Choferes")
        sub_cli, sub_cd, sub_ch = st.tabs(["👤 1. Clientes", "🏢 2. Centros de Distribución", "🚛 3. Choferes y Vehículos"])

        with sub_cli:
            clientes_l = consultar_sql("SELECT * FROM clientes ORDER BY nombre ASC")
            mapa_c_ed = {f"{c['nombre']} ({c['documento_id']})": c for c in clientes_l}
            sel_c = st.selectbox("Seleccionar Cliente para Editar o Borrar:", ["-- Nuevo Cliente --"] + list(mapa_c_ed.keys()))
            co = mapa_c_ed.get(sel_c)

            cn = st.text_input("Nombre o Razón Social", value=co["nombre"] if co else "", key="bd_cn").strip().upper()
            cr = st.text_input("Cédula o RIF", value=co["documento_id"] if co else "", key="bd_cr").strip().upper()
            lista_e = list(ESCALAS_PRECIO.keys())
            idx_ce = lista_e.index(co["escala_preferida"]) if (co and co["escala_preferida"] in ESCALAS_PRECIO) else 0
            ce = st.selectbox("Escala Preferida", lista_e, index=idx_ce, key="bd_ce")
            ct = st.text_input("Teléfono", value=co["telefono"] if co else "", key="bd_ct")
            cd_dir = st.text_input("Dirección", value=co["direccion"] if co else "", key="bd_cdir")

            bc1, bc2 = st.columns(2)
            if bc1.button("💾 Guardar / Actualizar Cliente") and cn:
                if co:
                    ejecutar_sql("UPDATE clientes SET nombre=?, documento_id=?, escala_preferida=?, telefono=?, direccion=? WHERE id=?", (cn, cr or "S/D", ce, ct, cd_dir, co["id"]))
                else:
                    ejecutar_sql("INSERT INTO clientes (nombre, documento_id, escala_preferida, telefono, direccion) VALUES (?, ?, ?, ?, ?)", (cn, cr or "S/D", ce, ct, cd_dir))
                st.rerun()
            if co and bc2.button("🗑️ Eliminar Cliente"):
                ejecutar_sql("DELETE FROM clientes WHERE id=?", (co["id"],))
                st.rerun()
            st.dataframe([dict(r) for r in clientes_l], use_container_width=True)

        with sub_cd:
            sel_d = st.selectbox("Seleccionar Centro para Editar o Borrar:", ["-- Nuevo Centro --"] + list(mapa_centros.keys()))
            do = mapa_centros.get(sel_d)
            dcod = st.text_input("Código Sede (Ej. CD-03)", value=do["codigo"] if do else "").strip().upper()
            dnom = st.text_input("Nombre del Centro", value=do["nombre"] if do else "").strip().upper()
            denc = st.text_input("Encargado", value=do["encargado"] if do else "")
            dtel = st.text_input("Teléfono Centro", value=do["telefono"] if do else "")
            dubi = st.text_input("Ubicación", value=do["ubicacion"] if do else "")

            bd1, bd2 = st.columns(2)
            if bd1.button("💾 Guardar / Actualizar Centro") and dcod and dnom:
                if do:
                    ejecutar_sql("UPDATE centros_distribucion SET codigo=?, nombre=?, encargado=?, telefono=?, ubicacion=? WHERE id=?", (dcod, dnom, denc, dtel, dubi, do["id"]))
                else:
                    ejecutar_sql("INSERT OR IGNORE INTO centros_distribucion (codigo, nombre, encargado, telefono, ubicacion) VALUES (?, ?, ?, ?, ?)", (dcod, dnom, denc, dtel, dubi))
                st.rerun()
            if do and bd2.button("🗑️ Eliminar Centro"):
                ejecutar_sql("DELETE FROM centros_distribucion WHERE id=?", (do["id"],))
                st.rerun()
            st.dataframe([dict(r) for r in centros_all], use_container_width=True)

        with sub_ch:
            ch_l = consultar_sql("SELECT * FROM choferes ORDER BY nombre ASC")
            mapa_ch_ed = {f"{ch['nombre']} ({ch['placa']})": ch for ch in ch_l}
            sel_h = st.selectbox("Seleccionar Chofer para Editar o Borrar:", ["-- Nuevo Chofer --"] + list(mapa_ch_ed.keys()))
            ho = mapa_ch_ed.get(sel_h)

            hnom = st.text_input("Nombre y Apellido del Chofer", value=ho["nombre"] if ho else "").strip().upper()
            hced = st.text_input("Cédula del Chofer", value=ho["cedula"] if ho else "").strip().upper()
            hpla = st.text_input("Placa del Vehículo", value=ho["placa"] if ho else "").strip().upper()
            htel = st.text_input("Teléfono del Chofer", value=ho["telefono"] if ho else "")
            hveh = st.text_input("Descripción del Camión / Vehículo", value=ho["vehiculo"] if ho else "")

            bh1, bh2 = st.columns(2)
            if bh1.button("💾 Guardar / Actualizar Chofer") and hnom and hced and hpla:
                if ho:
                    ejecutar_sql("UPDATE choferes SET nombre=?, cedula=?, placa=?, telefono=?, vehiculo=? WHERE id=?", (hnom, hced, hpla, htel, hveh, ho["id"]))
                else:
                    ejecutar_sql("INSERT INTO choferes (nombre, cedula, placa, telefono, vehiculo) VALUES (?, ?, ?, ?, ?)", (hnom, hced, hpla, htel, hveh))
                st.rerun()
            if ho and bh2.button("🗑️ Eliminar Chofer"):
                ejecutar_sql("DELETE FROM choferes WHERE id=?", (ho["id"],))
                st.rerun()
            st.dataframe([dict(r) for r in ch_l], use_container_width=True)

    # --------------------------------------------------------------------------
    # MÓDULO 7: HISTORIAL GENERAL DE DOCUMENTOS
    # --------------------------------------------------------------------------
    with tab_historial:
        st.subheader("📜 Historial General de Documentos")
        f_hist = st.selectbox("Filtrar por Categoría", ["Todos", "DESPACHO", "NOTA_CREDITO", "NOTA_DEBITO", "MOVIMIENTO"])
        if f_hist == "Todos":
            docs_h = consultar_sql("SELECT * FROM documentos ORDER BY id DESC")
        else:
            docs_h = consultar_sql("SELECT * FROM documentos WHERE tipo_doc=? ORDER BY id DESC", (f_hist,))

        if not docs_h:
            st.info("Todavía no hay documentos emitidos en esta categoría.")
        else:
            mapa_docs = {f"{d['correlativo']} | {d['entidad_destino']} | {d['total_kilos']:.2f} kg ({d['fecha']})": d for d in docs_h}
            doc_sel_h = st.selectbox("Selecciona un Documento para Descargar PDF/Excel o Eliminar:", list(mapa_docs.keys()))
            doc_row = mapa_docs[doc_sel_h]

            mostrar_descargas_doc(doc_row["id"], "hist")

            if st.button("🗑️ Eliminar este Documento (Revierte Stock y Saldo)"):
                items_d = consultar_sql("SELECT * FROM documento_items WHERE documento_id=?", (doc_row["id"],))
                with conectar_db() as conn:
                    if doc_row["afecto_saldo"] == 1:
                        if doc_row["tipo_doc"] == "NOTA_CREDITO" and doc_row["doc_identidad"]:
                            conn.execute("UPDATE centros_distribucion SET saldo_pendiente_ref = saldo_pendiente_ref + ? WHERE codigo=?", (doc_row["total_ref"], doc_row["doc_identidad"]))
                        elif doc_row["tipo_doc"] == "NOTA_DEBITO":
                            cod_or = doc_row["entidad_origen"].split(" - ")[0].strip()
                            conn.execute("UPDATE centros_distribucion SET saldo_pendiente_ref = saldo_pendiente_ref - ? WHERE codigo=?", (doc_row["total_ref"], cod_or))
                    if doc_row["afecto_inventario"] == 1:
                        for it in items_d:
                            kg_i = float(it["kilos"])
                            if doc_row["tipo_doc"] == "DESPACHO":
                                ajustar_stock_sede(conn, doc_row["entidad_origen"], it["producto"], +kg_i)
                            elif doc_row["tipo_doc"] in ("MOVIMIENTO", "NOTA_DEBITO"):
                                ajustar_stock_sede(conn, doc_row["entidad_origen"], it["producto"], +kg_i)
                                ajustar_stock_sede(conn, doc_row["entidad_destino"], it["producto"], -kg_i)
                    conn.execute("DELETE FROM documentos WHERE id=?", (doc_row["id"],))
                    conn.commit()
                st.success(f"Documento '{doc_row['correlativo']}' eliminado y stock restaurado.")
                st.rerun()

            st.dataframe([dict(d) for d in docs_h], use_container_width=True)


if __name__ == "__main__":
    ejecutar_interfaz_principal()