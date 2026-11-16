"""
================================================================================
SISTEMA INTEGRAL UNIFICADO: LISTAS DE PRECIOS, DESPACHOS, CENTROS E INVENTARIO
Empresa: AGROPECUARIA OPTIMA C.A
Versión: 5.5 Master (Checklist Manual de Disponibilidad en Listas de Precios)
Arquitectura: Python 3.12 + CustomTkinter + SQLite3 + Pillow + fpdf2 + openpyxl
Módulos integrados en una sola aplicación:
 1. Listas de Precios con Checklist Manual (☑ Disponible / ☐ No Disponible).
 2. Órdenes de Despacho (Clientes, Kilos x Precio Unitario y descuento de stock).
 3. Notas de Crédito (Exclusivo para Centros de Distribución).
 4. Notas de Débito y Movimientos entre Sedes (Con Preselector de Chofer y Rotación).
 5. Inventario por Sede (Existencias en vivo en KG, ajustes manuales y PDF/Excel).
 6. Bases de Datos Maestras (Clientes, Centros de Distribución y Choferes).
 7. Historial General de Documentos (Reimpresión 1 clic y reversión de stock).
================================================================================
"""

import os
import sys
import time
import sqlite3
import threading
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Tuple, Optional, Callable

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import customtkinter as ctk
from PIL import Image, ImageDraw, ImageChops

import urllib3
import requests
from bs4 import BeautifulSoup
from fpdf import FPDF
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.drawing.image import Image as ImagenExcel

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


# ==============================================================================
# PARÁMETROS AJUSTABLES DEL SISTEMA Y PALETA VERDE VIVO
# ==============================================================================
TITULO_VENTANA = "Sistema Integral: Listas de Precios, Despachos e Inventario | Agropecuaria Óptima C.A"
NOMBRE_EMPRESA = "AGROPECUARIA OPTIMA C.A"
SEDE_PLANTA_PRINCIPAL = "PLANTA / FINCA PRINCIPAL"
OPCION_CHOFER_MANUAL = "-- Escribir Manualmente --"
NOMBRE_IMAGEN_LOGO_DEFECTO = "WhatsApp Image 2026-11-10 at 7.02.36 PM.jpeg"

INTERVALO_TASAS_SEG = 900       # Actualización automática BCV cada 15 minutos
UMBRAL_STOCK_BAJO_KG = 20.0     # Kilos por debajo de los cuales se marca "Bajo Stock"
TOLERANCIA_RECORTE_JPEG = 28    # Sensibilidad para limpiar bordes del logo
MARGEN_LOGO_PX = 20             # Margen alrededor del logo en píxeles
ALTO_MAX_LOGO_PDF_MM = 24.0     # Alto máximo del logo en el encabezado PDF
ANCHO_MAX_LOGO_PDF_MM = 88.0    # Ancho máximo del logo en el encabezado PDF

# Prefijos de numeración correlativa para documentos operativos
PREFIJOS_DOCUMENTO = {
    "DESPACHO": "OD",
    "NOTA_CREDITO": "NC",
    "NOTA_DEBITO": "ND",
    "MOVIMIENTO": "MOV",
}

# Modalidades de exportación para el Módulo 1 (Listas de Precios)
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

# Paleta corporativa: (Color en Modo Claro, Color en Modo Oscuro Verde Vivo)
COLOR_FONDO_VENTANA = ("#F4F1EA", "#0D2618")   # Verde bosque profundo en modo oscuro
COLOR_TARJETA = ("#E9E3D5", "#143824")         # Verde esmeralda oscuro para tarjetas principales
COLOR_SUBTARJETA = ("#DFD8C8", "#1E4E33")      # Verde vivo medio para sub-paneles internos
COLOR_BORDE = ("#4F6835", "#34D399")           # Verde esmeralda brillante para resaltar bordes
COLOR_TEXTO_TITULO = ("#2D421E", "#6EE7B7")    # Verde luminoso para títulos en modo oscuro
COLOR_TEXTO_SUAVE = ("#4B5563", "#A7F3D0")     # Verde claro suave para subtítulos

COLOR_PRIMARIO = ("#3F5A2A", "#16A34A")        # Botones principales en verde vivo
COLOR_PRIMARIO_HOVER = ("#2D421E", "#15803D")
COLOR_ACENTO = "#D97706"                       # Dorado / Ámbar corporativo
COLOR_ACENTO_HOVER = "#B45309"
COLOR_NEUTRO = ("#475569", "#2A6F48")
COLOR_NEUTRO_HOVER = ("#334155", "#1E5235")
COLOR_PELIGRO = "#DC2626"                      # Rojo para botones de Eliminar
COLOR_PELIGRO_HOVER = "#B91C1C"

DIAS_SEMANA_ES = {
    0: "Lunes",
    1: "Martes",
    2: "Miércoles",
    3: "Jueves",
    4: "Viernes",
    5: "Sábado",
    6: "Domingo",
}

ESCALAS_PRECIO = {
    "Cliente Final": "precio_cliente",
    "Vendedor 10%": "precio_vendedor",
    "Al Mayor": "precio_mayor",
}


# ==============================================================================
# PASO 1: CONFIGURACIÓN DE RUTAS Y PROCESAMIENTO DEL LOGOTIPO
# ==============================================================================
def obtener_ruta_base() -> Path:
    """Detecta la carpeta real tanto en VS Code como compilado en .exe."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


RUTA_BASE: Path = obtener_ruta_base()
RUTA_DB: Path = RUTA_BASE / "sistema_finca_optima.db"
RUTA_LOGO: Path = RUTA_BASE / "logo_finca.png"
RUTA_ICONO: Path = RUTA_BASE / "icono_finca.ico"
DIR_PDF: Path = RUTA_BASE / "reportes_pdf"
DIR_EXCEL: Path = RUTA_BASE / "reportes_excel"

for carpeta in (DIR_PDF, DIR_EXCEL):
    carpeta.mkdir(parents=True, exist_ok=True)


def abrir_archivo(ruta: Path) -> None:
    """Abre automáticamente el documento PDF o Excel recién generado."""
    if ruta.exists() and sys.platform.startswith("win"):
        try:
            os.startfile(str(ruta))
        except Exception:
            pass


def obtener_marca_tiempo() -> Tuple[str, str, str]:
    """Devuelve (Día en español, Fecha DD/MM/AAAA, Hora 12h)."""
    ahora = datetime.now()
    return (
        DIAS_SEMANA_ES[ahora.weekday()],
        ahora.strftime("%d/%m/%Y"),
        ahora.strftime("%I:%M:%S %p"),
    )


def recortar_y_guardar_logo(ruta_origen: Path) -> bool:
    """Recorta los bordes sobrantes del logo y genera logo_finca.png e icono_finca.ico."""
    try:
        with Image.open(ruta_origen) as abierta:
            img_rgb = abierta.convert("RGB")
            w, h = img_rgb.size
            color_fondo = img_rgb.getpixel((min(5, w - 1), min(5, h - 1)))
            lienzo = Image.new("RGB", (w, h), color_fondo)
            dif = ImageChops.difference(img_rgb, lienzo).convert("L")
            caja = dif.point(lambda p: 255 if p > TOLERANCIA_RECORTE_JPEG else 0).getbbox()

            if caja:
                x1 = max(0, caja[0] - MARGEN_LOGO_PX)
                y1 = max(0, caja[1] - MARGEN_LOGO_PX)
                x2 = min(w, caja[2] + MARGEN_LOGO_PX)
                y2 = min(h, caja[3] + MARGEN_LOGO_PX)
                recortada = abierta.convert("RGBA").crop((x1, y1, x2, y2))
            else:
                recortada = abierta.convert("RGBA").copy()

        recortada.save(RUTA_LOGO, format="PNG")
        lado = max(recortada.size) + 24
        cuadrado = Image.new("RGBA", (lado, lado), (*color_fondo, 255))
        cuadrado.paste(
            recortada,
            ((lado - recortada.width) // 2, (lado - recortada.height) // 2),
            recortada,
        )
        cuadrado.save(RUTA_ICONO, format="ICO", sizes=[(256, 256), (128, 128), (64, 64), (32, 32)])
        return True
    except Exception:
        return False


def inicializar_logo() -> None:
    """Busca el logo oficial en carpetas vecinas o genera un escudo campestre."""
    if RUTA_LOGO.exists() and RUTA_ICONO.exists():
        return

    rutas_posibles = [
        RUTA_BASE / "logo_finca.png",
        RUTA_BASE.parent / "Lista de precios" / "logo_finca.png",
        RUTA_BASE.parent / "Despachos Finca" / "logo_finca.png",
        RUTA_BASE / NOMBRE_IMAGEN_LOGO_DEFECTO,
        Path.home() / "Downloads" / NOMBRE_IMAGEN_LOGO_DEFECTO,
        Path.home() / "Descargas" / NOMBRE_IMAGEN_LOGO_DEFECTO,
        Path.home() / "OneDrive" / "Escritorio" / NOMBRE_IMAGEN_LOGO_DEFECTO,
        Path.home() / "Desktop" / NOMBRE_IMAGEN_LOGO_DEFECTO,
    ]
    for candidata in rutas_posibles:
        if candidata.exists() and recortar_y_guardar_logo(candidata):
            return

    img = Image.new("RGBA", (440, 170), (63, 90, 42, 255))
    dib = ImageDraw.Draw(img)
    dib.rounded_rectangle((8, 8, 432, 162), radius=20, outline=(250, 204, 21, 255), width=4)
    img.save(RUTA_LOGO, format="PNG")
    img.save(RUTA_ICONO, format="ICO", sizes=[(256, 256), (64, 64), (32, 32)])


# ==============================================================================
# PASO 2: BASE DE DATOS UNIFICADA CON COLUMNA 'DISPONIBLE' (CHECKLIST)
# ==============================================================================
def conectar_db() -> sqlite3.Connection:
    conn = sqlite3.connect(RUTA_DB, timeout=15)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL;")
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


def reparar_esquema_db(conn: sqlite3.Connection) -> None:
    """
    Asegura que la tabla 'productos' tenga la columna 'disponible' para el
    checklist manual (1 = Disponible, 0 = No disponible) sin borrar datos previos.
    """
    cols_prod = {f["name"] for f in conn.execute("PRAGMA table_info(productos)").fetchall()}
    if "disponible" not in cols_prod:
        conn.execute("ALTER TABLE productos ADD COLUMN disponible INTEGER NOT NULL DEFAULT 1;")


def migrar_bases_anteriores(conn: sqlite3.Connection) -> None:
    """Rescata automáticamente datos de 'despachos_finca.db' o 'precios_finca.db' si existen."""
    posibles_dbs = [
        RUTA_BASE / "despachos_finca.db",
        RUTA_BASE / "precios_finca.db",
        RUTA_BASE.parent / "Lista de precios" / "precios_finca.db",
        RUTA_BASE.parent / "Despachos Finca" / "despachos_finca.db",
    ]
    for ruta_vieja in posibles_dbs:
        if not ruta_vieja.exists() or ruta_vieja.resolve() == RUTA_DB.resolve():
            continue
        try:
            with sqlite3.connect(ruta_vieja) as ext:
                ext.row_factory = sqlite3.Row
                nombres_tablas = {
                    r["name"] for r in ext.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
                }

                if "productos" in nombres_tablas:
                    cols_p_ext = {f["name"] for f in ext.execute("PRAGMA table_info(productos)")}
                    for p in ext.execute("SELECT * FROM productos"):
                        disp = int(p["disponible"]) if "disponible" in cols_p_ext else 1
                        conn.execute(
                            """INSERT INTO productos (nombre, rubro, precio_mayor, precio_vendedor, precio_cliente, disponible)
                               VALUES (?, ?, ?, ?, ?, ?)
                               ON CONFLICT(nombre) DO UPDATE SET
                                   rubro=excluded.rubro,
                                   precio_mayor=excluded.precio_mayor,
                                   precio_vendedor=excluded.precio_vendedor,
                                   precio_cliente=excluded.precio_cliente,
                                   disponible=excluded.disponible""",
                            (p["nombre"], p["rubro"], p["precio_mayor"], p["precio_vendedor"], p["precio_cliente"], disp),
                        )

                if "clientes" in nombres_tablas and conn.execute("SELECT COUNT(*) AS c FROM clientes").fetchone()["c"] == 0:
                    for c in ext.execute("SELECT * FROM clientes"):
                        conn.execute(
                            "INSERT INTO clientes (nombre, documento_id, escala_preferida, telefono, direccion) VALUES (?, ?, ?, ?, ?)",
                            (c["nombre"], c["documento_id"], c["escala_preferida"], c["telefono"], c["direccion"]),
                        )

                if "centros_distribucion" in nombres_tablas and conn.execute("SELECT COUNT(*) AS c FROM centros_distribucion").fetchone()["c"] == 0:
                    cols_cd_ext = {f["name"] for f in ext.execute("PRAGMA table_info(centros_distribucion)")}
                    for d in ext.execute("SELECT * FROM centros_distribucion"):
                        saldo = float(d["saldo_pendiente_ref"]) if "saldo_pendiente_ref" in cols_cd_ext else 0.0
                        conn.execute(
                            """INSERT OR IGNORE INTO centros_distribucion
                               (codigo, nombre, encargado, telefono, ubicacion, saldo_pendiente_ref)
                               VALUES (?, ?, ?, ?, ?, ?)""",
                            (d["codigo"], d["nombre"], d["encargado"], d["telefono"], d["ubicacion"], saldo),
                        )

                if "choferes" in nombres_tablas and conn.execute("SELECT COUNT(*) AS c FROM choferes").fetchone()["c"] == 0:
                    for ch in ext.execute("SELECT * FROM choferes"):
                        conn.execute(
                            "INSERT INTO choferes (nombre, cedula, placa, telefono, vehiculo) VALUES (?, ?, ?, ?, ?)",
                            (ch["nombre"], ch["cedula"], ch["placa"], ch["telefono"], ch["vehiculo"]),
                        )

                if "inventario_sedes" in nombres_tablas:
                    for inv in ext.execute("SELECT * FROM inventario_sedes"):
                        conn.execute(
                            """INSERT INTO inventario_sedes (sede, producto, kilos_disponibles, ultima_actualizacion)
                               VALUES (?, ?, ?, ?)
                               ON CONFLICT(sede, producto) DO UPDATE SET
                                   kilos_disponibles=excluded.kilos_disponibles,
                                   ultima_actualizacion=excluded.ultima_actualizacion""",
                            (inv["sede"], inv["producto"], inv["kilos_disponibles"], inv["ultima_actualizacion"]),
                        )
        except Exception:
            continue


def inicializar_base_de_datos() -> None:
    """Crea todas las tablas del sistema unificado e invoca la reparación de columnas."""
    es_nueva = not RUTA_DB.exists()
    with conectar_db() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS tasas_cambio (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha TEXT NOT NULL,
                hora TEXT NOT NULL,
                bcv_usd REAL NOT NULL,
                bcv_eur REAL NOT NULL,
                fuente TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS productos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL UNIQUE,
                rubro TEXT NOT NULL,
                precio_mayor REAL NOT NULL,
                precio_vendedor REAL NOT NULL,
                precio_cliente REAL NOT NULL,
                disponible INTEGER NOT NULL DEFAULT 1
            );

            CREATE TABLE IF NOT EXISTS clientes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL,
                documento_id TEXT NOT NULL,
                escala_preferida TEXT NOT NULL DEFAULT 'Cliente Final',
                telefono TEXT DEFAULT '',
                direccion TEXT DEFAULT ''
            );

            CREATE TABLE IF NOT EXISTS centros_distribucion (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo TEXT NOT NULL UNIQUE,
                nombre TEXT NOT NULL,
                encargado TEXT DEFAULT '',
                telefono TEXT DEFAULT '',
                ubicacion TEXT DEFAULT '',
                saldo_pendiente_ref REAL DEFAULT 0.0
            );

            CREATE TABLE IF NOT EXISTS choferes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL,
                cedula TEXT NOT NULL,
                placa TEXT NOT NULL,
                telefono TEXT DEFAULT '',
                vehiculo TEXT DEFAULT ''
            );

            CREATE TABLE IF NOT EXISTS inventario_sedes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sede TEXT NOT NULL,
                producto TEXT NOT NULL,
                kilos_disponibles REAL NOT NULL DEFAULT 0.0,
                ultima_actualizacion TEXT NOT NULL,
                UNIQUE(sede, producto)
            );

            CREATE TABLE IF NOT EXISTS documentos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                correlativo TEXT NOT NULL UNIQUE,
                tipo_doc TEXT NOT NULL,
                subtipo TEXT NOT NULL,
                entidad_origen TEXT NOT NULL,
                entidad_destino TEXT NOT NULL,
                doc_identidad TEXT DEFAULT '',
                referencia TEXT DEFAULT '',
                motivo TEXT DEFAULT '',
                chofer TEXT DEFAULT '',
                cedula_chofer TEXT DEFAULT '',
                placa_vehiculo TEXT DEFAULT '',
                afecto_saldo INTEGER DEFAULT 0,
                afecto_inventario INTEGER DEFAULT 1,
                observaciones TEXT DEFAULT '',
                modo_moneda TEXT NOT NULL,
                tasa_usd REAL NOT NULL,
                tasa_eur REAL NOT NULL,
                total_kilos REAL NOT NULL,
                total_ref REAL NOT NULL,
                total_bs REAL NOT NULL,
                dia_semana TEXT NOT NULL,
                fecha TEXT NOT NULL,
                hora TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS documento_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                documento_id INTEGER NOT NULL,
                producto TEXT NOT NULL,
                escala TEXT NOT NULL,
                kilos REAL NOT NULL,
                precio_unitario REAL NOT NULL,
                subtotal_ref REAL NOT NULL,
                FOREIGN KEY (documento_id) REFERENCES documentos(id) ON DELETE CASCADE
            );
            """
        )

        reparar_esquema_db(conn)

        if es_nueva:
            migrar_bases_anteriores(conn)

        if conn.execute("SELECT COUNT(*) AS c FROM productos").fetchone()["c"] == 0:
            conn.executemany(
                """INSERT INTO productos (nombre, rubro, precio_mayor, precio_vendedor, precio_cliente, disponible)
                   VALUES (?, ?, ?, ?, ?, ?)""",
                CATALOGO_BASE,
            )

        if conn.execute("SELECT COUNT(*) AS c FROM clientes").fetchone()["c"] == 0:
            conn.execute(
                """INSERT INTO clientes (nombre, documento_id, escala_preferida, telefono, direccion)
                   VALUES (?, ?, ?, ?, ?)""",
                ("CLIENTE MOSTRADOR GENERAL", "V-00000000", "Cliente Final", "", "Venta directa"),
            )

        if conn.execute("SELECT COUNT(*) AS c FROM centros_distribucion").fetchone()["c"] == 0:
            conn.executemany(
                """INSERT INTO centros_distribucion (codigo, nombre, encargado, telefono, ubicacion, saldo_pendiente_ref)
                   VALUES (?, ?, ?, ?, ?, ?)""",
                [
                    ("CD-01", "CENTRO DE DISTRIBUCION PRINCIPAL", "Encargado Sede 1", "", "Sede Central", 0.0),
                    ("CD-02", "CENTRO DE DISTRIBUCION SUCURSAL 2", "Encargado Sede 2", "", "Zona Comercial", 0.0),
                ],
            )

        if conn.execute("SELECT COUNT(*) AS c FROM choferes").fetchone()["c"] == 0:
            conn.execute(
                """INSERT INTO choferes (nombre, cedula, placa, telefono, vehiculo)
                   VALUES (?, ?, ?, ?, ?)""",
                ("CHOFER TRANSPORTE PLANTA", "V-12345678", "A00AA00", "0414-0000000", "Camión Cava Fría"),
            )

        if conn.execute("SELECT COUNT(*) AS c FROM tasas_cambio").fetchone()["c"] == 0:
            ahora = datetime.now()
            conn.execute(
                "INSERT INTO tasas_cambio (fecha, hora, bcv_usd, bcv_eur, fuente) VALUES (?, ?, ?, ?, ?)",
                (ahora.strftime("%Y-%m-%d"), ahora.strftime("%H:%M:%S"), 901.62, 1007.96, "Inicial"),
            )
        conn.commit()


def normalizar_nombre_sede(nombre_sede: str) -> str:
    limpio = str(nombre_sede).strip()
    if limpio.startswith(SEDE_PLANTA_PRINCIPAL):
        return SEDE_PLANTA_PRINCIPAL
    return limpio


def ajustar_stock_sede(conn: sqlite3.Connection, sede: str, producto: str, delta_kilos: float) -> None:
    sede_norm = normalizar_nombre_sede(sede)
    if not sede_norm or not producto:
        return
    _, fecha_str, hora_str = obtener_marca_tiempo()
    marca = f"{fecha_str} {hora_str}"
    conn.execute(
        """INSERT INTO inventario_sedes (sede, producto, kilos_disponibles, ultima_actualizacion)
           VALUES (?, ?, ?, ?)
           ON CONFLICT(sede, producto) DO UPDATE SET
               kilos_disponibles = ROUND(inventario_sedes.kilos_disponibles + excluded.kilos_disponibles, 2),
               ultima_actualizacion = excluded.ultima_actualizacion""",
        (sede_norm, producto.strip().upper(), round(delta_kilos, 2), marca),
    )


def fijar_stock_exacto_sede(sede: str, producto: str, kilos_exactos: float) -> None:
    sede_norm = normalizar_nombre_sede(sede)
    _, fecha_str, hora_str = obtener_marca_tiempo()
    marca = f"{fecha_str} {hora_str}"
    ejecutar_sql(
        """INSERT INTO inventario_sedes (sede, producto, kilos_disponibles, ultima_actualizacion)
           VALUES (?, ?, ?, ?)
           ON CONFLICT(sede, producto) DO UPDATE SET
               kilos_disponibles = excluded.kilos_disponibles,
               ultima_actualizacion = excluded.ultima_actualizacion""",
        (sede_norm, producto.strip().upper(), round(kilos_exactos, 2), marca),
    )


def eliminar_item_inventario_sede(sede: str, producto: str) -> None:
    sede_norm = normalizar_nombre_sede(sede)
    ejecutar_sql(
        "DELETE FROM inventario_sedes WHERE sede = ? AND producto = ?",
        (sede_norm, producto.strip().upper()),
    )


def generar_siguiente_correlativo(tipo_doc: str) -> str:
    prefijo = PREFIJOS_DOCUMENTO.get(tipo_doc, "DOC")
    fila = consultar_sql("SELECT COUNT(*) AS total FROM documentos WHERE tipo_doc = ?", (tipo_doc,))
    consecutivo = (fila[0]["total"] if fila else 0) + 1
    while True:
        candidato = f"{prefijo}-{consecutivo:04d}"
        existe = consultar_sql("SELECT 1 FROM documentos WHERE correlativo = ?", (candidato,))
        if not existe:
            return candidato
        consecutivo += 1


# ==============================================================================
# PASO 3: MOTOR DE TASAS BCV ($ USD Y € EUR) EN SEGUNDO PLANO
# ==============================================================================
class MotorTasas:
    def __init__(self, callback_ui: Optional[Callable[[Dict[str, float]], None]] = None):
        self.callback_ui = callback_ui
        self.activo = True

    def iniciar(self) -> None:
        threading.Thread(target=self._ciclo_web, daemon=True).start()

    def obtener_tasas(self) -> Dict[str, float]:
        filas = consultar_sql("SELECT bcv_usd, bcv_eur FROM tasas_cambio ORDER BY id DESC LIMIT 1")
        if filas:
            return {"bcv_usd": float(filas[0]["bcv_usd"]), "bcv_eur": float(filas[0]["bcv_eur"])}
        return {"bcv_usd": 901.62, "bcv_eur": 1007.96}

    def guardar_tasas(self, usd: float, eur: float, fuente: str = "Manual") -> None:
        ahora = datetime.now()
        ejecutar_sql(
            "INSERT INTO tasas_cambio (fecha, hora, bcv_usd, bcv_eur, fuente) VALUES (?, ?, ?, ?, ?)",
            (ahora.strftime("%Y-%m-%d"), ahora.strftime("%H:%M:%S"), round(usd, 2), round(eur, 2), fuente),
        )
        if self.callback_ui:
            self.callback_ui(self.obtener_tasas())

    def sincronizar_web(self) -> bool:
        actuales = self.obtener_tasas()
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
            r_bcv = requests.get(
                "https://www.bcv.org.ve/", headers={"User-Agent": "Mozilla/5.0"}, verify=False, timeout=10
            )
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
            self.guardar_tasas(usd, eur, fuente="Auto-BCV")
        return exito

    def _ciclo_web(self) -> None:
        while self.activo:
            self.sincronizar_web()
            for _ in range(INTERVALO_TASAS_SEG):
                if not self.activo:
                    return
                time.sleep(1)


# ==============================================================================
# PASO 4: GENERADORES PDF Y EXCEL (FILTRANDO SOLO PRODUCTOS DISPONIBLES ☑)
# ==============================================================================
def limpiar_pdf(texto: str) -> str:
    return str(texto).replace("€", "EUR").encode("latin-1", "replace").decode("latin-1")


def usa_tasa_euro(modo_moneda: str) -> bool:
    return ("Euro" in modo_moneda) or ("€" in modo_moneda)


def obtener_tasa_activa(modo_moneda: str, tasa_usd: float, tasa_eur: float) -> Tuple[float, str, str]:
    if usa_tasa_euro(modo_moneda):
        return tasa_eur, "EUR", "Euro BCV"
    return tasa_usd, "$", "Dólar BCV"


def formatear_monto_celda(valor_ref: float, modo_moneda: str, tasa_usd: float, tasa_eur: float) -> str:
    tasa, simb, _ = obtener_tasa_activa(modo_moneda, tasa_usd, tasa_eur)
    if modo_moneda.startswith("REF"):
        return f"{simb} {valor_ref:,.2f}"
    if modo_moneda.startswith("Bolívares"):
        return f"Bs {valor_ref * tasa:,.2f}"
    return f"{simb} {valor_ref:,.2f} / Bs {valor_ref * tasa:,.2f}"


def obtener_columnas_lista_precios(modalidad_lista: str) -> List[Tuple[str, str]]:
    """Determina qué columnas de precio se imprimen en la Lista de Precios oficial."""
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


def consultar_productos_para_lista(rubro_filtro: str, solo_disponibles: bool = True) -> List[sqlite3.Row]:
    """
    Obtiene los productos del catálogo filtrando por rubro y respetando el
    Checklist manual (disponible = 1) para que los no marcados no salgan en la lista.
    """
    if rubro_filtro == "Todos los Rubros":
        if solo_disponibles:
            return consultar_sql("SELECT * FROM productos WHERE disponible = 1 ORDER BY rubro ASC, nombre ASC")
        return consultar_sql("SELECT * FROM productos ORDER BY rubro ASC, nombre ASC")
    else:
        if solo_disponibles:
            return consultar_sql(
                "SELECT * FROM productos WHERE rubro = ? AND disponible = 1 ORDER BY nombre ASC",
                (rubro_filtro,),
            )
        return consultar_sql("SELECT * FROM productos WHERE rubro = ? ORDER BY nombre ASC", (rubro_filtro,))


def exportar_lista_precios_pdf(
    modalidad_lista: str,
    rubro_filtro: str,
    modo_moneda: str,
    tasa_usd: float,
    tasa_eur: float,
    solo_disponibles: bool = True,
) -> Path:
    """Genera el PDF corporativo de Lista de Precios Finca únicamente con los productos marcados (☑)."""
    productos = consultar_productos_para_lista(rubro_filtro, solo_disponibles)
    if not productos:
        raise ValueError(
            "No hay ningún producto marcado como '☑ Disponible' en la categoría seleccionada.\n\n"
            "Marca al menos un producto en el checklist de la tabla antes de generar la lista."
        )

    dia_s, fecha_s, hora_s = obtener_marca_tiempo()
    tasa_apl, _, nom_tasa = obtener_tasa_activa(modo_moneda, tasa_usd, tasa_eur)
    cols_precio = obtener_columnas_lista_precios(modalidad_lista)

    tag_archivo = modalidad_lista.split("(")[0].strip().replace(" ", "_").upper()
    ruta_pdf = DIR_PDF / f"LISTA_PRECIOS_{tag_archivo}_{fecha_s.replace('/', '')}.pdf"

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.add_page()

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
    pdf.set_font("Helvetica", "B", 13)
    pdf.cell(0, 9, limpiar_pdf(f"{NOMBRE_EMPRESA}  -  LISTA OFICIAL DE PRECIOS DISPONIBLES"), ln=True, align="C", fill=True)

    pdf.set_fill_color(217, 119, 6)
    pdf.set_font("Helvetica", "B", 9.2)
    cintillo = (
        f"MODALIDAD: {modalidad_lista.upper()}   |   "
        f"TASA {nom_tasa.upper()}: Bs {tasa_apl:,.2f}   |   "
        f"{dia_s.upper()} {fecha_s}"
    )
    pdf.cell(0, 6.5, limpiar_pdf(cintillo), ln=True, align="C", fill=True)
    pdf.ln(3)

    n_cols = len(cols_precio)
    ancho_precio = 36 if n_cols == 3 else (46 if n_cols == 2 else 60)
    ancho_rubro = 38
    ancho_prod = 190 - ancho_rubro - (ancho_precio * n_cols)

    pdf.set_fill_color(234, 228, 215)
    pdf.set_text_color(45, 66, 30)
    pdf.set_font("Helvetica", "B", 8.5)
    pdf.cell(ancho_prod, 8, "  PRODUCTO / CORTE DISPONIBLE", border=1, fill=True)
    pdf.cell(ancho_rubro, 8, "CATEGORIA", border=1, align="C", fill=True)
    for i, (titulo_col, _) in enumerate(cols_precio):
        salto = (i == n_cols - 1)
        pdf.cell(ancho_precio, 8, limpiar_pdf(titulo_col), border=1, align="R", fill=True, ln=salto)

    alternar = False
    for p in productos:
        pdf.set_fill_color(246, 244, 238) if alternar else pdf.set_fill_color(255, 255, 255)
        alternar = not alternar

        pdf.set_text_color(20, 20, 20)
        pdf.set_font("Helvetica", "B", 8.2)
        pdf.cell(ancho_prod, 7, limpiar_pdf(f"  {p['nombre']}")[:42], border=1, fill=True)
        pdf.set_font("Helvetica", "", 8)
        pdf.cell(ancho_rubro, 7, limpiar_pdf(p["rubro"])[:22], border=1, align="C", fill=True)

        pdf.set_font("Helvetica", "B", 8.2)
        for i, (_, campo_db) in enumerate(cols_precio):
            salto = (i == n_cols - 1)
            txt_monto = formatear_monto_celda(float(p[campo_db]), modo_moneda, tasa_usd, tasa_eur)
            pdf.cell(ancho_precio, 7, limpiar_pdf(f"{txt_monto} "), border=1, align="R", fill=True, ln=salto)

    pdf.ln(3)
    pdf.set_font("Helvetica", "I", 8)
    pdf.set_text_color(80, 80, 80)
    nota_pie = (
        f"Total productos disponibles en lista: {len(productos)}  |  "
        f"Precios por Kilogramo (KG)  |  Emitido el {dia_s} {fecha_s} a las {hora_s}."
    )
    pdf.cell(0, 5, limpiar_pdf(nota_pie), align="C")

    pdf.output(str(ruta_pdf))
    return ruta_pdf


def exportar_lista_precios_excel(
    modalidad_lista: str,
    rubro_filtro: str,
    modo_moneda: str,
    tasa_usd: float,
    tasa_eur: float,
    solo_disponibles: bool = True,
) -> Path:
    """Exporta la Lista de Precios oficial a formato Excel (.xlsx) solo con productos marcados (☑)."""
    productos = consultar_productos_para_lista(rubro_filtro, solo_disponibles)
    if not productos:
        raise ValueError(
            "No hay ningún producto marcado como '☑ Disponible' para exportar a Excel.\n\n"
            "Marca al menos un producto en el checklist de la tabla."
        )

    dia_s, fecha_s, hora_s = obtener_marca_tiempo()
    tasa_apl, simb, nom_tasa = obtener_tasa_activa(modo_moneda, tasa_usd, tasa_eur)
    cols_precio = obtener_columnas_lista_precios(modalidad_lista)

    tag_archivo = modalidad_lista.split("(")[0].strip().replace(" ", "_").upper()
    ruta_xlsx = DIR_EXCEL / f"LISTA_PRECIOS_{tag_archivo}_{fecha_s.replace('/', '')}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "Lista_de_Precios"

    ws.row_dimensions[1].height = 55
    if RUTA_LOGO.exists():
        try:
            with Image.open(RUTA_LOGO) as im:
                w_px, h_px = im.size
            logo_xl = ImagenExcel(str(RUTA_LOGO))
            logo_xl.height = 68
            logo_xl.width = int(68 * (w_px / max(h_px, 1)))
            ws.add_image(logo_xl, "A1")
        except Exception:
            pass

    ws["B1"] = f"{NOMBRE_EMPRESA} - LISTA DE PRECIOS DISPONIBLES ({modalidad_lista})"
    ws["B1"].font = Font(size=13, bold=True, color="3F5A2A")
    ws.append([f"Fecha: {dia_s} {fecha_s} ({hora_s}) | Moneda: {modo_moneda} | Tasa {nom_tasa}: Bs {tasa_apl:,.2f}"])
    ws.append([])

    encabezados = ["Producto / Corte Disponible", "Categoría / Rubro"]
    for titulo_col, _ in cols_precio:
        encabezados.append(f"{titulo_col} ({simb})")
        encabezados.append(f"{titulo_col} (Bs)")

    fila_enc = ws.max_row + 1
    ws.append(encabezados)

    relleno = PatternFill(start_color="3F5A2A", end_color="3F5A2A", fill_type="solid")
    fuente_b = Font(bold=True, color="FFFFFF")
    for idx in range(1, len(encabezados) + 1):
        c = ws.cell(row=fila_enc, column=idx)
        c.fill = relleno
        c.font = fuente_b
        c.alignment = Alignment(horizontal="center")

    for p in productos:
        fila_datos = [p["nombre"], p["rubro"]]
        for _, campo_db in cols_precio:
            val_ref = float(p[campo_db])
            fila_datos.append(round(val_ref, 2))
            fila_datos.append(round(val_ref * tasa_apl, 2))
        ws.append(fila_datos)

    for col in ws.columns:
        ancho = max(len(str(celda.value or "")) for celda in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(ancho + 4, 15), 42)

    wb.save(str(ruta_xlsx))
    return ruta_xlsx


class PlantillaPDFOperacion(FPDF):
    def __init__(self, doc: sqlite3.Row):
        super().__init__()
        self.doc = doc

    def header(self) -> None:
        if RUTA_LOGO.exists():
            try:
                with Image.open(RUTA_LOGO) as im:
                    w_px, h_px = im.size
                prop = w_px / max(h_px, 1)
                alto = ALTO_MAX_LOGO_PDF_MM
                ancho = alto * prop
                if ancho > ANCHO_MAX_LOGO_PDF_MM:
                    ancho = ANCHO_MAX_LOGO_PDF_MM
                    alto = ancho / max(prop, 0.01)
                self.image(str(RUTA_LOGO), x=(210.0 - ancho) / 2.0, y=6, w=ancho, h=alto)
                self.set_y(6 + alto + 3)
            except Exception:
                self.set_y(10)

        self.set_fill_color(63, 90, 42)
        self.set_text_color(255, 255, 255)
        self.set_font("Helvetica", "B", 12.5)
        titulo_barra = f"{NOMBRE_EMPRESA}  -  {self.doc['subtipo'].upper()}"
        self.cell(0, 9, limpiar_pdf(titulo_barra), ln=True, align="C", fill=True)

        self.set_fill_color(217, 119, 6)
        self.set_font("Helvetica", "B", 9.5)
        tasa_apl, _, nom_tasa = obtener_tasa_activa(
            self.doc["modo_moneda"], self.doc["tasa_usd"], self.doc["tasa_eur"]
        )
        cintillo = (
            f"DOCUMENTO NRO: {self.doc['correlativo']}   |   "
            f"MONEDA: {self.doc['modo_moneda']}   |   "
            f"TASA {nom_tasa.upper()}: Bs {tasa_apl:,.2f}"
        )
        self.cell(0, 6.5, limpiar_pdf(cintillo), ln=True, align="C", fill=True)
        self.ln(3)

    def footer(self) -> None:
        self.set_y(-17)
        self.set_draw_color(63, 90, 42)
        self.line(10, self.get_y(), 200, self.get_y())
        self.ln(1.5)
        self.set_font("Helvetica", "B", 8)
        self.set_text_color(63, 90, 42)
        pie = (
            f"Documento: {self.doc['correlativo']}  |  "
            f"Emitido el {self.doc['dia_semana']} {self.doc['fecha']} a las {self.doc['hora']}  |  "
            f"Pagina {self.page_no()}"
        )
        self.cell(0, 5, limpiar_pdf(pie), align="C")


def exportar_documento_pdf(documento_id: int) -> Path:
    cabeceras = consultar_sql("SELECT * FROM documentos WHERE id = ?", (documento_id,))
    if not cabeceras:
        raise ValueError("No se encontró el documento solicitado en la base de datos.")
    doc = cabeceras[0]
    items = consultar_sql("SELECT * FROM documento_items WHERE documento_id = ?", (documento_id,))

    ruta_pdf = DIR_PDF / f"{doc['correlativo']}_{doc['fecha'].replace('/', '')}.pdf"
    pdf = PlantillaPDFOperacion(doc)
    pdf.set_auto_page_break(auto=True, margin=22)
    pdf.add_page()

    pdf.set_fill_color(246, 244, 238)
    pdf.set_draw_color(63, 90, 42)
    pdf.set_text_color(25, 25, 25)
    pdf.set_font("Helvetica", "B", 9)

    if doc["tipo_doc"] == "DESPACHO":
        pdf.cell(125, 6.5, limpiar_pdf(f"  CLIENTE: {doc['entidad_destino']}"), border="LTR", fill=True)
        pdf.cell(65, 6.5, limpiar_pdf(f"  CEDULA / RIF: {doc['doc_identidad']}"), border="LTR", fill=True, ln=True)
        pdf.set_font("Helvetica", "", 8.5)
        pdf.cell(125, 6, limpiar_pdf(f"  DESPACHA: {doc['entidad_origen']}"), border="LBR", fill=True)
        pdf.cell(65, 6, limpiar_pdf(f"  DETALLE: {doc['motivo']}"), border="LBR", fill=True, ln=True)
    else:
        pdf.cell(95, 6.5, limpiar_pdf(f"  ORIGEN: {doc['entidad_origen']}"), border="LTR", fill=True)
        pdf.cell(95, 6.5, limpiar_pdf(f"  DESTINO: {doc['entidad_destino']}"), border="LTR", fill=True, ln=True)
        pdf.set_font("Helvetica", "", 8.5)
        pdf.cell(95, 6, limpiar_pdf(f"  MOTIVO: {doc['motivo']}"), border="LBR", fill=True)
        pdf.cell(95, 6, limpiar_pdf(f"  DOC. REFERENCIA: {doc['referencia'] or 'N/A'}"), border="LBR", fill=True, ln=True)

        if doc["chofer"] or doc["cedula_chofer"] or doc["placa_vehiculo"]:
            pdf.set_font("Helvetica", "B", 8.5)
            linea_trans = (
                f"  DATOS DE TRANSPORTE  ->  Chofer: {doc['chofer'] or 'N/A'}   |   "
                f"Cedula: {doc['cedula_chofer'] or 'N/A'}   |   "
                f"Placa Vehiculo: {doc['placa_vehiculo'] or 'N/A'}"
            )
            pdf.cell(190, 6.5, limpiar_pdf(linea_trans), border=1, fill=True, ln=True)

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

    alternar = False
    for it in items:
        pdf.set_fill_color(246, 244, 238) if alternar else pdf.set_fill_color(255, 255, 255)
        alternar = not alternar

        pdf.set_text_color(20, 20, 20)
        pdf.set_font("Helvetica", "B", 8.2)
        pdf.cell(78, 7, limpiar_pdf(f"  {it['producto']}")[:44], border=1, fill=True)

        pdf.set_font("Helvetica", "", 8)
        pdf.cell(28, 7, limpiar_pdf(it["escala"]), border=1, align="C", fill=True)
        pdf.cell(24, 7, f"{float(it['kilos']):,.2f} kg ", border=1, align="R", fill=True)

        p_unit = formatear_monto_celda(float(it["precio_unitario"]), doc["modo_moneda"], doc["tasa_usd"], doc["tasa_eur"])
        sub_txt = formatear_monto_celda(float(it["subtotal_ref"]), doc["modo_moneda"], doc["tasa_usd"], doc["tasa_eur"])
        pdf.cell(30, 7, limpiar_pdf(p_unit), border=1, align="R", fill=True)
        pdf.set_font("Helvetica", "B", 8.2)
        pdf.cell(30, 7, limpiar_pdf(sub_txt), border=1, align="R", fill=True, ln=True)

    pdf.ln(3)

    _, simb, nom_tasa = obtener_tasa_activa(doc["modo_moneda"], doc["tasa_usd"], doc["tasa_eur"])
    pdf.set_fill_color(63, 90, 42)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 9.5)
    pdf.cell(60, 8, f"  TOTAL PESO: {doc['total_kilos']:,.2f} KG", border=1, fill=True)
    pdf.cell(62, 8, limpiar_pdf(f"  TOTAL REF: {simb} {doc['total_ref']:,.2f}"), border=1, align="C", fill=True)
    pdf.cell(68, 8, limpiar_pdf(f"TOTAL {nom_tasa.upper()}: Bs {doc['total_bs']:,.2f}  "), border=1, align="R", fill=True, ln=True)

    pdf.ln(14)
    if pdf.get_y() < 255:
        pdf.set_text_color(60, 60, 60)
        pdf.set_draw_color(120, 120, 120)
        y_firma = pdf.get_y()
        pdf.line(22, y_firma, 88, y_firma)
        pdf.line(122, y_firma, 188, y_firma)
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.cell(95, 6, "ENTREGADO / EMITIDO POR", align="C")
        pdf.cell(95, 6, "RECIBIDO CONFORME (FIRMA Y SELLO)", align="C", ln=True)

    pdf.output(str(ruta_pdf))
    return ruta_pdf


def exportar_documento_excel(documento_id: int) -> Path:
    doc = consultar_sql("SELECT * FROM documentos WHERE id = ?", (documento_id,))[0]
    items = consultar_sql("SELECT * FROM documento_items WHERE documento_id = ?", (documento_id,))
    ruta_xlsx = DIR_EXCEL / f"{doc['correlativo']}_{doc['fecha'].replace('/', '')}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = doc["correlativo"]

    ws.row_dimensions[1].height = 55
    if RUTA_LOGO.exists():
        try:
            with Image.open(RUTA_LOGO) as im:
                w_px, h_px = im.size
            logo_xl = ImagenExcel(str(RUTA_LOGO))
            logo_xl.height = 68
            logo_xl.width = int(68 * (w_px / max(h_px, 1)))
            ws.add_image(logo_xl, "A1")
        except Exception:
            pass

    ws["B1"] = f"{NOMBRE_EMPRESA} - {doc['subtipo']} ({doc['correlativo']})"
    ws["B1"].font = Font(size=13, bold=True, color="3F5A2A")
    ws.append([f"Origen: {doc['entidad_origen']} | Destino: {doc['entidad_destino']} | Fecha: {doc['fecha']} {doc['hora']}"])
    ws.append([
        f"Detalle/Motivo: {doc['motivo']} | Chofer: {doc['chofer'] or 'N/A'} | "
        f"Cédula: {doc['cedula_chofer'] or 'N/A'} | Placa: {doc['placa_vehiculo'] or 'N/A'} | "
        f"Obs: {doc['observaciones']}"
    ])
    ws.append([])

    tasa_apl, simb, nom_tasa = obtener_tasa_activa(doc["modo_moneda"], doc["tasa_usd"], doc["tasa_eur"])
    fila_enc = ws.max_row + 1
    encabezados = [
        "Producto",
        "Escala Aplicada",
        "Peso (KG)",
        f"Precio Unitario ({simb})",
        f"Subtotal REF ({simb})",
        f"Subtotal Bs ({nom_tasa})",
    ]
    ws.append(encabezados)

    relleno = PatternFill(start_color="3F5A2A", end_color="3F5A2A", fill_type="solid")
    fuente_b = Font(bold=True, color="FFFFFF")
    for idx in range(1, len(encabezados) + 1):
        c = ws.cell(row=fila_enc, column=idx)
        c.fill = relleno
        c.font = fuente_b
        c.alignment = Alignment(horizontal="center")

    for it in items:
        sub_ref = float(it["subtotal_ref"])
        ws.append([
            it["producto"],
            it["escala"],
            round(float(it["kilos"]), 2),
            round(float(it["precio_unitario"]), 2),
            round(sub_ref, 2),
            round(sub_ref * tasa_apl, 2),
        ])

    ws.append([])
    ws.append(["TOTALES", "", round(doc["total_kilos"], 2), "", round(doc["total_ref"], 2), round(doc["total_bs"], 2)])

    for col in ws.columns:
        ancho = max(len(str(celda.value or "")) for celda in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(ancho + 4, 16), 44)

    wb.save(str(ruta_xlsx))
    return ruta_xlsx


def exportar_inventario_pdf(sede_filtro: str, filas_inv: List[Dict[str, object]], modo_moneda: str, t_usd: float, t_eur: float) -> Path:
    dia_s, fecha_s, hora_s = obtener_marca_tiempo()
    nombre_limpio = sede_filtro.replace(" ", "_").replace("/", "").replace("-", "")[:20]
    ruta_pdf = DIR_PDF / f"INVENTARIO_{nombre_limpio}_{fecha_s.replace('/', '')}.pdf"

    tasa_apl, simb, nom_tasa = obtener_tasa_activa(modo_moneda, t_usd, t_eur)

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.add_page()

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
    pdf.set_font("Helvetica", "B", 13)
    pdf.cell(0, 9, limpiar_pdf(f"{NOMBRE_EMPRESA}  -  REPORTE DE INVENTARIO POR SEDE"), ln=True, align="C", fill=True)

    pdf.set_fill_color(217, 119, 6)
    pdf.set_font("Helvetica", "B", 9.5)
    cintillo = f"SEDE: {sede_filtro.upper()}   |   FECHA: {dia_s} {fecha_s} ({hora_s})   |   TASA {nom_tasa.upper()}: Bs {tasa_apl:,.2f}"
    pdf.cell(0, 6.5, limpiar_pdf(cintillo), ln=True, align="C", fill=True)
    pdf.ln(4)

    pdf.set_fill_color(234, 228, 215)
    pdf.set_text_color(45, 66, 30)
    pdf.set_font("Helvetica", "B", 8.5)
    pdf.cell(62, 8, "  SEDE / CENTRO DIST.", border=1, fill=True)
    pdf.cell(54, 8, "  PRODUCTO / CORTE", border=1, fill=True)
    pdf.cell(24, 8, "STOCK (KG)", border=1, align="R", fill=True)
    pdf.cell(25, 8, f"VALOR ({simb})", border=1, align="R", fill=True)
    pdf.cell(25, 8, "ACTUALIZADO", border=1, align="C", fill=True, ln=True)

    tot_kg = 0.0
    tot_ref = 0.0
    alternar = False

    for r in filas_inv:
        pdf.set_fill_color(246, 244, 238) if alternar else pdf.set_fill_color(255, 255, 255)
        alternar = not alternar
        kg = float(r["kilos"])
        val_ref = float(r["valor_ref"])
        tot_kg += kg
        tot_ref += val_ref

        pdf.set_text_color(20, 20, 20)
        pdf.set_font("Helvetica", "", 8)
        pdf.cell(62, 7, limpiar_pdf(f"  {r['sede']}")[:34], border=1, fill=True)
        pdf.set_font("Helvetica", "B", 8)
        pdf.cell(54, 7, limpiar_pdf(f"  {r['producto']}")[:30], border=1, fill=True)
        pdf.cell(24, 7, f"{kg:,.2f} kg ", border=1, align="R", fill=True)
        pdf.set_font("Helvetica", "", 8)
        pdf.cell(25, 7, limpiar_pdf(f"{simb} {val_ref:,.2f} "), border=1, align="R", fill=True)
        pdf.cell(25, 7, limpiar_pdf(str(r["actualizado"])[:10]), border=1, align="C", fill=True, ln=True)

    pdf.ln(3)
    pdf.set_fill_color(63, 90, 42)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 9.5)
    pdf.cell(60, 8, f"  TOTAL STOCK: {tot_kg:,.2f} KG", border=1, fill=True)
    pdf.cell(62, 8, limpiar_pdf(f"  VALOR REF MAYOR: {simb} {tot_ref:,.2f}"), border=1, align="C", fill=True)
    pdf.cell(68, 8, limpiar_pdf(f"EQUIVALENTE: Bs {tot_ref * tasa_apl:,.2f}  "), border=1, align="R", fill=True, ln=True)

    pdf.output(str(ruta_pdf))
    return ruta_pdf


def exportar_inventario_excel(sede_filtro: str, filas_inv: List[Dict[str, object]], modo_moneda: str, t_usd: float, t_eur: float) -> Path:
    _, fecha_s, hora_s = obtener_marca_tiempo()
    nombre_limpio = sede_filtro.replace(" ", "_").replace("/", "").replace("-", "")[:20]
    ruta_xlsx = DIR_EXCEL / f"INVENTARIO_{nombre_limpio}_{fecha_s.replace('/', '')}.xlsx"

    tasa_apl, simb, nom_tasa = obtener_tasa_activa(modo_moneda, t_usd, t_eur)

    wb = Workbook()
    ws = wb.active
    ws.title = "Inventario_Sedes"

    ws["A1"] = f"{NOMBRE_EMPRESA} - INVENTARIO POR SEDE ({sede_filtro})"
    ws["A1"].font = Font(size=13, bold=True, color="3F5A2A")
    ws.append([f"Emitido: {fecha_s} {hora_s} | Tasa {nom_tasa}: Bs {tasa_apl:,.2f}"])
    ws.append([])

    encabezados = [
        "Sede / Centro de Distribución",
        "Producto / Corte",
        "Stock Disponible (KG)",
        "Estado",
        f"Precio Mayor ({simb})",
        f"Valor Estimado ({simb})",
        f"Valor en Bs ({nom_tasa})",
        "Última Actualización",
    ]
    fila_enc = ws.max_row + 1
    ws.append(encabezados)

    relleno = PatternFill(start_color="3F5A2A", end_color="3F5A2A", fill_type="solid")
    fuente_b = Font(bold=True, color="FFFFFF")
    for idx in range(1, len(encabezados) + 1):
        c = ws.cell(row=fila_enc, column=idx)
        c.fill = relleno
        c.font = fuente_b
        c.alignment = Alignment(horizontal="center")

    tot_kg = 0.0
    tot_ref = 0.0
    for r in filas_inv:
        kg = float(r["kilos"])
        val_ref = float(r["valor_ref"])
        tot_kg += kg
        tot_ref += val_ref
        ws.append([
            r["sede"],
            r["producto"],
            round(kg, 2),
            r["estado"],
            round(float(r["precio_mayor"]), 2),
            round(val_ref, 2),
            round(val_ref * tasa_apl, 2),
            r["actualizado"],
        ])

    ws.append([])
    ws.append(["TOTALES", "", round(tot_kg, 2), "", "", round(tot_ref, 2), round(tot_ref * tasa_apl, 2), ""])

    for col in ws.columns:
        ancho = max(len(str(celda.value or "")) for celda in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(ancho + 4, 15), 42)

    wb.save(str(ruta_xlsx))
    return ruta_xlsx


# ==============================================================================
# PASO 5: COMPONENTE REUTILIZABLE DE CARRITO (KILOS X PRECIO UNITARIO)
# ==============================================================================
class PanelCarritoProductos(ctk.CTkFrame):
    def __init__(self, master, titulo_panel: str, obtener_contexto_moneda: Callable[[], Tuple[str, float, float]]):
        super().__init__(master, fg_color=COLOR_TARJETA, border_width=1, border_color=COLOR_BORDE, corner_radius=12)
        self.obtener_contexto_moneda = obtener_contexto_moneda
        self.titulo_panel = titulo_panel
        self.lineas: List[Dict[str, object]] = []
        self.mapa_productos: Dict[str, sqlite3.Row] = {}

        self._construir_interfaz_carrito()
        self.recargar_productos()

    def _construir_interfaz_carrito(self) -> None:
        cab = ctk.CTkFrame(self, fg_color="transparent")
        cab.pack(fill="x", padx=16, pady=(12, 4))

        ctk.CTkLabel(
            cab,
            text=self.titulo_panel,
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=COLOR_TEXTO_TITULO,
        ).pack(side="left")

        self.lbl_previa = ctk.CTkLabel(
            cab,
            text="Vista previa línea: $ 0.00",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=COLOR_ACENTO,
        )
        self.lbl_previa.pack(side="right")

        barra = ctk.CTkFrame(self, fg_color=COLOR_SUBTARJETA, corner_radius=10)
        barra.pack(fill="x", padx=16, pady=6)
        for col_i, peso in enumerate((3, 2, 1, 1, 0, 0)):
            barra.grid_columnconfigure(col_i, weight=peso)

        ctk.CTkLabel(barra, text="Producto o Corte", font=ctk.CTkFont(size=12, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=10, pady=(8, 2)
        )
        self.cmb_producto = ctk.CTkComboBox(barra, height=32, command=lambda _: self.actualizar_precio_sugerido())
        self.cmb_producto.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 10))

        ctk.CTkLabel(barra, text="Escala Manual", font=ctk.CTkFont(size=12, weight="bold")).grid(
            row=0, column=1, sticky="w", padx=6, pady=(8, 2)
        )
        self.cmb_escala = ctk.CTkComboBox(
            barra,
            height=32,
            values=list(ESCALAS_PRECIO.keys()),
            command=lambda _: self.actualizar_precio_sugerido(),
        )
        self.cmb_escala.grid(row=1, column=1, sticky="ew", padx=6, pady=(0, 10))
        self.cmb_escala.set("Cliente Final")

        ctk.CTkLabel(barra, text="Kilos (KG)", font=ctk.CTkFont(size=12, weight="bold")).grid(
            row=0, column=2, sticky="w", padx=6, pady=(8, 2)
        )
        self.ent_kilos = ctk.CTkEntry(barra, width=105, height=32, placeholder_text="Ej. 15.50")
        self.ent_kilos.grid(row=1, column=2, sticky="ew", padx=6, pady=(0, 10))
        self.ent_kilos.bind("<KeyRelease>", lambda _: self._mostrar_previa_linea())
        self.ent_kilos.bind("<Return>", lambda _: self.agregar_linea())

        ctk.CTkLabel(barra, text="Precio Unit. REF", font=ctk.CTkFont(size=12, weight="bold")).grid(
            row=0, column=3, sticky="w", padx=6, pady=(8, 2)
        )
        self.ent_precio = ctk.CTkEntry(barra, width=105, height=32)
        self.ent_precio.grid(row=1, column=3, sticky="ew", padx=6, pady=(0, 10))
        self.ent_precio.bind("<KeyRelease>", lambda _: self._mostrar_previa_linea())
        self.ent_precio.bind("<Return>", lambda _: self.agregar_linea())

        ctk.CTkButton(
            barra,
            text="➕ Agregar",
            width=125,
            height=32,
            fg_color=COLOR_PRIMARIO,
            hover_color=COLOR_PRIMARIO_HOVER,
            font=ctk.CTkFont(weight="bold"),
            command=self.agregar_linea,
        ).grid(row=1, column=4, padx=(6, 4), pady=(0, 10))

        ctk.CTkButton(
            barra,
            text="🗑️ Eliminar Fila",
            width=125,
            height=32,
            fg_color=COLOR_PELIGRO,
            hover_color=COLOR_PELIGRO_HOVER,
            font=ctk.CTkFont(weight="bold"),
            command=self.quitar_seleccionado,
        ).grid(row=1, column=5, padx=(4, 10), pady=(0, 10))

        marco_tabla = ctk.CTkFrame(self, fg_color="transparent")
        marco_tabla.pack(fill="both", expand=True, padx=16, pady=4)

        cols = [
            ("nro", "#", 42),
            ("prod", "Producto / Corte", 250),
            ("esc", "Escala Aplicada", 125),
            ("kg", "Cantidad (KG)", 110),
            ("unit", "Precio Unit. REF", 120),
            ("sub_ref", "Subtotal REF", 125),
            ("sub_bs", "Subtotal Bs", 145),
        ]
        self.tabla = ttk.Treeview(marco_tabla, columns=[c[0] for c in cols], show="headings", selectmode="browse")
        scroll_y = ttk.Scrollbar(marco_tabla, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=scroll_y.set)

        for cid, texto, ancho in cols:
            self.tabla.heading(cid, text=texto)
            self.tabla.column(cid, width=ancho, anchor="w" if cid == "prod" else "center")

        scroll_y.pack(side="right", fill="y")
        self.tabla.pack(side="left", fill="both", expand=True)

        pie_tot = ctk.CTkFrame(self, fg_color=COLOR_SUBTARJETA, corner_radius=10)
        pie_tot.pack(fill="x", padx=16, pady=(6, 14))

        self.lbl_totales = ctk.CTkLabel(
            pie_tot,
            text="PESO TOTAL: 0.00 KG   |   TOTAL REF: $ 0.00   |   TOTAL BOLÍVARES: Bs 0.00",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=COLOR_ACENTO,
        )
        self.lbl_totales.pack(side="left", padx=14, pady=10)

        ctk.CTkButton(
            pie_tot,
            text="🧹 Vaciar Lista",
            width=130,
            height=30,
            fg_color=COLOR_NEUTRO,
            hover_color=COLOR_NEUTRO_HOVER,
            font=ctk.CTkFont(weight="bold"),
            command=self.solicitar_vaciar_carrito,
        ).pack(side="right", padx=12, pady=8)

    def fijar_escala(self, nombre_escala: str) -> None:
        if nombre_escala in ESCALAS_PRECIO:
            self.cmb_escala.set(nombre_escala)
            self.actualizar_precio_sugerido()

    def recargar_productos(self) -> None:
        filas = consultar_sql("SELECT * FROM productos ORDER BY nombre ASC")
        self.mapa_productos = {f["nombre"]: f for f in filas}
        nombres = list(self.mapa_productos.keys())
        if nombres:
            self.cmb_producto.configure(values=nombres)
            if self.cmb_producto.get() not in self.mapa_productos:
                self.cmb_producto.set(nombres[0])
            self.actualizar_precio_sugerido()

    def actualizar_precio_sugerido(self) -> None:
        prod = self.mapa_productos.get(self.cmb_producto.get())
        if not prod:
            return
        col_db = ESCALAS_PRECIO.get(self.cmb_escala.get(), "precio_cliente")
        precio = float(prod[col_db])
        self.ent_precio.delete(0, "end")
        self.ent_precio.insert(0, f"{precio:.2f}")
        self._mostrar_previa_linea()

    def _mostrar_previa_linea(self) -> None:
        try:
            kg = float(self.ent_kilos.get().replace(",", "."))
            pu = float(self.ent_precio.get().replace(",", "."))
            modo, t_usd, t_eur = self.obtener_contexto_moneda()
            _, simb, _ = obtener_tasa_activa(modo, t_usd, t_eur)
            self.lbl_previa.configure(text=f"Vista previa línea: {simb} {kg * pu:,.2f}")
        except ValueError:
            self.lbl_previa.configure(text="Vista previa línea: 0.00")

    def agregar_linea(self) -> bool:
        nombre_prod = self.cmb_producto.get().strip()
        if not nombre_prod:
            messagebox.showwarning("Producto faltante", "Selecciona un producto del menú desplegable.")
            return False
        try:
            kilos = float(self.ent_kilos.get().replace(",", "."))
            precio_u = float(self.ent_precio.get().replace(",", "."))
            if kilos <= 0 or precio_u < 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Cantidad inválida", "Escribe una cantidad de kilos mayor a 0 y un precio válido.")
            return False

        subtotal_ref = round(kilos * precio_u, 2)
        self.lineas.append({
            "producto": nombre_prod,
            "escala": self.cmb_escala.get(),
            "kilos": round(kilos, 2),
            "precio_unitario": round(precio_u, 2),
            "subtotal_ref": subtotal_ref,
        })
        self.ent_kilos.delete(0, "end")
        self._mostrar_previa_linea()
        self.refrescar_vista()
        return True

    def asegurar_lineas_listas(self) -> bool:
        texto_kg = self.ent_kilos.get().strip().replace(",", ".")
        if texto_kg:
            try:
                if float(texto_kg) > 0:
                    self.agregar_linea()
            except ValueError:
                pass
        return len(self.lineas) > 0

    def quitar_seleccionado(self) -> None:
        if not self.lineas:
            messagebox.showinfo("Lista vacía", "Todavía no hay productos agregados en esta lista.")
            return

        sel = self.tabla.selection()
        if sel:
            indice = self.tabla.index(sel[0])
            if 0 <= indice < len(self.lineas):
                eliminado = self.lineas.pop(indice)
                self.refrescar_vista()
                messagebox.showinfo("Fila Eliminada", f"Se quitó '{eliminado['producto']}' de la lista.")
                return

        ultimo = self.lineas[-1]
        if messagebox.askyesno(
            "Eliminar Producto de la Lista",
            f"No seleccionaste una fila en la tabla.\n\n"
            f"¿Deseas eliminar el último producto agregado ({ultimo['producto']} - {ultimo['kilos']} kg)?",
        ):
            self.lineas.pop()
            self.refrescar_vista()

    def solicitar_vaciar_carrito(self) -> None:
        if not self.lineas:
            return
        if messagebox.askyesno("Vaciar Lista", "¿Deseas quitar todos los productos cargados en esta tabla?"):
            self.limpiar_carrito()

    def limpiar_carrito(self) -> None:
        self.lineas.clear()
        self.refrescar_vista()

    def calcular_totales(self) -> Tuple[float, float, float]:
        modo, t_usd, t_eur = self.obtener_contexto_moneda()
        tasa_apl, _, _ = obtener_tasa_activa(modo, t_usd, t_eur)
        tot_kg = round(sum(float(x["kilos"]) for x in self.lineas), 2)
        tot_ref = round(sum(float(x["subtotal_ref"]) for x in self.lineas), 2)
        tot_bs = round(tot_ref * tasa_apl, 2)
        return tot_kg, tot_ref, tot_bs

    def refrescar_vista(self) -> None:
        self.tabla.delete(*self.tabla.get_children())
        modo, t_usd, t_eur = self.obtener_contexto_moneda()
        tasa_apl, simb, nom_tasa = obtener_tasa_activa(modo, t_usd, t_eur)

        for idx, item in enumerate(self.lineas, start=1):
            sub_ref = float(item["subtotal_ref"])
            tag = "par" if idx % 2 == 0 else "impar"
            self.tabla.insert(
                "",
                "end",
                tags=(tag,),
                values=(
                    idx,
                    item["producto"],
                    item["escala"],
                    f"{float(item['kilos']):,.2f} kg",
                    f"{simb} {float(item['precio_unitario']):,.2f}",
                    f"{simb} {sub_ref:,.2f}",
                    f"Bs {sub_ref * tasa_apl:,.2f}",
                ),
            )

        tot_kg, tot_ref, tot_bs = self.calcular_totales()
        self.lbl_totales.configure(
            text=f"PESO TOTAL: {tot_kg:,.2f} KG   |   TOTAL REF: {simb} {tot_ref:,.2f}   |   TOTAL {nom_tasa.upper()}: Bs {tot_bs:,.2f}"
        )


# ==============================================================================
# PASO 6: APLICACIÓN MAESTRA UNIFICADA (CON CHECKLIST EN MÓDULO 1)
# ==============================================================================
class AppSistemaIntegralFinca(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(TITULO_VENTANA)
        self.geometry("1480x900")
        self.minsize(1250, 760)
        self._aplicar_icono()

        self.modo_oscuro = True
        ctk.set_appearance_mode("Dark")
        self.configure(fg_color=COLOR_FONDO_VENTANA)

        self.ctk_logo: Optional[ctk.CTkImage] = None
        self.id_cliente_edit: Optional[int] = None
        self.id_centro_edit: Optional[int] = None
        self.id_chofer_edit: Optional[int] = None
        self.id_prod_edit: Optional[int] = None

        self.mapa_clientes: Dict[str, sqlite3.Row] = {}
        self.mapa_centros: Dict[str, sqlite3.Row] = {}
        self.mapa_choferes: Dict[str, sqlite3.Row] = {}
        self.cache_inventario_vista: List[Dict[str, object]] = []

        self.estilo_tv = ttk.Style()
        self.aplicar_estilo_tablas()

        self.motor_tasas = MotorTasas(callback_ui=self._actualizar_tasas_desde_hilo)

        self._construir_encabezado()
        self._construir_pestanas_modulos()

        self.motor_tasas.iniciar()
        self.refrescar_datos_maestros()

    def _aplicar_icono(self) -> None:
        if sys.platform.startswith("win") and RUTA_ICONO.exists():
            try:
                self.iconbitmap(str(RUTA_ICONO))
                self.after(250, lambda: self.iconbitmap(str(RUTA_ICONO)))
            except Exception:
                pass

    def aplicar_estilo_tablas(self) -> None:
        self.estilo_tv.theme_use("clam")
        bg_impar = "#143824" if self.modo_oscuro else "#FFFFFF"
        bg_par = "#1E4E33" if self.modo_oscuro else "#F4F1EA"
        fg = "#ECFDF5" if self.modo_oscuro else "#1B4332"
        color_cabecera = "#16A34A" if self.modo_oscuro else "#3F5A2A"

        self.estilo_tv.configure(
            "Treeview",
            background=bg_impar,
            foreground=fg,
            fieldbackground=bg_impar,
            rowheight=28,
            font=("Segoe UI", 9),
            borderwidth=0,
        )
        self.estilo_tv.configure(
            "Treeview.Heading",
            background=color_cabecera,
            foreground="#FFFFFF",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
        )
        self.estilo_tv.map(
            "Treeview",
            background=[("selected", COLOR_ACENTO)],
            foreground=[("selected", "#FFFFFF")],
        )

        tablas_existentes = []
        for attr in ("tv_precios", "tv_clientes", "tv_centros", "tv_choferes", "tv_historial", "tv_inventario"):
            if hasattr(self, attr):
                tablas_existentes.append(getattr(self, attr))
        for c_attr in ("carrito_despacho", "carrito_nc", "carrito_nd"):
            if hasattr(self, c_attr):
                tablas_existentes.append(getattr(self, c_attr).tabla)

        for tv in tablas_existentes:
            tv.tag_configure("impar", background=bg_impar, foreground=fg)
            tv.tag_configure("par", background=bg_par, foreground=fg)
            tv.tag_configure(
                "no_disponible",
                background="#27272A" if self.modo_oscuro else "#E5E7EB",
                foreground="#9CA3AF" if self.modo_oscuro else "#6B7280",
            )
            tv.tag_configure(
                "alerta",
                background="#3B1D1D" if self.modo_oscuro else "#FEE2E2",
                foreground="#FCA5A5" if self.modo_oscuro else "#B91C1C",
            )

    def alternar_modo_visual(self) -> None:
        self.modo_oscuro = not self.modo_oscuro
        ctk.set_appearance_mode("Dark" if self.modo_oscuro else "Light")
        if self.modo_oscuro:
            self.btn_modo.configure(text="☀️ Pasar a Modo Claro", fg_color="#D97706", hover_color="#B45309")
        else:
            self.btn_modo.configure(text="🌙 Pasar a Modo Oscuro", fg_color="#16A34A", hover_color="#15803D")
        self.aplicar_estilo_tablas()

    # --------------------------------------------------------------------------
    # ENCABEZADO SUPERIOR UNIFICADO
    # --------------------------------------------------------------------------
    def _construir_encabezado(self) -> None:
        header = ctk.CTkFrame(
            self, fg_color=COLOR_TARJETA, border_width=2, border_color=COLOR_BORDE, corner_radius=12
        )
        header.pack(fill="x", padx=14, pady=(10, 6))

        self.lbl_logo = ctk.CTkLabel(header, text="")
        self.lbl_logo.pack(side="left", padx=(12, 10), pady=8)
        self._cargar_miniatura_logo()

        bloque_tit = ctk.CTkFrame(header, fg_color="transparent")
        bloque_tit.pack(side="left", pady=8)

        fila_sup = ctk.CTkFrame(bloque_tit, fg_color="transparent")
        fila_sup.pack(anchor="w")
        ctk.CTkLabel(
            fila_sup,
            text="SISTEMA INTEGRAL: PRECIOS, DESPACHOS E INVENTARIO",
            font=ctk.CTkFont(size=17, weight="bold"),
            text_color=COLOR_TEXTO_TITULO,
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            fila_sup,
            text="🖼️ Cambiar Logo",
            width=110,
            height=26,
            fg_color=COLOR_PRIMARIO,
            hover_color=COLOR_PRIMARIO_HOVER,
            font=ctk.CTkFont(size=11, weight="bold"),
            command=self._cambiar_logo_manual,
        ).pack(side="left")

        ctk.CTkLabel(
            bloque_tit,
            text="Agropecuaria Óptima C.A  •  Listas con Checklist, Órdenes, Créditos, Débitos, Choferes y Stock por Sede",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=COLOR_TEXTO_SUAVE,
        ).pack(anchor="w", pady=(3, 0))

        controles = ctk.CTkFrame(header, fg_color=COLOR_SUBTARJETA, corner_radius=10)
        controles.pack(side="right", padx=12, pady=8)

        fila_der_1 = ctk.CTkFrame(controles, fg_color="transparent")
        fila_der_1.pack(fill="x", padx=10, pady=(6, 3))

        ctk.CTkLabel(fila_der_1, text="Moneda PDF:", font=ctk.CTkFont(size=12, weight="bold")).pack(
            side="left", padx=(0, 6)
        )
        self.cmb_moneda = ctk.CTkComboBox(
            fila_der_1,
            width=185,
            height=28,
            values=[
                "REF Tasa BCV ($)",
                "REF Tasa Euro (€)",
                "Bolívares (Tasa BCV $)",
                "Bolívares (Tasa Euro €)",
                "Doble: REF ($) + Bs BCV",
                "Doble: REF (€) + Bs Euro",
            ],
            command=lambda _: self._refrescar_todas_las_vistas_monetarias(),
        )
        self.cmb_moneda.pack(side="left", padx=(0, 10))
        self.cmb_moneda.set("REF Tasa BCV ($)")

        self.btn_modo = ctk.CTkButton(
            fila_der_1,
            text="☀️ Pasar a Modo Claro",
            width=165,
            height=28,
            fg_color="#D97706",
            hover_color="#B45309",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self.alternar_modo_visual,
        )
        self.btn_modo.pack(side="right")

        fila_der_2 = ctk.CTkFrame(controles, fg_color="transparent")
        fila_der_2.pack(fill="x", padx=10, pady=(2, 6))

        tasas = self.motor_tasas.obtener_tasas()
        ctk.CTkLabel(fila_der_2, text="💵 Tasa BCV ($):", font=ctk.CTkFont(size=12, weight="bold")).pack(
            side="left", padx=(0, 4)
        )
        self.ent_usd = ctk.CTkEntry(fila_der_2, width=78, height=28)
        self.ent_usd.pack(side="left", padx=(0, 10))
        self.ent_usd.insert(0, f"{tasas['bcv_usd']:.2f}")

        ctk.CTkLabel(fila_der_2, text="💶 Tasa Euro (€):", font=ctk.CTkFont(size=12, weight="bold")).pack(
            side="left", padx=(0, 4)
        )
        self.ent_eur = ctk.CTkEntry(fila_der_2, width=78, height=28)
        self.ent_eur.pack(side="left", padx=(0, 10))
        self.ent_eur.insert(0, f"{tasas['bcv_eur']:.2f}")

        ctk.CTkButton(
            fila_der_2,
            text="🔄 BCV",
            width=68,
            height=28,
            fg_color=COLOR_ACENTO,
            hover_color=COLOR_ACENTO_HOVER,
            font=ctk.CTkFont(size=11, weight="bold"),
            command=self._forzar_bcv_vivo,
        ).pack(side="right", padx=(4, 0))

        ctk.CTkButton(
            fila_der_2,
            text="💾 Fijar",
            width=68,
            height=28,
            fg_color=COLOR_PRIMARIO,
            hover_color=COLOR_PRIMARIO_HOVER,
            font=ctk.CTkFont(size=11, weight="bold"),
            command=self._guardar_tasa_manual,
        ).pack(side="right")

    def _cargar_miniatura_logo(self) -> None:
        if not RUTA_LOGO.exists():
            return
        try:
            with Image.open(RUTA_LOGO) as im:
                copia = im.convert("RGBA").copy()
            w, h = copia.size
            alto = 56
            ancho = min(max(int(alto * (w / max(h, 1))), 56), 165)
            self.ctk_logo = ctk.CTkImage(light_image=copia, dark_image=copia, size=(ancho, alto))
            self.lbl_logo.configure(image=self.ctk_logo, text="")
        except Exception:
            self.lbl_logo.configure(text="🚜")

    def _cambiar_logo_manual(self) -> None:
        ruta = filedialog.askopenfilename(
            title="Seleccionar Logo", filetypes=[("Imágenes", "*.jpeg *.jpg *.png *.webp")]
        )
        if ruta and recortar_y_guardar_logo(Path(ruta)):
            self._cargar_miniatura_logo()
            self._aplicar_icono()
            messagebox.showinfo("Logo Actualizado", "El logotipo fue actualizado correctamente.")

    def obtener_contexto_moneda(self) -> Tuple[str, float, float]:
        try:
            u = max(float(self.ent_usd.get().replace(",", ".")), 0.01)
            e = max(float(self.ent_eur.get().replace(",", ".")), 0.01)
        except ValueError:
            t = self.motor_tasas.obtener_tasas()
            u, e = t["bcv_usd"], t["bcv_eur"]
        return self.cmb_moneda.get(), u, e

    def _guardar_tasa_manual(self) -> None:
        try:
            u = float(self.ent_usd.get().replace(",", "."))
            e = float(self.ent_eur.get().replace(",", "."))
            self.motor_tasas.guardar_tasas(u, e, fuente="Manual")
            self._refrescar_todas_las_vistas_monetarias()
            messagebox.showinfo("Tasas Guardadas", f"BCV ($): Bs {u:.2f}\nEuro (€): Bs {e:.2f}")
        except ValueError:
            messagebox.showerror("Error", "Ingresa montos numéricos válidos para las tasas.")

    def _forzar_bcv_vivo(self) -> None:
        def tarea():
            ok = self.motor_tasas.sincronizar_web()
            msg = "Tasas sincronizadas con el BCV." if ok else "No hubo respuesta del BCV; se mantienen las tasas actuales."
            self.after(0, lambda: messagebox.showinfo("Sincronización BCV", msg))

        threading.Thread(target=tarea, daemon=True).start()

    def _actualizar_tasas_desde_hilo(self, tasas: Dict[str, float]) -> None:
        def pintar():
            self.ent_usd.delete(0, "end")
            self.ent_usd.insert(0, f"{tasas['bcv_usd']:.2f}")
            self.ent_eur.delete(0, "end")
            self.ent_eur.insert(0, f"{tasas['bcv_eur']:.2f}")
            self._refrescar_todas_las_vistas_monetarias()

        self.after(0, pintar)

    def _refrescar_todas_las_vistas_monetarias(self) -> None:
        for carrito in (self.carrito_despacho, self.carrito_nc, self.carrito_nd):
            carrito.refrescar_vista()
        self.refrescar_tabla_catalogo()
        self.refrescar_tabla_inventario()

    # --------------------------------------------------------------------------
    # PESTAÑAS DE LOS 7 MÓDULOS DEL SISTEMA UNIFICADO
    # --------------------------------------------------------------------------
    def _construir_pestanas_modulos(self) -> None:
        self.tabs = ctk.CTkTabview(
            self,
            fg_color="transparent",
            segmented_button_selected_color=COLOR_PRIMARIO,
            segmented_button_selected_hover_color=COLOR_PRIMARIO_HOVER,
        )
        self.tabs.pack(fill="both", expand=True, padx=14, pady=(0, 10))

        self.tab_listas = self.tabs.add("📋 1. Listas de Precios")
        self.tab_despachos = self.tabs.add("📦 2. Órdenes de Despacho")
        self.tab_credito = self.tabs.add("💳 3. Notas de Crédito (CD)")
        self.tab_debito_mov = self.tabs.add("🔄 4. Débitos y Movimientos")
        self.tab_inventario = self.tabs.add("📊 5. Inventario por Sede")
        self.tab_bases_datos = self.tabs.add("👥 6. Clientes, Centros y Choferes")
        self.tab_historial = self.tabs.add("📜 7. Historial General")

        self._armar_modulo_listas_precios()
        self._armar_modulo_despachos()
        self._armar_submodulo_notas_credito()
        self._armar_submodulo_debito_y_movimientos()
        self._armar_modulo_inventario()
        self._armar_modulo_bases_datos()
        self._armar_modulo_historial()
        self.aplicar_estilo_tablas()

    # ==========================================================================
    # MÓDULO 1: LISTAS DE PRECIOS CON CHECKLIST MANUAL DE DISPONIBILIDAD
    # ==========================================================================
    def _armar_modulo_listas_precios(self) -> None:
        izq = ctk.CTkFrame(
            self.tab_listas, width=405, fg_color=COLOR_TARJETA, border_width=1, border_color=COLOR_BORDE, corner_radius=12
        )
        izq.pack(side="left", fill="y", padx=(0, 10), pady=4)
        izq.pack_propagate(False)

        # Botones inferiores de emisión de Lista de Precios en PDF y Excel
        marco_exp = ctk.CTkFrame(izq, fg_color="transparent")
        marco_exp.pack(fill="x", side="bottom", padx=14, pady=10)

        ctk.CTkButton(
            marco_exp,
            text="📄 GENERAR LISTA DE DISPONIBLES (PDF)",
            height=38,
            fg_color=COLOR_ACENTO,
            hover_color=COLOR_ACENTO_HOVER,
            font=ctk.CTkFont(size=12, weight="bold"),
            command=lambda: self._emitir_lista_precios_oficial(False),
        ).pack(fill="x", pady=(0, 5))

        ctk.CTkButton(
            marco_exp,
            text="📊 Exportar Lista de Disponibles (Excel)",
            height=34,
            fg_color=COLOR_PRIMARIO,
            hover_color=COLOR_PRIMARIO_HOVER,
            font=ctk.CTkFont(weight="bold"),
            command=lambda: self._emitir_lista_precios_oficial(True),
        ).pack(fill="x")

        ctk.CTkLabel(
            izq, text="📋 GENERADOR DE LISTAS DE PRECIOS", font=ctk.CTkFont(size=14, weight="bold"), text_color=COLOR_TEXTO_TITULO
        ).pack(anchor="w", padx=16, pady=(10, 4))

        # Sub-tarjeta 1: Configuración de la Lista a Imprimir
        card_cfg = ctk.CTkFrame(izq, fg_color=COLOR_SUBTARJETA, corner_radius=10)
        card_cfg.pack(fill="x", padx=14, pady=3)

        ctk.CTkLabel(card_cfg, text="Precios Elegibles para el PDF / Excel", font=ctk.CTkFont(size=11, weight="bold")).pack(
            anchor="w", padx=12, pady=(6, 1)
        )
        self.cmb_modalidad_lista = ctk.CTkComboBox(card_cfg, height=28, values=MODALIDADES_LISTA_PRECIOS)
        self.cmb_modalidad_lista.pack(fill="x", padx=12, pady=1)
        self.cmb_modalidad_lista.set(MODALIDADES_LISTA_PRECIOS[0])

        ctk.CTkLabel(card_cfg, text="Filtrar por Categoría / Rubro", font=ctk.CTkFont(size=11, weight="bold")).pack(
            anchor="w", padx=12, pady=(4, 1)
        )
        self.cmb_filtro_rubro_lista = ctk.CTkComboBox(
            card_cfg,
            height=28,
            values=["Todos los Rubros"] + RUBROS_CATALOGO,
            command=lambda _: self.refrescar_tabla_catalogo(),
        )
        self.cmb_filtro_rubro_lista.pack(fill="x", padx=12, pady=1)
        self.cmb_filtro_rubro_lista.set("Todos los Rubros")

        self.chk_solo_disponibles_lista = ctk.CTkCheckBox(
            card_cfg,
            text="Incluir solo los productos marcados (☑ Disponible)",
            font=ctk.CTkFont(size=11, weight="bold"),
            fg_color=COLOR_PRIMARIO,
            hover_color=COLOR_PRIMARIO_HOVER,
        )
        self.chk_solo_disponibles_lista.pack(anchor="w", padx=12, pady=(6, 8))
        self.chk_solo_disponibles_lista.select()

        # Sub-tarjeta 2: Editor de Producto (Con check de disponibilidad)
        card_ed = ctk.CTkFrame(izq, fg_color=COLOR_SUBTARJETA, corner_radius=10)
        card_ed.pack(fill="x", padx=14, pady=4)
        for i in range(3):
            card_ed.grid_columnconfigure(i, weight=1)

        ctk.CTkLabel(card_ed, text="🥩 Agregar o Editar Corte del Catálogo", font=ctk.CTkFont(size=11, weight="bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", padx=10, pady=(6, 2)
        )

        ctk.CTkLabel(card_ed, text="Nombre del Producto", font=ctk.CTkFont(size=11, weight="bold")).grid(
            row=1, column=0, columnspan=2, sticky="w", padx=8
        )
        ctk.CTkLabel(card_ed, text="Categoría", font=ctk.CTkFont(size=11, weight="bold")).grid(
            row=1, column=2, sticky="w", padx=8
        )

        self.ent_p_nom = ctk.CTkEntry(card_ed, height=27, placeholder_text="Ej. CHULETA FRESCA")
        self.ent_p_nom.grid(row=2, column=0, columnspan=2, sticky="ew", padx=8, pady=(1, 4))
        self.cmb_p_rubro = ctk.CTkComboBox(card_ed, height=27, values=RUBROS_CATALOGO)
        self.cmb_p_rubro.grid(row=2, column=2, sticky="ew", padx=8, pady=(1, 4))
        self.cmb_p_rubro.set(RUBROS_CATALOGO[0])

        ctk.CTkLabel(card_ed, text="Al Mayor ($)", font=ctk.CTkFont(size=11, weight="bold")).grid(row=3, column=0, sticky="w", padx=8)
        ctk.CTkLabel(card_ed, text="Vendedor 10% ($)", font=ctk.CTkFont(size=11, weight="bold")).grid(row=3, column=1, sticky="w", padx=8)
        ctk.CTkLabel(card_ed, text="Cliente Final ($)", font=ctk.CTkFont(size=11, weight="bold")).grid(row=3, column=2, sticky="w", padx=8)

        self.ent_p_may = ctk.CTkEntry(card_ed, height=27, placeholder_text="0.00")
        self.ent_p_may.grid(row=4, column=0, sticky="ew", padx=8, pady=(1, 4))
        self.ent_p_ven = ctk.CTkEntry(card_ed, height=27, placeholder_text="0.00")
        self.ent_p_ven.grid(row=4, column=1, sticky="ew", padx=8, pady=(1, 4))
        self.ent_p_cli = ctk.CTkEntry(card_ed, height=27, placeholder_text="0.00")
        self.ent_p_cli.grid(row=4, column=2, sticky="ew", padx=8, pady=(1, 4))

        self.chk_prod_disp_form = ctk.CTkCheckBox(
            card_ed,
            text="Producto Disponible para la Lista (☑)",
            font=ctk.CTkFont(size=11, weight="bold"),
            fg_color=COLOR_PRIMARIO,
            hover_color=COLOR_PRIMARIO_HOVER,
        )
        self.chk_prod_disp_form.grid(row=5, column=0, columnspan=3, sticky="w", padx=8, pady=(2, 6))
        self.chk_prod_disp_form.select()

        bp = ctk.CTkFrame(card_ed, fg_color="transparent")
        bp.grid(row=6, column=0, columnspan=3, sticky="ew", padx=8, pady=(0, 8))
        for i in range(3):
            bp.grid_columnconfigure(i, weight=1)

        ctk.CTkButton(bp, text="💾 Guardar", height=28, fg_color=COLOR_PRIMARIO, hover_color=COLOR_PRIMARIO_HOVER, font=ctk.CTkFont(weight="bold"), command=self._guardar_producto_cat).grid(row=0, column=0, sticky="ew", padx=(0, 4))
        ctk.CTkButton(bp, text="🧹 Limpiar", height=28, fg_color=COLOR_NEUTRO, hover_color=COLOR_NEUTRO_HOVER, command=self._limpiar_form_producto).grid(row=0, column=1, sticky="ew", padx=4)
        ctk.CTkButton(bp, text="🗑️ Eliminar", height=28, fg_color=COLOR_PELIGRO, hover_color=COLOR_PELIGRO_HOVER, font=ctk.CTkFont(weight="bold"), command=self._eliminar_producto_cat).grid(row=0, column=2, sticky="ew", padx=(4, 0))

        # Sub-tarjeta 3: Ajuste Porcentual Rápido o Importación Externa
        card_mas = ctk.CTkFrame(izq, fg_color=COLOR_SUBTARJETA, corner_radius=10)
        card_mas.pack(fill="x", padx=14, pady=3)

        ctk.CTkLabel(card_mas, text="⚡ Ajuste Masivo (%) o Sincronización", font=ctk.CTkFont(size=11, weight="bold")).pack(
            anchor="w", padx=12, pady=(5, 2)
        )
        fila_m = ctk.CTkFrame(card_mas, fg_color="transparent")
        fila_m.pack(fill="x", padx=12, pady=(1, 6))

        self.ent_porcentaje_masivo = ctk.CTkEntry(fila_m, width=85, height=27, placeholder_text="Ej. 5 o -5")
        self.ent_porcentaje_masivo.pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            fila_m,
            text="Aplicar %",
            width=95,
            height=27,
            fg_color=COLOR_NEUTRO,
            hover_color=COLOR_NEUTRO_HOVER,
            font=ctk.CTkFont(size=11, weight="bold"),
            command=self._aplicar_porcentaje_catalogo,
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            fila_m,
            text="📂 Importar .DB",
            width=125,
            height=27,
            fg_color=COLOR_PRIMARIO,
            hover_color=COLOR_PRIMARIO_HOVER,
            font=ctk.CTkFont(size=11, weight="bold"),
            command=self._importar_precios_app_anterior,
        ).pack(side="right")

        # Panel Derecho: Checklist Interactivo de Disponibilidad + Tabla de Precios
        der = ctk.CTkFrame(
            self.tab_listas, fg_color=COLOR_TARJETA, border_width=1, border_color=COLOR_BORDE, corner_radius=12
        )
        der.pack(side="right", fill="both", expand=True, pady=4)

        cab_der = ctk.CTkFrame(der, fg_color="transparent")
        cab_der.pack(fill="x", padx=16, pady=(12, 4))

        ctk.CTkLabel(
            cab_der,
            text="☑️ CHECKLIST DE PRODUCTOS DISPONIBLES (HAZ CLIC EN LA CASILLA ☑ / ☐ PARA MARCAR O QUITAR)",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=COLOR_TEXTO_TITULO,
        ).pack(side="left")

        self.lbl_contador_disp = ctk.CTkLabel(
            cab_der,
            text="Disponibles: 0 de 0",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=COLOR_ACENTO,
        )
        self.lbl_contador_disp.pack(side="right")

        # Barra de botones rápidos del Checklist Manual
        barra_chk = ctk.CTkFrame(der, fg_color=COLOR_SUBTARJETA, corner_radius=10)
        barra_chk.pack(fill="x", padx=16, pady=4)

        ctk.CTkLabel(
            barra_chk,
            text="Control Rápido de Disponibilidad (sin borrar productos):",
            font=ctk.CTkFont(size=12, weight="bold"),
        ).pack(side="left", padx=(12, 10), pady=8)

        ctk.CTkButton(
            barra_chk,
            text="☑️ Marcar / Desmarcar Seleccionado",
            width=225,
            height=30,
            fg_color=COLOR_ACENTO,
            hover_color=COLOR_ACENTO_HOVER,
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self._alternar_disponibilidad_seleccionado,
        ).pack(side="left", padx=4, pady=8)

        ctk.CTkButton(
            barra_chk,
            text="⬜ Desmarcar Todos",
            width=145,
            height=30,
            fg_color=COLOR_NEUTRO,
            hover_color=COLOR_NEUTRO_HOVER,
            font=ctk.CTkFont(size=12, weight="bold"),
            command=lambda: self._cambiar_disponibilidad_masiva(0),
        ).pack(side="right", padx=(4, 12), pady=8)

        ctk.CTkButton(
            barra_chk,
            text="✅ Marcar Todos",
            width=140,
            height=30,
            fg_color=COLOR_PRIMARIO,
            hover_color=COLOR_PRIMARIO_HOVER,
            font=ctk.CTkFont(size=12, weight="bold"),
            command=lambda: self._cambiar_disponibilidad_masiva(1),
        ).pack(side="right", padx=4, pady=8)

        mc = ctk.CTkFrame(der, fg_color="transparent")
        mc.pack(fill="both", expand=True, padx=16, pady=(4, 14))
        cols_p = [
            ("disp", "Checklist Lista", 125),
            ("id", "ID", 38),
            ("nom", "Producto / Corte", 210),
            ("rub", "Categoría", 135),
            ("may", "Al Mayor ($)", 90),
            ("ven", "Vendedor 10% ($)", 105),
            ("cli", "Cliente Final ($)", 105),
            ("cli_bs", "Cliente Final (Bs)", 125),
        ]
        self.tv_precios = ttk.Treeview(mc, columns=[c[0] for c in cols_p], show="headings", selectmode="browse")
        sc_p = ttk.Scrollbar(mc, orient="vertical", command=self.tv_precios.yview)
        self.tv_precios.configure(yscrollcommand=sc_p.set)
        for cid, tit, anc in cols_p:
            self.tv_precios.heading(cid, text=tit)
            self.tv_precios.column(cid, width=anc, anchor="w" if cid == "nom" else "center")
        sc_p.pack(side="right", fill="y")
        self.tv_precios.pack(side="left", fill="both", expand=True)

        # Eventos para que el Checklist funcione con 1 clic en la primera columna, doble clic o barra espaciadora
        self.tv_precios.bind("<<TreeviewSelect>>", self._seleccionar_prod_cat)
        self.tv_precios.bind("<ButtonRelease-1>", self._al_hacer_clic_en_tabla_precios)
        self.tv_precios.bind("<Double-1>", lambda _: self._alternar_disponibilidad_seleccionado())
        self.tv_precios.bind("<space>", lambda _: self._alternar_disponibilidad_seleccionado())

    def _al_hacer_clic_en_tabla_precios(self, event) -> None:
        """
        Si el usuario hace clic directamente sobre la primera columna ('Checklist Lista'),
        marca o desmarca el producto al instante sin borrarlo.
        """
        region = self.tv_precios.identify("region", event.x, event.y)
        if region != "cell":
            return
        columna = self.tv_precios.identify_column(event.x)
        fila_id = self.tv_precios.identify_row(event.y)
        if columna == "#1" and fila_id:
            valores = self.tv_precios.item(fila_id, "values")
            if valores and len(valores) >= 2:
                pid = int(valores[1])
                self._invertir_disponibilidad_producto_db(pid)

    def _invertir_disponibilidad_producto_db(self, producto_id: int) -> None:
        """Cambia disponible de 1 a 0 (o de 0 a 1) en SQLite y refresca la vista."""
        filas = consultar_sql("SELECT disponible FROM productos WHERE id = ?", (producto_id,))
        if not filas:
            return
        nuevo_estado = 0 if int(filas[0]["disponible"]) == 1 else 1
        ejecutar_sql("UPDATE productos SET disponible = ? WHERE id = ?", (nuevo_estado, producto_id))
        if nuevo_estado == 1:
            self.chk_prod_disp_form.select()
        else:
            self.chk_prod_disp_form.deselect()
        self.refrescar_tabla_catalogo()

    def _alternar_disponibilidad_seleccionado(self) -> None:
        """Alterna la disponibilidad del producto marcado actualmente en la tabla."""
        sel = self.tv_precios.selection()
        pid = int(self.tv_precios.item(sel[0], "values")[1]) if sel else self.id_prod_edit
        if not pid:
            messagebox.showinfo(
                "Selecciona un Producto",
                "Haz clic sobre un producto en la tabla para marcarlo (☑ Disponible) o desmarcarlo (☐ No Disponible).",
            )
            return
        self._invertir_disponibilidad_producto_db(pid)

    def _cambiar_disponibilidad_masiva(self, nuevo_estado: int) -> None:
        """Marca todos (1) o desmarca todos (0) los productos de la vista actual sin borrarlos."""
        rubro_f = self.cmb_filtro_rubro_lista.get()
        if rubro_f == "Todos los Rubros":
            ejecutar_sql("UPDATE productos SET disponible = ?", (nuevo_estado,))
        else:
            ejecutar_sql("UPDATE productos SET disponible = ? WHERE rubro = ?", (nuevo_estado, rubro_f))
        self.refrescar_tabla_catalogo()

    def _emitir_lista_precios_oficial(self, en_excel: bool) -> None:
        try:
            modalidad = self.cmb_modalidad_lista.get()
            rubro = self.cmb_filtro_rubro_lista.get()
            solo_disp = (self.chk_solo_disponibles_lista.get() == 1)
            modo_mon, t_usd, t_eur = self.obtener_contexto_moneda()
            ruta = (
                exportar_lista_precios_excel(modalidad, rubro, modo_mon, t_usd, t_eur, solo_disp)
                if en_excel
                else exportar_lista_precios_pdf(modalidad, rubro, modo_mon, t_usd, t_eur, solo_disp)
            )
            messagebox.showinfo("Lista de Precios Creada", f"Documento generado con éxito en:\n\n{ruta}")
            abrir_archivo(ruta)
        except Exception as error:
            messagebox.showwarning("Aviso al Generar Lista", str(error))

    def _aplicar_porcentaje_catalogo(self) -> None:
        try:
            porc = float(self.ent_porcentaje_masivo.get().replace(",", "."))
        except ValueError:
            messagebox.showwarning("Porcentaje Inválido", "Escribe un número (ej. 5 para subir 5% o -5 para bajar 5%).")
            return
        if not messagebox.askyesno(
            "Confirmar Ajuste Masivo",
            f"¿Deseas aplicar un ajuste de {porc:+.2f}% a todos los precios del catálogo?",
        ):
            return
        factor = 1.0 + (porc / 100.0)
        ejecutar_sql(
            """UPDATE productos SET
               precio_mayor = ROUND(precio_mayor * ?, 2),
               precio_vendedor = ROUND(precio_vendedor * ?, 2),
               precio_cliente = ROUND(precio_cliente * ?, 2)""",
            (factor, factor, factor),
        )
        self.ent_porcentaje_masivo.delete(0, "end")
        for c in (self.carrito_despacho, self.carrito_nc, self.carrito_nd):
            c.recargar_productos()
        self.refrescar_datos_maestros()
        messagebox.showinfo("Precios Actualizados", f"Se aplicó el {porc:+.2f}% en todo el catálogo.")

    # ==========================================================================
    # MÓDULO 2: ÓRDENES DE DESPACHO
    # ==========================================================================
    def _armar_modulo_despachos(self) -> None:
        izq = ctk.CTkFrame(
            self.tab_despachos, width=385, fg_color=COLOR_TARJETA, border_width=1, border_color=COLOR_BORDE, corner_radius=12
        )
        izq.pack(side="left", fill="y", padx=(0, 10), pady=4)
        izq.pack_propagate(False)

        marco_emitir = ctk.CTkFrame(izq, fg_color="transparent")
        marco_emitir.pack(fill="x", side="bottom", padx=14, pady=12)

        ctk.CTkButton(
            marco_emitir,
            text="📄 EMITIR ORDEN DE DESPACHO (PDF)",
            height=40,
            fg_color=COLOR_ACENTO,
            hover_color=COLOR_ACENTO_HOVER,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=lambda: self._procesar_emision_despacho(False),
        ).pack(fill="x", pady=(0, 6))

        ctk.CTkButton(
            marco_emitir,
            text="📊 Exportar Orden en Excel (.xlsx)",
            height=36,
            fg_color=COLOR_PRIMARIO,
            hover_color=COLOR_PRIMARIO_HOVER,
            font=ctk.CTkFont(weight="bold"),
            command=lambda: self._procesar_emision_despacho(True),
        ).pack(fill="x")

        ctk.CTkLabel(
            izq, text="📦 DATOS DE LA ORDEN DE DESPACHO", font=ctk.CTkFont(size=14, weight="bold"), text_color=COLOR_TEXTO_TITULO
        ).pack(anchor="w", padx=16, pady=(14, 8))

        card_sel = ctk.CTkFrame(izq, fg_color=COLOR_SUBTARJETA, corner_radius=10)
        card_sel.pack(fill="x", padx=14, pady=4)

        ctk.CTkLabel(card_sel, text="Cliente Registrado", font=ctk.CTkFont(size=12, weight="bold")).pack(
            anchor="w", padx=12, pady=(8, 2)
        )
        self.cmb_cliente_desp = ctk.CTkComboBox(card_sel, height=32, command=self._al_elegir_cliente_despacho)
        self.cmb_cliente_desp.pack(fill="x", padx=12, pady=2)

        self.lbl_info_cliente = ctk.CTkLabel(
            card_sel, text="Documento: -   |   Escala: -", font=ctk.CTkFont(size=12, weight="bold"), text_color=COLOR_ACENTO
        )
        self.lbl_info_cliente.pack(anchor="w", padx=12, pady=(4, 8))

        card_nuevo = ctk.CTkFrame(izq, fg_color=COLOR_SUBTARJETA, corner_radius=10)
        card_nuevo.pack(fill="x", padx=14, pady=6)

        ctk.CTkLabel(card_nuevo, text="⚡ Registro Rápido de Cliente", font=ctk.CTkFont(size=12, weight="bold")).pack(
            anchor="w", padx=12, pady=(8, 4)
        )
        self.ent_rap_nom = ctk.CTkEntry(card_nuevo, height=30, placeholder_text="Nombre o Razón Social")
        self.ent_rap_nom.pack(fill="x", padx=12, pady=3)

        fila_r = ctk.CTkFrame(card_nuevo, fg_color="transparent")
        fila_r.pack(fill="x", padx=12, pady=3)
        fila_r.grid_columnconfigure(0, weight=1)
        fila_r.grid_columnconfigure(1, weight=1)

        self.ent_rap_rif = ctk.CTkEntry(fila_r, height=30, placeholder_text="RIF o Cédula")
        self.ent_rap_rif.grid(row=0, column=0, sticky="ew", padx=(0, 3))
        self.cmb_rap_esc = ctk.CTkComboBox(fila_r, height=30, values=list(ESCALAS_PRECIO.keys()))
        self.cmb_rap_esc.grid(row=0, column=1, sticky="ew", padx=(3, 0))
        self.cmb_rap_esc.set("Cliente Final")

        ctk.CTkButton(
            card_nuevo,
            text="💾 Guardar y Usar Cliente",
            height=30,
            fg_color=COLOR_PRIMARIO,
            hover_color=COLOR_PRIMARIO_HOVER,
            font=ctk.CTkFont(weight="bold"),
            command=self._alta_rapida_cliente,
        ).pack(fill="x", padx=12, pady=(6, 10))

        card_det = ctk.CTkFrame(izq, fg_color=COLOR_SUBTARJETA, corner_radius=10)
        card_det.pack(fill="x", padx=14, pady=4)

        ctk.CTkLabel(card_det, text="Sede que Despacha (Descuenta Stock)", font=ctk.CTkFont(size=12, weight="bold")).pack(
            anchor="w", padx=12, pady=(8, 2)
        )
        self.cmb_origen_desp = ctk.CTkComboBox(card_det, height=30, values=[SEDE_PLANTA_PRINCIPAL])
        self.cmb_origen_desp.pack(fill="x", padx=12, pady=2)

        ctk.CTkLabel(card_det, text="Observaciones de Entrega", font=ctk.CTkFont(size=12, weight="bold")).pack(
            anchor="w", padx=12, pady=(6, 2)
        )
        self.ent_obs_desp = ctk.CTkEntry(card_det, height=30, placeholder_text="Ej. Pago móvil / Retiro en sede")
        self.ent_obs_desp.pack(fill="x", padx=12, pady=2)

        self.chk_descontar_inv_desp = ctk.CTkCheckBox(
            card_det,
            text="Descontar kilos del inventario de la sede",
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color=COLOR_PRIMARIO,
            hover_color=COLOR_PRIMARIO_HOVER,
        )
        self.chk_descontar_inv_desp.pack(anchor="w", padx=12, pady=(8, 10))
        self.chk_descontar_inv_desp.select()

        self.carrito_despacho = PanelCarritoProductos(
            self.tab_despachos, "🛒 PRODUCTOS DEL DESPACHO (KILOS X PRECIO UNITARIO)", self.obtener_contexto_moneda
        )
        self.carrito_despacho.pack(side="right", fill="both", expand=True, pady=4)

    def _al_elegir_cliente_despacho(self, seleccion: str) -> None:
        cli = self.mapa_clientes.get(seleccion)
        if not cli:
            return
        self.lbl_info_cliente.configure(
            text=f"Documento: {cli['documento_id']}   |   Escala: {cli['escala_preferida']}"
        )
        self.carrito_despacho.fijar_escala(cli["escala_preferida"])

    def _alta_rapida_cliente(self) -> None:
        nom = self.ent_rap_nom.get().strip().upper()
        rif = self.ent_rap_rif.get().strip().upper() or "S/D"
        esc = self.cmb_rap_esc.get()
        if not nom:
            messagebox.showwarning("Falta Nombre", "Escribe el nombre del cliente para registrarlo.")
            return
        ejecutar_sql(
            "INSERT INTO clientes (nombre, documento_id, escala_preferida) VALUES (?, ?, ?)",
            (nom, rif, esc),
        )
        self.ent_rap_nom.delete(0, "end")
        self.ent_rap_rif.delete(0, "end")
        self.refrescar_datos_maestros()
        clave = f"{nom} ({rif})"
        if clave in self.mapa_clientes:
            self.cmb_cliente_desp.set(clave)
            self._al_elegir_cliente_despacho(clave)
        messagebox.showinfo("Cliente Registrado", f"El cliente '{nom}' fue guardado y seleccionado.")

    def _procesar_emision_despacho(self, exportar_excel: bool) -> None:
        try:
            if not self.carrito_despacho.asegurar_lineas_listas():
                messagebox.showwarning(
                    "Orden Vacía",
                    "Escribe los kilos en 'Kilos (KG)' y presiona '➕ Agregar' en la tabla derecha antes de emitir.",
                )
                return
            cli = self.mapa_clientes.get(self.cmb_cliente_desp.get())
            if not cli:
                messagebox.showwarning("Falta Cliente", "Selecciona o registra un cliente antes de emitir la orden.")
                return

            afectar_inv = 1 if self.chk_descontar_inv_desp.get() == 1 else 0
            escala_actual = self.carrito_despacho.cmb_escala.get()
            doc_id = self._guardar_documento_en_db(
                tipo_doc="DESPACHO",
                subtipo="ORDEN DE DESPACHO",
                origen=self.cmb_origen_desp.get(),
                destino=cli["nombre"],
                doc_identidad=cli["documento_id"],
                referencia="",
                motivo=f"Escala: {escala_actual}",
                afecto_inventario=afectar_inv,
                observaciones=self.ent_obs_desp.get().strip(),
                carrito=self.carrito_despacho,
            )
            ruta = exportar_documento_excel(doc_id) if exportar_excel else exportar_documento_pdf(doc_id)
            self.carrito_despacho.limpiar_carrito()
            self.ent_obs_desp.delete(0, "end")
            self.refrescar_datos_maestros()
            messagebox.showinfo("Orden Emitida con Éxito", f"Tu Orden de Despacho fue creada en:\n\n{ruta}")
            abrir_archivo(ruta)
        except Exception as error:
            messagebox.showerror("Error al emitir Orden de Despacho", str(error))

    # ==========================================================================
    # MÓDULO 3: NOTAS DE CRÉDITO A CENTROS DE DISTRIBUCIÓN
    # ==========================================================================
    def _armar_submodulo_notas_credito(self) -> None:
        izq = ctk.CTkFrame(
            self.tab_credito, width=385, fg_color=COLOR_TARJETA, border_width=1, border_color=COLOR_BORDE, corner_radius=12
        )
        izq.pack(side="left", fill="y", padx=(0, 10), pady=4)
        izq.pack_propagate(False)

        marco_btn = ctk.CTkFrame(izq, fg_color="transparent")
        marco_btn.pack(fill="x", side="bottom", padx=14, pady=14)

        ctk.CTkButton(
            marco_btn,
            text="📄 GENERAR NOTA DE CRÉDITO (PDF)",
            height=40,
            fg_color=COLOR_ACENTO,
            hover_color=COLOR_ACENTO_HOVER,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=lambda: self._procesar_nota_credito(False),
        ).pack(fill="x", pady=(0, 6))

        ctk.CTkButton(
            marco_btn,
            text="📊 Exportar Nota de Crédito (Excel)",
            height=36,
            fg_color=COLOR_PRIMARIO,
            hover_color=COLOR_PRIMARIO_HOVER,
            font=ctk.CTkFont(weight="bold"),
            command=lambda: self._procesar_nota_credito(True),
        ).pack(fill="x")

        ctk.CTkLabel(
            izq, text="💳 NOTA DE CRÉDITO A CENTRO DIST.", font=ctk.CTkFont(size=14, weight="bold"), text_color=COLOR_TEXTO_TITULO
        ).pack(anchor="w", padx=16, pady=(14, 8))

        card_nc = ctk.CTkFrame(izq, fg_color=COLOR_SUBTARJETA, corner_radius=10)
        card_nc.pack(fill="x", padx=14, pady=6)

        ctk.CTkLabel(card_nc, text="Centro de Distribución Beneficiario", font=ctk.CTkFont(size=12, weight="bold")).pack(
            anchor="w", padx=12, pady=(10, 2)
        )
        self.cmb_cd_credito = ctk.CTkComboBox(card_nc, height=32)
        self.cmb_cd_credito.pack(fill="x", padx=12, pady=2)

        ctk.CTkLabel(card_nc, text="Motivo de Acreditación", font=ctk.CTkFont(size=12, weight="bold")).pack(
            anchor="w", padx=12, pady=(8, 2)
        )
        self.cmb_motivo_nc = ctk.CTkComboBox(
            card_nc,
            height=32,
            values=[
                "Ajuste de Precio / Descuento Comercial",
                "Merma de Peso en Recepción",
                "Diferencia en Facturación / Guía",
                "Bonificación Operativa",
                "Otro Concepto",
            ],
        )
        self.cmb_motivo_nc.pack(fill="x", padx=12, pady=2)

        ctk.CTkLabel(card_nc, text="Documento o Guía de Referencia", font=ctk.CTkFont(size=12, weight="bold")).pack(
            anchor="w", padx=12, pady=(8, 2)
        )
        self.ent_ref_nc = ctk.CTkEntry(card_nc, height=32, placeholder_text="Ej. OD-0005 o Guía Física")
        self.ent_ref_nc.pack(fill="x", padx=12, pady=2)

        ctk.CTkLabel(card_nc, text="Observaciones", font=ctk.CTkFont(size=12, weight="bold")).pack(
            anchor="w", padx=12, pady=(8, 2)
        )
        self.ent_obs_nc = ctk.CTkEntry(card_nc, height=32, placeholder_text="Detalle adicional de la nota")
        self.ent_obs_nc.pack(fill="x", padx=12, pady=2)

        self.chk_afectar_saldo_nc = ctk.CTkCheckBox(
            card_nc,
            text="Descontar monto del saldo del Centro",
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color=COLOR_PRIMARIO,
            hover_color=COLOR_PRIMARIO_HOVER,
        )
        self.chk_afectar_saldo_nc.pack(anchor="w", padx=12, pady=(12, 12))
        self.chk_afectar_saldo_nc.select()

        self.carrito_nc = PanelCarritoProductos(
            self.tab_credito, "💳 RUBROS A ACREDITAR AL CENTRO DE DISTRIBUCIÓN", self.obtener_contexto_moneda
        )
        self.carrito_nc.fijar_escala("Al Mayor")
        self.carrito_nc.pack(side="right", fill="both", expand=True, pady=4)

    def _procesar_nota_credito(self, exportar_excel: bool) -> None:
        try:
            if not self.carrito_nc.asegurar_lineas_listas():
                messagebox.showwarning(
                    "Nota Vacía",
                    "Escribe los kilos en 'Kilos (KG)' y presiona '➕ Agregar' en la lista derecha antes de emitir.",
                )
                return
            cd = self.mapa_centros.get(self.cmb_cd_credito.get())
            if not cd:
                messagebox.showwarning("Falta Centro", "Selecciona un Centro de Distribución.")
                return

            _, tot_ref, _ = self.carrito_nc.calcular_totales()
            afectar = 1 if self.chk_afectar_saldo_nc.get() == 1 else 0
            if afectar:
                ejecutar_sql(
                    "UPDATE centros_distribucion SET saldo_pendiente_ref = saldo_pendiente_ref - ? WHERE id = ?",
                    (tot_ref, cd["id"]),
                )

            doc_id = self._guardar_documento_en_db(
                tipo_doc="NOTA_CREDITO",
                subtipo="NOTA DE CREDITO - CENTRO DE DISTRIBUCION",
                origen=SEDE_PLANTA_PRINCIPAL,
                destino=f"{cd['codigo']} - {cd['nombre']}",
                doc_identidad=cd["codigo"],
                referencia=self.ent_ref_nc.get().strip(),
                motivo=self.cmb_motivo_nc.get(),
                afecto_saldo=afectar,
                afecto_inventario=0,
                observaciones=self.ent_obs_nc.get().strip(),
                carrito=self.carrito_nc,
            )
            ruta = exportar_documento_excel(doc_id) if exportar_excel else exportar_documento_pdf(doc_id)
            self.carrito_nc.limpiar_carrito()
            self.ent_ref_nc.delete(0, "end")
            self.ent_obs_nc.delete(0, "end")
            self.refrescar_datos_maestros()
            messagebox.showinfo("Nota de Crédito Creada", f"Documento generado correctamente:\n\n{ruta}")
            abrir_archivo(ruta)
        except Exception as error:
            messagebox.showerror("Error al generar Nota de Crédito", str(error))

    # ==========================================================================
    # MÓDULO 4: DÉBITOS Y MOVIMIENTOS (CON PRESELECTOR DE CHOFER Y ROTACIÓN)
    # ==========================================================================
    def _armar_submodulo_debito_y_movimientos(self) -> None:
        izq = ctk.CTkFrame(
            self.tab_debito_mov, width=390, fg_color=COLOR_TARJETA, border_width=1, border_color=COLOR_BORDE, corner_radius=12
        )
        izq.pack(side="left", fill="y", padx=(0, 10), pady=4)
        izq.pack_propagate(False)

        marco_btn = ctk.CTkFrame(izq, fg_color="transparent")
        marco_btn.pack(fill="x", side="bottom", padx=14, pady=10)

        ctk.CTkButton(
            marco_btn,
            text="📄 EMITIR GUÍA Y ROTAR INVENTARIO (PDF)",
            height=38,
            fg_color=COLOR_ACENTO,
            hover_color=COLOR_ACENTO_HOVER,
            font=ctk.CTkFont(size=12, weight="bold"),
            command=lambda: self._procesar_nd_o_movimiento(False),
        ).pack(fill="x", pady=(0, 5))

        ctk.CTkButton(
            marco_btn,
            text="📊 Exportar Comprobante en Excel",
            height=34,
            fg_color=COLOR_PRIMARIO,
            hover_color=COLOR_PRIMARIO_HOVER,
            font=ctk.CTkFont(weight="bold"),
            command=lambda: self._procesar_nd_o_movimiento(True),
        ).pack(fill="x")

        ctk.CTkLabel(
            izq, text="🔄 DÉBITOS Y ROTACIÓN ENTRE SEDES", font=ctk.CTkFont(size=14, weight="bold"), text_color=COLOR_TEXTO_TITULO
        ).pack(anchor="w", padx=16, pady=(8, 4))

        card_op = ctk.CTkFrame(izq, fg_color=COLOR_SUBTARJETA, corner_radius=10)
        card_op.pack(fill="x", padx=14, pady=3)

        ctk.CTkLabel(card_op, text="Modalidad de Operación", font=ctk.CTkFont(size=11, weight="bold")).pack(
            anchor="w", padx=12, pady=(5, 1)
        )
        self.cmb_modo_nd_mov = ctk.CTkComboBox(
            card_op,
            height=27,
            values=[
                "Movimiento entre Centros de Distribución",
                "Nota de Débito (Devolución NO Vendida)",
            ],
            command=self._al_cambiar_modo_nd_mov,
        )
        self.cmb_modo_nd_mov.pack(fill="x", padx=12, pady=1)

        ctk.CTkLabel(card_op, text="Sede / Centro Origen (Resta Stock -)", font=ctk.CTkFont(size=11, weight="bold")).pack(
            anchor="w", padx=12, pady=(4, 1)
        )
        self.cmb_cd_origen = ctk.CTkComboBox(card_op, height=27)
        self.cmb_cd_origen.pack(fill="x", padx=12, pady=1)

        self.lbl_cd_destino = ctk.CTkLabel(
            card_op, text="Sede / Centro Destino (Suma Stock +)", font=ctk.CTkFont(size=11, weight="bold")
        )
        self.lbl_cd_destino.pack(anchor="w", padx=12, pady=(4, 1))
        self.cmb_cd_destino = ctk.CTkComboBox(card_op, height=27)
        self.cmb_cd_destino.pack(fill="x", padx=12, pady=1)

        ctk.CTkLabel(card_op, text="Motivo del Registro", font=ctk.CTkFont(size=11, weight="bold")).pack(
            anchor="w", padx=12, pady=(4, 1)
        )
        self.cmb_motivo_nd = ctk.CTkComboBox(
            card_op,
            height=27,
            values=[
                "Rotación de Inventario Frío entre Sedes",
                "Abastecimiento desde Planta a Centro",
                "Retorno de Mercancía NO Vendida",
                "Reasignación de Pedido",
                "Cargo por Devolución",
            ],
        )
        self.cmb_motivo_nd.pack(fill="x", padx=12, pady=(1, 6))

        card_tr = ctk.CTkFrame(izq, fg_color=COLOR_SUBTARJETA, corner_radius=10)
        card_tr.pack(fill="x", padx=14, pady=3)

        ctk.CTkLabel(
            card_tr, text="🚛 Datos de Transporte en Guía (Preseleccionar)", font=ctk.CTkFont(size=11, weight="bold")
        ).pack(anchor="w", padx=12, pady=(6, 2))

        self.cmb_chofer_pre = ctk.CTkComboBox(
            card_tr,
            height=28,
            values=[OPCION_CHOFER_MANUAL],
            command=self._al_preseleccionar_chofer,
        )
        self.cmb_chofer_pre.pack(fill="x", padx=12, pady=(0, 4))
        self.cmb_chofer_pre.set(OPCION_CHOFER_MANUAL)

        self.ent_trans_chofer = ctk.CTkEntry(card_tr, height=27, placeholder_text="Nombre y Apellido del Chofer")
        self.ent_trans_chofer.pack(fill="x", padx=12, pady=2)

        fila_tr = ctk.CTkFrame(card_tr, fg_color="transparent")
        fila_tr.pack(fill="x", padx=12, pady=2)
        fila_tr.grid_columnconfigure(0, weight=1)
        fila_tr.grid_columnconfigure(1, weight=1)

        self.ent_trans_cedula = ctk.CTkEntry(fila_tr, height=27, placeholder_text="Cédula Chofer")
        self.ent_trans_cedula.grid(row=0, column=0, sticky="ew", padx=(0, 3))
        self.ent_trans_placa = ctk.CTkEntry(fila_tr, height=27, placeholder_text="Placa Vehículo")
        self.ent_trans_placa.grid(row=0, column=1, sticky="ew", padx=(3, 0))

        self.ent_obs_nd = ctk.CTkEntry(card_tr, height=27, placeholder_text="Observaciones / Cadena de frío")
        self.ent_obs_nd.pack(fill="x", padx=12, pady=(2, 4))

        self.chk_afectar_saldo_nd = ctk.CTkCheckBox(
            card_tr,
            text="Afectar saldo contable del Centro (en ND)",
            font=ctk.CTkFont(size=11, weight="bold"),
            fg_color=COLOR_PRIMARIO,
            hover_color=COLOR_PRIMARIO_HOVER,
        )
        self.chk_afectar_saldo_nd.pack(anchor="w", padx=12, pady=(2, 6))
        self.chk_afectar_saldo_nd.select()

        self.carrito_nd = PanelCarritoProductos(
            self.tab_debito_mov,
            "🔄 MERCANCÍA EN ROTACIÓN / DEVOLUCIÓN (ACTUALIZA EL MÓDULO 5)",
            self.obtener_contexto_moneda,
        )
        self.carrito_nd.fijar_escala("Al Mayor")
        self.carrito_nd.pack(side="right", fill="both", expand=True, pady=4)

    def _al_preseleccionar_chofer(self, seleccion: str) -> None:
        for ent in (self.ent_trans_chofer, self.ent_trans_cedula, self.ent_trans_placa):
            ent.delete(0, "end")

        ch = self.mapa_choferes.get(seleccion)
        if not ch:
            return
        self.ent_trans_chofer.insert(0, ch["nombre"])
        self.ent_trans_cedula.insert(0, ch["cedula"])
        self.ent_trans_placa.insert(0, ch["placa"])

    def _al_cambiar_modo_nd_mov(self, modo_elegido: str) -> None:
        nombres_cd = list(self.mapa_centros.keys())
        todas_sedes = [SEDE_PLANTA_PRINCIPAL] + nombres_cd
        if "Devolución" in modo_elegido:
            self.lbl_cd_destino.configure(text="Sede que Recibe Devolución (Suma Stock +)")
            self.cmb_cd_origen.configure(values=nombres_cd if nombres_cd else [SEDE_PLANTA_PRINCIPAL])
            if nombres_cd:
                self.cmb_cd_origen.set(nombres_cd[0])
            self.cmb_cd_destino.configure(values=todas_sedes)
            self.cmb_cd_destino.set(SEDE_PLANTA_PRINCIPAL)
            self.cmb_motivo_nd.set("Retorno de Mercancía NO Vendida")
        else:
            self.lbl_cd_destino.configure(text="Centro de Distribución Destino (Suma Stock +)")
            self.cmb_cd_origen.configure(values=todas_sedes)
            if nombres_cd:
                self.cmb_cd_origen.set(nombres_cd[0])
            self.cmb_cd_destino.configure(values=todas_sedes)
            if len(nombres_cd) > 1:
                self.cmb_cd_destino.set(nombres_cd[1])
            elif nombres_cd:
                self.cmb_cd_destino.set(nombres_cd[0])
            self.cmb_motivo_nd.set("Rotación de Inventario Frío entre Sedes")

    def _procesar_nd_o_movimiento(self, exportar_excel: bool) -> None:
        try:
            if not self.carrito_nd.asegurar_lineas_listas():
                messagebox.showwarning(
                    "Lista de Productos Vacía",
                    "Escribe la cantidad en 'Kilos (KG)' y presiona '➕ Agregar' en la tabla derecha antes de emitir la guía.",
                )
                return

            origen = normalizar_nombre_sede(self.cmb_cd_origen.get())
            destino = normalizar_nombre_sede(self.cmb_cd_destino.get())
            if not origen or not destino or origen == destino:
                messagebox.showwarning(
                    "Revisa Origen y Destino",
                    "La Sede Origen y la Sede Destino deben ser diferentes para poder rotar el inventario.",
                )
                return

            es_devolucion = "Devolución" in self.cmb_modo_nd_mov.get()
            tipo_doc = "NOTA_DEBITO" if es_devolucion else "MOVIMIENTO"
            subtipo = (
                "NOTA DE DEBITO - DEVOLUCION MERCANCIA NO VENDIDA"
                if es_devolucion
                else "GUIA DE MOVIMIENTO ENTRE CENTROS DE DISTRIBUCION"
            )

            tot_kg, tot_ref, _ = self.carrito_nd.calcular_totales()
            afectar_saldo = 1 if (es_devolucion and self.chk_afectar_saldo_nd.get() == 1) else 0
            if afectar_saldo:
                cd_obj = self.mapa_centros.get(origen)
                if cd_obj:
                    ejecutar_sql(
                        "UPDATE centros_distribucion SET saldo_pendiente_ref = saldo_pendiente_ref + ? WHERE id = ?",
                        (tot_ref, cd_obj["id"]),
                    )

            placa_txt = self.ent_trans_placa.get().strip().upper()
            doc_id = self._guardar_documento_en_db(
                tipo_doc=tipo_doc,
                subtipo=subtipo,
                origen=origen,
                destino=destino,
                doc_identidad="",
                referencia=placa_txt,
                motivo=self.cmb_motivo_nd.get(),
                chofer=self.ent_trans_chofer.get().strip().upper(),
                cedula_chofer=self.ent_trans_cedula.get().strip().upper(),
                placa_vehiculo=placa_txt,
                afecto_saldo=afectar_saldo,
                afecto_inventario=1,
                observaciones=self.ent_obs_nd.get().strip(),
                carrito=self.carrito_nd,
            )
            ruta = exportar_documento_excel(doc_id) if exportar_excel else exportar_documento_pdf(doc_id)
            self.carrito_nd.limpiar_carrito()
            self.ent_obs_nd.delete(0, "end")
            self.refrescar_datos_maestros()
            messagebox.showinfo(
                "Guía Emitida e Inventario Rotado",
                f"Se transfirieron {tot_kg:,.2f} kg exitosamente:\n"
                f"• Restado en: {origen}\n"
                f"• Sumado en: {destino}\n\n"
                f"El nuevo stock ya está reflejado en el Módulo 5.\nArchivo creado en:\n{ruta}",
            )
            abrir_archivo(ruta)
        except Exception as error:
            messagebox.showerror("Error al emitir Guía / Comprobante", str(error))

    # ==========================================================================
    # MÓDULO 5: INVENTARIO POR SEDE
    # ==========================================================================
    def _armar_modulo_inventario(self) -> None:
        izq = ctk.CTkFrame(
            self.tab_inventario, width=395, fg_color=COLOR_TARJETA, border_width=1, border_color=COLOR_BORDE, corner_radius=12
        )
        izq.pack(side="left", fill="y", padx=(0, 10), pady=4)
        izq.pack_propagate(False)

        marco_rep = ctk.CTkFrame(izq, fg_color="transparent")
        marco_rep.pack(fill="x", side="bottom", padx=14, pady=12)

        ctk.CTkButton(
            marco_rep,
            text="📄 EXPORTAR INVENTARIO EN PDF",
            height=40,
            fg_color=COLOR_ACENTO,
            hover_color=COLOR_ACENTO_HOVER,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=lambda: self._exportar_reporte_inventario(False),
        ).pack(fill="x", pady=(0, 6))

        ctk.CTkButton(
            marco_rep,
            text="📊 Exportar Inventario en Excel (.xlsx)",
            height=36,
            fg_color=COLOR_PRIMARIO,
            hover_color=COLOR_PRIMARIO_HOVER,
            font=ctk.CTkFont(weight="bold"),
            command=lambda: self._exportar_reporte_inventario(True),
        ).pack(fill="x")

        ctk.CTkLabel(
            izq, text="📊 CONTROL DE INVENTARIO POR SEDE", font=ctk.CTkFont(size=14, weight="bold"), text_color=COLOR_TEXTO_TITULO
        ).pack(anchor="w", padx=16, pady=(14, 8))

        card_f = ctk.CTkFrame(izq, fg_color=COLOR_SUBTARJETA, corner_radius=10)
        card_f.pack(fill="x", padx=14, pady=4)

        ctk.CTkLabel(card_f, text="Filtrar Inventario por Sede", font=ctk.CTkFont(size=12, weight="bold")).pack(
            anchor="w", padx=12, pady=(8, 2)
        )
        self.cmb_filtro_sede_inv = ctk.CTkComboBox(
            card_f, height=32, values=["Todas las Sedes"], command=lambda _: self.refrescar_tabla_inventario()
        )
        self.cmb_filtro_sede_inv.pack(fill="x", padx=12, pady=2)
        self.cmb_filtro_sede_inv.set("Todas las Sedes")

        ctk.CTkLabel(card_f, text="Mostrar Productos", font=ctk.CTkFont(size=12, weight="bold")).pack(
            anchor="w", padx=12, pady=(6, 2)
        )
        self.cmb_filtro_estado_inv = ctk.CTkComboBox(
            card_f,
            height=30,
            values=[
                "Con Movimiento o Stock Registrado",
                "Catálogo Completo (22 Productos en la Sede)",
                "Solo Productos con Stock Positivo (Mas de 0 KG)",
                "Solo Alertas (Bajo Stock o Negativo)",
            ],
            command=lambda _: self.refrescar_tabla_inventario(),
        )
        self.cmb_filtro_estado_inv.pack(fill="x", padx=12, pady=(2, 10))
        self.cmb_filtro_estado_inv.set("Con Movimiento o Stock Registrado")

        card_aj = ctk.CTkFrame(izq, fg_color=COLOR_SUBTARJETA, corner_radius=10)
        card_aj.pack(fill="x", padx=14, pady=6)

        ctk.CTkLabel(card_aj, text="⚖️ Carga Inicial / Ajuste Manual de Stock", font=ctk.CTkFont(size=12, weight="bold")).pack(
            anchor="w", padx=12, pady=(8, 4)
        )

        ctk.CTkLabel(card_aj, text="Sede o Centro de Distribución", font=ctk.CTkFont(size=11, weight="bold")).pack(
            anchor="w", padx=12, pady=(2, 0)
        )
        self.cmb_inv_sede_aj = ctk.CTkComboBox(card_aj, height=30, values=[SEDE_PLANTA_PRINCIPAL])
        self.cmb_inv_sede_aj.pack(fill="x", padx=12, pady=2)

        ctk.CTkLabel(card_aj, text="Producto / Corte", font=ctk.CTkFont(size=11, weight="bold")).pack(
            anchor="w", padx=12, pady=(4, 0)
        )
        self.cmb_inv_prod_aj = ctk.CTkComboBox(card_aj, height=30)
        self.cmb_inv_prod_aj.pack(fill="x", padx=12, pady=2)

        fila_aj = ctk.CTkFrame(card_aj, fg_color="transparent")
        fila_aj.pack(fill="x", padx=12, pady=4)
        fila_aj.grid_columnconfigure(0, weight=2)
        fila_aj.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(fila_aj, text="Tipo de Operación", font=ctk.CTkFont(size=11, weight="bold")).grid(
            row=0, column=0, sticky="w"
        )
        ctk.CTkLabel(fila_aj, text="Kilos (KG)", font=ctk.CTkFont(size=11, weight="bold")).grid(
            row=0, column=1, sticky="w", padx=(6, 0)
        )

        self.cmb_inv_tipo_aj = ctk.CTkComboBox(
            fila_aj,
            height=30,
            values=[
                "Sumar Entrada (+ KG)",
                "Restar Salida (- KG)",
                "Fijar Stock Exacto (= KG)",
            ],
        )
        self.cmb_inv_tipo_aj.grid(row=1, column=0, sticky="ew", padx=(0, 3), pady=2)
        self.ent_inv_kilos_aj = ctk.CTkEntry(fila_aj, height=30, placeholder_text="Ej. 100.00")
        self.ent_inv_kilos_aj.grid(row=1, column=1, sticky="ew", padx=(3, 0), pady=2)

        b_aj = ctk.CTkFrame(card_aj, fg_color="transparent")
        b_aj.pack(fill="x", padx=12, pady=(6, 10))
        for i in range(3):
            b_aj.grid_columnconfigure(i, weight=1)

        ctk.CTkButton(
            b_aj,
            text="💾 Aplicar",
            height=32,
            fg_color=COLOR_PRIMARIO,
            hover_color=COLOR_PRIMARIO_HOVER,
            font=ctk.CTkFont(weight="bold"),
            command=self._aplicar_ajuste_manual_inventario,
        ).grid(row=0, column=0, sticky="ew", padx=(0, 4))

        ctk.CTkButton(
            b_aj,
            text="🧹 En 0 KG",
            height=32,
            fg_color=COLOR_NEUTRO,
            hover_color=COLOR_NEUTRO_HOVER,
            font=ctk.CTkFont(weight="bold"),
            command=self._poner_stock_en_cero,
        ).grid(row=0, column=1, sticky="ew", padx=4)

        ctk.CTkButton(
            b_aj,
            text="🗑️ Eliminar",
            height=32,
            fg_color=COLOR_PELIGRO,
            hover_color=COLOR_PELIGRO_HOVER,
            font=ctk.CTkFont(weight="bold"),
            command=self._eliminar_registro_inventario,
        ).grid(row=0, column=2, sticky="ew", padx=(4, 0))

        der = ctk.CTkFrame(
            self.tab_inventario, fg_color=COLOR_TARJETA, border_width=1, border_color=COLOR_BORDE, corner_radius=12
        )
        der.pack(side="right", fill="both", expand=True, pady=4)

        cab_der = ctk.CTkFrame(der, fg_color="transparent")
        cab_der.pack(fill="x", padx=16, pady=(12, 6))

        ctk.CTkLabel(
            cab_der,
            text="🏭 EXISTENCIAS EN VIVO POR SEDE Y CENTRO DE DISTRIBUCIÓN",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=COLOR_TEXTO_TITULO,
        ).pack(side="left")

        ctk.CTkButton(
            cab_der,
            text="🔄 Actualizar Vista",
            width=135,
            height=30,
            fg_color=COLOR_NEUTRO,
            hover_color=COLOR_NEUTRO_HOVER,
            font=ctk.CTkFont(weight="bold"),
            command=self.refrescar_tabla_inventario,
        ).pack(side="right", padx=(6, 0))

        ctk.CTkButton(
            cab_der,
            text="🗑️ Eliminar Fila Seleccionada",
            width=195,
            height=30,
            fg_color=COLOR_PELIGRO,
            hover_color=COLOR_PELIGRO_HOVER,
            font=ctk.CTkFont(weight="bold"),
            command=self._eliminar_registro_inventario,
        ).pack(side="right")

        marco_tv = ctk.CTkFrame(der, fg_color="transparent")
        marco_tv.pack(fill="both", expand=True, padx=16, pady=4)

        cols_inv = [
            ("nro", "#", 40),
            ("sede", "Sede / Centro de Distribución", 230),
            ("prod", "Producto / Corte", 210),
            ("kg", "Stock Disponible (KG)", 135),
            ("est", "Estado", 125),
            ("val_ref", "Valor Mayor REF", 120),
            ("val_bs", "Valor Mayor Bs", 130),
            ("act", "Último Movimiento", 130),
        ]
        self.tv_inventario = ttk.Treeview(
            marco_tv, columns=[c[0] for c in cols_inv], show="headings", selectmode="browse"
        )
        sc_inv = ttk.Scrollbar(marco_tv, orient="vertical", command=self.tv_inventario.yview)
        self.tv_inventario.configure(yscrollcommand=sc_inv.set)

        for cid, texto, ancho in cols_inv:
            self.tv_inventario.heading(cid, text=texto)
            self.tv_inventario.column(cid, width=ancho, anchor="w" if cid in ("sede", "prod") else "center")

        sc_inv.pack(side="right", fill="y")
        self.tv_inventario.pack(side="left", fill="both", expand=True)
        self.tv_inventario.bind("<<TreeviewSelect>>", self._al_seleccionar_fila_inventario)

        pie_inv = ctk.CTkFrame(der, fg_color=COLOR_SUBTARJETA, corner_radius=10)
        pie_inv.pack(fill="x", padx=16, pady=(6, 14))

        self.lbl_resumen_inventario = ctk.CTkLabel(
            pie_inv,
            text="STOCK TOTAL EN SEDE: 0.00 KG   |   VALOR ESTIMADO (AL MAYOR): $ 0.00   |   BOLÍVARES: Bs 0.00",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=COLOR_ACENTO,
        )
        self.lbl_resumen_inventario.pack(side="left", padx=14, pady=10)

    def _al_seleccionar_fila_inventario(self, _e=None) -> None:
        sel = self.tv_inventario.selection()
        if not sel:
            return
        valores = self.tv_inventario.item(sel[0], "values")
        if len(valores) >= 3:
            self.cmb_inv_sede_aj.set(valores[1])
            self.cmb_inv_prod_aj.set(valores[2])

    def _aplicar_ajuste_manual_inventario(self) -> None:
        sede = self.cmb_inv_sede_aj.get().strip()
        prod = self.cmb_inv_prod_aj.get().strip()
        if not sede or not prod:
            messagebox.showwarning("Faltan Datos", "Selecciona la Sede y el Producto a ajustar.")
            return
        try:
            kilos = float(self.ent_inv_kilos_aj.get().replace(",", "."))
        except ValueError:
            messagebox.showwarning("Cantidad Inválida", "Escribe una cantidad numérica válida en Kilos (ej. 120.50).")
            return

        modo_op = self.cmb_inv_tipo_aj.get()
        if "Fijar" in modo_op:
            fijar_stock_exacto_sede(sede, prod, kilos)
            accion_txt = f"Se fijó el stock exacto en {kilos:,.2f} kg"
        elif "Restar" in modo_op:
            with conectar_db() as conn:
                ajustar_stock_sede(conn, sede, prod, -abs(kilos))
                conn.commit()
            accion_txt = f"Se restaron {abs(kilos):,.2f} kg"
        else:
            with conectar_db() as conn:
                ajustar_stock_sede(conn, sede, prod, +abs(kilos))
                conn.commit()
            accion_txt = f"Se sumaron {abs(kilos):,.2f} kg"

        self.ent_inv_kilos_aj.delete(0, "end")
        self.refrescar_tabla_inventario()
        messagebox.showinfo("Inventario Actualizado", f"{accion_txt} de '{prod}' en:\n{sede}")

    def _poner_stock_en_cero(self) -> None:
        sel = self.tv_inventario.selection()
        if sel:
            valores = self.tv_inventario.item(sel[0], "values")
            sede, prod = valores[1], valores[2]
        else:
            sede = self.cmb_inv_sede_aj.get().strip()
            prod = self.cmb_inv_prod_aj.get().strip()

        if not sede or not prod:
            messagebox.showwarning("Selecciona un Registro", "Haz clic sobre una fila de la tabla de inventario primero.")
            return

        if messagebox.askyesno(
            "Confirmar Stock en Cero",
            f"¿Deseas dejar en 0.00 KG el stock de '{prod}' en '{sede}'?",
        ):
            fijar_stock_exacto_sede(sede, prod, 0.0)
            self.refrescar_tabla_inventario()
            messagebox.showinfo("Stock en Cero", f"El stock de '{prod}' en '{sede}' quedó en 0.00 KG.")

    def _eliminar_registro_inventario(self) -> None:
        sel = self.tv_inventario.selection()
        if sel:
            valores = self.tv_inventario.item(sel[0], "values")
            sede, prod = valores[1], valores[2]
        else:
            sede = self.cmb_inv_sede_aj.get().strip()
            prod = self.cmb_inv_prod_aj.get().strip()

        if not sede or not prod:
            messagebox.showwarning(
                "Selecciona una Fila",
                "Haz clic sobre el producto que deseas quitar en la tabla de inventario y luego pulsa '🗑️ Eliminar'.",
            )
            return

        if messagebox.askyesno(
            "Eliminar Registro de Inventario",
            f"¿Estás seguro de eliminar por completo el registro de:\n\n"
            f"• Producto: {prod}\n"
            f"• Sede: {sede}\n\n"
            "Desaparecerá de la tabla de existencias inmediatamente.",
        ):
            eliminar_item_inventario_sede(sede, prod)
            self.refrescar_tabla_inventario()
            messagebox.showinfo("Registro Eliminado", f"Se eliminó '{prod}' del inventario de '{sede}'.")

    def _exportar_reporte_inventario(self, en_excel: bool) -> None:
        try:
            if not self.cache_inventario_vista:
                messagebox.showwarning("Sin Datos en Vista", "No hay registros de inventario en la vista actual para exportar.")
                return
            sede_f = self.cmb_filtro_sede_inv.get()
            modo, t_usd, t_eur = self.obtener_contexto_moneda()
            ruta = (
                exportar_inventario_excel(sede_f, self.cache_inventario_vista, modo, t_usd, t_eur)
                if en_excel
                else exportar_inventario_pdf(sede_f, self.cache_inventario_vista, modo, t_usd, t_eur)
            )
            messagebox.showinfo("Reporte de Inventario Listo", f"Archivo generado correctamente:\n\n{ruta}")
            abrir_archivo(ruta)
        except Exception as error:
            messagebox.showerror("Error al exportar inventario", str(error))

    # ==========================================================================
    # MÓDULO 6: BASES DE DATOS TRIPLES (CLIENTES | CENTROS | CHOFERES)
    # ==========================================================================
    def _armar_modulo_bases_datos(self) -> None:
        self.tab_bases_datos.grid_columnconfigure(0, weight=1, uniform="bd_triple")
        self.tab_bases_datos.grid_columnconfigure(1, weight=1, uniform="bd_triple")
        self.tab_bases_datos.grid_columnconfigure(2, weight=1, uniform="bd_triple")
        self.tab_bases_datos.grid_rowconfigure(0, weight=1)

        # Columna 1: Clientes
        col_cli = ctk.CTkFrame(
            self.tab_bases_datos, fg_color=COLOR_TARJETA, border_width=1, border_color=COLOR_BORDE, corner_radius=12
        )
        col_cli.grid(row=0, column=0, sticky="nsew", padx=(0, 5), pady=4)

        ctk.CTkLabel(
            col_cli, text="👤 1. BASE DE DATOS: CLIENTES", font=ctk.CTkFont(size=13, weight="bold"), text_color=COLOR_TEXTO_TITULO
        ).pack(anchor="w", padx=12, pady=(10, 4))

        fc = ctk.CTkFrame(col_cli, fg_color=COLOR_SUBTARJETA, corner_radius=10)
        fc.pack(fill="x", padx=10, pady=4)
        fc.grid_columnconfigure(0, weight=1)
        fc.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(fc, text="Nombre o Razón Social", font=ctk.CTkFont(size=11, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=6, pady=(6, 0)
        )
        ctk.CTkLabel(fc, text="Cédula o RIF", font=ctk.CTkFont(size=11, weight="bold")).grid(
            row=0, column=1, sticky="w", padx=6, pady=(6, 0)
        )

        self.ent_cli_nombre = ctk.CTkEntry(fc, height=28, placeholder_text="Ej. Inversiones Los Llanos")
        self.ent_cli_nombre.grid(row=1, column=0, sticky="ew", padx=6, pady=(2, 4))
        self.ent_cli_rif = ctk.CTkEntry(fc, height=28, placeholder_text="Ej. J-12345678-9")
        self.ent_cli_rif.grid(row=1, column=1, sticky="ew", padx=6, pady=(2, 4))

        ctk.CTkLabel(fc, text="Escala Preferida", font=ctk.CTkFont(size=11, weight="bold")).grid(
            row=2, column=0, sticky="w", padx=6
        )
        ctk.CTkLabel(fc, text="Teléfono", font=ctk.CTkFont(size=11, weight="bold")).grid(
            row=2, column=1, sticky="w", padx=6
        )

        self.cmb_cli_esc = ctk.CTkComboBox(fc, height=28, values=list(ESCALAS_PRECIO.keys()))
        self.cmb_cli_esc.grid(row=3, column=0, sticky="ew", padx=6, pady=(2, 4))
        self.ent_cli_tel = ctk.CTkEntry(fc, height=28, placeholder_text="0414-0000000")
        self.ent_cli_tel.grid(row=3, column=1, sticky="ew", padx=6, pady=(2, 4))

        ctk.CTkLabel(fc, text="Dirección del Cliente", font=ctk.CTkFont(size=11, weight="bold")).grid(
            row=4, column=0, columnspan=2, sticky="w", padx=6
        )
        self.ent_cli_dir = ctk.CTkEntry(fc, height=28, placeholder_text="Ciudad / Zona / Local")
        self.ent_cli_dir.grid(row=5, column=0, columnspan=2, sticky="ew", padx=6, pady=(2, 8))

        bc = ctk.CTkFrame(col_cli, fg_color="transparent")
        bc.pack(fill="x", padx=10, pady=4)
        for i in range(4):
            bc.grid_columnconfigure(i, weight=1)

        ctk.CTkButton(bc, text="➕ Guardar", height=30, fg_color=COLOR_PRIMARIO, hover_color=COLOR_PRIMARIO_HOVER, font=ctk.CTkFont(size=11, weight="bold"), command=self._guardar_cliente).grid(row=0, column=0, sticky="ew", padx=(0, 3))
        ctk.CTkButton(bc, text="✏️ Editar", height=30, fg_color=COLOR_ACENTO, hover_color=COLOR_ACENTO_HOVER, font=ctk.CTkFont(size=11, weight="bold"), command=self._actualizar_cliente).grid(row=0, column=1, sticky="ew", padx=3)
        ctk.CTkButton(bc, text="🧹 Limpiar", height=30, fg_color=COLOR_NEUTRO, hover_color=COLOR_NEUTRO_HOVER, font=ctk.CTkFont(size=11), command=self._limpiar_form_cliente).grid(row=0, column=2, sticky="ew", padx=3)
        ctk.CTkButton(bc, text="🗑️ Borrar", height=30, fg_color=COLOR_PELIGRO, hover_color=COLOR_PELIGRO_HOVER, font=ctk.CTkFont(size=11, weight="bold"), command=self._eliminar_cliente).grid(row=0, column=3, sticky="ew", padx=(3, 0))

        m_tvc = ctk.CTkFrame(col_cli, fg_color="transparent")
        m_tvc.pack(fill="both", expand=True, padx=10, pady=(4, 10))
        cols_c = [
            ("id", "ID", 34),
            ("nom", "Cliente", 145),
            ("rif", "RIF / CI", 82),
            ("esc", "Escala", 85),
            ("tel", "Teléfono", 82),
        ]
        self.tv_clientes = ttk.Treeview(m_tvc, columns=[c[0] for c in cols_c], show="headings", selectmode="browse")
        sc_c = ttk.Scrollbar(m_tvc, orient="vertical", command=self.tv_clientes.yview)
        self.tv_clientes.configure(yscrollcommand=sc_c.set)
        for cid, tit, anc in cols_c:
            self.tv_clientes.heading(cid, text=tit)
            self.tv_clientes.column(cid, width=anc, anchor="w" if cid == "nom" else "center")
        sc_c.pack(side="right", fill="y")
        self.tv_clientes.pack(side="left", fill="both", expand=True)
        self.tv_clientes.bind("<<TreeviewSelect>>", self._seleccionar_cliente_tabla)

        # Columna 2: Centros de Distribución
        col_cd = ctk.CTkFrame(
            self.tab_bases_datos, fg_color=COLOR_TARJETA, border_width=1, border_color=COLOR_BORDE, corner_radius=12
        )
        col_cd.grid(row=0, column=1, sticky="nsew", padx=5, pady=4)

        ctk.CTkLabel(
            col_cd, text="🏢 2. CENTROS DE DISTRIBUCIÓN", font=ctk.CTkFont(size=13, weight="bold"), text_color=COLOR_TEXTO_TITULO
        ).pack(anchor="w", padx=12, pady=(10, 4))

        fd = ctk.CTkFrame(col_cd, fg_color=COLOR_SUBTARJETA, corner_radius=10)
        fd.pack(fill="x", padx=10, pady=4)
        fd.grid_columnconfigure(0, weight=1)
        fd.grid_columnconfigure(1, weight=2)

        ctk.CTkLabel(fd, text="Código Sede", font=ctk.CTkFont(size=11, weight="bold")).grid(row=0, column=0, sticky="w", padx=6, pady=(6, 0))
        ctk.CTkLabel(fd, text="Nombre del Centro", font=ctk.CTkFont(size=11, weight="bold")).grid(row=0, column=1, sticky="w", padx=6, pady=(6, 0))

        self.ent_cd_cod = ctk.CTkEntry(fd, height=28, placeholder_text="Ej. CD-03")
        self.ent_cd_cod.grid(row=1, column=0, sticky="ew", padx=6, pady=(2, 4))
        self.ent_cd_nom = ctk.CTkEntry(fd, height=28, placeholder_text="Nombre de la Sucursal")
        self.ent_cd_nom.grid(row=1, column=1, sticky="ew", padx=6, pady=(2, 4))

        ctk.CTkLabel(fd, text="Encargado", font=ctk.CTkFont(size=11, weight="bold")).grid(row=2, column=0, sticky="w", padx=6)
        ctk.CTkLabel(fd, text="Teléfono", font=ctk.CTkFont(size=11, weight="bold")).grid(row=2, column=1, sticky="w", padx=6)

        self.ent_cd_enc = ctk.CTkEntry(fd, height=28, placeholder_text="Responsable")
        self.ent_cd_enc.grid(row=3, column=0, sticky="ew", padx=6, pady=(2, 4))
        self.ent_cd_tel = ctk.CTkEntry(fd, height=28, placeholder_text="0414-0000000")
        self.ent_cd_tel.grid(row=3, column=1, sticky="ew", padx=6, pady=(2, 4))

        ctk.CTkLabel(fd, text="Ubicación / Zona", font=ctk.CTkFont(size=11, weight="bold")).grid(
            row=4, column=0, columnspan=2, sticky="w", padx=6
        )
        self.ent_cd_ubi = ctk.CTkEntry(fd, height=28, placeholder_text="Dirección o Zona Comercial")
        self.ent_cd_ubi.grid(row=5, column=0, columnspan=2, sticky="ew", padx=6, pady=(2, 8))

        bd = ctk.CTkFrame(col_cd, fg_color="transparent")
        bd.pack(fill="x", padx=10, pady=4)
        for i in range(4):
            bd.grid_columnconfigure(i, weight=1)

        ctk.CTkButton(bd, text="➕ Guardar", height=30, fg_color=COLOR_PRIMARIO, hover_color=COLOR_PRIMARIO_HOVER, font=ctk.CTkFont(size=11, weight="bold"), command=self._guardar_centro).grid(row=0, column=0, sticky="ew", padx=(0, 3))
        ctk.CTkButton(bd, text="✏️ Editar", height=30, fg_color=COLOR_ACENTO, hover_color=COLOR_ACENTO_HOVER, font=ctk.CTkFont(size=11, weight="bold"), command=self._actualizar_centro).grid(row=0, column=1, sticky="ew", padx=3)
        ctk.CTkButton(bd, text="🧹 Limpiar", height=30, fg_color=COLOR_NEUTRO, hover_color=COLOR_NEUTRO_HOVER, font=ctk.CTkFont(size=11), command=self._limpiar_form_centro).grid(row=0, column=2, sticky="ew", padx=3)
        ctk.CTkButton(bd, text="🗑️ Borrar", height=30, fg_color=COLOR_PELIGRO, hover_color=COLOR_PELIGRO_HOVER, font=ctk.CTkFont(size=11, weight="bold"), command=self._eliminar_centro).grid(row=0, column=3, sticky="ew", padx=(3, 0))

        m_tvd = ctk.CTkFrame(col_cd, fg_color="transparent")
        m_tvd.pack(fill="both", expand=True, padx=10, pady=(4, 10))
        cols_d = [
            ("id", "ID", 34),
            ("cod", "Cód.", 58),
            ("nom", "Centro Dist.", 145),
            ("enc", "Encargado", 95),
            ("saldo", "Saldo REF", 82),
        ]
        self.tv_centros = ttk.Treeview(m_tvd, columns=[c[0] for c in cols_d], show="headings", selectmode="browse")
        sc_d = ttk.Scrollbar(m_tvd, orient="vertical", command=self.tv_centros.yview)
        self.tv_centros.configure(yscrollcommand=sc_d.set)
        for cid, tit, anc in cols_d:
            self.tv_centros.heading(cid, text=tit)
            self.tv_centros.column(cid, width=anc, anchor="w" if cid == "nom" else "center")
        sc_d.pack(side="right", fill="y")
        self.tv_centros.pack(side="left", fill="both", expand=True)
        self.tv_centros.bind("<<TreeviewSelect>>", self._seleccionar_centro_tabla)

        # Columna 3: Choferes y Vehículos
        col_ch = ctk.CTkFrame(
            self.tab_bases_datos, fg_color=COLOR_TARJETA, border_width=1, border_color=COLOR_BORDE, corner_radius=12
        )
        col_ch.grid(row=0, column=2, sticky="nsew", padx=(5, 0), pady=4)

        ctk.CTkLabel(
            col_ch, text="🚛 3. BASE DE DATOS: CHOFERES", font=ctk.CTkFont(size=13, weight="bold"), text_color=COLOR_TEXTO_TITULO
        ).pack(anchor="w", padx=12, pady=(10, 4))

        fch = ctk.CTkFrame(col_ch, fg_color=COLOR_SUBTARJETA, corner_radius=10)
        fch.pack(fill="x", padx=10, pady=4)
        fch.grid_columnconfigure(0, weight=1)
        fch.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(fch, text="Nombre y Apellido del Chofer", font=ctk.CTkFont(size=11, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=6, pady=(6, 0)
        )
        ctk.CTkLabel(fch, text="Cédula del Chofer", font=ctk.CTkFont(size=11, weight="bold")).grid(
            row=0, column=1, sticky="w", padx=6, pady=(6, 0)
        )

        self.ent_ch_nom = ctk.CTkEntry(fch, height=28, placeholder_text="Ej. Carlos Pérez")
        self.ent_ch_nom.grid(row=1, column=0, sticky="ew", padx=6, pady=(2, 4))
        self.ent_ch_ced = ctk.CTkEntry(fch, height=28, placeholder_text="Ej. V-15888999")
        self.ent_ch_ced.grid(row=1, column=1, sticky="ew", padx=6, pady=(2, 4))

        ctk.CTkLabel(fch, text="Placa del Vehículo", font=ctk.CTkFont(size=11, weight="bold")).grid(
            row=2, column=0, sticky="w", padx=6
        )
        ctk.CTkLabel(fch, text="Teléfono", font=ctk.CTkFont(size=11, weight="bold")).grid(
            row=2, column=1, sticky="w", padx=6
        )

        self.ent_ch_pla = ctk.CTkEntry(fch, height=28, placeholder_text="Ej. A25BC99")
        self.ent_ch_pla.grid(row=3, column=0, sticky="ew", padx=6, pady=(2, 4))
        self.ent_ch_tel = ctk.CTkEntry(fch, height=28, placeholder_text="0412-0000000")
        self.ent_ch_tel.grid(row=3, column=1, sticky="ew", padx=6, pady=(2, 4))

        ctk.CTkLabel(fch, text="Descripción del Vehículo / Camión", font=ctk.CTkFont(size=11, weight="bold")).grid(
            row=4, column=0, columnspan=2, sticky="w", padx=6
        )
        self.ent_ch_veh = ctk.CTkEntry(fch, height=28, placeholder_text="Ej. NPR Termoking Cava Fría")
        self.ent_ch_veh.grid(row=5, column=0, columnspan=2, sticky="ew", padx=6, pady=(2, 8))

        bch = ctk.CTkFrame(col_ch, fg_color="transparent")
        bch.pack(fill="x", padx=10, pady=4)
        for i in range(4):
            bch.grid_columnconfigure(i, weight=1)

        ctk.CTkButton(bch, text="➕ Guardar", height=30, fg_color=COLOR_PRIMARIO, hover_color=COLOR_PRIMARIO_HOVER, font=ctk.CTkFont(size=11, weight="bold"), command=self._guardar_chofer).grid(row=0, column=0, sticky="ew", padx=(0, 3))
        ctk.CTkButton(bch, text="✏️ Editar", height=30, fg_color=COLOR_ACENTO, hover_color=COLOR_ACENTO_HOVER, font=ctk.CTkFont(size=11, weight="bold"), command=self._actualizar_chofer).grid(row=0, column=1, sticky="ew", padx=3)
        ctk.CTkButton(bch, text="🧹 Limpiar", height=30, fg_color=COLOR_NEUTRO, hover_color=COLOR_NEUTRO_HOVER, font=ctk.CTkFont(size=11), command=self._limpiar_form_chofer).grid(row=0, column=2, sticky="ew", padx=3)
        ctk.CTkButton(bch, text="🗑️ Borrar", height=30, fg_color=COLOR_PELIGRO, hover_color=COLOR_PELIGRO_HOVER, font=ctk.CTkFont(size=11, weight="bold"), command=self._eliminar_chofer).grid(row=0, column=3, sticky="ew", padx=(3, 0))

        m_tvch = ctk.CTkFrame(col_ch, fg_color="transparent")
        m_tvch.pack(fill="both", expand=True, padx=10, pady=(4, 10))
        cols_ch = [
            ("id", "ID", 34),
            ("nom", "Chofer", 140),
            ("ced", "Cédula", 82),
            ("pla", "Placa", 75),
            ("veh", "Vehículo", 95),
        ]
        self.tv_choferes = ttk.Treeview(m_tvch, columns=[c[0] for c in cols_ch], show="headings", selectmode="browse")
        sc_ch = ttk.Scrollbar(m_tvch, orient="vertical", command=self.tv_choferes.yview)
        self.tv_choferes.configure(yscrollcommand=sc_ch.set)
        for cid, tit, anc in cols_ch:
            self.tv_choferes.heading(cid, text=tit)
            self.tv_choferes.column(cid, width=anc, anchor="w" if cid in ("nom", "veh") else "center")
        sc_ch.pack(side="right", fill="y")
        self.tv_choferes.pack(side="left", fill="both", expand=True)
        self.tv_choferes.bind("<<TreeviewSelect>>", self._seleccionar_chofer_tabla)

    # ==========================================================================
    # MÓDULO 7: HISTORIAL GENERAL DE DOCUMENTOS
    # ==========================================================================
    def _armar_modulo_historial(self) -> None:
        p_hist = ctk.CTkFrame(
            self.tab_historial, fg_color=COLOR_TARJETA, border_width=1, border_color=COLOR_BORDE, corner_radius=12
        )
        p_hist.pack(fill="both", expand=True, pady=4)

        cab_h = ctk.CTkFrame(p_hist, fg_color="transparent")
        cab_h.pack(fill="x", padx=16, pady=(12, 6))
        ctk.CTkLabel(
            cab_h,
            text="📜 HISTORIAL GENERAL DE ÓRDENES, NOTAS DE CRÉDITO, NOTAS DE DÉBITO Y MOVIMIENTOS",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=COLOR_TEXTO_TITULO,
        ).pack(side="left")

        barra_h = ctk.CTkFrame(p_hist, fg_color=COLOR_SUBTARJETA, corner_radius=10)
        barra_h.pack(fill="x", padx=16, pady=4)

        ctk.CTkLabel(barra_h, text="Filtrar por Tipo de Documento:", font=ctk.CTkFont(weight="bold")).pack(
            side="left", padx=(12, 6), pady=8
        )
        self.cmb_filtro_hist = ctk.CTkComboBox(
            barra_h,
            width=185,
            height=30,
            values=["Todos", "DESPACHO", "NOTA_CREDITO", "NOTA_DEBITO", "MOVIMIENTO"],
            command=lambda _: self.refrescar_tabla_historial(),
        )
        self.cmb_filtro_hist.pack(side="left", padx=4, pady=8)
        self.cmb_filtro_hist.set("Todos")

        ctk.CTkButton(
            barra_h,
            text="🗑️ Eliminar Documento (Revierte Stock/Saldo)",
            width=260,
            height=30,
            fg_color=COLOR_PELIGRO,
            hover_color=COLOR_PELIGRO_HOVER,
            font=ctk.CTkFont(weight="bold"),
            command=self._eliminar_documento_historial,
        ).pack(side="right", padx=(6, 12), pady=8)

        ctk.CTkButton(
            barra_h,
            text="📊 Exportar Excel",
            width=135,
            height=30,
            fg_color=COLOR_PRIMARIO,
            hover_color=COLOR_PRIMARIO_HOVER,
            font=ctk.CTkFont(weight="bold"),
            command=lambda: self._reimprimir_historial(True),
        ).pack(side="right", padx=6, pady=8)

        ctk.CTkButton(
            barra_h,
            text="📄 Reimprimir PDF",
            width=145,
            height=30,
            fg_color=COLOR_ACENTO,
            hover_color=COLOR_ACENTO_HOVER,
            font=ctk.CTkFont(weight="bold"),
            command=lambda: self._reimprimir_historial(False),
        ).pack(side="right", padx=6, pady=8)

        m_h = ctk.CTkFrame(p_hist, fg_color="transparent")
        m_h.pack(fill="both", expand=True, padx=16, pady=(6, 14))
        cols_h = [
            ("id", "ID", 42),
            ("cor", "Correlativo", 95),
            ("tipo", "Categoría", 115),
            ("orig", "Sede Origen", 185),
            ("dest", "Cliente / Sede Destino", 210),
            ("chof", "Chofer / Placa", 160),
            ("kg", "Total KG", 90),
            ("ref", "Total REF", 95),
            ("bs", "Total Bs", 120),
            ("fec", "Fecha y Hora", 130),
        ]
        self.tv_historial = ttk.Treeview(m_h, columns=[c[0] for c in cols_h], show="headings", selectmode="browse")
        sc_h = ttk.Scrollbar(m_h, orient="vertical", command=self.tv_historial.yview)
        self.tv_historial.configure(yscrollcommand=sc_h.set)
        for cid, tit, anc in cols_h:
            self.tv_historial.heading(cid, text=tit)
            self.tv_historial.column(cid, width=anc, anchor="w" if cid in ("orig", "dest", "chof") else "center")
        sc_h.pack(side="right", fill="y")
        self.tv_historial.pack(side="left", fill="both", expand=True)

    # --------------------------------------------------------------------------
    # FUNCIONES CRUD DE PRODUCTOS, CLIENTES, CENTROS, CHOFERES E HISTORIAL
    # --------------------------------------------------------------------------
    def _seleccionar_prod_cat(self, _e=None) -> None:
        sel = self.tv_precios.selection()
        if not sel:
            return
        valores = self.tv_precios.item(sel[0], "values")
        if not valores or len(valores) < 2:
            return
        pid = int(valores[1])
        filas = consultar_sql("SELECT * FROM productos WHERE id=?", (pid,))
        if not filas:
            return
        p = filas[0]
        self.id_prod_edit = pid
        for e, val in (
            (self.ent_p_nom, p["nombre"]),
            (self.ent_p_may, f"{p['precio_mayor']:.2f}"),
            (self.ent_p_ven, f"{p['precio_vendedor']:.2f}"),
            (self.ent_p_cli, f"{p['precio_cliente']:.2f}"),
        ):
            e.delete(0, "end")
            e.insert(0, val)
        if p["rubro"] in RUBROS_CATALOGO:
            self.cmb_p_rubro.set(p["rubro"])
        if int(p["disponible"]) == 1:
            self.chk_prod_disp_form.select()
        else:
            self.chk_prod_disp_form.deselect()

    def _limpiar_form_producto(self) -> None:
        self.id_prod_edit = None
        self.tv_precios.selection_remove(self.tv_precios.selection())
        for e in (self.ent_p_nom, self.ent_p_may, self.ent_p_ven, self.ent_p_cli):
            e.delete(0, "end")
        self.chk_prod_disp_form.select()

    def _guardar_producto_cat(self) -> None:
        nom = self.ent_p_nom.get().strip().upper()
        rub = self.cmb_p_rubro.get().strip() or RUBROS_CATALOGO[0]
        disp = 1 if self.chk_prod_disp_form.get() == 1 else 0
        if not nom:
            messagebox.showwarning("Falta Producto", "Escribe el nombre del producto o selecciónalo en la tabla.")
            return
        try:
            pm = float(self.ent_p_may.get().replace(",", "."))
            pv = float(self.ent_p_ven.get().replace(",", "."))
            pc = float(self.ent_p_cli.get().replace(",", "."))
        except ValueError:
            messagebox.showerror("Precios Inválidos", "Verifica que los 3 precios sean números válidos.")
            return

        ejecutar_sql(
            """INSERT INTO productos (nombre, rubro, precio_mayor, precio_vendedor, precio_cliente, disponible)
               VALUES (?, ?, ?, ?, ?, ?)
               ON CONFLICT(nombre) DO UPDATE SET
                   rubro=excluded.rubro,
                   precio_mayor=excluded.precio_mayor,
                   precio_vendedor=excluded.precio_vendedor,
                   precio_cliente=excluded.precio_cliente,
                   disponible=excluded.disponible""",
            (nom, rub, round(pm, 2), round(pv, 2), round(pc, 2), disp),
        )
        for c in (self.carrito_despacho, self.carrito_nc, self.carrito_nd):
            c.recargar_productos()
        self.refrescar_datos_maestros()
        messagebox.showinfo("Catálogo Actualizado", f"Se guardaron los datos de '{nom}' correctamente.")

    def _eliminar_producto_cat(self) -> None:
        sel = self.tv_precios.selection()
        pid = int(self.tv_precios.item(sel[0], "values")[1]) if sel else self.id_prod_edit
        if not pid:
            messagebox.showwarning(
                "Selecciona un Producto",
                "Haz clic sobre un producto de la tabla derecha antes de pulsar '🗑️ Eliminar'.",
            )
            return
        filas = consultar_sql("SELECT nombre FROM productos WHERE id=?", (pid,))
        nom = filas[0]["nombre"] if filas else f"ID {pid}"
        if messagebox.askyesno(
            "Confirmar Eliminación Permanente",
            f"¿Deseas eliminar definitivamente '{nom}' de la base de datos?\n\n"
            "(Nota: Si solo quieres ocultarlo de la lista de hoy, usa la casilla ☑ / ☐ sin eliminarlo).",
        ):
            ejecutar_sql("DELETE FROM productos WHERE id=?", (pid,))
            ejecutar_sql("DELETE FROM inventario_sedes WHERE producto=?", (nom,))
            self._limpiar_form_producto()
            for c in (self.carrito_despacho, self.carrito_nc, self.carrito_nd):
                c.recargar_productos()
            self.refrescar_datos_maestros()
            messagebox.showinfo("Producto Eliminado", f"Se eliminó '{nom}' del catálogo.")

    def _importar_precios_app_anterior(self) -> None:
        ruta_db_ext = filedialog.askopenfilename(
            title="Selecciona una base de datos (.db) para importar productos",
            filetypes=[("Base de datos SQLite", "*.db")],
        )
        if not ruta_db_ext:
            return
        try:
            with sqlite3.connect(ruta_db_ext) as ext_conn:
                ext_conn.row_factory = sqlite3.Row
                cols_ext = {f["name"] for f in ext_conn.execute("PRAGMA table_info(productos)")}
                filas = ext_conn.execute("SELECT * FROM productos").fetchall()
            with conectar_db() as local_conn:
                for r in filas:
                    disp = int(r["disponible"]) if "disponible" in cols_ext else 1
                    local_conn.execute(
                        """INSERT INTO productos (nombre, rubro, precio_mayor, precio_vendedor, precio_cliente, disponible)
                           VALUES (?, ?, ?, ?, ?, ?)
                           ON CONFLICT(nombre) DO UPDATE SET
                               rubro=excluded.rubro,
                               precio_mayor=excluded.precio_mayor,
                               precio_vendedor=excluded.precio_vendedor,
                               precio_cliente=excluded.precio_cliente,
                               disponible=excluded.disponible""",
                        (r["nombre"], r["rubro"], r["precio_mayor"], r["precio_vendedor"], r["precio_cliente"], disp),
                    )
                local_conn.commit()
            for c in (self.carrito_despacho, self.carrito_nc, self.carrito_nd):
                c.recargar_productos()
            self.refrescar_datos_maestros()
            messagebox.showinfo("Sincronización Exitosa", f"Se importaron {len(filas)} productos actualizados.")
        except Exception as e:
            messagebox.showerror("Error al importar", str(e))

    def _obtener_id_cliente_seleccionado(self) -> Optional[int]:
        sel = self.tv_clientes.selection()
        if sel:
            return int(self.tv_clientes.item(sel[0], "values")[0])
        return self.id_cliente_edit

    def _guardar_cliente(self) -> None:
        nom = self.ent_cli_nombre.get().strip().upper()
        if not nom:
            messagebox.showwarning("Falta Nombre", "Escribe el nombre del cliente antes de guardar.")
            return
        ejecutar_sql(
            """INSERT INTO clientes (nombre, documento_id, escala_preferida, telefono, direccion)
               VALUES (?, ?, ?, ?, ?)""",
            (
                nom,
                self.ent_cli_rif.get().strip().upper() or "S/D",
                self.cmb_cli_esc.get(),
                self.ent_cli_tel.get().strip(),
                self.ent_cli_dir.get().strip(),
            ),
        )
        self._limpiar_form_cliente()
        self.refrescar_datos_maestros()
        messagebox.showinfo("Cliente Guardado", f"Se registró correctamente a '{nom}'.")

    def _actualizar_cliente(self) -> None:
        cid = self._obtener_id_cliente_seleccionado()
        if not cid:
            messagebox.showwarning("Selecciona un Cliente", "Haz clic sobre un cliente de la tabla inferior para actualizarlo.")
            return
        nom = self.ent_cli_nombre.get().strip().upper()
        if not nom:
            messagebox.showwarning("Falta Nombre", "El nombre del cliente no puede quedar vacío.")
            return
        ejecutar_sql(
            "UPDATE clientes SET nombre=?, documento_id=?, escala_preferida=?, telefono=?, direccion=? WHERE id=?",
            (
                nom,
                self.ent_cli_rif.get().strip().upper() or "S/D",
                self.cmb_cli_esc.get(),
                self.ent_cli_tel.get().strip(),
                self.ent_cli_dir.get().strip(),
                cid,
            ),
        )
        self.refrescar_datos_maestros()
        messagebox.showinfo("Cliente Actualizado", f"Los datos de '{nom}' fueron actualizados.")

    def _eliminar_cliente(self) -> None:
        cid = self._obtener_id_cliente_seleccionado()
        if not cid:
            messagebox.showwarning(
                "Selecciona un Cliente",
                "Primero haz clic sobre la fila del cliente en la tabla de abajo y luego presiona '🗑️ Borrar'.",
            )
            return
        filas = consultar_sql("SELECT nombre FROM clientes WHERE id=?", (cid,))
        nombre_cli = filas[0]["nombre"] if filas else f"ID {cid}"
        if messagebox.askyesno("Confirmar Eliminación", f"¿Estás seguro de eliminar al cliente '{nombre_cli}'?"):
            ejecutar_sql("DELETE FROM clientes WHERE id=?", (cid,))
            self._limpiar_form_cliente()
            self.refrescar_datos_maestros()
            messagebox.showinfo("Cliente Eliminado", f"El cliente '{nombre_cli}' fue eliminado de la base de datos.")

    def _limpiar_form_cliente(self) -> None:
        self.id_cliente_edit = None
        self.tv_clientes.selection_remove(self.tv_clientes.selection())
        for ent in (self.ent_cli_nombre, self.ent_cli_rif, self.ent_cli_tel, self.ent_cli_dir):
            ent.delete(0, "end")

    def _seleccionar_cliente_tabla(self, _e=None) -> None:
        sel = self.tv_clientes.selection()
        if not sel:
            return
        cid = int(self.tv_clientes.item(sel[0], "values")[0])
        filas = consultar_sql("SELECT * FROM clientes WHERE id=?", (cid,))
        if not filas:
            return
        c = filas[0]
        self.id_cliente_edit = cid
        for ent in (self.ent_cli_nombre, self.ent_cli_rif, self.ent_cli_tel, self.ent_cli_dir):
            ent.delete(0, "end")
        self.ent_cli_nombre.insert(0, c["nombre"])
        self.ent_cli_rif.insert(0, c["documento_id"])
        self.cmb_cli_esc.set(c["escala_preferida"])
        self.ent_cli_tel.insert(0, c["telefono"])
        self.ent_cli_dir.insert(0, c["direccion"])

    def _obtener_id_centro_seleccionado(self) -> Optional[int]:
        sel = self.tv_centros.selection()
        if sel:
            return int(self.tv_centros.item(sel[0], "values")[0])
        return self.id_centro_edit

    def _guardar_centro(self) -> None:
        cod = self.ent_cd_cod.get().strip().upper()
        nom = self.ent_cd_nom.get().strip().upper()
        if not cod or not nom:
            messagebox.showwarning("Datos Incompletos", "Escribe tanto el Código (ej. CD-03) como el Nombre del Centro.")
            return
        try:
            ejecutar_sql(
                "INSERT INTO centros_distribucion (codigo, nombre, encargado, telefono, ubicacion) VALUES (?, ?, ?, ?, ?)",
                (cod, nom, self.ent_cd_enc.get().strip(), self.ent_cd_tel.get().strip(), self.ent_cd_ubi.get().strip()),
            )
            self._limpiar_form_centro()
            self.refrescar_datos_maestros()
            messagebox.showinfo("Centro Registrado", f"El centro '{cod} - {nom}' fue guardado con éxito.")
        except sqlite3.IntegrityError:
            messagebox.showerror("Código Duplicado", f"Ya existe un Centro de Distribución con el código '{cod}'.")

    def _actualizar_centro(self) -> None:
        did = self._obtener_id_centro_seleccionado()
        if not did:
            messagebox.showwarning("Selecciona un Centro", "Haz clic sobre un Centro de Distribución en la tabla para actualizarlo.")
            return
        cod = self.ent_cd_cod.get().strip().upper()
        nom = self.ent_cd_nom.get().strip().upper()
        if not cod or not nom:
            messagebox.showwarning("Datos Requeridos", "El Código y el Nombre no pueden estar vacíos.")
            return

        antiguo = consultar_sql("SELECT codigo, nombre FROM centros_distribucion WHERE id=?", (did,))
        etiqueta_vieja = f"{antiguo[0]['codigo']} - {antiguo[0]['nombre']}" if antiguo else ""
        etiqueta_nueva = f"{cod} - {nom}"

        with conectar_db() as conn:
            conn.execute(
                "UPDATE centros_distribucion SET codigo=?, nombre=?, encargado=?, telefono=?, ubicacion=? WHERE id=?",
                (cod, nom, self.ent_cd_enc.get().strip(), self.ent_cd_tel.get().strip(), self.ent_cd_ubi.get().strip(), did),
            )
            if etiqueta_vieja and etiqueta_vieja != etiqueta_nueva:
                conn.execute("UPDATE inventario_sedes SET sede=? WHERE sede=?", (etiqueta_nueva, etiqueta_vieja))
            conn.commit()

        self.refrescar_datos_maestros()
        messagebox.showinfo("Centro Actualizado", f"Se actualizaron los datos de '{etiqueta_nueva}'.")

    def _eliminar_centro(self) -> None:
        did = self._obtener_id_centro_seleccionado()
        if not did:
            messagebox.showwarning(
                "Selecciona un Centro",
                "Primero haz clic sobre la fila del Centro de Distribución en la tabla inferior y luego presiona '🗑️ Borrar'.",
            )
            return
        filas = consultar_sql("SELECT codigo, nombre FROM centros_distribucion WHERE id=?", (did,))
        etiqueta = f"{filas[0]['codigo']} - {filas[0]['nombre']}" if filas else f"ID {did}"
        if messagebox.askyesno("Confirmar Eliminación", f"¿Deseas eliminar el Centro de Distribución '{etiqueta}'?"):
            ejecutar_sql("DELETE FROM centros_distribucion WHERE id=?", (did,))
            ejecutar_sql("DELETE FROM inventario_sedes WHERE sede=?", (etiqueta,))
            self._limpiar_form_centro()
            self.refrescar_datos_maestros()
            messagebox.showinfo("Centro Eliminado", f"Se eliminó '{etiqueta}' correctamente.")

    def _limpiar_form_centro(self) -> None:
        self.id_centro_edit = None
        self.tv_centros.selection_remove(self.tv_centros.selection())
        for ent in (self.ent_cd_cod, self.ent_cd_nom, self.ent_cd_enc, self.ent_cd_tel, self.ent_cd_ubi):
            ent.delete(0, "end")

    def _seleccionar_centro_tabla(self, _e=None) -> None:
        sel = self.tv_centros.selection()
        if not sel:
            return
        did = int(self.tv_centros.item(sel[0], "values")[0])
        filas = consultar_sql("SELECT * FROM centros_distribucion WHERE id=?", (did,))
        if not filas:
            return
        d = filas[0]
        self.id_centro_edit = did
        for ent in (self.ent_cd_cod, self.ent_cd_nom, self.ent_cd_enc, self.ent_cd_tel, self.ent_cd_ubi):
            ent.delete(0, "end")
        self.ent_cd_cod.insert(0, d["codigo"])
        self.ent_cd_nom.insert(0, d["nombre"])
        self.ent_cd_enc.insert(0, d["encargado"])
        self.ent_cd_tel.insert(0, d["telefono"])
        self.ent_cd_ubi.insert(0, d["ubicacion"])

    def _obtener_id_chofer_seleccionado(self) -> Optional[int]:
        sel = self.tv_choferes.selection()
        if sel:
            return int(self.tv_choferes.item(sel[0], "values")[0])
        return self.id_chofer_edit

    def _guardar_chofer(self) -> None:
        nom = self.ent_ch_nom.get().strip().upper()
        ced = self.ent_ch_ced.get().strip().upper()
        pla = self.ent_ch_pla.get().strip().upper()
        if not nom or not ced or not pla:
            messagebox.showwarning(
                "Datos del Chofer Incompletos",
                "Escribe el Nombre del Chofer, su Cédula y la Placa del Vehículo para guardarlo.",
            )
            return
        ejecutar_sql(
            "INSERT INTO choferes (nombre, cedula, placa, telefono, vehiculo) VALUES (?, ?, ?, ?, ?)",
            (nom, ced, pla, self.ent_ch_tel.get().strip(), self.ent_ch_veh.get().strip()),
        )
        self._limpiar_form_chofer()
        self.refrescar_datos_maestros()
        messagebox.showinfo("Chofer Guardado", f"El chofer '{nom}' (Placa: {pla}) quedó registrado para preselección.")

    def _actualizar_chofer(self) -> None:
        chid = self._obtener_id_chofer_seleccionado()
        if not chid:
            messagebox.showwarning("Selecciona un Chofer", "Haz clic sobre un chofer en la tabla inferior para editarlo.")
            return
        nom = self.ent_ch_nom.get().strip().upper()
        ced = self.ent_ch_ced.get().strip().upper()
        pla = self.ent_ch_pla.get().strip().upper()
        if not nom or not ced or not pla:
            messagebox.showwarning("Datos Requeridos", "Nombre, Cédula y Placa no pueden quedar vacíos.")
            return
        ejecutar_sql(
            "UPDATE choferes SET nombre=?, cedula=?, placa=?, telefono=?, vehiculo=? WHERE id=?",
            (nom, ced, pla, self.ent_ch_tel.get().strip(), self.ent_ch_veh.get().strip(), chid),
        )
        self.refrescar_datos_maestros()
        messagebox.showinfo("Chofer Actualizado", f"Se actualizaron los datos del chofer '{nom}'.")

    def _eliminar_chofer(self) -> None:
        chid = self._obtener_id_chofer_seleccionado()
        if not chid:
            messagebox.showwarning(
                "Selecciona un Chofer",
                "Primero haz clic sobre la fila del chofer en la tabla de abajo y luego presiona '🗑️ Borrar'.",
            )
            return
        filas = consultar_sql("SELECT nombre, placa FROM choferes WHERE id=?", (chid,))
        etiqueta = f"{filas[0]['nombre']} ({filas[0]['placa']})" if filas else f"ID {chid}"
        if messagebox.askyesno("Confirmar Eliminación", f"¿Estás seguro de eliminar al chofer '{etiqueta}'?"):
            ejecutar_sql("DELETE FROM choferes WHERE id=?", (chid,))
            self._limpiar_form_chofer()
            self.refrescar_datos_maestros()
            messagebox.showinfo("Chofer Eliminado", f"Se eliminó a '{etiqueta}' de la base de datos.")

    def _limpiar_form_chofer(self) -> None:
        self.id_chofer_edit = None
        self.tv_choferes.selection_remove(self.tv_choferes.selection())
        for ent in (self.ent_ch_nom, self.ent_ch_ced, self.ent_ch_pla, self.ent_ch_tel, self.ent_ch_veh):
            ent.delete(0, "end")

    def _seleccionar_chofer_tabla(self, _e=None) -> None:
        sel = self.tv_choferes.selection()
        if not sel:
            return
        chid = int(self.tv_choferes.item(sel[0], "values")[0])
        filas = consultar_sql("SELECT * FROM choferes WHERE id=?", (chid,))
        if not filas:
            return
        ch = filas[0]
        self.id_chofer_edit = chid
        for ent in (self.ent_ch_nom, self.ent_ch_ced, self.ent_ch_pla, self.ent_ch_tel, self.ent_ch_veh):
            ent.delete(0, "end")
        self.ent_ch_nom.insert(0, ch["nombre"])
        self.ent_ch_ced.insert(0, ch["cedula"])
        self.ent_ch_pla.insert(0, ch["placa"])
        self.ent_ch_tel.insert(0, ch["telefono"])
        self.ent_ch_veh.insert(0, ch["vehiculo"])

    def _reimprimir_historial(self, en_excel: bool) -> None:
        try:
            sel = self.tv_historial.selection()
            if not sel:
                messagebox.showwarning("Selecciona un Documento", "Haz clic en un documento del historial primero.")
                return
            doc_id = int(self.tv_historial.item(sel[0], "values")[0])
            ruta = exportar_documento_excel(doc_id) if en_excel else exportar_documento_pdf(doc_id)
            abrir_archivo(ruta)
        except Exception as error:
            messagebox.showerror("Error al reimprimir documento", str(error))

    def _eliminar_documento_historial(self) -> None:
        sel = self.tv_historial.selection()
        if not sel:
            messagebox.showwarning(
                "Selecciona un Documento",
                "Haz clic sobre una fila del historial de documentos y luego presiona '🗑️ Eliminar Documento'.",
            )
            return
        doc_id = int(self.tv_historial.item(sel[0], "values")[0])
        filas = consultar_sql("SELECT * FROM documentos WHERE id=?", (doc_id,))
        if not filas:
            return
        doc = filas[0]
        if not messagebox.askyesno(
            "Confirmar Eliminación de Documento",
            f"¿Deseas eliminar el documento '{doc['correlativo']}' ({doc['entidad_destino']})?\n\n"
            "El sistema revertirá automáticamente su impacto en saldos e inventario.",
        ):
            return

        items_doc = consultar_sql("SELECT * FROM documento_items WHERE documento_id=?", (doc_id,))

        with conectar_db() as conn:
            if doc["afecto_saldo"] == 1:
                if doc["tipo_doc"] == "NOTA_CREDITO" and doc["doc_identidad"]:
                    conn.execute(
                        "UPDATE centros_distribucion SET saldo_pendiente_ref = saldo_pendiente_ref + ? WHERE codigo = ?",
                        (doc["total_ref"], doc["doc_identidad"]),
                    )
                elif doc["tipo_doc"] == "NOTA_DEBITO":
                    cod_origen = doc["entidad_origen"].split(" - ")[0].strip()
                    conn.execute(
                        "UPDATE centros_distribucion SET saldo_pendiente_ref = saldo_pendiente_ref - ? WHERE codigo = ?",
                        (doc["total_ref"], cod_origen),
                    )

            if doc["afecto_inventario"] == 1:
                for it in items_doc:
                    kg = float(it["kilos"])
                    prod = it["producto"]
                    if doc["tipo_doc"] == "DESPACHO":
                        ajustar_stock_sede(conn, doc["entidad_origen"], prod, +kg)
                    elif doc["tipo_doc"] in ("MOVIMIENTO", "NOTA_DEBITO"):
                        ajustar_stock_sede(conn, doc["entidad_origen"], prod, +kg)
                        ajustar_stock_sede(conn, doc["entidad_destino"], prod, -kg)

            conn.execute("DELETE FROM documentos WHERE id=?", (doc_id,))
            conn.commit()

        self.refrescar_datos_maestros()
        messagebox.showinfo("Documento Eliminado", f"El registro '{doc['correlativo']}' fue eliminado y el stock fue restaurado.")

    # --------------------------------------------------------------------------
    # GUARDADO TRANSACCIONAL Y REFRESCO GENERAL DE TABLAS
    # --------------------------------------------------------------------------
    def _guardar_documento_en_db(
        self,
        tipo_doc: str,
        subtipo: str,
        origen: str,
        destino: str,
        doc_identidad: str,
        referencia: str,
        motivo: str,
        chofer: str = "",
        cedula_chofer: str = "",
        placa_vehiculo: str = "",
        afecto_saldo: int = 0,
        afecto_inventario: int = 1,
        observaciones: str = "",
        carrito: PanelCarritoProductos = None,
    ) -> int:
        correlativo = generar_siguiente_correlativo(tipo_doc)
        modo_moneda, tasa_usd, tasa_eur = self.obtener_contexto_moneda()
        tot_kg, tot_ref, tot_bs = carrito.calcular_totales()
        dia_sem, fecha_str, hora_str = obtener_marca_tiempo()

        origen_norm = normalizar_nombre_sede(origen)
        destino_norm = normalizar_nombre_sede(destino) if tipo_doc != "DESPACHO" else destino.strip()

        columnas = (
            "correlativo",
            "tipo_doc",
            "subtipo",
            "entidad_origen",
            "entidad_destino",
            "doc_identidad",
            "referencia",
            "motivo",
            "chofer",
            "cedula_chofer",
            "placa_vehiculo",
            "afecto_saldo",
            "afecto_inventario",
            "observaciones",
            "modo_moneda",
            "tasa_usd",
            "tasa_eur",
            "total_kilos",
            "total_ref",
            "total_bs",
            "dia_semana",
            "fecha",
            "hora",
        )
        valores = (
            correlativo,
            tipo_doc,
            subtipo,
            origen_norm,
            destino_norm,
            doc_identidad,
            referencia,
            motivo,
            chofer,
            cedula_chofer,
            placa_vehiculo,
            afecto_saldo,
            afecto_inventario,
            observaciones,
            modo_moneda,
            tasa_usd,
            tasa_eur,
            tot_kg,
            tot_ref,
            tot_bs,
            dia_sem,
            fecha_str,
            hora_str,
        )
        comodines = ", ".join(["?"] * len(valores))
        sql_cabecera = f"INSERT INTO documentos ({', '.join(columnas)}) VALUES ({comodines})"

        with conectar_db() as conn:
            cur = conn.execute(sql_cabecera, valores)
            doc_id = cur.lastrowid

            for linea in carrito.lineas:
                prod_nom = str(linea["producto"])
                kilos_lin = float(linea["kilos"])
                conn.execute(
                    """INSERT INTO documento_items (documento_id, producto, escala, kilos, precio_unitario, subtotal_ref)
                       VALUES (?, ?, ?, ?, ?, ?)""",
                    (
                        doc_id,
                        prod_nom,
                        linea["escala"],
                        kilos_lin,
                        linea["precio_unitario"],
                        linea["subtotal_ref"],
                    ),
                )

                if afecto_inventario == 1:
                    if tipo_doc == "DESPACHO":
                        ajustar_stock_sede(conn, origen_norm, prod_nom, -kilos_lin)
                    elif tipo_doc in ("MOVIMIENTO", "NOTA_DEBITO"):
                        ajustar_stock_sede(conn, origen_norm, prod_nom, -kilos_lin)
                        ajustar_stock_sede(conn, destino_norm, prod_nom, +kilos_lin)

            conn.commit()
        return doc_id

    def refrescar_tabla_inventario(self) -> None:
        if not hasattr(self, "tv_inventario"):
            return
        self.tv_inventario.delete(*self.tv_inventario.get_children())
        self.cache_inventario_vista.clear()

        modo_mon, t_usd, t_eur = self.obtener_contexto_moneda()
        tasa_apl, simb, nom_tasa = obtener_tasa_activa(modo_mon, t_usd, t_eur)

        sede_filtro = self.cmb_filtro_sede_inv.get()
        estado_filtro = self.cmb_filtro_estado_inv.get()

        prods_db = consultar_sql("SELECT nombre, precio_mayor FROM productos ORDER BY id ASC")
        mapa_precio_mayor = {p["nombre"]: float(p["precio_mayor"]) for p in prods_db}

        if sede_filtro == "Todas las Sedes":
            registros_db = consultar_sql("SELECT * FROM inventario_sedes ORDER BY sede ASC, producto ASC")
        else:
            registros_db = consultar_sql(
                "SELECT * FROM inventario_sedes WHERE sede=? ORDER BY producto ASC", (sede_filtro,)
            )

        filas_procesadas: List[Dict[str, object]] = []

        if "Catálogo Completo" in estado_filtro:
            sedes_objetivo = (
                [SEDE_PLANTA_PRINCIPAL] + list(self.mapa_centros.keys())
                if sede_filtro == "Todas las Sedes"
                else [sede_filtro]
            )
            mapa_existente = {(r["sede"], r["producto"]): r for r in registros_db}
            for s_nom in sedes_objetivo:
                for p in prods_db:
                    p_nom = p["nombre"]
                    reg = mapa_existente.get((s_nom, p_nom))
                    kg = float(reg["kilos_disponibles"]) if reg else 0.0
                    act = reg["ultima_actualizacion"] if reg else "Sin movimientos"
                    filas_procesadas.append({"sede": s_nom, "producto": p_nom, "kilos": kg, "actualizado": act})
        else:
            for r in registros_db:
                filas_procesadas.append({
                    "sede": r["sede"],
                    "producto": r["producto"],
                    "kilos": float(r["kilos_disponibles"]),
                    "actualizado": r["ultima_actualizacion"],
                })

        tot_kg = 0.0
        tot_ref = 0.0
        contador = 0

        for item in filas_procesadas:
            kg = round(float(item["kilos"]), 2)
            if "Positivo" in estado_filtro and kg <= 0:
                continue
            if "Alertas" in estado_filtro and kg >= UMBRAL_STOCK_BAJO_KG:
                continue

            if kg <= 0:
                estado_txt = "🔴 Sin Stock" if kg == 0 else "🔴 Stock Negativo"
                es_alerta = True
            elif kg < UMBRAL_STOCK_BAJO_KG:
                estado_txt = "🟡 Bajo Stock"
                es_alerta = True
            else:
                estado_txt = "🟢 En Stock"
                es_alerta = False

            p_mayor = mapa_precio_mayor.get(str(item["producto"]), 0.0)
            val_ref = round(kg * p_mayor, 2)
            val_bs = round(val_ref * tasa_apl, 2)

            tot_kg += kg
            tot_ref += val_ref
            contador += 1

            registro_cache = {
                "sede": item["sede"],
                "producto": item["producto"],
                "kilos": kg,
                "estado": estado_txt,
                "precio_mayor": p_mayor,
                "valor_ref": val_ref,
                "actualizado": item["actualizado"],
            }
            self.cache_inventario_vista.append(registro_cache)

            tag_fila = "alerta" if es_alerta else ("par" if contador % 2 == 0 else "impar")
            self.tv_inventario.insert(
                "",
                "end",
                tags=(tag_fila,),
                values=(
                    contador,
                    item["sede"],
                    item["producto"],
                    f"{kg:,.2f} kg",
                    estado_txt,
                    f"{simb} {val_ref:,.2f}",
                    f"Bs {val_bs:,.2f}",
                    item["actualizado"],
                ),
            )

        self.lbl_resumen_inventario.configure(
            text=(
                f"STOCK TOTAL EN VISTA: {tot_kg:,.2f} KG   |   "
                f"VALOR ESTIMADO (MAYOR): {simb} {tot_ref:,.2f}   |   "
                f"TOTAL {nom_tasa.upper()}: Bs {tot_ref * tasa_apl:,.2f}"
            )
        )

    def refrescar_datos_maestros(self) -> None:
        # 1. Clientes
        self.tv_clientes.delete(*self.tv_clientes.get_children())
        filas_cli = consultar_sql("SELECT * FROM clientes ORDER BY nombre ASC")
        self.mapa_clientes = {f"{c['nombre']} ({c['documento_id']})": c for c in filas_cli}
        for idx, c in enumerate(filas_cli, start=1):
            tag = "par" if idx % 2 == 0 else "impar"
            self.tv_clientes.insert(
                "",
                "end",
                tags=(tag,),
                values=(c["id"], c["nombre"], c["documento_id"], c["escala_preferida"], c["telefono"]),
            )

        claves_cli = list(self.mapa_clientes.keys())
        if claves_cli:
            self.cmb_cliente_desp.configure(values=claves_cli)
            if self.cmb_cliente_desp.get() not in self.mapa_clientes:
                self.cmb_cliente_desp.set(claves_cli[0])
            self._al_elegir_cliente_despacho(self.cmb_cliente_desp.get())

        # 2. Centros de Distribución
        self.tv_centros.delete(*self.tv_centros.get_children())
        filas_cd = consultar_sql("SELECT * FROM centros_distribucion ORDER BY codigo ASC")
        self.mapa_centros = {f"{d['codigo']} - {d['nombre']}": d for d in filas_cd}
        for idx, d in enumerate(filas_cd, start=1):
            tag = "par" if idx % 2 == 0 else "impar"
            saldo = float(d["saldo_pendiente_ref"] or 0.0)
            self.tv_centros.insert(
                "",
                "end",
                tags=(tag,),
                values=(d["id"], d["codigo"], d["nombre"], d["encargado"], f"$ {saldo:,.2f}"),
            )

        claves_cd = list(self.mapa_centros.keys())
        todas_las_sedes = [SEDE_PLANTA_PRINCIPAL] + claves_cd

        self.cmb_origen_desp.configure(values=todas_las_sedes)
        if self.cmb_origen_desp.get() not in todas_las_sedes:
            self.cmb_origen_desp.set(SEDE_PLANTA_PRINCIPAL)

        if claves_cd:
            self.cmb_cd_credito.configure(values=claves_cd)
            if self.cmb_cd_credito.get() not in self.mapa_centros:
                self.cmb_cd_credito.set(claves_cd[0])

        self._al_cambiar_modo_nd_mov(self.cmb_modo_nd_mov.get())

        # 3. Choferes y Vehículos
        self.tv_choferes.delete(*self.tv_choferes.get_children())
        filas_ch = consultar_sql("SELECT * FROM choferes ORDER BY nombre ASC")
        self.mapa_choferes = {f"{ch['nombre']} | Placa: {ch['placa']}": ch for ch in filas_ch}
        for idx, ch in enumerate(filas_ch, start=1):
            tag = "par" if idx % 2 == 0 else "impar"
            self.tv_choferes.insert(
                "",
                "end",
                tags=(tag,),
                values=(ch["id"], ch["nombre"], ch["cedula"], ch["placa"], ch["vehiculo"]),
            )

        opciones_chofer = [OPCION_CHOFER_MANUAL] + list(self.mapa_choferes.keys())
        self.cmb_chofer_pre.configure(values=opciones_chofer)
        if self.cmb_chofer_pre.get() not in opciones_chofer:
            self.cmb_chofer_pre.set(opciones_chofer[1] if len(opciones_chofer) > 1 else OPCION_CHOFER_MANUAL)
        if self.cmb_chofer_pre.get() in self.mapa_choferes:
            self._al_preseleccionar_chofer(self.cmb_chofer_pre.get())

        # 4. Selectores de Inventario
        opciones_filtro_inv = ["Todas las Sedes"] + todas_las_sedes
        self.cmb_filtro_sede_inv.configure(values=opciones_filtro_inv)
        if self.cmb_filtro_sede_inv.get() not in opciones_filtro_inv:
            self.cmb_filtro_sede_inv.set("Todas las Sedes")

        self.cmb_inv_sede_aj.configure(values=todas_las_sedes)
        if self.cmb_inv_sede_aj.get() not in todas_las_sedes:
            self.cmb_inv_sede_aj.set(claves_cd[0] if claves_cd else SEDE_PLANTA_PRINCIPAL)

        nombres_prod = [p["nombre"] for p in consultar_sql("SELECT nombre FROM productos ORDER BY nombre ASC")]
        if nombres_prod:
            self.cmb_inv_prod_aj.configure(values=nombres_prod)
            if self.cmb_inv_prod_aj.get() not in nombres_prod:
                self.cmb_inv_prod_aj.set(nombres_prod[0])

        self.refrescar_tabla_catalogo()
        self.refrescar_tabla_historial()
        self.refrescar_tabla_inventario()

    def refrescar_tabla_catalogo(self) -> None:
        if not hasattr(self, "tv_precios"):
            return
        self.tv_precios.delete(*self.tv_precios.get_children())
        modo_mon, t_usd, t_eur = self.obtener_contexto_moneda()
        tasa_apl, simb, _ = obtener_tasa_activa(modo_mon, t_usd, t_eur)

        rubro_f = self.cmb_filtro_rubro_lista.get() if hasattr(self, "cmb_filtro_rubro_lista") else "Todos los Rubros"
        if rubro_f == "Todos los Rubros":
            filas = consultar_sql("SELECT * FROM productos ORDER BY id ASC")
        else:
            filas = consultar_sql("SELECT * FROM productos WHERE rubro=? ORDER BY id ASC", (rubro_f,))

        disponibles_cont = 0
        for idx, p in enumerate(filas, start=1):
            disp = int(p["disponible"]) == 1
            if disp:
                disponibles_cont += 1
                texto_check = "☑ Disponible"
                tag = "par" if idx % 2 == 0 else "impar"
            else:
                texto_check = "☐ No disponible"
                tag = "no_disponible"

            p_cli = float(p["precio_cliente"])
            self.tv_precios.insert(
                "",
                "end",
                tags=(tag,),
                values=(
                    texto_check,
                    p["id"],
                    p["nombre"],
                    p["rubro"],
                    f"{simb} {p['precio_mayor']:.2f}",
                    f"{simb} {p['precio_vendedor']:.2f}",
                    f"{simb} {p_cli:.2f}",
                    f"Bs {p_cli * tasa_apl:,.2f}",
                ),
            )

        if hasattr(self, "lbl_contador_disp"):
            self.lbl_contador_disp.configure(
                text=f"Disponibles en Lista: {disponibles_cont} de {len(filas)} productos"
            )

    def refrescar_tabla_historial(self) -> None:
        if not hasattr(self, "tv_historial"):
            return
        self.tv_historial.delete(*self.tv_historial.get_children())
        filtro = self.cmb_filtro_hist.get()
        docs = (
            consultar_sql("SELECT * FROM documentos ORDER BY id DESC")
            if filtro == "Todos"
            else consultar_sql("SELECT * FROM documentos WHERE tipo_doc=? ORDER BY id DESC", (filtro,))
        )
        for idx, d in enumerate(docs, start=1):
            tag = "par" if idx % 2 == 0 else "impar"
            chof_txt = f"{d['chofer']} ({d['placa_vehiculo']})" if d["chofer"] or d["placa_vehiculo"] else "-"
            self.tv_historial.insert(
                "",
                "end",
                tags=(tag,),
                values=(
                    d["id"],
                    d["correlativo"],
                    d["tipo_doc"],
                    d["entidad_origen"],
                    d["entidad_destino"],
                    chof_txt,
                    f"{d['total_kilos']:,.2f} kg",
                    f"${d['total_ref']:,.2f}",
                    f"Bs {d['total_bs']:,.2f}",
                    f"{d['fecha']} {d['hora']}",
                ),
            )


if __name__ == "__main__":
    inicializar_logo()
    inicializar_base_de_datos()
    app = AppSistemaIntegralFinca()
    app.mainloop()