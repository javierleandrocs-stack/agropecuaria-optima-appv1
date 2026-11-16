"""
================================================================================
EXPORTADOR OFICIAL (.EXE + .ZIP) ADAPTADO PARA 'Agropecuaria.py'
Carpeta del Proyecto: Programa Agropecuaria - Agropecuaria Óptima C.A
Propósito:
 1. Consolida los archivos temporales (.db-wal y .db-shm) en 'sistema_finca_optima.db'.
 2. Toma 'Agropecuaria.py' y lo convierte en un ejecutable .EXE independiente.
 3. Crea la carpeta instaladora y el archivo .ZIP listos para la otra laptop.
================================================================================
"""

import os
import sys
import shutil
import sqlite3
import tempfile
import subprocess
from pathlib import Path
import customtkinter as ctk
from tkinter import messagebox
from PIL import Image, ImageDraw


# ==============================================================================
# PARÁMETROS AJUSTABLES POR EL USUARIO
# ==============================================================================
# Apunta exactamente al nombre que le pusiste a tu archivo en Visual Studio Code
NOMBRE_ARCHIVO_CODIGO = "Agropecuaria.py"

# Nombre que tendrán el archivo .exe, la carpeta instaladora y el archivo .zip
NOMBRE_EJECUTABLE = "Sistema_Integral_Finca"
NOMBRE_CARPETA_INSTALADOR = "Instalador_Sistema_Integral_Finca"
NOMBRE_ARCHIVO_ZIP = "Sistema_Integral_Finca_Portable"

# Si es True, lleva tus precios, checklist, clientes, centros, choferes e inventario
INCLUIR_BASE_DE_DATOS_ACTUAL = True

# Si es True, crea también el botón de acceso directo en el Escritorio de tu PC actual
CREAR_ICONO_EN_ESTA_PC = True

# Si es True, abre la carpeta mostrándote el archivo .zip apenas termine
ABRIR_CARPETA_AL_TERMINAR = True

RUTA_BASE: Path = Path(__file__).resolve().parent
RUTA_LOGO: Path = RUTA_BASE / "logo_finca.png"
RUTA_ICONO: Path = RUTA_BASE / "icono_finca.ico"
RUTA_DB_UNIFICADA: Path = RUTA_BASE / "sistema_finca_optima.db"


# ==============================================================================
# PASO 1: CONSOLIDAR LA BASE DE DATOS (.DB-WAL / .DB-SHM) Y VERIFICAR PYINSTALLER
# ==============================================================================
def consolidar_base_de_datos() -> None:
    """
    Si existen archivos temporales 'sistema_finca_optima.db-wal' o '.db-shm',
    guarda todos sus cambios dentro de 'sistema_finca_optima.db' antes de copiarla.
    """
    if not RUTA_DB_UNIFICADA.exists():
        return
    try:
        with sqlite3.connect(RUTA_DB_UNIFICADA, timeout=10) as conn:
            conn.execute("PRAGMA wal_checkpoint(FULL);")
            conn.commit()
    except Exception:
        pass


def asegurar_pyinstaller() -> None:
    """Verifica que PyInstaller esté instalado en tu Python 3.12."""
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        subprocess.run(
            [sys.executable, "-m", "pip", "install", "pyinstaller"],
            check=True,
        )


# ==============================================================================
# PASO 2: LOCALIZAR 'Agropecuaria.py' Y RECURSOS GRÁFICOS
# ==============================================================================
def localizar_codigo_unificado() -> Path:
    """Encuentra 'Agropecuaria.py' ignorando 'exportar.py' y otros auxiliares."""
    for nombre_posible in (NOMBRE_ARCHIVO_CODIGO, "sistema_integral_finca.py", "despachos_finca.py"):
        candidato = RUTA_BASE / nombre_posible
        if candidato.exists():
            return candidato

    ignorados = {
        Path(__file__).name,
        "exportar.py",
        "exportar_sistema_integral.py",
        "exportar_portable.py",
        "exportar-portable.py",
        "copiar_a_pendrive.py",
        "verificar_pendrive.py",
    }
    candidatos = [p for p in RUTA_BASE.glob("*.py") if p.name not in ignorados]
    if not candidatos:
        raise FileNotFoundError(
            f"No se encontró '{NOMBRE_ARCHIVO_CODIGO}' en:\n{RUTA_BASE}"
        )
    candidatos.sort(key=lambda arch: arch.stat().st_size, reverse=True)
    return candidatos[0]


def asegurar_logo_e_icono() -> None:
    """Asegura que 'logo_finca.png' e 'icono_finca.ico' estén listos."""
    if not RUTA_LOGO.exists():
        img = Image.new("RGBA", (440, 170), (63, 90, 42, 255))
        dib = ImageDraw.Draw(img)
        dib.rounded_rectangle((8, 8, 432, 162), radius=20, outline=(250, 204, 21, 255), width=4)
        img.save(RUTA_LOGO, format="PNG")

    if not RUTA_ICONO.exists():
        with Image.open(RUTA_LOGO) as abierta:
            copia = abierta.convert("RGBA").copy()
        lado = max(copia.size) + 24
        cuadrado = Image.new("RGBA", (lado, lado), (63, 90, 42, 255))
        cuadrado.paste(copia, ((lado - copia.width) // 2, (lado - copia.height) // 2), copia)
        cuadrado.save(RUTA_ICONO, format="ICO", sizes=[(256, 256), (128, 128), (64, 64), (32, 32)])


# ==============================================================================
# PASO 3: COMPILAR EL EJECUTABLE (.EXE) Y ARMAR EL PAQUETE PORTABLE
# ==============================================================================
def compilar_ejecutable(script_principal: Path, carpeta_destino: Path) -> Path:
    """Compila 'Agropecuaria.py' en un archivo .EXE autónomo fuera de OneDrive."""
    if carpeta_destino.exists():
        shutil.rmtree(carpeta_destino)
    carpeta_destino.mkdir(parents=True, exist_ok=True)

    carpeta_temp = Path(tempfile.gettempdir()) / "build_agropecuaria_temp"
    if carpeta_temp.exists():
        shutil.rmtree(carpeta_temp, ignore_errors=True)
    carpeta_temp.mkdir(parents=True, exist_ok=True)

    comando = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--noconsole",
        "--onefile",
        "--collect-all",
        "customtkinter",
        "--hidden-import",
        "PIL._tkinter_finder",
        f"--workpath={carpeta_temp}",
        f"--specpath={carpeta_temp}",
        f"--distpath={carpeta_destino}",
        f"--icon={RUTA_ICONO}",
        f"--name={NOMBRE_EJECUTABLE}",
        str(script_principal),
    ]

    resultado = subprocess.run(comando, cwd=str(RUTA_BASE))
    ruta_exe = carpeta_destino / f"{NOMBRE_EJECUTABLE}.exe"

    if resultado.returncode != 0 or not ruta_exe.exists():
        raise RuntimeError(
            "No se pudo completar la compilación del archivo .exe. "
            "Verifica que hayas cerrado primero la ventana del programa Agropecuaria."
        )
    return ruta_exe


def preparar_carpeta_instaladora(carpeta_destino: Path) -> Path:
    """Copia el logo, el icono, la base de datos consolidada y el creador de icono."""
    shutil.copy2(RUTA_LOGO, carpeta_destino / "logo_finca.png")
    shutil.copy2(RUTA_ICONO, carpeta_destino / "icono_finca.ico")

    if INCLUIR_BASE_DE_DATOS_ACTUAL and RUTA_DB_UNIFICADA.exists():
        consolidar_base_de_datos()
        shutil.copy2(RUTA_DB_UNIFICADA, carpeta_destino / "sistema_finca_optima.db")

    codigo_vbs = (
        'Set fso = CreateObject("Scripting.FileSystemObject")\r\n'
        'carpeta = fso.GetParentFolderName(WScript.ScriptFullName)\r\n'
        'Set shell = CreateObject("WScript.Shell")\r\n'
        'escritorio = shell.SpecialFolders("Desktop")\r\n'
        'Set lnk = shell.CreateShortcut(escritorio & "\\Sistema Integral Finca.lnk")\r\n'
        f'lnk.TargetPath = carpeta & "\\{NOMBRE_EJECUTABLE}.exe"\r\n'
        'lnk.WorkingDirectory = carpeta\r\n'
        'lnk.IconLocation = carpeta & "\\icono_finca.ico,0"\r\n'
        'lnk.Description = "Sistema Integral: Precios, Despachos e Inventario - Agropecuaria Optima C.A"\r\n'
        'lnk.Save\r\n'
        'MsgBox "Icono \'Sistema Integral Finca\' creado en el Escritorio con exito.", 64, "Instalacion Lista"\r\n'
    )
    ruta_vbs = carpeta_destino / "Instalar_Icono_En_Escritorio.vbs"
    ruta_vbs.write_text(codigo_vbs, encoding="latin-1")

    guia_txt = (
        "PASOS PARA INSTALAR EN LA OTRA LAPTOP (SIN PYTHON):\n"
        "===================================================\n"
        "1. Copia toda esta carpeta dentro de 'Documentos' en la otra laptop.\n"
        "2. Haz doble clic en 'Instalar_Icono_En_Escritorio'.\n"
        "3. Abre el icono desde el Escritorio (Si sale aviso azul de Windows,\n"
        "   pulsa 'Mas informacion' -> 'Ejecutar de todas formas').\n"
    )
    (carpeta_destino / "LEEME_PASOS_INSTALACION.txt").write_text(guia_txt, encoding="utf-8")
    return ruta_vbs


def comprimir_en_zip(carpeta_instaladora: Path) -> Path:
    """Empaqueta toda la carpeta instaladora en un archivo .ZIP."""
    base_zip = RUTA_BASE / NOMBRE_ARCHIVO_ZIP
    zip_previo = RUTA_BASE / f"{NOMBRE_ARCHIVO_ZIP}.zip"
    if zip_previo.exists():
        zip_previo.unlink()

    ruta_zip = shutil.make_archive(
        base_name=str(base_zip),
        format="zip",
        root_dir=str(RUTA_BASE),
        base_dir=carpeta_instaladora.name,
    )
    return Path(ruta_zip)


# ==============================================================================
# PASO 4: VENTANA ASISTENTE PARA INICIAR LA EXPORTACIÓN CON 1 CLIC
# ==============================================================================
class VentanaExportador(ctk.CTk):
    """Ventana visual que te guía y confirma cuando el .EXE y el .ZIP están listos."""

    def __init__(self):
        super().__init__()
        self.title("Empaquetador a .EXE - Programa Agropecuaria")
        self.geometry("560x260")
        ctk.set_appearance_mode("Dark")
        self.configure(fg_color="#0D2618")

        ctk.CTkLabel(
            self,
            text="📦 Convertir 'Agropecuaria.py' en Ejecutable (.EXE)",
            font=ctk.CTkFont(size=17, weight="bold"),
            text_color="#6EE7B7",
        ).pack(pady=(20, 8))

        self.lbl_estado = ctk.CTkLabel(
            self,
            text="Asegúrate de haber cerrado la ventana principal del programa\n"
                 "y presiona el botón naranja para fabricar el instalador:",
            font=ctk.CTkFont(size=13),
        )
        self.lbl_estado.pack(pady=8)

        self.btn_iniciar = ctk.CTkButton(
            self,
            text="🚀 FABRICAR ARCHIVO .EXE Y PAQUETE .ZIP AHORA",
            width=440,
            height=44,
            fg_color="#D97706",
            hover_color="#B45309",
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.ejecutar_proceso,
        )
        self.btn_iniciar.pack(pady=18)

    def ejecutar_proceso(self) -> None:
        self.btn_iniciar.configure(state="disabled", text="⏳ Empaquetando Python y módulos (~1 minuto)...")
        self.lbl_estado.configure(
            text="Por favor espera unos segundos mientras se construye el ejecutable...",
            text_color="#FBBF24",
        )
        self.update()

        try:
            consolidar_base_de_datos()
            asegurar_pyinstaller()
            script_unificado = localizar_codigo_unificado()
            asegurar_logo_e_icono()

            carpeta_instaladora = RUTA_BASE / NOMBRE_CARPETA_INSTALADOR
            compilar_ejecutable(script_unificado, carpeta_instaladora)
            ruta_vbs = preparar_carpeta_instaladora(carpeta_instaladora)
            ruta_zip = comprimir_en_zip(carpeta_instaladora)

            if CREAR_ICONO_EN_ESTA_PC and sys.platform.startswith("win"):
                subprocess.run(["wscript", str(ruta_vbs)], check=False)

            if ABRIR_CARPETA_AL_TERMINAR and sys.platform.startswith("win"):
                subprocess.Popen(["explorer", f"/select,{ruta_zip}"])

            self.lbl_estado.configure(text="✅ ¡Proceso terminado con éxito!", text_color="#34D399")
            messagebox.showinfo(
                "¡Sistema Integral Empaquetado con Éxito!",
                f"Tu archivo 'Agropecuaria.py' ya fue convertido en .EXE:\n\n"
                f"1. Carpeta para copiar a pendrive:\n   {carpeta_instaladora}\n\n"
                f"2. Archivo .ZIP listo:\n   {ruta_zip}",
            )
            self.destroy()
        except Exception as error:
            self.btn_iniciar.configure(state="normal", text="🚀 REINTENTAR EXPORTACIÓN")
            self.lbl_estado.configure(text="⚠️ Ocurrió un aviso durante el proceso.", text_color="#FCA5A5")
            messagebox.showerror("Error durante la exportación", str(error))


if __name__ == "__main__":
    app = VentanaExportador()
    app.mainloop()