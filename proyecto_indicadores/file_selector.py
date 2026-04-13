"""
file_selector.py
─────────────────────────────────────────────────────────────────────────────
Diálogos nativos del sistema operativo para seleccionar archivos y carpetas.
Usa tkinter (incluido en Python, sin dependencias extra).

Muestra tres diálogos en secuencia:
  1. Seleccionar el archivo Maestro.xlsx
  2. Seleccionar la carpeta con los archivos individuales
  3. Seleccionar dónde guardar el maestro actualizado

Si el usuario cancela cualquier diálogo obligatorio, el programa se detiene
con un mensaje claro.
─────────────────────────────────────────────────────────────────────────────
"""

import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox


def _make_root() -> tk.Tk:
    """
    Crea una ventana Tk oculta que aparece al frente de todas las demás.
    Es necesario crearla antes de cada diálogo para que este tome foco.
    """
    root = tk.Tk()
    root.withdraw()                  # Ocultar ventana principal
    root.attributes("-topmost", True)  # Siempre al frente
    root.lift()
    root.focus_force()
    return root


def select_maestro(default_path: Path | None = None) -> Path:
    """
    Abre un diálogo para seleccionar el archivo Maestro.xlsx.

    Parámetros
    ----------
    default_path : Path | None
        Ruta inicial del diálogo (la carpeta donde abre por defecto).
        Si es None, abre en el directorio actual.

    Retorna
    -------
    Path  Ruta al archivo seleccionado.

    Sale del programa si el usuario cancela.
    """
    root = _make_root()

    initial_dir = str(default_path.parent) if default_path and default_path.exists() else "/"

    path_str = filedialog.askopenfilename(
        parent=root,
        title="Paso 1 de 3 — Selecciona el archivo Maestro.xlsx",
        initialdir=initial_dir,
        filetypes=[
            ("Archivos Excel", "*.xlsx"),
            ("Todos los archivos", "*.*"),
        ],
    )
    root.destroy()

    if not path_str:
        _cancelado("Maestro.xlsx")

    return Path(path_str)


def select_individuales_folder(default_path: Path | None = None) -> Path:
    """
    Abre un diálogo para seleccionar la carpeta con los archivos individuales.

    Parámetros
    ----------
    default_path : Path | None
        Carpeta inicial del diálogo. Si es None, abre en el directorio actual.

    Retorna
    -------
    Path  Ruta a la carpeta seleccionada.

    Sale del programa si el usuario cancela.
    """
    root = _make_root()

    initial_dir = str(default_path) if default_path and default_path.exists() else "/"

    folder_str = filedialog.askdirectory(
        parent=root,
        title="Paso 2 de 3 — Selecciona la carpeta con los archivos individuales",
        initialdir=initial_dir,
        mustexist=True,
    )
    root.destroy()

    if not folder_str:
        _cancelado("carpeta de individuales")

    return Path(folder_str)


def select_output_file(default_path: Path | None = None) -> Path:
    """
    Abre un diálogo para elegir dónde guardar el maestro actualizado.

    Si el usuario cancela, se usa la ruta por defecto de config.py
    (no es un paso obligatorio).

    Parámetros
    ----------
    default_path : Path | None
        Ruta sugerida (nombre y carpeta inicial).

    Retorna
    -------
    Path  Ruta de salida elegida o la ruta por defecto si se cancela.
    """
    root = _make_root()

    initial_dir = str(default_path.parent) if default_path else "/"
    initial_file = default_path.name if default_path else "Maestro_actualizado.xlsx"

    path_str = filedialog.asksaveasfilename(
        parent=root,
        title="Paso 3 de 3 — ¿Dónde guardar el Maestro actualizado?",
        initialdir=initial_dir,
        initialfile=initial_file,
        defaultextension=".xlsx",
        filetypes=[
            ("Archivos Excel", "*.xlsx"),
            ("Todos los archivos", "*.*"),
        ],
    )
    root.destroy()

    if not path_str:
        # El usuario canceló → usar ruta por defecto
        return default_path
    return Path(path_str)


def confirm_dry_run() -> bool:
    """
    Pregunta al usuario si desea ejecutar en modo simulación (dry-run).

    Retorna
    -------
    bool  True = simulación, False = ejecución real.
    """
    root = _make_root()
    respuesta = messagebox.askyesno(
        parent=root,
        title="Modo de ejecución",
        message=(
            "¿Deseas ejecutar en modo SIMULACIÓN?\n\n"
            "• SÍ  → Simula el proceso sin guardar cambios (recomendado para revisar primero).\n"
            "• NO  → Ejecuta el proceso real y actualiza el Maestro."
        ),
        icon=messagebox.QUESTION,
    )
    root.destroy()
    return respuesta


def select_metas_folder(default_path: "Path | None" = None) -> "Path | None":
    """
    Abre un diálogo para seleccionar la carpeta con los archivos de metas.
    Si el usuario cancela, retorna None (la carga de metas se omite).

    Parámetros
    ----------
    default_path : Path | None
        Carpeta inicial del diálogo.

    Retorna
    -------
    Path | None  Carpeta seleccionada, o None si se cancela.
    """
    root = _make_root()

    initial_dir = str(default_path) if default_path and default_path.exists() else "/"

    folder_str = filedialog.askdirectory(
        parent=root,
        title="Selecciona la carpeta con los archivos de METAS (cancela para omitir)",
        initialdir=initial_dir,
        mustexist=True,
    )
    root.destroy()

    return Path(folder_str) if folder_str else None


def _cancelado(nombre: str):
    """Muestra mensaje de cancelación y termina el programa."""
    root = _make_root()
    messagebox.showwarning(
        parent=root,
        title="Selección cancelada",
        message=f"No se seleccionó {nombre}.\nEl proceso fue cancelado.",
    )
    root.destroy()
    print(f"[CANCELADO] No se seleccionó {nombre}. El programa se detuvo.")
    sys.exit(0)
