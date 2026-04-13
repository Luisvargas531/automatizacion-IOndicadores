"""
config.py
─────────────────────────────────────────────────────────────────────────────
Configuración central del proyecto.
Modifica SOLO este archivo para cambiar rutas, modos y opciones.
No toques la lógica de los demás módulos para ajustar el comportamiento.
─────────────────────────────────────────────────────────────────────────────
"""

from pathlib import Path

# ─────────────────────────────────────────────
# RUTAS PRINCIPALES
# ─────────────────────────────────────────────

BASE_DIR = Path(__file__).parent

# Ruta al archivo maestro (fuente de verdad que se actualiza)
MAESTRO_PATH = BASE_DIR / "data" / "Maestro.xlsx"

# Carpeta donde están los archivos individuales de indicadores
INDIVIDUALES_FOLDER = BASE_DIR / "data" / "individuales"

# Ruta donde se guardará el maestro actualizado
OUTPUT_PATH = BASE_DIR / "output" / "Maestro_actualizado.xlsx"

# Carpeta donde se guardan los backups automáticos del maestro
BACKUP_FOLDER = BASE_DIR / "backups"

# Carpeta donde se guardan los archivos de log
LOGS_FOLDER = BASE_DIR / "logs"

# Carpeta donde se guarda el reporte final
REPORTS_FOLDER = BASE_DIR / "reports"

# ─────────────────────────────────────────────
# OPCIONES DE COMPORTAMIENTO
# ─────────────────────────────────────────────

# DRY_RUN = True  → solo simula, NO guarda cambios al disco
# DRY_RUN = False → ejecuta y guarda cambios reales
DRY_RUN = False

# INSERT_NEW_RECORDS = True  → inserta filas nuevas si la llave no existe en el maestro
# INSERT_NEW_RECORDS = False → solo actualiza; lo no encontrado va al reporte
INSERT_NEW_RECORDS = True

# ─────────────────────────────────────────────
# LLAVES Y COLUMNAS REQUERIDAS
# ─────────────────────────────────────────────

# Llave de negocio: combinación exacta para el cruce
KEY_COLUMNS = ["Equipo", "Clave", "Indicador"]

# Columnas mínimas que debe tener el maestro para que el proceso arranque
MAESTRO_REQUIRED_COLS = [
    "Fecha", "Anio", "Mes", "Periodo_YYYYMM",
    "Equipo", "Clave", "Indicador",
    "Ejecucion", "Origen"
]

# Columnas mínimas obligatorias en cada archivo individual
# "Clave" es opcional: si no existe, el cruce usará solo Equipo + Indicador
INDIVIDUAL_REQUIRED_COLS = [
    "Equipo", "Indicador", "Ejecucion"
]

# ─────────────────────────────────────────────
# FORMATO DEL CAMPO MES
# ─────────────────────────────────────────────
# "numero" → almacena el mes como entero: 1, 2, 3 … 12
# "texto"  → almacena el mes como nombre en español: enero, febrero … diciembre
MES_FORMAT = "numero"

# ─────────────────────────────────────────────
# CELDA DE FECHA EN ARCHIVOS INDIVIDUALES
# ─────────────────────────────────────────────
# Celda donde se lee la fecha de realización en cada archivo individual.
# Si esa celda está vacía, el sistema intentará extraer la fecha del nombre
# del archivo buscando un patrón YYYYMM (ej: "..._202603.xlsx" → 2026-03-01).
DATE_CELL = "H2"

# ─────────────────────────────────────────────
# HOJAS DE EXCEL
# ─────────────────────────────────────────────

# Nombre o índice de la hoja del maestro. None = primera hoja
MAESTRO_SHEET = None

# Nombre o índice de la hoja de cada archivo individual. None = primera hoja
INDIVIDUAL_SHEET = None

# Fila donde empieza el encabezado de la tabla en los archivos individuales
# 0 = primera fila (comportamiento por defecto de pandas)
INDIVIDUAL_HEADER_ROW = 0

# ─────────────────────────────────────────────
# CARGA DE METAS
# ─────────────────────────────────────────────

# Carpeta donde están los archivos de metas anuales por equipo
# Patrón esperado: METAS_<EQUIPO>_<ANIO>.xlsx  (ej. METAS_ANALISIS_CUANTITATIVO_2026.xlsx)
METAS_FOLDER = BASE_DIR / "data" / "metas"

# Patrón glob para detectar archivos de metas dentro de METAS_FOLDER
METAS_FILE_PATTERN = "METAS_*.xlsx"

# Nombre (o substring) de la hoja de metas dentro de cada archivo
METAS_SHEET = "🎯 Metas Anuales"

# OVERWRITE_METAS = False → si la celda ya tiene meta, NO se sobreescribe
# OVERWRITE_METAS = True  → sobreescribe aunque ya tenga valor
OVERWRITE_METAS = False

# Cómo se distribuye la meta en indicadores con Periodicidad = "Anual":
#   "replicate"          → Meta_Periodo = Meta_Anual en TODOS los meses del año
#   "closing_month_only" → Meta_Periodo = Meta_Anual SOLO en diciembre (mes 12)
ANNUAL_GOAL_MODE = "replicate"
