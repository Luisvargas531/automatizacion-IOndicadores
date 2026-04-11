"""
logger_utils.py
─────────────────────────────────────────────────────────────────────────────
Configuración del sistema de logging.
Escribe simultáneamente en consola y en un archivo .log con timestamp.
─────────────────────────────────────────────────────────────────────────────
"""

import logging
import sys
from datetime import datetime
from pathlib import Path

from config import LOGS_FOLDER


def setup_logger(name: str = "indicadores") -> logging.Logger:
    """
    Crea y configura el logger principal del proyecto.

    Parámetros
    ----------
    name : str
        Nombre del logger (usado en los mensajes).

    Retorna
    -------
    logging.Logger
        Logger listo para usar en todos los módulos.
    """
    LOGS_FOLDER.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = LOGS_FOLDER / f"proceso_{timestamp}.log"

    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)

    # Evitar duplicar handlers si setup_logger se llama más de una vez
    if logger.handlers:
        return logger

    formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)-8s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    # Handler de consola (INFO en adelante)
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)

    # Handler de archivo (DEBUG en adelante: incluye detalles técnicos)
    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)

    logger.addHandler(console_handler)
    logger.addHandler(file_handler)

    logger.info(f"Logger iniciado. Archivo de log: {log_file}")
    return logger
