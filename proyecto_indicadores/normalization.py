"""
normalization.py
─────────────────────────────────────────────────────────────────────────────
Funciones de normalización de texto para la llave de cruce.

Objetivo: hacer que "  Equipo A " == "equipo a" == "Equipo  A" == "Équipo A"
          durante la comparación, SIN alterar los valores originales del maestro.

Las funciones de este módulo trabajan sobre COPIAS de los valores; nunca
modifican el DataFrame original en su columna de datos visible.
─────────────────────────────────────────────────────────────────────────────
"""

import re
import unicodedata

import pandas as pd


def normalize_value(value) -> str:
    """
    Normaliza un valor de texto para comparación:
      1. Convierte a string (None/NaN → cadena vacía)
      2. Strip de espacios extremos
      3. Colapsa espacios internos múltiples → un solo espacio
      4. Convierte a minúsculas
      5. Elimina tildes y diacríticos (á→a, ñ→n, ü→u, etc.)
      6. Elimina caracteres de control / no imprimibles

    Parámetros
    ----------
    value : any
        Valor crudo leído del Excel.

    Retorna
    -------
    str
        Cadena normalizada para comparación.
    """
    if pd.isna(value) or value is None:
        return ""

    text = str(value).strip()

    # Colapsar espacios múltiples
    text = re.sub(r"\s+", " ", text)

    # Minúsculas
    text = text.lower()

    # Eliminar diacríticos (tildes, diéresis, etc.)
    # NFD descompone el carácter en su base + marca; Mn = Non-Spacing Mark
    nfd = unicodedata.normalize("NFD", text)
    text = "".join(c for c in nfd if unicodedata.category(c) != "Mn")

    # Eliminar caracteres de control no imprimibles
    text = re.sub(r"[\x00-\x1f\x7f]", "", text)

    return text


def add_normalized_key_columns(
    df: pd.DataFrame,
    key_cols: list[str],
    prefix: str = "_norm_"
) -> pd.DataFrame:
    """
    Agrega al DataFrame columnas auxiliares con los valores de la llave
    ya normalizados. Las columnas originales NO se modifican.

    Ejemplo: key_cols=["Equipo","Clave","Indicador"]
             genera columnas "_norm_Equipo", "_norm_Clave", "_norm_Indicador"

    Parámetros
    ----------
    df : pd.DataFrame
        DataFrame al que se agregarán las columnas normalizadas.
    key_cols : list[str]
        Nombres de las columnas que forman la llave.
    prefix : str
        Prefijo para las columnas auxiliares (default "_norm_").

    Retorna
    -------
    pd.DataFrame
        El mismo DataFrame con columnas adicionales de normalización.
    """
    df = df.copy()
    for col in key_cols:
        if col in df.columns:
            df[f"{prefix}{col}"] = df[col].apply(normalize_value)
        else:
            # Si la columna no existe, se rellena con cadena vacía
            df[f"{prefix}{col}"] = ""
    return df


def build_normalized_tuple(row: pd.Series, key_cols: list[str], prefix: str = "_norm_") -> tuple:
    """
    Construye la tupla de llave normalizada a partir de una fila.

    Parámetros
    ----------
    row : pd.Series
        Fila del DataFrame.
    key_cols : list[str]
        Columnas que forman la llave.
    prefix : str
        Prefijo de las columnas normalizadas.

    Retorna
    -------
    tuple
        Tupla de strings normalizados, ej: ("equipo a", "ind-001", "ventas")
    """
    return tuple(row[f"{prefix}{col}"] for col in key_cols)
