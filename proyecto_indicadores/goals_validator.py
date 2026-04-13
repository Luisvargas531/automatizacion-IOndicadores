"""
goals_validator.py
─────────────────────────────────────────────────────────────────────────────
Valida el estado de las metas en el maestro para un Equipo + Año dado.

Estados posibles:
  COMPLETO → todos los registros del Equipo+Año tienen Meta_Anual no nulo
  PARCIAL  → al menos uno tiene valor y al menos uno es nulo
  VACIO    → ningún registro tiene Meta_Anual
─────────────────────────────────────────────────────────────────────────────
"""

import logging
from typing import Literal

import pandas as pd

from normalization import normalize_value

logger = logging.getLogger("indicadores")

MetaState = Literal["COMPLETO", "PARCIAL", "VACIO"]


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _to_anio_int(val) -> int | None:
    try:
        return int(float(str(val).strip()))
    except Exception:
        return None


def _has_value(val) -> bool:
    """Retorna True si val no es None, NaN ni cadena vacía."""
    if val is None:
        return False
    if isinstance(val, float) and pd.isna(val):
        return False
    return str(val).strip() != ""


# ─────────────────────────────────────────────────────────────────────────────
# FUNCIONES PRINCIPALES
# ─────────────────────────────────────────────────────────────────────────────

def check_meta_state(
    maestro_df: pd.DataFrame,
    equipo: str,
    anio: int,
    meta_col: str = "Meta_Anual",
) -> MetaState:
    """
    Determina si el maestro tiene metas cargadas para un Equipo + Año.

    Parámetros
    ----------
    maestro_df : pd.DataFrame  DataFrame completo del maestro.
    equipo     : str           Nombre del equipo (se normaliza para comparar).
    anio       : int           Año a verificar.
    meta_col   : str           Columna que representa la meta de período.

    Retorna
    -------
    MetaState : "COMPLETO", "PARCIAL" o "VACIO"
    """
    if meta_col not in maestro_df.columns:
        return "VACIO"

    equipo_norm = normalize_value(equipo)

    mask_equipo = maestro_df["Equipo"].apply(
        lambda v: normalize_value(str(v)) == equipo_norm
    )
    mask_anio = maestro_df["Anio"].apply(_to_anio_int) == anio

    subset = maestro_df[mask_equipo & mask_anio]

    if subset.empty:
        return "VACIO"

    n_total  = len(subset)
    n_filled = subset[meta_col].apply(_has_value).sum()

    if n_filled == 0:
        return "VACIO"
    if n_filled == n_total:
        return "COMPLETO"
    return "PARCIAL"


def summarize_meta_states(
    maestro_df: pd.DataFrame,
    equipos: list[str],
    anio: int,
    meta_col: str = "Meta_Anual",
) -> dict[str, MetaState]:
    """
    Retorna el estado de metas para una lista de equipos en un año dado.

    Retorna
    -------
    dict[str, MetaState]  {equipo: "COMPLETO" | "PARCIAL" | "VACIO"}
    """
    return {
        equipo: check_meta_state(maestro_df, equipo, anio, meta_col)
        for equipo in equipos
    }
