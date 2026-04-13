"""
goals_updater.py
─────────────────────────────────────────────────────────────────────────────
Aplica metas al maestro desde un DataFrame de metas normalizado.

Columnas que escribe en el maestro:
  Meta_Anual   → objetivo anual global del indicador (igual para todos los meses)
  Meta_Anual → objetivo específico del período:
                   • Mensual : valor del mes correspondiente (Meta_01..Meta_12)
                   • Anual / replicate          : Meta_Anual en todos los meses
                   • Anual / closing_month_only : Meta_Anual solo en diciembre

Reglas:
  - overwrite_metas = False (default) → no sobreescribe celdas que ya tienen valor
  - overwrite_metas = True            → sobreescribe siempre
─────────────────────────────────────────────────────────────────────────────
"""

import logging
import warnings

import pandas as pd

from normalization import normalize_value

# Suprimir FutureWarning de pandas sobre concat con columnas all-NA
warnings.filterwarnings(
    "ignore",
    message="The behavior of DataFrame concatenation with empty or all-NA entries",
    category=FutureWarning,
)

logger = logging.getLogger("indicadores")

# Nombre canónico por mes (debe coincidir con goals_reader.MES_CANONICAL)
_MES_CANONICAL: dict[int, str] = {n: f"Meta_{n:02d}" for n in range(1, 13)}

_MES_TEXTO_A_NUM: dict[str, int] = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
    "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
    "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12,
}

_MESES_ES: dict[int, str] = {v: k for k, v in _MES_TEXTO_A_NUM.items()}


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _to_mes_int(val) -> int | None:
    """Convierte 1, 1.0, '1', 'enero' → int 1..12, o None si no se puede."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip().lower()
    if s in _MES_TEXTO_A_NUM:
        return _MES_TEXTO_A_NUM[s]
    try:
        n = int(float(s))
        return n if 1 <= n <= 12 else None
    except Exception:
        return None


def _to_anio_int(val) -> int | None:
    try:
        return int(float(str(val).strip()))
    except Exception:
        return None


def _has_value(val) -> bool:
    """True si val no es None, NaN ni cadena vacía."""
    if val is None:
        return False
    if isinstance(val, float) and pd.isna(val):
        return False
    return str(val).strip() != ""


def _is_anual(periodicidad: str) -> bool:
    """True si la periodicidad indica frecuencia anual."""
    p = normalize_value(periodicidad)
    return "anual" in p or "annual" in p or "yearly" in p


def _ensure_meta_columns(maestro_df: pd.DataFrame) -> pd.DataFrame:
    """Agrega Meta_Anual al maestro si no existe."""
    if "Meta_Anual" not in maestro_df.columns:
        maestro_df["Meta_Anual"] = None
        logger.info("  [METAS] Columna 'Meta_Anual' creada en el maestro.")
    return maestro_df


# ─────────────────────────────────────────────────────────────────────────────
# FUNCIÓN PRINCIPAL
# ─────────────────────────────────────────────────────────────────────────────

def _meta_periodo_for_mes(
    goal_row: pd.Series,
    mes_num: int,
    es_anual: bool,
    meta_anual,
    annual_goal_mode: str,
) -> object:
    """Calcula el valor de Meta_Anual para un mes dado."""
    if es_anual:
        if annual_goal_mode == "closing_month_only":
            return meta_anual if mes_num == 12 else None
        return meta_anual  # "replicate"
    v = goal_row.get(_MES_CANONICAL[mes_num], None)
    return None if not _has_value(v) else v


def apply_metas(
    maestro_df: pd.DataFrame,
    goals_df: pd.DataFrame,
    anio: int,
    filename: str,
    overwrite_metas: bool = False,
    annual_goal_mode: str = "replicate",
    dry_run: bool = False,
) -> tuple[pd.DataFrame, list[dict]]:
    """
    Aplica las metas del goals_df al maestro para el año indicado.

    Parámetros
    ----------
    maestro_df       : DataFrame completo del maestro (sin filtro de período).
    goals_df         : DataFrame de metas (output de goals_reader.read_meta_file).
    anio             : Año al que corresponden las metas.
    filename         : Nombre del archivo de metas (para trazabilidad).
    overwrite_metas  : Si True, sobreescribe valores existentes.
    annual_goal_mode : "replicate" | "closing_month_only"
    dry_run          : Si True, simula sin modificar el maestro.

    Retorna
    -------
    tuple[pd.DataFrame, list[dict]]
        maestro_df (actualizado o sin cambios si dry_run) + log de operaciones
    """
    from config import MES_FORMAT

    meta_log: list[dict] = []
    all_new_rows: list[pd.Series] = []

    maestro_df = _ensure_meta_columns(maestro_df)

    for _, goal_row in goals_df.iterrows():
        equipo       = str(goal_row.get("Equipo", "")).strip()
        indicador    = str(goal_row.get("Indicador", "")).strip()
        periodicidad = str(goal_row.get("Periodicidad", "")).strip()
        meta_anual   = goal_row.get("Meta_Anual", None)
        if not _has_value(meta_anual):
            meta_anual = None

        equipo_norm    = normalize_value(equipo)
        indicador_norm = normalize_value(indicador)
        es_anual = _is_anual(periodicidad)

        # Localizar filas del maestro para Equipo + Indicador + Año
        mask = maestro_df.apply(
            lambda row: (
                normalize_value(str(row.get("Equipo", ""))) == equipo_norm
                and normalize_value(str(row.get("Indicador", ""))) == indicador_norm
                and _to_anio_int(row.get("Anio", None)) == anio
            ),
            axis=1,
        )
        target_rows = maestro_df[mask]

        # ── Registrar qué meses ya existen ──────────────────────────────
        existing_months: set[int] = set()
        for idx in target_rows.index:
            mes_val = maestro_df.at[idx, "Mes"] if "Mes" in maestro_df.columns else None
            mn = _to_mes_int(mes_val)
            if mn:
                existing_months.add(mn)

        # ── Actualizar meses que ya existen ──────────────────────────────
        for maestro_idx in target_rows.index:
            mes_raw = maestro_df.at[maestro_idx, "Mes"] if "Mes" in maestro_df.columns else None
            mes_num = _to_mes_int(mes_raw)

            meta_periodo = _meta_periodo_for_mes(
                goal_row, mes_num, es_anual, meta_anual, annual_goal_mode
            ) if mes_num else None

            ya_tiene_anual   = _has_value(maestro_df.at[maestro_idx, "Meta_Anual"])
            ya_tiene_periodo = _has_value(maestro_df.at[maestro_idx, "Meta_Anual"])

            aplicar_anual   = (meta_anual   is not None) and (overwrite_metas or not ya_tiene_anual)
            aplicar_periodo = (meta_periodo is not None) and (overwrite_metas or not ya_tiene_periodo)

            if not aplicar_anual and not aplicar_periodo:
                meta_log.append({
                    "accion": "OMITIDO", "archivo": filename,
                    "Equipo": equipo, "Indicador": indicador,
                    "Anio": anio, "Mes": mes_num,
                    "maestro_fila": maestro_idx + 2,
                    "Meta_Anual_nueva": meta_anual,
                    "Meta_Anual_nueva": meta_periodo,
                    "Periodicidad": periodicidad,
                    "nota": "Ya tiene meta y overwrite_metas=False",
                    "simulado": dry_run,
                })
                continue

            accion = "META_SIMULADA" if dry_run else "META_APLICADA"
            meta_log.append({
                "accion": accion, "archivo": filename,
                "Equipo": equipo, "Indicador": indicador,
                "Anio": anio, "Mes": mes_num,
                "maestro_fila": maestro_idx + 2,
                "Meta_Anual_nueva":  meta_anual   if aplicar_anual   else "(sin cambio)",
                "Meta_Anual_nueva": meta_periodo if aplicar_periodo else "(sin cambio)",
                "Periodicidad": periodicidad,
                "simulado": dry_run,
            })

            if not dry_run:
                if aplicar_anual:
                    maestro_df.at[maestro_idx, "Meta_Anual"] = meta_anual
                if aplicar_periodo:
                    maestro_df.at[maestro_idx, "Meta_Anual"] = meta_periodo
                logger.info(
                    f"  [META] {equipo} | {indicador} | Año={anio} Mes={mes_num} "
                    f"| Meta_Anual={meta_periodo}"
                )

        # ── Crear filas para meses que faltan en el maestro ──────────────
        missing_months = set(range(1, 13)) - existing_months

        if missing_months:
            # Buscar plantilla: fila más reciente del mismo Equipo+Indicador
            mask_any = maestro_df.apply(
                lambda row: (
                    normalize_value(str(row.get("Equipo", ""))) == equipo_norm
                    and normalize_value(str(row.get("Indicador", ""))) == indicador_norm
                ),
                axis=1,
            )
            all_rows_for_ind = maestro_df[mask_any]

            if not all_rows_for_ind.empty:
                def _safe_periodo(v):
                    try: return int(float(str(v)))
                    except: return 0
                template_idx = all_rows_for_ind["Periodo_YYYYMM"].apply(_safe_periodo).idxmax()
                template = all_rows_for_ind.loc[template_idx].copy()
            else:
                template = pd.Series({col: None for col in maestro_df.columns})
                template["Equipo"]    = equipo
                template["Indicador"] = indicador

            for mes_num in sorted(missing_months):
                nueva = template.copy()
                nueva["Anio"]          = anio
                nueva["Mes"]           = mes_num if MES_FORMAT == "numero" else _MESES_ES[mes_num]
                nueva["Periodo_YYYYMM"] = anio * 100 + mes_num
                nueva["Ejecucion"]     = None
                nueva["Origen"]        = f"Meta {filename}"
                nueva["Meta_Anual"]    = meta_anual
                nueva["Meta_Anual"]  = _meta_periodo_for_mes(
                    goal_row, mes_num, es_anual, meta_anual, annual_goal_mode
                )

                accion = "PERIODO_SIMULADO" if dry_run else "PERIODO_CREADO"
                meta_log.append({
                    "accion": accion, "archivo": filename,
                    "Equipo": equipo, "Indicador": indicador,
                    "Anio": anio, "Mes": mes_num,
                    "Periodo_YYYYMM": anio * 100 + mes_num,
                    "Meta_Anual_nueva":  meta_anual,
                    "Meta_Anual_nueva": nueva["Meta_Anual"],
                    "Periodicidad": periodicidad,
                    "nota": "Fila creada desde archivo de metas (sin ejecución aún)",
                    "simulado": dry_run,
                })

                if not dry_run:
                    all_new_rows.append(nueva)
                    logger.info(
                        f"  [META-NEW] {equipo} | {indicador} | "
                        f"Periodo={anio * 100 + mes_num} | Meta_Anual={nueva['Meta_Anual']}"
                    )
                else:
                    logger.info(
                        f"  [DRY-RUN META-NEW] {equipo} | {indicador} | "
                        f"Periodo={anio * 100 + mes_num}"
                    )

    # Agregar todas las filas nuevas de una sola vez
    if all_new_rows and not dry_run:
        nuevas_df = pd.DataFrame(all_new_rows, columns=maestro_df.columns)
        maestro_df = pd.concat([maestro_df, nuevas_df], ignore_index=True)
        logger.info(f"  → {len(all_new_rows)} fila(s) de período creada(s) desde metas.")

    return maestro_df, meta_log
