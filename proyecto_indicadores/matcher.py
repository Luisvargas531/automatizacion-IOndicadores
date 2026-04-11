"""
matcher.py
─────────────────────────────────────────────────────────────────────────────
Lógica de cruce entre el archivo individual y el maestro.

Responsabilidades:
  1. Detectar duplicados en el maestro (por llave normalizada)
  2. Detectar duplicados en el individual (por llave normalizada)
  3. Realizar el cruce exacto y clasificar registros en:
       - matches_exactos   : 1 coincidencia → se puede actualizar
       - no_encontrados    : 0 coincidencias en el maestro
       - ambiguos          : >1 coincidencias en el maestro (no se actualiza)

La normalización se aplica SOLO para comparar; los valores originales
del maestro se preservan intactos en la columna visible.
─────────────────────────────────────────────────────────────────────────────
"""

import logging
from dataclasses import dataclass, field

import pandas as pd

from normalization import add_normalized_key_columns, build_normalized_tuple

logger = logging.getLogger("indicadores")


@dataclass
class MatchResult:
    """Resultado del proceso de cruce para un archivo individual."""

    # Filas del individual con su índice en el maestro donde actualizar
    # lista de dicts: {"ind_idx": i, "maestro_idx": j, "key_tuple": (...)}
    matches_exactos: list = field(default_factory=list)

    # Filas del individual que no tienen coincidencia en el maestro
    # lista de dicts con los datos de la fila individual
    no_encontrados: list = field(default_factory=list)

    # Filas del individual con más de una coincidencia en el maestro
    ambiguos: list = field(default_factory=list)

    # Llaves duplicadas detectadas dentro del maestro
    duplicados_maestro: list = field(default_factory=list)

    # Llaves duplicadas detectadas dentro del individual
    duplicados_individual: list = field(default_factory=list)


def find_duplicates(df: pd.DataFrame, key_cols: list[str], source: str) -> list[dict]:
    """
    Detecta filas con llave duplicada (usando columnas normalizadas).

    Parámetros
    ----------
    df       : pd.DataFrame  DataFrame con columnas _norm_ ya agregadas.
    key_cols : list[str]     Columnas originales de la llave.
    source   : str           Nombre del archivo (para logs y reporte).

    Retorna
    -------
    list[dict]
        Lista de registros con información de cada duplicado detectado.
    """
    norm_cols = [f"_norm_{c}" for c in key_cols]
    duplicados = df[df.duplicated(subset=norm_cols, keep=False)]

    result = []
    for idx, row in duplicados.iterrows():
        result.append({
            "fuente": source,
            "fila_excel": idx + 2,  # +2 porque idx=0 → fila 2 en Excel (encabezado en fila 1)
            "Equipo": row.get("Equipo", ""),
            "Clave": row.get("Clave", ""),
            "Indicador": row.get("Indicador", ""),
        })

    if result:
        logger.warning(
            f"  [DUPLICADO] {len(result)} filas con llave duplicada en '{source}'"
        )
    return result


def cross_match(
    maestro_df: pd.DataFrame,
    individual_df: pd.DataFrame,
    key_cols: list[str],
    filename: str,
    _effective_keys: list[str] | None = None,
) -> MatchResult:
    """
    Cruza el archivo individual contra el maestro usando la llave normalizada.

    Pasos:
      1. Agrega columnas _norm_ a ambos DataFrames.
      2. Detecta duplicados en maestro y en individual.
      3. Para cada fila del individual:
           - Busca coincidencias en el maestro por tupla normalizada.
           - Si 0 coincidencias → no_encontrado.
           - Si 1 coincidencia exacta (y NO está en duplicados del maestro) → match.
           - Si >1 coincidencias → ambiguo.
      4. Las filas del individual con llave duplicada se tratan como ambiguas.

    Parámetros
    ----------
    maestro_df    : pd.DataFrame  DataFrame del maestro (sin normalizar aún).
    individual_df : pd.DataFrame  DataFrame del individual (sin normalizar aún).
    key_cols      : list[str]     ["Equipo", "Clave", "Indicador"]
    filename      : str           Nombre del archivo individual (para trazabilidad).

    Retorna
    -------
    MatchResult
    """
    result = MatchResult()

    # ── Determinar llaves efectivas ──────────────────────────────────────
    # Si alguna columna de la llave no existe en el individual, usar solo
    # las que existen en AMBOS lados. Se registra advertencia.
    if _effective_keys is not None:
        key_cols = _effective_keys
    else:
        available_in_ind = [c for c in key_cols if c in individual_df.columns]
        if len(available_in_ind) < len(key_cols):
            missing_keys = [c for c in key_cols if c not in individual_df.columns]
            logger.warning(
                f"  [LLAVE] Columnas de llave ausentes en '{filename}': {missing_keys}. "
                f"Se usará llave reducida: {available_in_ind}. "
                f"ATENCIÓN: mayor riesgo de ambigüedad."
            )
            key_cols = available_in_ind

        if not key_cols:
            logger.error(f"  [LLAVE] Ninguna columna de llave existe en '{filename}'. Saltando.")
            return result

    # Normalizar llaves
    maestro_norm = add_normalized_key_columns(maestro_df, key_cols)
    individual_norm = add_normalized_key_columns(individual_df, key_cols)

    norm_cols = [f"_norm_{c}" for c in key_cols]

    # ── Detectar duplicados ──────────────────────────────────────────────
    result.duplicados_maestro = find_duplicates(
        maestro_norm, key_cols, "Maestro"
    )
    result.duplicados_individual = find_duplicates(
        individual_norm, key_cols, filename
    )

    # Llaves duplicadas en maestro → no se actualizan aunque haya match
    maestro_dup_keys = set(
        tuple(r[f"_norm_{c}"] for c in key_cols)
        for _, r in maestro_norm[
            maestro_norm.duplicated(subset=norm_cols, keep=False)
        ].iterrows()
    )

    # Llaves duplicadas en individual → se tratan como ambiguas
    ind_dup_keys = set(
        tuple(r[f"_norm_{c}"] for c in key_cols)
        for _, r in individual_norm[
            individual_norm.duplicated(subset=norm_cols, keep=False)
        ].iterrows()
    )

    # Construir índice del maestro: tupla_norm → lista de índices
    maestro_index: dict[tuple, list[int]] = {}
    for idx, row in maestro_norm.iterrows():
        key = build_normalized_tuple(row, key_cols)
        maestro_index.setdefault(key, []).append(idx)

    # ── Diagnóstico: mostrar muestras de llaves en ambos lados ──────────
    logger.info(f"  [DIAG] Columnas del individual: {list(individual_df.columns)}")
    logger.info(f"  [DIAG] Filas en individual: {len(individual_norm)}")
    logger.info(f"  [DIAG] Filas en maestro: {len(maestro_norm)}")

    # Mostrar las primeras 3 llaves del individual (valores originales + normalizados)
    for i, (_, row) in enumerate(individual_norm.head(3).iterrows()):
        orig = tuple(row.get(c, "") for c in key_cols)
        norm = build_normalized_tuple(row, key_cols)
        logger.info(f"  [DIAG] Individual fila {i+2} | original={orig} | normalizado={norm}")

    # Mostrar las primeras 3 llaves del maestro
    for i, (_, row) in enumerate(maestro_norm.head(3).iterrows()):
        orig = tuple(row.get(c, "") for c in key_cols)
        norm = build_normalized_tuple(row, key_cols)
        logger.info(f"  [DIAG] Maestro    fila {i+2} | original={orig} | normalizado={norm}")

    # ── Cruzar cada fila del individual ─────────────────────────────────
    for ind_idx, ind_row in individual_norm.iterrows():
        key = build_normalized_tuple(ind_row, key_cols)

        # Fila con llave duplicada en el individual → ambiguo
        if key in ind_dup_keys:
            result.ambiguos.append({
                "archivo": filename,
                "fila_excel": ind_idx + 2,
                "razon": "Llave duplicada en el archivo individual",
                "Equipo": ind_row.get("Equipo", ""),
                "Clave": ind_row.get("Clave", ""),
                "Indicador": ind_row.get("Indicador", ""),
                "Ejecucion": ind_row.get("Ejecucion", ""),
            })
            continue

        maestro_matches = maestro_index.get(key, [])

        if len(maestro_matches) == 0:
            # No existe en el maestro
            result.no_encontrados.append({
                "archivo": filename,
                "fila_excel": ind_idx + 2,
                "Equipo": ind_row.get("Equipo", ""),
                "Clave": ind_row.get("Clave", ""),
                "Indicador": ind_row.get("Indicador", ""),
                "Ejecucion": ind_row.get("Ejecucion", ""),
            })
            logger.debug(
                f"  [NO ENCONTRADO] {key} en '{filename}'"
            )

        elif len(maestro_matches) == 1:
            maestro_idx = maestro_matches[0]

            if key in maestro_dup_keys:
                # La coincidencia tiene duplicado en el maestro → ambiguo
                result.ambiguos.append({
                    "archivo": filename,
                    "fila_excel": ind_idx + 2,
                    "razon": "Llave duplicada en el maestro",
                    "Equipo": ind_row.get("Equipo", ""),
                    "Clave": ind_row.get("Clave", ""),
                    "Indicador": ind_row.get("Indicador", ""),
                    "Ejecucion": ind_row.get("Ejecucion", ""),
                })
                logger.warning(
                    f"  [AMBIGUO-MAESTRO] {key} — duplicado en maestro, no se actualiza."
                )
            else:
                result.matches_exactos.append({
                    "ind_idx": ind_idx,
                    "maestro_idx": maestro_idx,
                    "key_tuple": key,
                    "Equipo": ind_row.get("Equipo", ""),
                    "Clave": ind_row.get("Clave", ""),
                    "Indicador": ind_row.get("Indicador", ""),
                    "Ejecucion": ind_row.get("Ejecucion", ""),
                })
                logger.debug(f"  [MATCH] {key}")

        else:
            # Más de una coincidencia en el maestro → ambiguo
            result.ambiguos.append({
                "archivo": filename,
                "fila_excel": ind_idx + 2,
                "razon": f"Múltiples coincidencias en maestro ({len(maestro_matches)})",
                "Equipo": ind_row.get("Equipo", ""),
                "Clave": ind_row.get("Clave", ""),
                "Indicador": ind_row.get("Indicador", ""),
                "Ejecucion": ind_row.get("Ejecucion", ""),
            })
            logger.warning(
                f"  [AMBIGUO-MULTI] {key} tiene {len(maestro_matches)} "
                f"coincidencias en el maestro."
            )

    logger.info(
        f"  Cruce '{filename}': "
        f"matches={len(result.matches_exactos)} | "
        f"no_encontrados={len(result.no_encontrados)} | "
        f"ambiguos={len(result.ambiguos)}"
    )
    return result
