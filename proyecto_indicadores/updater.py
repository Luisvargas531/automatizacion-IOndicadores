"""
updater.py
─────────────────────────────────────────────────────────────────────────────
Aplica actualizaciones e inserciones sobre el DataFrame del maestro.

Responsabilidades:
  - apply_updates  : actualiza filas existentes con match exacto.
  - insert_records : inserta filas nuevas cuando el período no existe en el
                     maestro, copiando los metadatos del período más reciente
                     disponible para ese mismo Equipo + Indicador.
  - Respeta DRY_RUN: si está activo, simula sin modificar.
─────────────────────────────────────────────────────────────────────────────
"""

import logging

import pandas as pd

from normalization import normalize_value

logger = logging.getLogger("indicadores")

UPDATE_COLS = ["Ejecucion", "Fecha", "Anio", "Mes", "Periodo_YYYYMM", "Origen"]


def apply_updates(
    maestro_df: pd.DataFrame,
    individual_df: pd.DataFrame,
    matches: list[dict],
    date_fields: dict,
    filename: str,
    dry_run: bool = False,
) -> tuple[pd.DataFrame, list[dict]]:
    """
    Actualiza filas del maestro que tienen match exacto con el individual.
    """
    update_log = []

    if dry_run:
        logger.info(f"  [DRY-RUN] Se simularían {len(matches)} actualizaciones de '{filename}'.")

    for match in matches:
        ind_idx = match["ind_idx"]
        maestro_idx = match["maestro_idx"]
        ejecucion = individual_df.at[ind_idx, "Ejecucion"]

        update_log.append({
            "accion": "ACTUALIZADO",
            "archivo": filename,
            "maestro_fila": maestro_idx + 2,
            "Equipo": match.get("Equipo", ""),
            "Indicador": match.get("Indicador", ""),
            "Ejecucion_nueva": ejecucion,
            "Fecha": date_fields["Fecha"],
            "Anio": date_fields["Anio"],
            "Mes": date_fields["Mes"],
            "Periodo_YYYYMM": date_fields["Periodo_YYYYMM"],
            "simulado": dry_run,
        })

        if not dry_run:
            maestro_df.at[maestro_idx, "Ejecucion"] = ejecucion
            maestro_df.at[maestro_idx, "Fecha"] = date_fields["Fecha"]
            maestro_df.at[maestro_idx, "Anio"] = date_fields["Anio"]
            maestro_df.at[maestro_idx, "Mes"] = date_fields["Mes"]
            maestro_df.at[maestro_idx, "Periodo_YYYYMM"] = date_fields["Periodo_YYYYMM"]
            maestro_df.at[maestro_idx, "Origen"] = filename

            logger.info(
                f"  [UPDATE] {match.get('Equipo')} | {match.get('Indicador')} "
                f"| Ejecucion={ejecucion}"
            )

    return maestro_df, update_log


def insert_records(
    maestro_df: pd.DataFrame,
    individual_df: pd.DataFrame,
    no_encontrados: list[dict],
    date_fields: dict,
    filename: str,
    dry_run: bool = False,
) -> tuple[pd.DataFrame, list[dict]]:
    """
    Inserta filas nuevas en el maestro para registros del individual que
    no tienen fila correspondiente en el período actual.

    Estrategia de inserción:
      1. Busca en el maestro completo (sin filtro de período) las filas del
         mismo Equipo + Indicador.
      2. Toma la fila del período más reciente como plantilla.
      3. Copia todos sus metadatos (Clave, Agrupador_Simple, SubAgrupador,
         Unidad, Periodicidad, Meta_Anual, etc.).
      4. Sobreescribe: Ejecucion, Fecha, Anio, Mes, Periodo_YYYYMM, Origen.
      5. Si no existe ninguna fila previa, crea una fila mínima con los
         datos disponibles.

    Parámetros
    ----------
    maestro_df    : DataFrame completo del maestro (SIN filtro de período).
    individual_df : DataFrame del archivo individual.
    no_encontrados: Lista de dicts del MatchResult.no_encontrados.
    date_fields   : {"Fecha":..., "Anio":..., "Mes":..., "Periodo_YYYYMM":...}
    filename      : Nombre del archivo individual.
    dry_run       : Si True, simula sin modificar.

    Retorna
    -------
    tuple[pd.DataFrame, list[dict]]
        - maestro_df con las nuevas filas agregadas (o sin cambios si dry_run).
        - Lista de registros insertados para el reporte.
    """
    insert_log = []
    new_rows = []

    for record in no_encontrados:
        equipo = record.get("Equipo", "")
        indicador = record.get("Indicador", "")
        ejecucion = record.get("Ejecucion", "")

        equipo_norm = normalize_value(equipo)
        indicador_norm = normalize_value(indicador)

        # Buscar filas previas del mismo Equipo + Indicador en el maestro completo
        mask = maestro_df.apply(
            lambda row: (
                normalize_value(row.get("Equipo", "")) == equipo_norm and
                normalize_value(row.get("Indicador", "")) == indicador_norm
            ),
            axis=1
        )
        existing_rows = maestro_df[mask]

        if not existing_rows.empty:
            # Tomar la fila del período más reciente como plantilla
            def _safe_period(val):
                try:
                    return int(float(str(val)))
                except Exception:
                    return 0

            periodo_vals = existing_rows["Periodo_YYYYMM"].apply(_safe_period)
            template_idx = periodo_vals.idxmax()
            nueva_fila = existing_rows.loc[template_idx].copy()
            origen = f"Plantilla copiada de período {_safe_period(nueva_fila.get('Periodo_YYYYMM', ''))}"
        else:
            # Sin plantilla: crear fila mínima con las columnas del maestro
            nueva_fila = pd.Series({col: None for col in maestro_df.columns})
            nueva_fila["Equipo"] = equipo
            nueva_fila["Indicador"] = indicador
            origen = "Fila nueva sin plantilla previa"
            logger.warning(
                f"  [INSERT] No hay plantilla para {equipo} | {indicador}. "
                f"Se creará fila mínima."
            )

        # Actualizar campos del período actual
        nueva_fila["Ejecucion"] = ejecucion
        nueva_fila["Fecha"] = date_fields["Fecha"]
        nueva_fila["Anio"] = date_fields["Anio"]
        nueva_fila["Mes"] = date_fields["Mes"]
        nueva_fila["Periodo_YYYYMM"] = date_fields["Periodo_YYYYMM"]
        nueva_fila["Origen"] = filename

        insert_log.append({
            "accion": "INSERTADO",
            "archivo": filename,
            "Equipo": equipo,
            "Indicador": indicador,
            "Ejecucion_nueva": ejecucion,
            "Fecha": date_fields["Fecha"],
            "Anio": date_fields["Anio"],
            "Mes": date_fields["Mes"],
            "Periodo_YYYYMM": date_fields["Periodo_YYYYMM"],
            "nota": origen,
            "simulado": dry_run,
        })

        if not dry_run:
            new_rows.append(nueva_fila)
            logger.info(
                f"  [INSERT] {equipo} | {indicador} | "
                f"Periodo={date_fields['Periodo_YYYYMM']} | Ejecucion={ejecucion}"
            )
        else:
            logger.info(
                f"  [DRY-RUN INSERT] {equipo} | {indicador} | "
                f"Periodo={date_fields['Periodo_YYYYMM']}"
            )

    if new_rows and not dry_run:
        nuevas_df = pd.DataFrame(new_rows, columns=maestro_df.columns)
        maestro_df = pd.concat([maestro_df, nuevas_df], ignore_index=True)
        logger.info(f"  → {len(new_rows)} fila(s) insertada(s) en el maestro.")

    return maestro_df, insert_log
