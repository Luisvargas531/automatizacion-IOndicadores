"""
reporting.py
─────────────────────────────────────────────────────────────────────────────
Generación del reporte final en Excel con múltiples hojas.

Hojas del reporte:
  1. Resumen              → conteo global por archivo y totales
  2. Actualizados         → registros efectivamente actualizados (o simulados)
  3. No_Encontrados       → registros del individual sin match en el maestro
  4. Ambiguos             → registros no actualizados por ambigüedad
  5. Duplicados_Maestro   → llaves duplicadas detectadas en el maestro
  6. Duplicados_Individual→ llaves duplicadas detectadas en los individuales
  7. Errores_Fecha        → archivos donde H2 no se pudo leer como fecha
  8. Errores_Generales    → cualquier otro error durante el proceso
─────────────────────────────────────────────────────────────────────────────
"""

import logging
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import pandas as pd

logger = logging.getLogger("indicadores")


@dataclass
class ReportCollector:
    """
    Acumula todos los eventos del proceso para generar el reporte final.
    Se inicializa una sola vez en main.py y se pasa a cada módulo.
    """

    # Archivos procesados exitosamente: {nombre: {matches, no_encontrados, ...}}
    archivos_procesados: list = field(default_factory=list)

    # Registros actualizados (o que se actualizarían en dry-run)
    actualizados: list = field(default_factory=list)

    # Registros del individual sin match en el maestro
    no_encontrados: list = field(default_factory=list)

    # Registros ambiguos (no actualizados)
    ambiguos: list = field(default_factory=list)

    # Llaves duplicadas en el maestro
    duplicados_maestro: list = field(default_factory=list)

    # Llaves duplicadas en cada individual
    duplicados_individual: list = field(default_factory=list)

    # Archivos con error al leer H2
    errores_fecha: list = field(default_factory=list)

    # Errores generales (columnas faltantes, archivos inaccesibles, etc.)
    errores_generales: list = field(default_factory=list)

    def add_archivo_procesado(
        self,
        filename: str,
        matches: int,
        no_encontrados: int,
        ambiguos: int,
        duplicados_maestro: int,
        duplicados_individual: int,
    ):
        self.archivos_procesados.append({
            "archivo": filename,
            "matches_exactos": matches,
            "no_encontrados": no_encontrados,
            "ambiguos": ambiguos,
            "duplicados_maestro": duplicados_maestro,
            "duplicados_individual": duplicados_individual,
        })

    def add_error_fecha(self, filename: str, detalle: str):
        self.errores_fecha.append({"archivo": filename, "detalle": detalle})

    def add_error_general(self, filename: str, detalle: str):
        self.errores_generales.append({"archivo": filename, "detalle": detalle})


def generate_report(collector: ReportCollector, reports_folder: Path, dry_run: bool = False) -> Path:
    """
    Genera el reporte final en Excel con todas las hojas de trazabilidad.

    Parámetros
    ----------
    collector      : ReportCollector  Objeto con todos los datos acumulados.
    reports_folder : Path             Carpeta donde se guarda el reporte.
    dry_run        : bool             Si True, agrega sufijo "_SIMULACION" al nombre.

    Retorna
    -------
    Path  Ruta del reporte generado.
    """
    reports_folder.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    suffix = "_SIMULACION" if dry_run else ""
    report_path = reports_folder / f"Reporte_Indicadores{suffix}_{timestamp}.xlsx"

    # ── Construir DataFrames ──────────────────────────────────────────────

    # Hoja 1: Resumen
    df_resumen = pd.DataFrame(collector.archivos_procesados) if collector.archivos_procesados else pd.DataFrame(
        columns=["archivo", "matches_exactos", "no_encontrados", "ambiguos",
                 "duplicados_maestro", "duplicados_individual"]
    )
    if not df_resumen.empty:
        totales = {
            "archivo": "TOTAL",
            "matches_exactos": df_resumen["matches_exactos"].sum(),
            "no_encontrados": df_resumen["no_encontrados"].sum(),
            "ambiguos": df_resumen["ambiguos"].sum(),
            "duplicados_maestro": df_resumen["duplicados_maestro"].sum(),
            "duplicados_individual": df_resumen["duplicados_individual"].sum(),
        }
        df_resumen = pd.concat(
            [df_resumen, pd.DataFrame([totales])], ignore_index=True
        )

    # Hoja 2: Actualizados
    df_act = pd.DataFrame(collector.actualizados) if collector.actualizados else pd.DataFrame(
        columns=["archivo", "maestro_fila", "Equipo", "Clave", "Indicador",
                 "Ejecucion_nueva", "Fecha", "Anio", "Mes", "Periodo_YYYYMM", "simulado"]
    )

    # Hoja 3: No encontrados
    df_nf = pd.DataFrame(collector.no_encontrados) if collector.no_encontrados else pd.DataFrame(
        columns=["archivo", "fila_excel", "Equipo", "Clave", "Indicador", "Ejecucion"]
    )

    # Hoja 4: Ambiguos
    df_amb = pd.DataFrame(collector.ambiguos) if collector.ambiguos else pd.DataFrame(
        columns=["archivo", "fila_excel", "razon", "Equipo", "Clave", "Indicador", "Ejecucion"]
    )

    # Hoja 5: Duplicados Maestro
    df_dup_mae = pd.DataFrame(collector.duplicados_maestro) if collector.duplicados_maestro else pd.DataFrame(
        columns=["fuente", "fila_excel", "Equipo", "Clave", "Indicador"]
    )

    # Hoja 6: Duplicados Individual
    df_dup_ind = pd.DataFrame(collector.duplicados_individual) if collector.duplicados_individual else pd.DataFrame(
        columns=["fuente", "fila_excel", "Equipo", "Clave", "Indicador"]
    )

    # Hoja 7: Errores de Fecha
    df_err_fecha = pd.DataFrame(collector.errores_fecha) if collector.errores_fecha else pd.DataFrame(
        columns=["archivo", "detalle"]
    )

    # Hoja 8: Errores Generales
    df_err_gen = pd.DataFrame(collector.errores_generales) if collector.errores_generales else pd.DataFrame(
        columns=["archivo", "detalle"]
    )

    # ── Escribir Excel ───────────────────────────────────────────────────
    with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
        _write_sheet(writer, df_resumen, "Resumen")
        _write_sheet(writer, df_act, "Actualizados")
        _write_sheet(writer, df_nf, "No_Encontrados")
        _write_sheet(writer, df_amb, "Ambiguos")
        _write_sheet(writer, df_dup_mae, "Duplicados_Maestro")
        _write_sheet(writer, df_dup_ind, "Duplicados_Individual")
        _write_sheet(writer, df_err_fecha, "Errores_Fecha")
        _write_sheet(writer, df_err_gen, "Errores_Generales")

    logger.info(f"Reporte generado: {report_path}")
    return report_path


def _write_sheet(writer: pd.ExcelWriter, df: pd.DataFrame, sheet_name: str):
    """Escribe un DataFrame como hoja en el Excel con formato básico."""
    df.to_excel(writer, sheet_name=sheet_name, index=False)

    # Ajuste automático del ancho de columnas
    worksheet = writer.sheets[sheet_name]
    for col_cells in worksheet.columns:
        max_len = max(
            (len(str(cell.value)) if cell.value is not None else 0)
            for cell in col_cells
        )
        # Límite razonable para no deformar la hoja
        worksheet.column_dimensions[col_cells[0].column_letter].width = min(max_len + 4, 60)
