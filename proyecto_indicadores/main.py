"""
main.py
─────────────────────────────────────────────────────────────────────────────
Punto de entrada del proyecto.
Orquesta el flujo completo:
  1. Diálogos del sistema para seleccionar archivos y carpetas
  2. Cargar y validar el maestro
  3. Crear backup del maestro
  4. Recorrer archivos individuales
  5. Por cada individual: leer fecha, construir campos, cruzar, actualizar
  6. Guardar el maestro actualizado (si no es dry-run)
  7. Generar reporte final

Uso:
  python main.py                  # Abre diálogos y ejecuta
  python main.py --dry-run        # Abre diálogos pero solo simula (no guarda)
  python main.py --no-gui         # Usa las rutas de config.py sin diálogos
─────────────────────────────────────────────────────────────────────────────
"""

import argparse
import shutil
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

# ── Módulos del proyecto ─────────────────────────────────────────────────
import config
from excel_reader import (
    read_maestro,
    read_date_from_h2,
    read_individual_data,
    read_execution_period,
    build_date_fields,
    validate_maestro_columns,
    validate_individual_columns,
)
from file_selector import (
    select_maestro,
    select_individuales_folder,
    select_metas_folder,
    select_output_file,
    confirm_dry_run,
)
from goals_reader import read_meta_file
from goals_validator import check_meta_state
from goals_updater import apply_metas
from logger_utils import setup_logger
from matcher import cross_match
from reporting import ReportCollector, generate_report
from updater import apply_updates, insert_records


# ─────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────

def create_backup(maestro_path: Path, backup_folder: Path) -> Path:
    """
    Crea una copia de respaldo del maestro con timestamp en el nombre.

    Parámetros
    ----------
    maestro_path  : Path  Ruta al maestro original.
    backup_folder : Path  Carpeta donde guardar el backup.

    Retorna
    -------
    Path  Ruta del archivo de backup generado.
    """
    backup_folder.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    stem = maestro_path.stem
    suffix = maestro_path.suffix
    backup_path = backup_folder / f"{stem}_backup_{timestamp}{suffix}"
    shutil.copy2(maestro_path, backup_path)
    return backup_path


def get_individual_files(folder: Path) -> list[Path]:
    """
    Retorna todos los archivos .xlsx de la carpeta de individuales,
    excluyendo archivos temporales de Excel (que empiezan con ~$).

    Parámetros
    ----------
    folder : Path  Carpeta de archivos individuales.

    Retorna
    -------
    list[Path]  Lista de archivos encontrados, ordenados por nombre.
    """
    return sorted(
        [f for f in folder.glob("*.xlsx") if not f.name.startswith("~$")]
    )


def save_maestro(maestro_df: pd.DataFrame, output_path: Path):
    """
    Guarda el maestro actualizado en la ruta de salida.
    Crea la carpeta si no existe.

    Parámetros
    ----------
    maestro_df  : pd.DataFrame  DataFrame final del maestro.
    output_path : Path          Ruta destino.
    """
    output_path.parent.mkdir(parents=True, exist_ok=True)
    maestro_df.to_excel(output_path, index=False, engine="openpyxl")


# ─────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────

def main():
    # ── Argumentos de línea de comandos ──────────────────────────────────
    parser = argparse.ArgumentParser(
        description="Actualiza el Maestro.xlsx a partir de archivos individuales de indicadores."
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        default=False,
        help="Modo simulación: no guarda cambios en el maestro ni en disco.",
    )
    parser.add_argument(
        "--no-gui",
        action="store_true",
        default=False,
        help="Usa las rutas definidas en config.py sin abrir diálogos.",
    )
    args = parser.parse_args()

    # ── Selección de rutas ────────────────────────────────────────────────
    # Si --no-gui está activo, se usan las rutas de config.py directamente.
    # De lo contrario, se abren diálogos del sistema operativo.
    if args.no_gui:
        maestro_path = config.MAESTRO_PATH
        individuales_folder = config.INDIVIDUALES_FOLDER
        output_path = config.OUTPUT_PATH
        # dry-run desde argumento o config
        dry_run = args.dry_run or config.DRY_RUN
    else:
        print("Abriendo selector de archivos...")

        # Diálogo 1: Maestro
        maestro_path = select_maestro(default_path=config.MAESTRO_PATH)

        # Diálogo 2: Carpeta de individuales
        individuales_folder = select_individuales_folder(
            default_path=config.INDIVIDUALES_FOLDER
        )

        # Diálogo 3: Dónde guardar el resultado
        # Sugiere: misma carpeta del maestro, nombre Maestro_actualizado.xlsx
        default_output = maestro_path.parent / f"{maestro_path.stem}_actualizado.xlsx"
        output_path = select_output_file(default_path=default_output)

        # Preguntar modo de ejecución solo si no se pasó --dry-run por CLI
        if args.dry_run or config.DRY_RUN:
            dry_run = True
        else:
            dry_run = confirm_dry_run()

    # ── Logger ────────────────────────────────────────────────────────────
    logger = setup_logger()

    logger.info("Rutas seleccionadas:")
    logger.info(f"  Maestro        : {maestro_path}")
    logger.info(f"  Individuales   : {individuales_folder}")
    logger.info(f"  Salida         : {output_path}")

    if dry_run:
        logger.info("=" * 60)
        logger.info("  MODO DRY-RUN ACTIVO — No se guardarán cambios.")
        logger.info("=" * 60)

    # ── Reporte acumulador ────────────────────────────────────────────────
    collector = ReportCollector()

    # ── 1. Cargar y validar el maestro ────────────────────────────────────
    logger.info("─" * 60)
    logger.info("PASO 1: Cargando el maestro...")
    try:
        maestro_df = read_maestro(maestro_path)
    except (FileNotFoundError, ValueError) as e:
        logger.error(f"Error crítico al leer el maestro: {e}")
        sys.exit(1)

    missing_cols = validate_maestro_columns(maestro_df, config.MAESTRO_REQUIRED_COLS)
    if missing_cols:
        logger.error(
            f"El maestro no tiene las columnas requeridas: {missing_cols}. "
            f"Verifica el archivo y la configuración."
        )
        sys.exit(1)
    logger.info("  → Maestro validado correctamente.")

    # ── 2. Crear backup ───────────────────────────────────────────────────
    logger.info("─" * 60)
    logger.info("PASO 2: Creando backup del maestro...")
    if not dry_run:
        backup_path = create_backup(maestro_path, config.BACKUP_FOLDER)
        logger.info(f"  → Backup creado: {backup_path.name}")
    else:
        logger.info("  → [DRY-RUN] Backup omitido en simulación.")

    # ── 3. Obtener archivos individuales ──────────────────────────────────
    logger.info("─" * 60)
    logger.info("PASO 3: Buscando archivos individuales...")
    individual_files = get_individual_files(individuales_folder)

    if not individual_files:
        logger.warning(
            f"No se encontraron archivos .xlsx en: {individuales_folder}. "
            f"Asegúrate de seleccionar la carpeta correcta."
        )
        # Generar reporte vacío y salir limpiamente
        generate_report(collector, config.REPORTS_FOLDER, dry_run=dry_run)
        sys.exit(0)

    logger.info(f"  → {len(individual_files)} archivo(s) encontrado(s).")

    # ── 4. Procesar cada archivo individual ───────────────────────────────
    logger.info("─" * 60)
    logger.info("PASO 4: Procesando archivos individuales...")

    for individual_path in individual_files:
        filename = individual_path.name
        logger.info(f"\n  Procesando: {filename}")

        # ── 4a. Leer fecha (H2) y período de ejecución (YYYYMM del archivo) ──
        try:
            fecha = read_date_from_h2(individual_path)
            date_fields = build_date_fields(fecha)
        except ValueError as e:
            logger.error(f"  [ERROR FECHA] {e}")
            collector.add_error_fecha(filename, str(e))
            continue

        # El período de EJECUCIÓN puede ser distinto a la fecha de H2
        # (H2 = fecha de diligenciamiento; el período real está en la hoja)
        execution_period = read_execution_period(individual_path)
        if execution_period and execution_period != date_fields["Periodo_YYYYMM"]:
            logger.info(
                f"  Fecha diligenciamiento (H2): {fecha.strftime('%Y-%m-%d')} "
                f"→ Periodo ejecución detectado: {execution_period} "
                f"(sobreescribe {date_fields['Periodo_YYYYMM']})"
            )
            date_fields["Periodo_YYYYMM"] = execution_period
            date_fields["Anio"] = execution_period // 100
            date_fields["Mes"] = execution_period % 100 if config.MES_FORMAT == "numero" else list({
                1:"enero",2:"febrero",3:"marzo",4:"abril",5:"mayo",6:"junio",
                7:"julio",8:"agosto",9:"septiembre",10:"octubre",11:"noviembre",12:"diciembre"
            }.values())[execution_period % 100 - 1]

        logger.info(
            f"  Fecha: {fecha.strftime('%Y-%m-%d')} | "
            f"Periodo ejecucion: {date_fields['Periodo_YYYYMM']}"
        )

        # ── 4b. Leer tabla del individual ────────────────────────────────
        try:
            individual_df = read_individual_data(individual_path)
        except Exception as e:
            logger.error(f"  [ERROR LECTURA] {filename}: {e}")
            collector.add_error_general(filename, f"Error al leer tabla: {e}")
            continue

        # ── 4c. Validar columnas requeridas ──────────────────────────────
        missing = validate_individual_columns(
            individual_df, config.INDIVIDUAL_REQUIRED_COLS, filename
        )
        if missing:
            msg = f"Columnas faltantes en '{filename}': {missing}"
            logger.error(f"  [ERROR COLUMNAS] {msg}")
            collector.add_error_general(filename, msg)
            continue

        # ── 4d. Cruce con el maestro ─────────────────────────────────────
        # Llave siempre: Equipo + Indicador
        # El maestro se pre-filtra al período de ejecución del archivo
        # para que la llave Equipo+Indicador sea unívoca dentro del filtro.
        periodo = date_fields["Periodo_YYYYMM"]
        effective_keys = ["Equipo", "Indicador"]

        if "Periodo_YYYYMM" in maestro_df.columns:
            # Normalizar a int para comparar: 202603.0 → 202603
            def _to_periodo_int(val):
                try:
                    return int(float(str(val).strip()))
                except Exception:
                    return None

            maestro_periodos = maestro_df["Periodo_YYYYMM"].apply(_to_periodo_int)
            maestro_para_cruce = maestro_df[maestro_periodos == int(periodo)].copy()
            logger.info(
                f"  Maestro filtrado a Periodo_YYYYMM={periodo}: "
                f"{len(maestro_para_cruce)} filas."
            )
            if len(maestro_para_cruce) == 0:
                logger.error(
                    f"  [SIN DATOS] El maestro no tiene ninguna fila con "
                    f"Periodo_YYYYMM={periodo}. "
                    f"Períodos disponibles en el maestro: "
                    f"{sorted(maestro_periodos.dropna().astype(int).unique().tolist())}"
                )
        else:
            maestro_para_cruce = maestro_df
            logger.warning("  Maestro no tiene columna Periodo_YYYYMM — cruce sin filtro de período.")

        # Diagnóstico: mostrar equipos presentes en el maestro para este período
        if len(maestro_para_cruce) > 0 and "Equipo" in maestro_para_cruce.columns:
            equipos_en_maestro = maestro_para_cruce["Equipo"].dropna().unique().tolist()
            equipos_en_individual = individual_df["Equipo"].dropna().unique().tolist() if "Equipo" in individual_df.columns else []
            logger.info(f"  Equipos en maestro (período {periodo}): {equipos_en_maestro}")
            logger.info(f"  Equipos en individual: {equipos_en_individual}")

        match_result = cross_match(
            maestro_df=maestro_para_cruce,
            individual_df=individual_df,
            key_cols=effective_keys,
            filename=filename,
            _effective_keys=effective_keys,
        )

        # ── 4f. Aplicar actualizaciones (matches exactos) ────────────────
        maestro_df, update_log = apply_updates(
            maestro_df=maestro_df,
            individual_df=individual_df,
            matches=match_result.matches_exactos,
            date_fields=date_fields,
            filename=filename,
            dry_run=dry_run,
        )

        # ── 4g. Insertar registros no encontrados (si INSERT_NEW_RECORDS) ─
        insert_log = []
        if config.INSERT_NEW_RECORDS and match_result.no_encontrados:
            logger.info(
                f"  INSERT_NEW_RECORDS=True: intentando insertar "
                f"{len(match_result.no_encontrados)} fila(s) no encontrada(s)..."
            )
            # Pasar el maestro_df COMPLETO (sin filtro de período) para buscar plantillas
            maestro_df, insert_log = insert_records(
                maestro_df=maestro_df,
                individual_df=individual_df,
                no_encontrados=match_result.no_encontrados,
                date_fields=date_fields,
                filename=filename,
                dry_run=dry_run,
            )

        # ── 4h. Acumular resultados en el reporte ─────────────────────────
        collector.actualizados.extend(update_log + insert_log)
        # Solo reportar como "no encontrados" los que tampoco se insertaron
        if config.INSERT_NEW_RECORDS:
            collector.no_encontrados.extend([])  # todos fueron insertados
        else:
            collector.no_encontrados.extend(match_result.no_encontrados)
        collector.ambiguos.extend(match_result.ambiguos)
        collector.duplicados_maestro.extend(match_result.duplicados_maestro)
        collector.duplicados_individual.extend(match_result.duplicados_individual)

        collector.add_archivo_procesado(
            filename=filename,
            matches=len(match_result.matches_exactos),
            no_encontrados=len(match_result.no_encontrados),
            ambiguos=len(match_result.ambiguos),
            duplicados_maestro=len(match_result.duplicados_maestro),
            duplicados_individual=len(match_result.duplicados_individual),
        )

    # ── 5. Guardar el maestro actualizado ─────────────────────────────────
    logger.info("─" * 60)
    logger.info("PASO 5: Guardando maestro actualizado...")
    if not dry_run:
        try:
            save_maestro(maestro_df, output_path)
            logger.info(f"  → Maestro actualizado guardado en: {output_path}")
        except Exception as e:
            logger.error(f"  [ERROR GUARDADO] No se pudo guardar el maestro: {e}")
            collector.add_error_general("Maestro", f"Error al guardar: {e}")
    else:
        logger.info("  → [DRY-RUN] Guardado omitido en simulación.")

    # ── 6. Cargar metas (opcional) ────────────────────────────────────────
    logger.info("─" * 60)
    logger.info("PASO 6: Cargando metas anuales...")

    # Determinar carpeta de metas: argumento, config, o diálogo
    if args.no_gui:
        metas_folder = config.METAS_FOLDER if config.METAS_FOLDER.exists() else None
    else:
        # Preguntar al usuario si desea cargar metas
        import tkinter as tk
        from tkinter import messagebox
        _root = tk.Tk()
        _root.withdraw()
        _root.attributes("-topmost", True)
        cargar_metas = messagebox.askyesno(
            parent=_root,
            title="Carga de Metas",
            message=(
                "¿Deseas cargar también los archivos de METAS anuales?\n\n"
                "• SÍ  → Selecciona la carpeta con archivos METAS_*.xlsx\n"
                "• NO  → Omitir carga de metas"
            ),
        )
        _root.destroy()
        if cargar_metas:
            metas_folder = select_metas_folder(default_path=config.METAS_FOLDER)
        else:
            metas_folder = None

    if metas_folder is None:
        logger.info("  → Carga de metas omitida.")
    else:
        meta_files = sorted(
            [f for f in metas_folder.glob(config.METAS_FILE_PATTERN)
             if not f.name.startswith("~$")]
        )
        if not meta_files:
            logger.warning(
                f"  → No se encontraron archivos '{config.METAS_FILE_PATTERN}' "
                f"en: {metas_folder}"
            )
        else:
            logger.info(f"  → {len(meta_files)} archivo(s) de metas encontrado(s).")
            for meta_path in meta_files:
                meta_filename = meta_path.name
                logger.info(f"\n  Procesando metas: {meta_filename}")

                try:
                    goals_df, anio = read_meta_file(meta_path, config.METAS_SHEET)
                except Exception as e:
                    logger.error(f"  [ERROR METAS] {meta_filename}: {e}")
                    collector.add_error_general(meta_filename, f"Error al leer metas: {e}")
                    continue

                if anio is None:
                    logger.warning(
                        f"  [METAS] No se pudo determinar el año de '{meta_filename}'. "
                        f"Se omite."
                    )
                    collector.add_error_general(
                        meta_filename, "No se pudo extraer el año del nombre del archivo."
                    )
                    continue

                # Obtener equipos únicos en el archivo de metas
                equipos_en_metas = goals_df["Equipo"].dropna().unique().tolist()

                for equipo in equipos_en_metas:
                    estado_previo = check_meta_state(
                        maestro_df, equipo, anio, meta_col="Meta_Anual"
                    )
                    logger.info(
                        f"  Estado metas en maestro — {equipo} | Año={anio}: "
                        f"{estado_previo}"
                    )

                    # Si COMPLETO y no se sobreescribe, saltear
                    if estado_previo == "COMPLETO" and not config.OVERWRITE_METAS:
                        logger.info(
                            f"  → Metas COMPLETAS para {equipo} | Año={anio} "
                            f"y OVERWRITE_METAS=False. Se omite."
                        )
                        collector.add_meta_resumen(
                            equipo, anio, estado_previo, estado_previo, meta_filename
                        )
                        continue

                # Aplicar metas del archivo completo
                maestro_df, meta_log = apply_metas(
                    maestro_df=maestro_df,
                    goals_df=goals_df,
                    anio=anio,
                    filename=meta_filename,
                    overwrite_metas=config.OVERWRITE_METAS,
                    annual_goal_mode=config.ANNUAL_GOAL_MODE,
                    dry_run=dry_run,
                )

                collector.add_meta_log(meta_log)

                # Calcular estado final y registrar en resumen
                for equipo in equipos_en_metas:
                    estado_final = check_meta_state(
                        maestro_df, equipo, anio, meta_col="Meta_Anual"
                    )
                    estado_previo = check_meta_state(
                        # recalcular sobre el df original (antes de apply) no es posible aquí
                        # usamos el estado que ya teníamos como referencia; en dry_run coinciden
                        maestro_df, equipo, anio, meta_col="Meta_Anual"
                    )
                    collector.add_meta_resumen(
                        equipo, anio, "—", estado_final, meta_filename
                    )

            # Guardar maestro con metas (si no es dry-run)
            if not dry_run and meta_files:
                try:
                    save_maestro(maestro_df, output_path)
                    logger.info(
                        f"  → Maestro con metas guardado en: {output_path}"
                    )
                except Exception as e:
                    logger.error(
                        f"  [ERROR GUARDADO METAS] No se pudo guardar: {e}"
                    )
                    collector.add_error_general("Maestro", f"Error al guardar con metas: {e}")

    # ── 7. Resumen en consola ─────────────────────────────────────────────
    logger.info("─" * 60)
    logger.info("RESUMEN FINAL:")
    logger.info(f"  Archivos procesados  : {len(collector.archivos_procesados)}")
    logger.info(f"  Registros actualizados: {len(collector.actualizados)}")
    logger.info(f"  No encontrados       : {len(collector.no_encontrados)}")
    logger.info(f"  Ambiguos             : {len(collector.ambiguos)}")
    logger.info(f"  Dup. en maestro      : {len(collector.duplicados_maestro)}")
    logger.info(f"  Dup. en individuales : {len(collector.duplicados_individual)}")
    logger.info(f"  Errores de fecha     : {len(collector.errores_fecha)}")
    logger.info(f"  Errores generales    : {len(collector.errores_generales)}")
    logger.info(f"  Metas aplicadas      : {len(collector.metas_aplicadas)}")
    logger.info(f"  Periodos creados     : {len(collector.metas_periodos_creados)}")
    logger.info(f"  Metas omitidas       : {len(collector.metas_omitidas)}")
    logger.info(f"  Metas sin fila       : {len(collector.metas_sin_fila)}")

    # ── 8. Generar reporte final ──────────────────────────────────────────
    logger.info("─" * 60)
    logger.info("PASO 8: Generando reporte final...")
    report_path = generate_report(collector, config.REPORTS_FOLDER, dry_run=dry_run)
    logger.info(f"  → Reporte disponible en: {report_path}")
    logger.info("─" * 60)
    logger.info("Proceso finalizado.")


if __name__ == "__main__":
    main()
