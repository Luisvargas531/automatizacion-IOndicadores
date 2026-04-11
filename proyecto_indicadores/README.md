# Proyecto Indicadores — Automatización del Maestro

Automatiza la actualización del archivo `Maestro.xlsx` a partir de múltiples archivos individuales de indicadores.

---

## Estructura del proyecto

```
proyecto_indicadores/
├── main.py              ← Punto de entrada, orquesta el flujo completo
├── config.py            ← Rutas, opciones y parámetros (editar aquí)
├── excel_reader.py      ← Lectura de Excel, fecha H2, construcción de campos
├── normalization.py     ← Normalización de llaves para comparación robusta
├── matcher.py           ← Cruce individual ↔ maestro, detección de duplicados
├── updater.py           ← Aplicación de actualizaciones (real o simulado)
├── reporting.py         ← Generación del reporte final en Excel
├── logger_utils.py      ← Configuración de logging (consola + archivo)
├── requirements.txt     ← Dependencias Python
│
├── data/
│   ├── Maestro.xlsx         ← Coloca aquí el archivo maestro
│   └── individuales/        ← Coloca aquí los archivos individuales (.xlsx)
│
├── output/
│   └── Maestro_actualizado.xlsx   ← Resultado generado automáticamente
│
├── backups/
│   └── Maestro_backup_YYYYMMDD_HHMMSS.xlsx   ← Backups automáticos
│
├── logs/
│   └── proceso_YYYYMMDD_HHMMSS.log   ← Logs detallados
│
└── reports/
    └── Reporte_Indicadores_YYYYMMDD_HHMMSS.xlsx   ← Reporte final
```

---

## Instalación

### 1. Requisitos previos
- Python 3.10 o superior
- pip

### 2. Instalar dependencias
```bash
cd proyecto_indicadores
pip install -r requirements.txt
```

---

## Configuración

Abre `config.py` y ajusta las rutas si necesitas cambiarlas:

```python
MAESTRO_PATH         = BASE_DIR / "data" / "Maestro.xlsx"
INDIVIDUALES_FOLDER  = BASE_DIR / "data" / "individuales"
OUTPUT_PATH          = BASE_DIR / "output" / "Maestro_actualizado.xlsx"
```

Opciones de comportamiento:

| Parámetro          | Valor por defecto | Descripción                                    |
|--------------------|-------------------|------------------------------------------------|
| `DRY_RUN`          | `False`           | Si `True`, simula sin guardar cambios           |
| `INSERT_NEW_RECORDS` | `False`         | Si `True`, inserta filas nuevas al maestro      |
| `MES_FORMAT`       | `"numero"`        | `"numero"` = 1..12, `"texto"` = enero..diciembre |
| `KEY_COLUMNS`      | `["Equipo","Clave","Indicador"]` | Llave de cruce (no cambiar) |

---

## Dónde colocar los archivos

| Archivo          | Ubicación                            |
|------------------|--------------------------------------|
| Maestro          | `data/Maestro.xlsx`                  |
| Individuales     | `data/individuales/*.xlsx`           |

Requisitos del **maestro**:
- Columnas obligatorias: `Fecha`, `Anio`, `Mes`, `Periodo_YYYYMM`, `Equipo`, `Clave`, `Indicador`, `Ejecucion`, `Origen`

Requisitos de cada **archivo individual**:
- Celda `H2` debe contener una fecha válida (DD/MM/YYYY, YYYY-MM-DD o DD-MM-YYYY)
- Columnas obligatorias: `Equipo`, `Clave`, `Indicador`, `Ejecucion`

---

## Ejecución

### Ejecución real (modifica el maestro)
```bash
python main.py
```

### Modo simulación (DRY-RUN — no guarda cambios)
```bash
python main.py --dry-run
```

El modo `--dry-run` también se puede activar permanentemente en `config.py`:
```python
DRY_RUN = True
```

---

## Archivos generados

| Archivo                                        | Descripción                                     |
|------------------------------------------------|-------------------------------------------------|
| `output/Maestro_actualizado.xlsx`              | Maestro con los datos actualizados              |
| `backups/Maestro_backup_YYYYMMDD_HHMMSS.xlsx`  | Copia de seguridad creada antes de modificar    |
| `logs/proceso_YYYYMMDD_HHMMSS.log`             | Log técnico detallado (DEBUG + INFO)            |
| `reports/Reporte_Indicadores_YYYYMMDD_HHMMSS.xlsx` | Reporte con 8 hojas de trazabilidad         |

### Hojas del reporte

| Hoja                   | Contenido                                              |
|------------------------|--------------------------------------------------------|
| `Resumen`              | Conteo por archivo + totales globales                  |
| `Actualizados`         | Registros actualizados (o que se actualizarían)        |
| `No_Encontrados`       | Registros sin match en el maestro                      |
| `Ambiguos`             | Registros con ambigüedad (no actualizados)             |
| `Duplicados_Maestro`   | Llaves duplicadas detectadas en el maestro             |
| `Duplicados_Individual`| Llaves duplicadas en los archivos individuales         |
| `Errores_Fecha`        | Archivos con error al leer la celda H2                 |
| `Errores_Generales`    | Cualquier otro error durante el proceso                |

---

## Lógica de negocio (resumen)

1. **Llave de cruce**: `Equipo` + `Clave` + `Indicador` — exacta, normalizada.
2. **Normalización**: se aplica solo para comparar (strip, minúsculas, sin tildes, espacios simples). Los valores originales del maestro se conservan.
3. **Actualización**: solo si hay exactamente 1 coincidencia en el maestro y 1 en el individual.
4. **No actualiza si**: hay duplicados en el maestro, duplicados en el individual, o múltiples matches.
5. **No inserta filas nuevas** por defecto (`INSERT_NEW_RECORDS = False`).
6. **Campos actualizados**: `Ejecucion`, `Fecha`, `Anio`, `Mes`, `Periodo_YYYYMM`, `Origen`.
7. **Origen**: registra el nombre del archivo individual del cual provino el dato.

---

## Supuestos asumidos

| # | Supuesto |
|---|----------|
| 1 | La hoja activa (primera) del maestro es la que contiene los datos. |
| 2 | La hoja activa (primera) de cada individual contiene los datos. |
| 3 | El encabezado de la tabla en los individuales está en la fila 1. |
| 4 | La celda H2 de cada individual contiene la fecha de realización en formato reconocible. |
| 5 | La columna `Ejecucion` existe con ese nombre exacto en todos los individuales. |
| 6 | El campo `Mes` en el maestro se almacena como número entero (configurable en `MES_FORMAT`). |
| 7 | Los archivos temporales de Excel (`~$*.xlsx`) se ignoran automáticamente. |
| 8 | La llave de negocio es únicamente: `Equipo` + `Clave` + `Indicador`. |
| 9 | Si se activa `INSERT_NEW_RECORDS = True`, las filas nuevas solo tendrán datos de las columnas de la llave + `Ejecucion` + campos de fecha + `Origen`; el resto quedará en blanco. |

---

## Plan de pruebas

### Caso 1: Ejecución normal exitosa
- **Setup**: Maestro con 5 registros. Individual con 3 registros que coinciden.
- **Esperado**: 3 actualizaciones en `Actualizados`, 0 errores.

### Caso 2: Archivo sin celda H2
- **Setup**: Individual con H2 vacía.
- **Esperado**: Archivo aparece en hoja `Errores_Fecha`. Resto del proceso continúa.

### Caso 3: Archivo sin columna `Ejecucion`
- **Setup**: Individual con columnas `Equipo`, `Clave`, `Indicador` pero sin `Ejecucion`.
- **Esperado**: Archivo aparece en `Errores_Generales`. Proceso sigue con otros archivos.

### Caso 4: Duplicados en el individual
- **Setup**: Individual con dos filas con misma combinación `Equipo + Clave + Indicador`.
- **Esperado**: Ambas filas en `Duplicados_Individual` y `Ambiguos`. No se actualiza ninguna.

### Caso 5: Duplicados en el maestro
- **Setup**: Maestro con dos filas con misma llave.
- **Esperado**: Registradas en `Duplicados_Maestro`. No se actualiza la llave duplicada.

### Caso 6: Registro no encontrado
- **Setup**: Individual con una llave que no existe en el maestro.
- **Esperado**: Registro aparece en `No_Encontrados`. El maestro no se modifica para esa fila.

### Caso 7: Dry-run
- **Setup**: Cualquier configuración válida.
- **Esperado**: Reporte con sufijo `_SIMULACION`. El maestro original NO se modifica. No se crea backup.

### Caso 8: Fecha en formato string
- **Setup**: H2 contiene `"15/03/2026"` como texto.
- **Esperado**: Se parsea correctamente → `Fecha=2026-03-15`, `Anio=2026`, `Mes=3`, `Periodo_YYYYMM=202603`.

### Caso 9: Múltiples archivos individuales
- **Setup**: 3 archivos individuales con fechas distintas.
- **Esperado**: Se procesan secuencialmente. Cada uno usa su propia fecha de H2. El resumen muestra los 3 archivos.

### Caso 10: Carpeta de individuales vacía
- **Setup**: `data/individuales/` sin archivos `.xlsx`.
- **Esperado**: Aviso en consola. Se genera reporte vacío. No se modifica el maestro.
