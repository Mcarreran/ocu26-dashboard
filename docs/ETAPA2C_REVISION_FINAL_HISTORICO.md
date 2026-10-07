# OCU26 — Etapa 2C: revisión final del histórico y depuración pre-SharePoint (2026-10-07)

Rama `etapa2-ocu26`. Sin push. Script: `scripts/cierre_historico_2c.py` (reutiliza el
staging y el escritor quirúrgico de `staging_import_historico.py`). Tests:
`tests/test_cierre_historico_2c.py`. SharePoint y dashboards todavía sin tocar.

## Base canónica

| Estado | SHA-256 | Filas CAMPANAS | OT |
|---|---|---|---|
| Pre-carga (`OCU26_CORTES/CORTE_PRE_CARGA_2026-10-06/`) | `4e1ff067…adde2d` | 15.333 | |
| Post 2B.2 (copia congelada en `OCU26_IMPORT_HISTORICO/PRE_SHAREPOINT/`) | `1c41aea8…33089c` | 16.365 | 633 |
| **Post 2C (vigente, `input/OCU26_BASE_DATOS.xlsx`)** | `b5015d08…1e5ea0` | **16.382** | 633 |

Solo cambian `xl/worksheets/sheet2.xml` y `xl/tables/table2.xml`; MAESTRO_ELEMENTOS y
PARAMETROS idénticos. `FechaHoraCarga` máxima: 2026-10-07 16:01:44.

## Flujo

1. Exige el SHA post 2B.2. Recalcula el staging 2B (mismas reglas) contra la base
   **actual** y reconcilia por firma exacta (hoja, código, OT, meses, celdas, fechas)
   el universo REVISAR del staging 2B.2 (458 asignaciones, 136 OT). Si no reconcilia, aborta.
2. Evalúa cada asignación con evidencia a nivel de celda de la fuente y filas de CAMPANAS.
3. Control de coherencia: los días nuevos de una operación no pueden contradecir a la
   misma OT cargada en sus otros elementos (si la OT está cargada en otros elementos,
   debe cubrir esos días; sin otros elementos el control es neutro). En estáticos,
   ningún día nuevo puede pisar otra OT (cargada o propuesta).
4. Valida en memoria (`validar_estado_final` de 2B + alcance/ElementoID prohibidos) y,
   con `--aplicar`, congela la copia POST-2B.2, escribe en temporal, verifica y reemplaza.

## Reglas 2C

| Regla | Resultado | Condición |
|---|---|---|
| `NUEVO_PERIODO_DISJUNTO_DE_OT_YA_CARGADA` (regla 2B) | INSERT | Al recalcular contra la base actual, el otro periodo de la OT ya está cargado. |
| `EXTENSION_FIN_DEMOSTRADA` | UPDATE FechaFin + ClaveNegocio (+ Estado por regla) | Extensión 2B de UNA fila (mismo INICIO) bloqueada solo por dudas en meses que esa fila ya cubre; los meses nuevos tienen celdas propias con el mismo INICIO, FIN no decreciente y presencia hasta el mes del FIN. Nunca cambia FechaInicio. |
| `PERIODO_DISJUNTO_EXPLICITO` | INSERT | Tramo sin contacto con las filas de la misma OT + elemento; todas las celdas de sus meses son propias, con el mismo par INICIO/FIN, y hay presencia en todos ellos. |
| `SIN_OT_CUBIERTA_POR_LA_MISMA_FILA` | Cerrado sin cambios | La propia fila muestra una única OT con la misma pauta y fechas, ya cargada cubriendo esos meses. No se reconstruye ninguna OT. |
| `PRESENCIA_YA_CUBIERTA_EN_CAMPANAS` | Cerrado sin cambios | **Decisión de negocio del usuario (2026-10-07)**: si CAMPANAS ya cubre todos los meses en que la grilla muestra la OT y la fuente solo declara un INICIO anterior o un FIN posterior que la grilla no muestra, prevalece CAMPANAS (`--cerrar-presencia-cubierta`). |

Se mantienen sin excepción: celdas con varias OT, extensiones que cambiarían FechaInicio,
extensiones con varias filas candidatas, filas sin OT o sin código, elementos con
confianza BAJA/MEDIA (no se crean elementos) y solapamientos en estáticos → REVISAR.

## Resultado

| Resolución | Asignaciones |
|---|---|
| UPDATE `EXTENSION_FIN_DEMOSTRADA` | 13 (13 filas existentes) |
| INSERT `PERIODO_DISJUNTO_EXPLICITO` | 15 |
| INSERT `NUEVO_PERIODO_DISJUNTO_DE_OT_YA_CARGADA` | 2 |
| Cerrado sin cambios (`PRESENCIA_YA_CUBIERTA_EN_CAMPANAS`) | 96 |
| Cerrado sin cambios (`SIN_OT_CUBIERTA_POR_LA_MISMA_FILA`) | 1 |
| **Pendientes humanos** | **331** (108 OT): 301 EQUIPO, 30 USUARIO |

Una extensión candidata quedó en REVISAR por el control de coherencia (la OT figura en
61 elementos con un hueco entre dos periodos que la extensión cubriría).

Pendientes por motivo: celdas compartidas 85, FIN no confirmado por la grilla 80, INICIO
anterior a la presencia 65, fechas contradictorias entre bloques 44, elemento sin
equivalencia 16, solapamiento en estático 8, fila sin código 8, extensión que cambiaría
FechaInicio 8, elemento ambiguo 6, sin OT 5, extensión ambigua 3, fecha indeterminada 2,
contradice a la misma OT 1.

## Archivos (fuera del repo, `OCU26_IMPORT_HISTORICO/PRE_SHAREPOINT/`)

- `OCU26_BASE_DATOS_POST_2B2_PRE_2C.xlsx` + `MANIFEST_POST_2B2_PRE_2C.json` (copia congelada).
- `OCU26_OPERACIONES_PROPUESTAS_2C.csv` (dry-run) y `OCU26_OPERACIONES_APLICADAS_2C.csv`.
- `OCU26_PENDIENTES_PRE_SHAREPOINT.xlsx` (hoja PENDIENTES + RESUMEN).
- `OCU26_REVISION_2C.xlsx` (las 458 con su resolución, regla y evidencia).
- `RESUMEN_DRYRUN_ETAPA2C.md` y `RESUMEN_ETAPA2C.md`.

## Ejecución

```
python scripts/cierre_historico_2c.py --fuente <OCUPACION_2026.xlsx> \
    --staging-2b2 <STAGING_PRE_IMPORT/OCU26_HISTORICO_STAGING.xlsx> --salida <PRE_SHAREPOINT> \
    [--base <copia.xlsx>] [--cerrar-presencia-cubierta] [--aplicar]
```

Sin `--aplicar` es dry-run. Con `--base` apuntando a una copia se ensaya sin tocar la base
canónica (la copia congelada solo se crea sobre la base canónica). El script exige el SHA
post 2B.2: sobre la base post 2C ya no reconcilia (las 30 operaciones son YA_EXISTE).

## Pendiente

Revisión humana de `OCU26_PENDIENTES_PRE_SHAREPOINT.xlsx`; luego subida a SharePoint y
regeneración de dashboards (etapas siguientes).
