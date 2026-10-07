# OCU26 — Etapa 2B / 2B.1: histórico OCUPACIÓN 2026 aplicado a CAMPANAS (2026-10-07)

Rama `etapa2-ocu26`. Sin push. Script: `scripts/staging_import_historico.py`
(tests: `tests/test_staging_import_historico.py`).

## Base canónica

| | SHA-256 | Filas CAMPANAS |
|---|---|---|
| Pre-carga (rollback en `OCU26_CORTES/CORTE_PRE_CARGA_2026-10-06/`) | `4e1ff067…adde2d` | 15.333 |
| Post Etapa 2B.1 (`input/OCU26_BASE_DATOS.xlsx`) | `a1e207df…fb0c30` | 16.331 |

Solo cambió la hoja CAMPANAS (`xl/worksheets/sheet2.xml` + `xl/tables/table2.xml`
del zip; el resto de las partes del xlsx se copian sin cambios). La fecha
máxima de `FechaHoraCarga` pasa a 2026-10-07 (antes 2026-08-18).

## Operaciones aplicadas

- 998 INSERT (`CargaID` HIST-00023127 … HIST-00024124,
  `UsuarioCarga = MIGRACION_OCUPACION_2026`,
  `FuenteCarga = "Migración histórica - OCUPACION_2026"`, `EstadoValidacion = OK`,
  `Estado` = Finalizada si FechaFin < fecha de carga, si no Activa).
- 94 UPDATE de extensión: solo `FechaFin` y `ClaveNegocio` de filas existentes.
- Log completo: `OCU26_IMPORT_HISTORICO/STAGING_PRE_IMPORT/OCU26_HISTORICO_OPERACIONES_APLICADAS.csv`.

## Reglas de negocio confirmadas

- OT = IDCampaña (solo el número). `B <n>` = bonificada → IDCampaña `<n>` + nota en Observaciones.
- `REM-TS n` (y slots) = `REM-DB-n` si existe en el maestro (REM-DB-2 y REM-DB-4 no existen → REVISAR).
- `PALS-3600seg` (Alsina) dado de baja → NO_IMPORTAR.
- Extensión: misma OT + elemento, mismo inicio (o continuidad contigua), FIN posterior, sin contradicciones → UPDATE.
- FIN cuyo mes solo aparece en celdas con otra OT → no atribuible → REVISAR.
- En elementos estáticos, días nuevos que pisan otra campaña cargada → REVISAR.

## Pendientes

492 asignaciones en REVISAR (detalle en `REPORTE_APLICACION_ETAPA2B1.md`).
`Estado` de las filas extendidas no se tocó (36 quedan `Finalizada` con FechaFin ≥ 2026-10-07).
`FuenteCarga` nueva no figura en el vocabulario de PARAMETROS (no validado por `validate_input`).
