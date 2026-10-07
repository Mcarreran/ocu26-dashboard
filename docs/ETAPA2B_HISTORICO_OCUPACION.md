# OCU26 — Etapa 2B / 2B.1 / 2B.2: histórico OCUPACIÓN 2026 aplicado a CAMPANAS (2026-10-07)

Rama `etapa2-ocu26`. Sin push. Scripts: `scripts/staging_import_historico.py`
(staging + aplicación 2B.1) y `scripts/cierre_historico_2b2.py` (cierre 2B.2).
Tests: `tests/test_staging_import_historico.py`, `tests/test_cierre_historico_2b2.py`.

## Base canónica

| Estado | SHA-256 | Filas CAMPANAS |
|---|---|---|
| Pre-carga (rollback en `OCU26_CORTES/CORTE_PRE_CARGA_2026-10-06/`) | `4e1ff067…adde2d` | 15.333 |
| Post 2B.1 | `a1e207df…fb0c30` | 16.331 |
| **Post 2B.2 (vigente, `input/OCU26_BASE_DATOS.xlsx`)** | `1c41aea8…33089c` | **16.365** |

Solo cambia la hoja CAMPANAS (`xl/worksheets/sheet2.xml` + `xl/tables/table2.xml`
del zip; el resto de las partes del xlsx se copian sin cambios). MAESTRO_ELEMENTOS
y PARAMETROS idénticos a la pre-carga. `FechaHoraCarga` máxima: 2026-10-07.

## Operaciones acumuladas

- 1.032 filas nuevas `UsuarioCarga = MIGRACION_OCUPACION_2026`
  (998 en 2B.1 + 34 Remeros en 2B.2), `FuenteCarga = Migración histórica`,
  `EstadoValidacion = OK`, `Estado` = Finalizada si FechaFin < fecha de carga, si no Activa.
- 94 filas existentes extendidas (FechaFin + ClaveNegocio; Estado por la regla, Reservada intacta).
- 46 filas REM de 2B.1 reasignadas al soporte posicional (ElementoID + ClaveNegocio, mismo CargaID).
- Logs: `OCU26_IMPORT_HISTORICO/STAGING_PRE_IMPORT/OCU26_HISTORICO_OPERACIONES_APLICADAS.csv` (2B.1)
  y `OCU26_HISTORICO_OPERACIONES_2B2.csv` (2B.2).

## Reglas de negocio confirmadas

- OT = IDCampaña (solo el número). `B <n>` = bonificada → IDCampaña `<n>` + nota en Observaciones.
- Remeros, regla posicional: REM-TS 1..6 = soportes Remeros Digital del maestro ordenados por
  número (1 / 3 / 5 / 6 / 8 / 10) → TS1=REM-DB-1, TS2=3, TS3=5, TS4=6, TS5=8, TS6=10.
- `PALS-3600seg` (Alsina) dado de baja → NO_IMPORTAR.
- Extensión: misma OT + elemento, mismo inicio (o continuidad contigua), FIN posterior → UPDATE.
- FIN cuyo mes solo aparece en celdas con otra OT → no atribuible → REVISAR.
- En elementos estáticos, días nuevos que pisan otra campaña cargada → REVISAR.
- `FuenteCarga` usa el vocabulario de PARAMETROS; la procedencia va en `UsuarioCarga`.

## Pendientes

458 asignaciones en REVISAR (136 OT); detalle en `REPORTE_CIERRE_ETAPA2B2.md`.
