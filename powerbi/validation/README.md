# Validación Gate 4B — DAX vs Python

`validate_gate4b.py` reproduce en pandas, sobre los mismos `output/*.parquet`
que consumiría Power BI, exactamente las fórmulas documentadas en
`powerbi/dax/03_calendario.dax` y `powerbi/dax/04_digital.dax` (mismo filtro
de universo, mismas columnas, mismo orden de exclusión por política/capacidad
desconocida/fecha incompleta), y las compara contra
`MetricsEngine.query()` (Gate 3B, fuente de verdad).

No se reabre `metrics_engine.py`/`semantic_model.py` del lado "Power BI" de
la comparación: ese lado solo usa pandas + los Parquet, igual que lo haría
una medida DAX real sobre el modelo relacional.

## Ejecutar

```bash
python powerbi/validation/validate_gate4b.py
```

## Resultado (2026-08-09, sobre el Excel fuente del commit `7b54220`)

```
GATE4B_VALIDATION 8/8 PASS
```

| # | Caso (prompt Gate 4B Sec.13) | Resultado |
|---|---|---|
| 1 | Cencosud Estático — `ocupacion_calendario_pct` | PASS (10.909…%, PARTIAL) |
| 2 | AA2000 — `ocupacion_calendario_pct` | PASS (NO_APLICA — `CompletitudMaestro`=PARCIAL) |
| 3 | AA2000 — `actividad_sobre_registrados_pct` | PASS (OK) |
| 4 | Pantallas LED — `fill_rate_slots` | PASS (16.426…%, OK) |
| 5 | YPF Digital — `fill_rate_slots` | PASS (NO_APLICA — 100% del grupo bloqueado por política) |
| 6 | Pantallas LED + YPF (mixto) — `fill_rate_slots` | PASS (16.426…%, PARTIAL — YPF excluido, LED computado) |
| 7 | London Supply — `fill_rate_slots` | PASS (PARTIAL — capacidad `REQUIERE_CONFIRMACION` en 6 elementos) |
| 8 | Campañas activas — `IDCampaña` en múltiples `ElementoID` contada 1 vez | PASS (`Cargas`=9503 despliegues, ejemplo real: 1 `IDCampaña` en 2268 `ElementoID` → `DISTINCTCOUNT(IDCampaña)`=1) |

Tolerancia numérica: `1e-6`. Un `FAIL` en cualquier caso es bloqueante para
Gate 4B (indicaría que el DAX documentado no reproduce a Python) y debe
resolverse antes de avanzar a Gate 5.
