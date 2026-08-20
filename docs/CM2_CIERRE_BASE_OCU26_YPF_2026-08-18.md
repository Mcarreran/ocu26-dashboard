# CM2 — Cierre de la integración YPF y promoción de la base OCU26 (2026-08-18)

Registro de cierre de la integración de la base YPF validada (Etapa 1 + Etapa 2) dentro de OCU26, su promoción a base oficial, y la actualización local de los seis tableros TV1–TV6.

## 1. Base oficial promovida

- **Ruta**: `input/OCU26_BASE_DATOS.xlsx`
- **SHA-256**: `d3b780898af3b7ca36ccd3ca351be478cc9fd31d560876a4f8f684d48a69f050`
- **Origen**: copia exacta de `Pendientes/OCU26_YPF_INTEGRACION/output/OCU26_BASE_DATOS_CON_YPF_FINAL_2026-08-18.xlsx` (byte a byte idéntica, verificado por SHA-256 dos veces).

## 2. Respaldo PRE_YPF

- **Ruta**: `Pendientes/OCU26_YPF_INTEGRACION/backup/OCU26_BASE_DATOS_PRE_YPF_2026-08-18.xlsx`
- **SHA-256**: `2f165e12ad90c2f05963367a8e1717dcd1006e9ce669f976afa6e57470aca2cd`
- Copia del `input/OCU26_BASE_DATOS.xlsx` anterior a la promoción, confirmada byte a byte idéntica antes de sobrescribir el input.

## 3. Conteos finales

| Hoja | Filas |
|---|---|
| MAESTRO_ELEMENTOS | **5.304** |
| CAMPANAS | **15.333** |
| PARAMETROS | **23** |

0 ElementoID/CargaID vacíos o duplicados. 0 campañas con ElementoID huérfano.

## 4. Reemplazo del bloque legacy YPF (IDCampaña 10000–10009)

Autorizado por decisión de negocio explícita: las 7.790 filas antiguas (`UsuarioCarga=Migración histórica`) eran la misma actividad real ya recalculada por YPF Etapa 2 con IDs más granulares (20000+). Se retiraron íntegramente y se reemplazaron:

- **Retiradas**: 7.790 filas, preservadas completas en la auditoría, clasificadas:
  - `REEMPLAZADA_POR_YPF_ETAPA2_EXACTA`: 7.090
  - `LEGACY_EXCLUIDA_POR_REGLA`: 511 (Dermaglos Enzo, excluida por la propia fuente YPF por fechas incompletas)
  - `LEGACY_FUERA_CATALOGO`: 169 (elementos tipo `276 - TT - 45`, fuera del catálogo YPF validado)
  - `LEGACY_APIE_30943_DIGITAL_INVALIDA`: 20
- **Conservadas sin cambios**: 1.717 campañas no-YPF (verificado celda por celda).
- **Insertadas**: 13.616 filas de `BASE CAMPAÑAS` (YPF Etapa 2), 0 rechazadas.
- **Verificación**: `9.507 − 7.790 + 13.616 = 15.333` (exacto, sin forzar).

## 5. APIE 30943

Resultado final: exclusivamente **`30943-FB-1`** y **`30943-FB-2`**, cero campañas digitales.

Se retiraron de MAESTRO_ELEMENTOS 17 elementos digitales legacy (`30943-MB-1/2`, `30943-PPUNTER-1..5`, `30943-TT-1..10`) no pertenecientes al catálogo YPF Etapa 1 validado. Antes del retiro se confirmó (0 referencias) que ninguna fila de CAMPANAS los usaba. Quedan preservados completos en la auditoría FINAL, hoja `ELEMENTOS_RETIRADOS_APIE_30943`, clasificación `ELEMENTO_LEGACY_APIE_30943_INVALIDO`.

## 6. Corrección aplicada en `scripts/transform_data.py`

Bug preexistente de Gate 2 (línea ~291): la validación de passthrough de `CAMPANAS.FilaOrigen` comparaba `list(...) != list(...)`, y `NaN == NaN` es `False` en Python → falso positivo de "cambio" apenas la columna tuvo blancos legítimos a escala (13.616 filas nuevas YPF con `FilaOrigen` vacío).

```python
# Antes
if list(campanas["FilaOrigen"]) != list(campanas_raw["FilaOrigen"]):
    raise TransformError("CAMPANAS: el orden/contenido de FilaOrigen cambió")

# Después
origen_transformado = campanas["FilaOrigen"].reset_index(drop=True)
origen_original = campanas_raw["FilaOrigen"].reset_index(drop=True)
if not origen_transformado.equals(origen_original):
    raise TransformError("CAMPANAS: el orden/contenido de FilaOrigen cambió")
```

Se revisaron los otros 3 usos del mismo patrón (`MAESTRO.ElementoID`, `CAMPANAS.CargaID`, `CAMPANAS.ElementoID`) y no se modificaron: son campos obligatorios ya bloqueados por Gate 1 si vinieran vacíos, sin riesgo de NaN.

## 7. Pruebas nuevas agregadas

`tests/test_transform_data.py` — 8 pruebas nuevas, todas verdes:

1. NaN en la misma posición (passthrough real) → no bloquea
2. Valor distinto → bloquea
3. NaN vs. valor real → bloquea
4. Cambio de orden → bloquea
5. Diferencia de cantidad de filas → bloquea
6. NaN=NaN a nivel de la primitiva (complemento de 1) → no bloquea
7. Respaldo PRE_YPF sigue pasando el Gate
8. Base FINAL YPF pasa el Gate, `FilaOrigen` passthrough intacto (≥13.616 NaN)

`tests/test_transform_data.py` completo: **39/39 passed**.

## 8. Artefactos de tableros regenerados

6 HTML + 6 JSON, únicos artefactos de tablero que cambiaron (verificado por hash contra el estado previo; templates, assets y capa Parquet de Power BI quedaron bit a bit idénticos):

- `tv1.html`, `tv2.html`, `tv3.html`, `tv4.html`, `tv5.html`, `tv6.html`
- `output/tv1_data.json`, `output/tv2_data.json`, `output/tv3_data.json`, `output/tv4_data.json`, `output/tv5_data.json`, `output/tv6_data.json`

## 9. Resultados de builds y validadores

| Paso | Resultado | Exit code |
|---|---|---|
| `scripts/validate_input.py` (sobre input promovido) | `VALID_WITH_WARNINGS`, 0 errores | 0 |
| `Pendientes/OCU26_YPF_INTEGRACION/scripts/validate_ocu26_ypf_final.py` | `VALID` (30/30 controles) | 0 |
| `scripts/build_tv1_dashboard.py` … `build_tv6_dashboard.py` | Los 6 sin errores | 0 (los 6) |
| `pytest tests/test_transform_data.py` | 39 passed | 0 |
| `pytest tests/` (suite completa, post-promoción) | 422 passed, 24 failed | 1 |

## 10. Recalibración de las 19 pruebas afectadas por la promoción (2026-08-19)

Las 19 fallas causadas por la promoción (distintas de las 5 preexistentes de TV5) fueron investigadas una por una, recalculadas y **recalibradas**. Resultado final de la suite completa: **441 passed / 5 failed** — quedan exclusivamente las 5 fallas preexistentes de TV5, sin ninguna causada por la promoción.

### 10.1 — Causa raíz (dos fuentes de cambio, ambas autorizadas)

1. **Reemplazo íntegro de campañas YPF** (bloque legacy 10000–10009 → YPF Etapa 2): afecta cualquier métrica cuyo universo incluya el circuito YPF.
2. **4 filas CENCOSUD ya presentes en `FINAL_V2` y nunca antes promovidas** (KFC × 3 + MISHKA × 1, `Subcircuito=CENCOSUD`, `Medio=Estático`), confirmadas por comparación directa entre `FINAL_V2` (9.507 filas) y el input real pre-promoción (9.503 filas). Al promover la base junto con YPF, estas 4 filas —ya autorizadas en una fase anterior del proyecto (Gate 4A/4B) pero nunca antes llevadas a producción— entraron en vigencia por primera vez, afectando también métricas de circuitos no-YPF (CENCOSUD forma parte del alcance de TV1/TV3/TV4/TV6).

Ninguna de las 19 diferencias involucra pérdida de datos, error de integridad referencial, ni fórmula de negocio incorrecta.

### 10.2 — Tabla de recalibración

| Archivo de test | Prueba | Métrica | Valor PRE_YPF | Valor nuevo | Cálculo de verificación | Causa |
|---|---|---|---|---|---|---|
| test_build_tv1_dashboard.py | test_metric_status_preserved_in_production_output | kpis.estatico.status | PARTIAL | **OK** | Recalculado vía `build_tv1_data` sobre el input promovido | YPF + CENCOSUD (universo combinado) |
| test_build_tv1_dashboard.py | test_campanas_unchanged_by_apie_correction | kpis.campanas_unicas.ytd | 425 | **443** | Independiente (pandas directo, overlap+distinct IDCampaña) = 443 ✓ | YPF + CENCOSUD |
| test_build_tv2_dashboard.py | test_tv1_pipeline_still_builds_successfully | kpis.core_comercial.value | 964 | **1064** | Independiente (513 no-YPF + 551 estaciones YPF surrogate) = 1064 ✓ | YPF (catálogo: +983/−17 elementos) |
| test_build_tv3_dashboard.py | test_reconciliacion_tv1_matches_tv1_estatico_scope | reconciliacion_tv1.ocupacion_dia_based_pct | 29.4% | **35.7%** | Recalculado vía `build_tv3_data`; denominador_dias idéntico (11.842), numerador 3.482→4.227 | CENCOSUD (fila MISHKA) |
| test_build_tv3_dashboard.py | test_tv1_pipeline_still_builds_successfully | kpis.core_comercial.value | 964 | **1064** | Igual a TV2 | YPF |
| test_build_tv4_dashboard.py | test_tv1_pipeline_still_builds_successfully | kpis.core_comercial.value | 964 | **1064** | Igual a TV2 | YPF |
| test_build_tv4_dashboard.py | test_tv3_pipeline_still_builds_successfully | kpis.ocupacion_calendario.activos | 119 | **146** | Igual a TV3 | CENCOSUD |
| test_build_tv5_dashboard.py | test_production_formato_lider_julio_is_punteras | formato_lider.actual.campanas_unicas | 9 | **17** | Recalculado vía `build_tv5_data` (TV5 = 100% circuito YPF) | YPF |
| test_build_tv5_dashboard.py | test_production_formato_lider_junio_no_longer_ties (renombrado; antes `..._ties_resolved_by_elementos_activos`) | formato_lider.anterior (label / campanas_unicas / empate) | Torres / 4 / empate=True | **Punteras / 7 / empate=False** | Recalculado vía `build_tv5_data`; ya no hay empate real en junio con datos productivos | YPF (Etapa 2 reconstruyó junio) |
| test_build_tv6_dashboard.py | test_kpi_cards_unchanged_julio_scope | 5 KPI cards (julio) | 11 / 86 / 2 / GCBA(2271) / 77.6% / 8203 | **10 / 95 / 15 / TAGGIFY(7) / 67.4% / 11497** | Recalculado vía `build_tv6_data` | YPF (Agencia=Marca=Cliente=Campaña) |
| test_build_tv6_dashboard.py | test_universo_exposes_julio_and_hist_separately | universo.julio.activaciones_totales | 8203 | **11497** | Igual a kpi_cards | YPF + CENCOSUD |
| test_build_tv6_dashboard.py | test_rankings_use_accumulated_hist_scope_not_julio | ranking.marcas[0] / relación agencias | GCBA(2356) / `len(nombres) > agencias_activas` | **SEGURIDAD VIAL(3065)** / reescrito (ver 10.3) | Independiente: `compute_agencias` recalculado = 15 ✓ | YPF |
| test_build_tv6_dashboard.py | test_matriz_row_sums_match_hist_not_julio | fila GCBA en matriz | GCBA existía (2356) | **GCBA ya no existe** (reescrito, ver 10.3) | Independiente: SEGURIDAD VIAL recalculada = 3065 ✓ | YPF |
| test_build_tv6_dashboard.py | test_tv1_pipeline_still_builds_successfully | kpis.core_comercial.value | 964 | **1064** | Igual a TV2 | YPF |
| test_build_tv6_dashboard.py | test_tv3_pipeline_still_builds_successfully | kpis.ocupacion_calendario.activos | 119 | **146** | Igual a TV3 | CENCOSUD |
| test_build_tv6_dashboard.py | test_tv4_pipeline_still_builds_successfully | kpis.actividad_actual.campanas_unicas | 77 | **82** | Recalculado vía `build_tv4_data` | CENCOSUD (TV4_CIRCUITOS incluye CENCOSUD) |
| test_export_data.py | test_25_idcampana_vacias_se_conservan_en_produccion (renombrado; antes `test_26_...`) | filas con IDCampaña vacío | 26 | **25** | Recalculado vía `production_pipeline` | 1 fila del bloque legacy retirado tenía IDCampaña vacío |
| test_export_data.py | test_bridge_volumen_produccion | len(bridge) | 881.210 | **1.185.499** | Independiente (suma directa de días por CargaID, sin `build_bridge_campana_dia`) = 1.185.499 ✓ | YPF + CENCOSUD (más filas, más días) |
| test_export_data.py | test_fact_metricas_diaria_volumen_produccion | len(fact) / digital / no-digital | 573.675 / 520.648 / 53.027 | **737.475 / 677.894 / 59.581** | Independiente (pares únicos ElementoID+día por explosión de rango) = exacto en los 3 valores ✓ | YPF + CENCOSUD |

### 10.3 — Los 3 casos reescritos (no solo un cambio de valor)

- **`test_production_formato_lider_junio_no_longer_ties`** (TV5): en junio ya no hay empate real entre formatos con la base productiva (antes 4 vs 4, ahora Punteras lidera con margen propio). Se actualizó para reflejar la nueva realidad. La cobertura de la lógica de desempate en sí **se conserva íntegra** en el test sintético ya existente `test_formato_lider_tie_resolved_by_elementos_activos_not_activaciones` (datos controlados, independiente del Excel productivo, provoca un empate deliberado 3 vs 1 elementos y verifica que se resuelve por elementos activos).
- **`test_rankings_use_accumulated_hist_scope_not_julio`** (TV6): se eliminó la relación `len(nombres) > agencias_activas` (dejó de ser un invariante válido: compara magnitudes de distinto alcance/forma). Se reemplazó por 3 controles verificables e independientes: (1) `agencias_activas` coincide con la cantidad real de agencias distintas de julio, recalculada de forma independiente = 15; (2) el ranking no incluye agencias vacías/no identificadas; (3) el total de activaciones HIST es estrictamente mayor al de julio (prueba de que no están mezclados).
- **`test_matriz_row_sums_match_hist_not_julio`** (TV6): ya no busca la fila `"GCBA"` (desapareció legítimamente). Se reescribió de forma general: verifica que las filas de la matriz son exactamente las del ranking HIST (sin filas fantasma), que cada fila suma su propio total HIST, y que `"GCBA"` no aparece. Como verificación puntual, confirma que `"SEGURIDAD VIAL"` existe y recalcula su total de forma independiente desde el Excel (= 3.065).

## 11. Archivos modificados y generados hoy

**Modificados (tracked por Git):**
- `input/OCU26_BASE_DATOS.xlsx` (promoción)
- `scripts/transform_data.py` (fix Gate 2 — única línea de lógica de negocio tocada en todo el proceso)
- `tests/test_transform_data.py` (8 pruebas nuevas de regresión del fix)
- `tests/test_export_data.py` (2 pruebas recalibradas)

**Modificados (untracked, ya existían como trabajo no versionado):**
- `tests/test_build_tv1_dashboard.py`, `test_build_tv2_dashboard.py`, `test_build_tv3_dashboard.py`, `test_build_tv4_dashboard.py`, `test_build_tv5_dashboard.py`, `test_build_tv6_dashboard.py` — 17 pruebas recalibradas (ver sección 10). Ningún script de negocio (`build_tvN_dashboard.py`, `metrics_engine.py`, `semantic_model.py`, `export_data.py`) fue tocado — confirmado por hash antes/después de re-ejecutar los 6 builds.

**Generados (fuera de Git / dentro de `Pendientes/`):**
- `Pendientes/OCU26_YPF_INTEGRACION/output/OCU26_BASE_DATOS_CON_YPF_CANDIDATA_2026-08-18.xlsx`
- `Pendientes/OCU26_YPF_INTEGRACION/output/OCU26_BASE_DATOS_CON_YPF_FINAL_2026-08-18.xlsx`
- `Pendientes/OCU26_YPF_INTEGRACION/output/OCU26_AUDITORIA_INTEGRACION_YPF_2026-08-18.xlsx`
- `Pendientes/OCU26_YPF_INTEGRACION/output/OCU26_AUDITORIA_INTEGRACION_YPF_FINAL_2026-08-18.xlsx`
- `Pendientes/OCU26_YPF_INTEGRACION/backup/OCU26_BASE_DATOS_PRE_YPF_2026-08-18.xlsx` + logs/hashes de evidencia
- `Pendientes/OCU26_YPF_INTEGRACION/scripts/` — `merge_ypf_common.py`, `audit_ocu26_ypf.py`, `build_ocu26_ypf.py`, `validate_ocu26_ypf.py`, `build_ocu26_ypf_final.py`, `audit_ocu26_ypf_final.py`, `validate_ocu26_ypf_final.py`

**Regenerados (artefactos de tablero, no modificados manualmente):**
- `tv1.html` … `tv6.html`, `output/tv1_data.json` … `output/tv6_data.json`

## 12. Estado de Git

```
 M input/OCU26_BASE_DATOS.xlsx
 M scripts/transform_data.py
 M tests/test_export_data.py
 M tests/test_transform_data.py
```
Sin cambios fuera de estos cuatro archivos tracked. `Pendientes/` permanece untracked como al inicio; `tests/test_build_tv1..6_dashboard.py` también son untracked (ya lo eran antes de empezar) y fueron modificados como parte de la recalibración.

## 13. Confirmación explícita

**No hubo commit, push, PR ni deploy.** No se tocó Power BI (capa Parquet de `output/` bit a bit idéntica, ni siquiera se invocó `export_data.py`), ni el formulario, ni SharePoint, ni hosting. Ningún script de negocio fue modificado durante la recalibración (solo tests); `input/OCU26_BASE_DATOS.xlsx` no cambió al re-ejecutar los 6 builds tras recalibrar (confirmado por hash). Los 6 `tvN.html`/`tvN_data.json` sí cambiaron de hash al regenerarse, pero únicamente en el campo `meta.generado` (timestamp de cada corrida); el contenido de datos es idéntico, garantizado por input y scripts de negocio bit a bit sin cambios.

## 14. Resultado final de la suite y próximos pasos

**Suite completa: 441 passed / 5 failed** (objetivo cumplido exacto). Las 5 fallas restantes son exclusivamente las preexistentes de TV5, documentadas como deuda técnica separada (sección 10.1 original, no relacionada con YPF ni con esta promoción).

Próximos pasos sugeridos:
1. **Investigar y corregir** las 5 fallas preexistentes de TV5 (firma de `compute_mapa()`, nombres de estado `PARTIAL`/`A_VALIDAR`/`REQUIERE_CONFIRMACION`, literales hardcodeados) — independiente de la integración YPF.
2. **Revisar visualmente** TV1–TV6 en navegador (no solo vía tests) para confirmar que el rediseño de datos se ve correcto en pantalla.
3. **Formulario, Power BI y SharePoint**: siguen sin actualizar; requieren su propio pase de promoción cuando se autorice.
