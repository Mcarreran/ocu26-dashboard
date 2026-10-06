# OCU26 — Etapa 2A: reglas centrales, corte pre-carga y baseline (2026-10-06)

Rama: `etapa2-ocu26` (creada desde `main` @ `9acc152`). Sin push. Sin publicación.

## 1. Tres fechas que no deben confundirse

| Concepto | Valor | Significado |
|---|---|---|
| `DATASET_FILE_VERSION` | 2026-09-23 | Versión del archivo Excel canónico |
| `DATA_LOADED_THROUGH` | 2026-08-18 | Última `FechaHoraCarga` real (no hay cargas posteriores) |
| `BASELINE_REPORT_PERIOD` | 2026-07 | Período analítico del baseline (los builders siguen con julio fijo) |

## 2. Base de entrada (input configurable)

Orden de resolución, común a todo el pipeline (`scripts/validate_input.py::resolve_input_path`):

1. argumento explícito `--file <ruta>` (o parámetro `path` de las funciones);
2. variable de entorno `OCU26_INPUT_PATH`;
3. fallback `input/OCU26_BASE_DATOS.xlsx` dentro del repositorio.

`input/`, `input_aux/`, `output/` y los formatos de datos (`*.xlsx`, `*.csv`, `*.parquet`, `*.db`, `*.sqlite`, …) están en `.gitignore`: las bases nunca se versionan (el repo es público). Plantilla de variables: `.env.example` (no se crea ni versiona un `.env` real).

> Precaución: `input/OCU26_BASE_DATOS.xlsx` sigue versionado en `main` y en las ramas históricas. Hacer checkout de esas ramas sobrescribe el archivo local ignorado. La base canónica 23-sep tiene copia en `C:\brand plus\OCU26_CORTES\CORTE_PRE_CARGA_2026-10-06\`.

## 3. Reglas de capacidad — fuente única

Todo vive en `config/business_semantics.json` y se resuelve en `scripts/semantic_model.py`. Los builders no definen tasas propias.

| Regla | Dónde |
|---|---|
| Pantalla LED 20 · Tótem 10 · Triedro 10 · Puente LED 10 · Patio de Comidas 10 | `digital_capacity.slots_profiles` |
| AA2000 digital 10 · Remeros digital 10 (formato sin perfil propio) | `digital_capacity.slots_por_circuito_digital` |
| YPF: 5 espacios base por estación (sin capacidad por elemento) | `digital_capacity.capacidad_por_estacion.YPF` |
| Segundos comerciales 72.000 (el `SegundosDia` histórico del Excel no gobierna) | `digital_capacity.default_segundos_comerciales` |
| `UNI-PUENTELED-1` → PUENTE_LED · `REM-DB-1/3/5/6/8/10` → TOTEM | `formato_negocio.elemento_formato_overrides` |
| "Patio de Comidas" (solo Medio Digital) → PATIO_COMIDAS | `formato_negocio.descripcion_keyword_rules` |
| Versión de reglas | `digital_capacity.reglas_version` = `ETAPA2A-2026-10-06` |

Precedencia de `SlotsComerciales`: override por ElementoID → capacidad por estación (YPF: vacío) → perfil por formato → slots por circuito digital → capacidad del Excel solo para circuitos listados en `use_legacy_source_if_positive_for_circuitos` (London) → `REQUIERE_CONFIRMACION`.

API central (`semantic_model`): `capacidad_espacio_digital(row)`, `categoria_espacio_digital(...)`, `capacidad_slots_numerica(...)`, `slots_por_formato(...)`, `espacios_base_por_estacion("YPF")`, `reglas_version()`, `CATEGORIAS_ESPACIO_DIGITAL`, `ETIQUETA_CATEGORIA_ESPACIO_DIGITAL`.

Consumidores: TV1 (`_espacio_capacidad_digital`), TV2 (`_espacio_capacidad_digital`), TV5 (`_espacio_capacidad_digital_tv5`, sobrecapacidad), `metrics_engine` / `export_data` / Power BI (vía `SlotsComerciales`).

## 4. YPF — ocupación por simultaneidad real

`scripts/metrics_engine.py`:

- `max_campanas_simultaneas(intervals)`
- `ocupacion_simultanea_por_estacion(asignaciones, period_start, period_end, espacios_base, station_col=..., campaign_col=..., start_col=..., end_col=...)`
- `resumen_ocupacion_simultanea(por_estacion)`

Regla: catálogo = 5 por estación; ocupación = máximo de campañas reales simultáneas en el período, **sin tope** (7 simultáneas → 7). Nunca "campañas distintas del mes" (A,B 01–15/07 + C,D 16–31/07 → 2, no 4).

Consumidores: TV1 (`_ypf_digital_ocupados_por_estacion`, llave StationKey), TV4 (`_ocupacion_periodo`, APIE), TV5 (`build_tv5_ypf_snapshot`, día de corte).

No se decidió todavía (anomalías visibles en el baseline): diferencias de universo por `RevisionMaestro`, APIE 819 (dos direcciones), 311177 vs 31177, y los 15 APIE que la llave de TV1 parte en dos.

## 5. Tests actualizados en Etapa 2A (motivo)

| Test | Cambio | Motivo |
|---|---|---|
| `test_transform_data.py` (5 tests FilaOrigen/NaN) | Restaurados desde `cierre` | Cobertura de regresión perdida en el PR limpio |
| `test_reglas_capacidad.py` (nuevo) | 14 reglas + función YPF | Pruebas centrales de las reglas de Etapa 2A |
| TV1: `uses_registered_capacity` (Tripstore/Triedro) | → regla central 10 | Regla confirmada; el Excel ya no gobierna |
| TV1: `valor_de_referencia_actual` (3.958) | → identidades + suma por regla central | Cifra de reglas anteriores |
| TV1: `sobrecarga_ypf_7_campanas` (5 ocupados) | → 7 ocupados | YPF sin tope |
| TV1: digital incluye Tripstore (= 200) / AA2000 = Tripstore / nota apertura | → elementos × 10, + otros digitales, nota nueva | Patio de Comidas y EZEPAW con regla |
| TV1: fixtures "sin regla" y "capacidad desconocida" | Pasan a formato desconocido de CENCOSUD | Patio de Comidas y AA2000 ya tienen regla |
| TV2: 990/570/220/200, 281/709, junio 380, 78/13, 56 tótems, 196, evolución fija | → suma por regla de config, identidades, conteos derivados | Cifras de reglas anteriores |
| TV2: unidad `_espacio_capacidad_digital` | Filas resueltas por semantic_model | La capacidad sale de `SlotsComerciales` |
| TV2: validación "exactamente 6 Remeros" | → sin excepción dispersa | Regla centralizada en config |
| TV2: reconciliación con TV1 | Control TV1 en memoria (tmp) | No leer `output/tv1_data.json` productivo |
| TV2: serie de evolución fija [206 … 281] | → igual a TV1 vigente + coherencia con tarjetas | Cifras de reglas anteriores |
| TV2: Aeroparque/Ezeiza 3/7/10 | → Ezeiza incluye EZEPAW005/011 | AA2000 digital = 10 |
| TV2: token legacy "321" | Retirado de la lista | Coincide por casualidad con los ocupados reales de julio |
| TV2: nuevo | Control de otra base/versión → se omite con aviso | `_reconcile_evolution_with_tv1` valida `meta.input_sha256` y `meta.reglas_version` de TV1 |
| metrics: `slots_comerciales_puente_led_13` / `totem_20` | → 10 / 10 | Regla central |
| TV4: guard `_PROTECTED_SHA256` | Sin builders TV1/2/3/6 ni tests TV1/2 | Modificados a propósito en esta etapa |
| TV4: sanity 1.294 / 2.443 / 53,0 / 107 | → identidades | Regla anterior (campañas distintas del mes) |
| TV4: nuevos | 3+3 no simultáneas = 3; 7 simultáneas = 7 | Regla definitiva YPF |
| semantic: Puente 13 / LED legacy 40 / cero → sentinel | → 10 / 20 / AA2000 = 10 | Reglas Etapa 2A, base 23-sep |
| export: 8 capacidades desconocidas | → solo London USH; YPF vacío | Reglas Etapa 2A |

## 6. Corte pre-carga y baseline

Fuera del repo: `C:\brand plus\OCU26_CORTES\CORTE_PRE_CARGA_2026-10-06\` — copia RAW de la base, `MANIFEST_CORTE_PRE_CARGA.json`, `RESUMEN_CORTE_PRE_CARGA.md`, `BASELINE_METRICAS_JULIO_2026.{json,md}` (calculado en memoria, sin generar HTML).

## 7. Pendientes

Período dinámico (julio fijo en los 6 builders y templates), decisión de `RevisionMaestro`, APIE 819, 311177/31177, llave de estación de TV1, normalización de marcas (BURGER/BURGUER KING), orquestador único de builders, deploy reproducible, importación de campañas históricas jul–sep.
