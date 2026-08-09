# OCU26 — Gate 4B — Modelo Power BI Desktop Free

Capa de **consumo** Power BI sobre el data mart de Gate 4A
(`output/*.parquet`). No agrega reglas de negocio nuevas: relaciones,
filtros, agregaciones, time intelligence y DAX simple sobre columnas ya
resueltas por Python (Gates 1–4A).

Compatible 100% con **Power BI Desktop Free** (Get Data > Parquet/Folder,
modelo relacional estándar, DAX estándar). No usa Power BI Service, Fabric,
Premium, ni fuentes que no sean archivos locales.

## Contenido

```
powerbi/
  README.md              # este archivo: modelo, relaciones, universos
  power_query/            # M: parámetro de ruta + 1 query por tabla
  dax/                     # medidas DAX, organizadas por familia (.dax = texto)
  model/                   # documentación de relaciones/cardinalidad
  validation/              # script Python que reproduce las medidas DAX
                           # sobre los Parquet y las compara contra MetricsEngine
```

No existe (todavía) ningún `.pbix`/`.pbip`/`.tmdl` en el repo. Estos son
artefactos de **soporte** para construir el modelo manualmente en Power BI
Desktop en minutos, no un proyecto binario ni un `.pbip` autoportante — Claude
no puede generar ni editar un `.pbix` binario.

## 1. Tablas (grain, fuente: Gate 4A)

| Tabla | Grain | Rol |
|---|---|---|
| `DIM_ELEMENTOS` | 1 fila / `ElementoID` | Dimensión de inventario |
| `FACT_CAMPANAS` | 1 fila / `CargaID` | Hecho comercial (campaña × elemento) |
| `DIM_CALENDARIO` | 1 fila / `Fecha` | Dimensión de tiempo (rango dinámico, sin huecos) |
| `BRIDGE_CAMPANA_DIA` | 1 fila / (`CargaID`, `Fecha`) vigente | Puente de vigencia temporal |
| `FACT_METRICAS_DIARIA` | 1 fila / (`ElementoID`, `Fecha`) **con actividad** | Hecho diario WIDE, sparse |

`IDCampaña` **no** es clave: es un atributo de agrupación comercial en
`FACT_CAMPANAS`. `CargaID` es la grain real (campaña × elemento).

Un día sin fila en `FACT_METRICAS_DIARIA` = actividad 0, **no** "día
inexistente". Cualquier medida "sobre todos los días del período" debe
iterar `DIM_CALENDARIO`, no asumir que `FACT_METRICAS_DIARIA` es densa.

## 2. Relaciones

Todas **single-direction** (filtro de 1 hacia N), todas **activas**, ninguna
bidireccional. Diseño en "doble estrella" con `BRIDGE_CAMPANA_DIA` como hoja
pura (nunca origina relaciones hacia otra tabla → sin loops posibles).

| # | Desde (1) | Hacia (N) | Clave | Cross-filter |
|---|---|---|---|---|
| R1 | `DIM_ELEMENTOS` | `FACT_CAMPANAS` | `ElementoID` | Single → |
| R2 | `DIM_ELEMENTOS` | `FACT_METRICAS_DIARIA` | `ElementoID` | Single → |
| R3 | `DIM_CALENDARIO` | `FACT_METRICAS_DIARIA` | `Fecha` | Single → |
| R4 | `DIM_CALENDARIO` | `BRIDGE_CAMPANA_DIA` | `Fecha` | Single → |
| R5 | `FACT_CAMPANAS` | `BRIDGE_CAMPANA_DIA` | `CargaID` | Single → |

`DIM_CALENDARIO` se marca como **tabla de fechas oficial** (Mark as Date
Table, columna `Fecha`) — es contigua y sin huecos por construcción
(`pd.date_range` en Gate 4A), condición necesaria para time intelligence
nativo (`SAMEPERIODLASTYEAR`, `DATEADD`, etc.).

### Cómo se resuelve un filtro "campañas activas por circuito/ciudad/cliente/marca en un rango"

`DIM_ELEMENTOS` (atributos de circuito/ciudad) → `FACT_CAMPANAS` (atributos
de cliente/marca/campaña, R1) → `BRIDGE_CAMPANA_DIA` (R5) ← `DIM_CALENDARIO`
(rango de fechas, R4). Cadena de 3 saltos, todos single-direction: el filtro
de cualquier combinación de esas dimensiones llega a `BRIDGE_CAMPANA_DIA`
sin ambigüedad y sin reimplementar el overlap de fechas en DAX (ya resuelto
por Gate 4A al construir la bridge).

### Relaciones que NO se crean (y por qué)

- **`DIM_ELEMENTOS` ↔ `BRIDGE_CAMPANA_DIA` directa** — redundante: ya existe
  el camino `DIM_ELEMENTOS → FACT_CAMPANAS → BRIDGE_CAMPANA_DIA` (R1+R5).
  Agregarla crea dos caminos entre las mismas tablas → ambigüedad que Power
  BI resuelve desactivando una relación de todos modos.
- **Cualquier relación con cross-filter `Both`** — evitado en todo el
  modelo. Con `BRIDGE_CAMPANA_DIA` como hoja pura, no hace falta
  bidireccionalidad para que los filtros "campañas activas por X" funcionen:
  el problema de ambigüedad/many-to-many que normalmente motiva
  bidireccional ya está resuelto por el diseño en estrella con bridge.
- **`FACT_METRICAS_DIARIA` ↔ `FACT_CAMPANAS`/`BRIDGE_CAMPANA_DIA` directa** —
  grains distintos (elemento×día resuelto vs. campaña×día vigente); no hace
  falta: comparten `DIM_ELEMENTOS`/`DIM_CALENDARIO` como conectores comunes.
- **`FACT_CAMPANAS[IDCampaña]` como relación** — no es clave (ver Sección
  0/3), nunca se usa como lado "1" de ninguna relación.

### Por qué no hay loops ni many-to-many ambiguos

`BRIDGE_CAMPANA_DIA` solo recibe relaciones (R4, R5), nunca las origina: no
hay forma de que un filtro vuelva a entrar por ella hacia otra tabla. Las
únicas dos tablas con más de una relación entrante activa son
`FACT_METRICAS_DIARIA` (R2+R3, ambas desde dimensiones puras, sin camino
alternativo entre `DIM_ELEMENTOS` y `DIM_CALENDARIO`) y `BRIDGE_CAMPANA_DIA`
(R4+R5, idem). Ningún camino cruza dos veces la misma tabla → sin loops.

## 3. Universos

`OPERATIVO_GENERAL`, `PERFORMANCE_CORE`, `COMPLETO_HISTORICO` **no** se
materializan como tablas separadas. Se implementan como **filtro DAX sobre
flags ya presentes en `DIM_ELEMENTOS`**:

| Universo | Filtro |
|---|---|
| `OPERATIVO_GENERAL` | `DIM_ELEMENTOS[IncluyeConteoGeneral] = TRUE` |
| `PERFORMANCE_CORE` | `DIM_ELEMENTOS[IncluyePerformanceCore] = TRUE` |
| `COMPLETO_HISTORICO` | sin filtro |

Implementación recomendada: **medidas separadas por universo** (no un
slicer con `USERELATIONSHIP`/`TREATAS` — sería sobre-ingeniería para Gate
4B). Ver `dax/00_universos.dax`. Vistas ejecutivas/comerciales usan
`PERFORMANCE_CORE` por defecto; APSA queda fuera del core y del conteo
general por defecto (ya resuelto en `DIM_ELEMENTOS.IncluyePerformanceCore` /
`IncluyeConteoGeneral` = `False` para `CircuitoNegocio = APSA`, Gate 3A).

## 4. MetricStatus — cómo se reconstruye sin lógica por circuito

Ningún DAX contiene `IF CircuitoNegocio = "YPF"` ni equivalentes. Toda la
semántica vive en flags de `DIM_ELEMENTOS` / `FACT_METRICAS_DIARIA` ya
resueltos por Gate 3/4A:

| Situación Python (`metrics_engine.py`) | Flag Gate 4A ya exportado | Uso en DAX |
|---|---|---|
| Circuito con política "no aplica" (ej. YPF) | `DIM_ELEMENTOS.PolicyBloqueadaSlotSeconds` | Si el 100% del grupo tiene el flag → `NO_APLICA`; si una parte → excluir esos elementos y marcar `PARTIAL` |
| `CoberturaCatalogo`/`CompletitudMaestro` ≠ `COMPLETO` (métricas estáticas: `ocupacion_calendario_pct`) | `DIM_ELEMENTOS.CoberturaCatalogo`, `CompletitudMaestro` | Si algún elemento del grupo no es `COMPLETO` → `NO_APLICA` (valor en blanco) |
| `CompletitudMaestro` ≠ `COMPLETO` (métricas digitales: `fill_rate_slots`, etc.) | `DIM_ELEMENTOS.CompletitudMaestro` | Solo agrega `PARTIAL`, el valor SÍ se calcula (asimetría real respecto al caso estático — ver Sección 8 abajo) |
| Capacidad de slots desconocida (`SlotsComerciales = REQUIERE_CONFIRMACION`) | `DIM_ELEMENTOS.CapacidadSlotsDesconocida` (+ `SlotsComercialesValor` nullable) | Excluir del numerador/denominador y marcar `PARTIAL` |
| Fechas de campaña incompletas (afecta cualquier período, no depende del rango elegido) | `DIM_ELEMENTOS.FechaIncompletaCalendario` (familia calendario) / `FechaIncompletaDigital` (familia digital) | Si algún elemento del grupo tiene el flag → `PARTIAL` |
| `SalidasVendidas` vacío dentro del período elegido (nunca se asume 0) | `FACT_METRICAS_DIARIA.HasSalidasIndeterminada = TRUE` para algún día del período | `REQUIERE_CONFIRMACION`, valor en blanco (no 0) |
| Grupo vacío / sin capacidad conocida | `DIVIDE` devuelve blanco naturalmente | `NO_APLICA` |

Precedencia de severidad (igual que `_STATUS_RANK` en Python): `OK` (0) <
`PARTIAL` (1) < `REQUIERE_CONFIRMACION` (2) < `NO_APLICA` (3). Cuando varias
condiciones aplican, **gana la más severa** — igual que
`metrics_engine._combine_status`. Ver `dax/06_metric_status.dax`.

**Nunca**: `BLANK() → 0`, `NO_APLICA → 0`, `REQUIERE_CONFIRMACION → 0`. Las
medidas usan `DIVIDE(...)` (blanco natural en división por 0) y `IF(...,
BLANK(), ...)` explícito — nunca `+0` ni `COALESCE(..., 0)` sobre estas
métricas.

## 5. Por qué el promedio diario digital se simplifica a una sola división

`metrics_engine._digital_period_activity` calcula, por `ElementoID`, la
suma de slots/segundos ocupados en el período dividida por `n_dias` del
período (promedio diario), y luego **suma esos promedios** entre los
elementos del grupo. Como `n_dias` es el mismo para todos los elementos del
grupo (es el largo del período, no de actividad por elemento):

```
Σ_elementos( Σ_días(SlotsOcupadosDia) / n_dias )  ≡  Σ_elementos_y_días(SlotsOcupadosDia) / n_dias
```

Es decir: **sumar primero y dividir una sola vez al final da el mismo
resultado matemático** que promediar por elemento y sumar. Esto es lo que
permite implementar `slots_ocupados`/`segundos_vendidos` en DAX como una
`SUM` simple sobre `FACT_METRICAS_DIARIA` filtrada al período, dividida por
`DATEDIFF` del período — sin iterar por `ElementoID` en DAX. Validado
numéricamente en `validation/validate_gate4b.py`.

## 6. Campañas activas / evitar doble conteo

`IDCampaña` puede repetirse en muchos `ElementoID`/`CargaID`. Reglas de
conteo (ver `dax/02_comercial.dax`):

| Qué se cuenta | Función |
|---|---|
| Filas campaña×elemento (`CargaID`) | `DISTINCTCOUNT(FACT_CAMPANAS[CargaID])` o `COUNTROWS(FACT_CAMPANAS)` (grain 1:1) |
| Campañas comerciales únicas (`IDCampaña`) | `DISTINCTCOUNT(FACT_CAMPANAS[IDCampaña])` — nunca `COUNTROWS` |
| Elementos por campaña | `DISTINCTCOUNT(FACT_CAMPANAS[ElementoID])`, filtrado por `IDCampaña` |
| Circuitos por campaña | `DISTINCTCOUNT(DIM_ELEMENTOS[CircuitoNegocio])` vía R1, filtrado por `IDCampaña` |
| Campañas/clientes/marcas **activas en un rango** | Sobre `BRIDGE_CAMPANA_DIA` filtrada por `DIM_CALENDARIO`, `DISTINCTCOUNT` de `IDCampaña`/`Cliente`/`Marca` (traídos de `FACT_CAMPANAS` vía R5) — nunca reimplementar el overlap de fechas en DAX |

## 7. Campos físicos (`b`/`h`/`q`/`m2`)

`b`, `h`, `m2` quedan como texto original (trazabilidad) **y** como
`BValor`/`HValor`/`M2Valor` numéricos nullable ya calculados por Gate 4A —
usar las columnas `*Valor` para cualquier análisis numérico/visual.
**`q` no tiene columna numérica derivada** (`QValor` no existe): su
semántica de negocio no está confirmada. Power BI no debe inventarle
significado — se mantiene solo como texto, igual que
`DimensionOptico`/`DimensionTotal`/`Observaciones`.

## 8. Validación contra Python

Ver `validation/validate_gate4b.py` y `validation/README.md` — 8/8 casos
comparados numéricamente contra `MetricsEngine.query()` usando exactamente
las fórmulas documentadas en `dax/`, reproducidas en pandas sobre los
mismos Parquet que consumiría Power BI (no se reabre el motor de negocio).

## 9. Qué falta para Gate 5

Con este modelo, `dax/` y la validación en verde, el modelo está listo para
que Gate 5 construya visuales/páginas (Dirección, Comercial, Ocupación,
YPF, Inventario, Calidad del dato) y las 6 TVs. Gate 4B no incluye diseño
visual ni layout de reportes.
