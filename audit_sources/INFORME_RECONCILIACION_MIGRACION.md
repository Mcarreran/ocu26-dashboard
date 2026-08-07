# Informe de Reconciliación de la Migración OCU26

**Fecha de auditoría:** 2026-08-06
**Auditor:** Reconstrucción independiente vía Claude Code, usando exclusivamente el intérprete `.venv` (Python 3.12.10, pandas 3.0.5, openpyxl 3.1.5)
**Alcance:** Verificación de contenido, no de nombres de archivo. No se modificó, corrigió, completó ni eliminó ningún dato en los tres archivos Excel.

---

## 1. Resumen ejecutivo

Se reconstruyó de forma independiente la migración `V3 auditada → OCU26_BASE_DATOS.xlsx + OCU26_PENDIENTES_VALIDACION.xlsx`, sin asumir que la migración original fue correcta y sin usar nombres de archivo como prueba de identidad.

**Resultado principal: la reconciliación cierra de forma EXACTA, con doble control independiente, tanto para el maestro de elementos como para las campañas.** No se detectó ninguna fila "sin destino", ningún ElementoID huérfano, ninguna duplicación en la base final, y ningún registro de `PENDIENTES` que no corresponda a una situación real y verificable en la V3.

El único hallazgo material es que **cuatro columnas del maestro (`TipoInventario`, `AplicaCantidad`, `RevisionMaestro`, `Proveedor`) están vacías tanto en la V3 auditada como en la base final** — no porque se hayan perdido en la migración, sino porque en la V3 son fórmulas de Excel cuyo **valor cacheado nunca fue calculado/guardado** (el archivo se guardó sin recalcular). La lógica de esas fórmulas es determinística y está documentada íntegramente en la Sección 7; puede reconstruirse con precisión antes de avanzar con Power BI.

Todas las demás diferencias observadas (formato de fecha en `FechaFin`, derivación de `ModalidadPauta`/`TipoCargaDeclarado`/`EstadoValidacion`/`ClaveNegocio`/`CargaID`) son transformaciones documentadas y verificables, no pérdidas de datos.

---

## 2. Identidad de archivos

No se asumió identidad por nombre. Se calculó SHA-256, tamaño, hojas y tablas de Excel de cada archivo tal como existe en el repositorio.

| Archivo | Ruta | Tamaño (bytes) | SHA-256 |
|---|---|---|---|
| `Base_ocupacion_26_FINAL_CON_YPF_AUDITADA_V3.xlsx` | `audit_sources/` | 9.954.204 | `b6b2481f4c596dcecdcea614ce4cbd3be234a1a0c0c9a246bfbe2bef1b80f3c8` |
| `OCU26_BASE_DATOS.xlsx` | `input/` | 1.591.230 | `2f165e12ad90c2f05963367a8e1717dcd1006e9ce669f976afa6e57470aca2cd` |
| `OCU26_PENDIENTES_VALIDACION.xlsx` | `audit_sources/` | 16.890 | `a0091dded07bf075e39c1c7ba4f710471f71c53bb8965bcdc192beb52b2651f1` |

Los tres nombres reales encontrados coinciden exactamente con los nombres esperados por el prompt de auditoría. No se detectaron discrepancias de nombre.

**Hojas y tablas — V3 auditada (7 hojas, sin Tablas de Excel estructuradas; todos los datos están en rangos planos):**

| # | Hoja | Dimensión usada | Filas con dato real |
|---|---|---|---|
| 1 | BASE_MAESTRA_ELEMENTOS | A1:Z10000 (rango físico) | 4.384 |
| 2 | BASE_CAMPAÑAS | A1:AZ12000 (rango físico) | 9.506 |
| 3 | CONTROL_DISPONIBILIDAD | A1:AN10004 (rango físico) | 15 |
| 4 | AUX_CASCADA | A1:CAF2056 (rango físico) | 2.055 |
| 5 | LOG_CORRECCIONES | A1:H28 (rango físico) | 26 |
| 6 | REVISION_IDS_ALERTADOS | A1:K26 (rango físico) | 24 |
| 7 | VALIDACION_FINAL | A1:D48 (rango físico) | 46 |

**Hojas y tablas — OCU26_BASE_DATOS.xlsx (3 hojas, con Tablas de Excel estructuradas):**

| Hoja | Tabla | Rango |
|---|---|---|
| MAESTRO_ELEMENTOS | `tblElementos` | A1:Z4339 (4.338 filas de datos) |
| CAMPANAS | `tblCampanas` | A1:AD9504 (9.503 filas de datos) |
| PARAMETROS | `tblParametros` | A1:B24 |

**Hojas y tablas — OCU26_PENDIENTES_VALIDACION.xlsx (1 hoja):**

| Hoja | Tabla | Rango |
|---|---|---|
| PENDIENTES | `tblPendientes` | A1:M158 (157 filas de datos) |

---

## 3. Estructura de la V3 (conteos reproducidos directamente desde los datos, no desde VALIDACION_FINAL)

**BASE_MAESTRA_ELEMENTOS**
- Filas con datos: **4.384**
- ElementoID vacíos: **0**
- ElementoID únicos: **4.356**
- ElementoID duplicados (valores que aparecen más de una vez): **22 IDs distintos, involucrando 50 filas en total**
- Campos: `TipoCatalogo, Ciudad, Medio, CircuitoDashboard, Subcircuito, Ubicacion, ElementoID, Nivel, Descripcion, Resolucion, DimensionOptico, DimensionTotal, Observaciones, Material, TipoInstalacion, Original, b, h, q, m2, CapacidadSlotsReel, SegundosDia, TipoInventario, AplicaCantidad, RevisionMaestro, Proveedor` (26 columnas)

**BASE_CAMPAÑAS**
- Filas con datos: **9.506**
- ElementoID vacíos: **0**
- IDCampaña vacíos: **26** (9.480 con IDCampaña)
- IDCampaña únicos: **440**
- Campos: 52 columnas, incluyendo controles internos de la V3 (`AlertaCantidad`, `SlotsEquivVendidos`, `FillRateReal`, `Lista_*`, `FechaFinCalc`, etc.) que **no** tienen equivalente 1:1 en la base final por diseño (son cálculos de control de la planilla origen, no datos de negocio a migrar).

**Hojas de control/auditoría (no son fuente de filas de migración):**
- `CONTROL_DISPONIBILIDAD`: panel de KPIs/disponibilidad, 15 filas de encabezados y fórmulas resumen.
- `AUX_CASCADA`: soporte técnico de listas desplegables en cascada de Excel (rangos con nombre `rng_00001`…`rng_02055`), sin relación con el contenido de negocio.
- `LOG_CORRECCIONES`: 26 filas, bitácora manual de correcciones aplicadas por el auditor humano antes de esta auditoría.
- `REVISION_IDS_ALERTADOS`: 24 filas, bitácora de revisión de IDs.
- `VALIDACION_FINAL`: 46 filas de texto narrativo. **No se usó como evidencia primaria**, conforme a la instrucción; todos los conteos de este informe se recalcularon directamente de `BASE_MAESTRA_ELEMENTOS` y `BASE_CAMPAÑAS`.

---

## 4. Reconciliación del maestro (BASE_MAESTRA_ELEMENTOS → MAESTRO_ELEMENTOS + PENDIENTES)

Clave usada: `ElementoID`, con verificación adicional por número de fila física de Excel (`FilaOrigen`) para distinguir filas idénticas de filas conflictivas dentro de un mismo ID duplicado.

| Categoría | Filas fuente (V3) | ElementoID afectados | Destino | Evidencia |
|---|---|---|---|---|
| **MIGRADA_DIRECTAMENTE** | 4.334 | 4.334 (ID único en V3, sin conflicto) | `MAESTRO_ELEMENTOS` | Comparación campo por campo: 100% idénticas (ver Sección 6) |
| **MIGRADA_DIRECTAMENTE (representante de grupo idéntico)** | 4 | 4 (una fila por grupo) | `MAESTRO_ELEMENTOS` | Fila conservada de cada grupo `DUPLICADO_IDENTICO_CONSOLIDADO` |
| **DUPLICADO_IDENTICO_CONSOLIDADO** | 5 | 4 IDs (`EZEPAW005` ×2, `3327 - TT - 1`, `67 - TT - 2`, `938 - TT - 1`) | `PENDIENTES` (EstadoResolucion=CONSOLIDADO); **no** duplica en `MAESTRO_ELEMENTOS` | `PENDIENTES` filas con `TipoPendiente=DUPLICADO_IDENTICO_CONSOLIDADO`, `Motivo="Duplicado idéntico; se conservó una sola fila"` |
| **ID_DUPLICADO_CONFLICTIVO_EN_PENDIENTES** | 41 | 18 IDs (datos divergentes entre copias, ej. `UNI-ASC -1`, `UNI-ASC -2`, `PLOM-CB-1`, …) | `PENDIENTES` únicamente; **excluidas por completo** de `MAESTRO_ELEMENTOS` | `PENDIENTES` filas con `TipoPendiente=ID_DUPLICADO_CONFLICTIVO`, `EstadoResolucion=PENDIENTE` |
| **EXCLUIDA_CON_JUSTIFICACION** | 0 (ya cubierto arriba) | — | — | — |
| **SIN_DESTINO** | **0** | — | — | Verificado por dos métodos independientes (Sección 12) |
| **OTRA_TRANSFORMACION_DOCUMENTADA (overlay informativo)** | 10 (subconjunto de las 4.334+4 ya migradas) | 10 IDs digitales con `CapacidadSlotsReel=0` | Migradas normalmente a `MAESTRO_ELEMENTOS`, **además** marcadas en `PENDIENTES` con `TipoPendiente=PENDIENTE_CAPACIDAD` | No es una categoría de destino distinta; es una alerta de calidad de dato superpuesta a una fila ya migrada |

**Cierre de la ecuación:**

```
4.384 (V3) = 4.334 (directa, ID único) + 4 (representante de grupo idéntico)
           + 5 (consolidada, no migra)  + 41 (conflictiva, excluida)
           = 4.384   ✔ EXACTO
```

```
4.338 (MAESTRO_ELEMENTOS final) = 4.334 + 4   ✔ EXACTO (coincide con el conteo real de filas de la hoja)
```

**Controles adicionales:**
- ElementoID vacíos en `MAESTRO_ELEMENTOS` final: **0**
- ElementoID duplicados en `MAESTRO_ELEMENTOS` final: **0**
- Todo ElementoID de la V3 que no llegó al maestro final tiene una fila de `PENDIENTES` que lo explica (41/41 = 100%). Ningún ID desaparece silenciosamente.
- Nota de auditoría menor (no es error): las filas 551 y 552 (copias redundantes de `EZEPAW005`) llevan **dos** motivos simultáneos en `PENDIENTES` (`DUPLICADO_IDENTICO_CONSOLIDADO` y `PENDIENTE_CAPACIDAD`). Esto es válido —son dos observaciones distintas sobre la misma fila física descartada— pero conviene documentarlo para que no se lea como inconsistencia al auditar `PENDIENTES` por conteo de filas simple.

---

## 5. Reconciliación de campañas (BASE_CAMPAÑAS → CAMPANAS + PENDIENTES)

Método principal: `FilaOrigen` (número de fila física de la V3, presente en el 100% de las filas de `CAMPANAS` final: 9.503/9.503, sin duplicados). Método de verificación secundaria: clave lógica `IDCampaña + ElementoID + FechaInicio + FechaFin + HoraInicio + HoraFin` (confirma la misma partición, ver Sección 12).

| Categoría | Filas fuente (V3) | Destino | Evidencia |
|---|---|---|---|
| **MIGRADA_DIRECTAMENTE** | 9.405 | `CAMPANAS`, `EstadoValidacion="OK"` | Sin fila correspondiente en `PENDIENTES` |
| **MIGRADA_CON_TRANSFORMACION_DOCUMENTADA** | (subconjunto transversal, ver Sección 6) | `CAMPANAS` | `FechaFin`, `ModalidadPauta`, `TipoCargaDeclarado`, `ClaveNegocio`, `CargaID` se derivan/reformatean para todas las filas migradas; no cambia el conteo de filas |
| **PENDIENTE_HISTORICO** | 80 | `CAMPANAS` (migrada) **y** `PENDIENTES` | `EstadoValidacion="PENDIENTE_HISTORICO"`, motivo `FechaInicio vacía` / `FechaInicio vacía; Estado vacío` |
| **PENDIENTE_DUPLICADO** | 11 | `CAMPANAS` (migrada) **y** `PENDIENTES` | Motivo "Duplicado exacto de activación" |
| **PENDIENTE_COLISION** | 7 | `CAMPANAS` (migrada) **y** `PENDIENTES` | Motivo "Colisión de ClaveNegocio con datos diferentes" |
| **PENDIENTE_ID_CONFLICTIVO** | 3 | **Excluida de `CAMPANAS`**, sólo en `PENDIENTES` | Motivo "ElementoID conflictivo excluido del maestro" (son las 3 activaciones que apuntaban a un ElementoID de la categoría `ID_DUPLICADO_CONFLICTIVO` del maestro) |
| **EXCLUIDA_CON_JUSTIFICACION** | 3 (=PENDIENTE_ID_CONFLICTIVO) | — | — |
| **SIN_DESTINO** | **0** | — | Verificado por dos métodos independientes |

**Cierre de la ecuación:**

```
9.506 (V3) = 9.405 (directa, sin flag) + 80 (histórico, migrada+flag) + 11 (duplicado, migrada+flag)
           + 7 (colisión, migrada+flag) + 3 (id conflictivo, excluida)
           = 9.506   ✔ EXACTO
```

```
9.503 (CAMPANAS final) = 9.405 + 80 + 11 + 7   ✔ EXACTO
```

**Controles adicionales:**
- `FilaOrigen` en `CAMPANAS` final que no corresponde a ninguna fila real de la V3: **0**
- `FilaOrigen` en `PENDIENTES` (tipo CAMPANA) que no corresponde a fila real de la V3: **0**
- ElementoID en `CAMPANAS` final sin correspondencia en `MAESTRO_ELEMENTOS` final (huérfanos): **0** de 2.935 ElementoID distintos usados en campañas.
- `EstadoValidacion` en `CAMPANAS` final es 100% consistente con la pertenencia a `PENDIENTES`: **ninguna** fila con `EstadoValidacion="OK"` aparece en `PENDIENTES`, y **todas** las filas con `EstadoValidacion` distinto de "OK" tienen su contraparte exacta en `PENDIENTES`. No se detectó ningún caso de "excluida y migrada como OK" simultáneamente.

---

## 6. Tabla de transformaciones por campo

### MAESTRO_ELEMENTOS (26 campos, comparación exacta campo por campo sobre las 4.338 filas finales)

| Campo | Origen en V3 | Tipo de transformación | Coincidencia V3↔Final |
|---|---|---|---|
| TipoCatalogo, Ciudad, Medio, CircuitoDashboard, Subcircuito, Ubicacion, ElementoID, Nivel, Descripcion, Resolucion, DimensionOptico, DimensionTotal, Observaciones, Material, TipoInstalacion, Original, b, h, q, m2 | Columna homónima en `BASE_MAESTRA_ELEMENTOS` | **COPIA_DIRECTA** (valor estático, no fórmula) | 100% idéntico, 0 diferencias en las 4.338 filas comparadas |
| CapacidadSlotsReel, SegundosDia | Columna homónima (valor estático) | **COPIA_DIRECTA** | 100% idéntico valor a valor (incluye la mezcla int/texto, ver Sección 8) |
| TipoInventario | Columna homónima, pero **es fórmula** en V3 (`=IF($G="","",IF(OR(SEARCH("CARRO"/"STOPPER"/"FLOORGRAPHIC"/"CUBRE"/"ALARMA",…)),"Flexible gráfico",IF(Medio="Digital","Digital","Físico estático")))`) | **VACIA_POR_FALTA_DE_DATO_CACHEADO** — la fórmula nunca fue calculada/guardada en la V3 (`data_only=True` devuelve `None` en el 100% de las 4.384 filas de la V3, no sólo en la base final) | 0/4.384 en V3, 0/4.338 en final — **ambas vacías por igual**, no hay pérdida diferencial |
| AplicaCantidad | Fórmula `=IF($G="","",IF(TipoInventario="Flexible gráfico","SI","NO"))` | Igual que arriba | 0/4.384 en V3, 0/4.338 en final |
| RevisionMaestro | Fórmula de control cruzado (duplicados / capacidad / coherencia) | Igual que arriba | 0/4.384 en V3, 0/4.338 en final |
| Proveedor | Columna existe en V3 pero **nunca tuvo datos** (ni fórmula ni valor en ninguna de las 4.384 filas) | **VACIA_INTENCIONAL / VACIA_EN_ORIGEN** | 0/4.384 en V3, 0/4.338 en final — campo vacío ya en el origen, no es pérdida de migración |

### CAMPANAS (30 campos)

| Campo | Origen en V3 | Tipo de transformación | Observación |
|---|---|---|---|
| ElementoID, IDCampaña, Campaña, Cliente, Marca, Agencia, Proveedor, FechaInicio, Estado, DuracionSpotSeg, SalidasVendidas, CANJE, PROGRAMATICA, Observaciones | Columna homónima en `BASE_CAMPAÑAS` | **COPIA_DIRECTA** | 100% idéntico tras normalizar NaN/NaN (0 diferencias reales en 9.503 filas) |
| FechaFin | Columna `FechaFin` de V3 | **DERIVADA** (reformateo datetime→fecha) + regla especial: cuando V3 contiene el texto `"INDET"` (5 filas), se traduce a `FechaFin` vacía + `FechaIndefinida="Si"` | 5 filas usan la regla INDET→FechaIndefinida; el resto (9.498) son el mismo valor de fecha, sólo cambia el formato de serialización |
| FechaIndefinida | No existe como columna en V3; se deriva del valor `"INDET"` en `FechaFin` de V3 | **NUEVA_DERIVADA** | 5 "Si", 9.443 "No", 55 vacíos (heredados de `FechaFin` vacío en V3) |
| CantidadUnidades | Columna homónima en V3 | **COPIA_DIRECTA** | Vacía en el 100% de ambas (V3: 0/9.506 con dato; final: 0/9.503) — campo vacío ya en el origen |
| TipoCargaDeclarado | Deriva 1:1 de `Medio` de V3 (`Digital→Digital`, `Estático→Estático`) | **DERIVADA**, regla verificada sin excepciones (crosstab 9.141/362, cero cruces) | 100% consistente |
| ModalidadPauta | No existe con datos en V3 (columna `ModalidadPauta` de V3 está 100% vacía, 0/9.506) | **NUEVA_DERIVADA** — se calcula en la base final (`"Slot / Reel normal"` 9.141 filas, `"No aplica"` 362 filas); el conteo de 362 coincide exactamente con las filas sin `DuracionSpotSeg`/`SalidasVendidas` | No es pérdida: el campo de V3 estaba vacío; el valor final es un cálculo nuevo, coherente con otros campos |
| HoraInicio, HoraFin | Columna homónima en V3 | **COPIA_DIRECTA** | V3 sólo tenía 3/9.506 filas con dato real; final conserva exactamente esas 3/9.503. No hay pérdida — el campo ya era casi enteramente vacío en origen |
| TipoExclusividad | No hay columna homóloga con datos en V3 (`EsExclusividad` existe pero está 100% vacía: 0/9.506) | **VACIA_EN_ORIGEN** | 0/9.503 en final — coherente con el origen, no es pérdida |
| CargaID | No existe en V3 | **NUEVA_CONTROL** — identificador secuencial `HIST-00000001`…`HIST-00009506` asignado en el orden original de la V3 | Útil para trazabilidad, no reemplaza a `FilaOrigen` |
| ClaveNegocio | No existe en V3 como campo único; es una clave compuesta (`IDCampaña|…|FechaInicio|FechaFin||`) | **NUEVA_DERIVADA** (clave de negocio para detectar colisiones/duplicados) | Base de los tipos `PENDIENTE_COLISION`/`PENDIENTE_DUPLICADO` |
| FechaHoraCarga | No existe en V3 | **VACIA_INTENCIONAL** (reservada para cargas futuras vía formulario/proceso operativo) | 100% vacía (9.503/9.503) — coherente con `UsuarioCarga`/`FuenteCarga` fijos en `"Migración histórica"` |
| UsuarioCarga, FuenteCarga | No existen en V3 | **NUEVA_CONTROL**, valor fijo `"Migración histórica"` para las 9.503 filas | Marca de procedencia de la carga |
| EstadoValidacion | No existe en V3; se deriva de la pertenencia a `PENDIENTES` | **NUEVA_CONTROL/DERIVADA** — 100% consistente con `PENDIENTES` (ver Sección 5) | `OK` 9.405, resto = TipoPendiente |
| ObservacionValidacion | No existe en V3 | **NUEVA_DERIVADA**, texto explicativo sólo en las 80 filas `PENDIENTE_HISTORICO` | 9.423/9.503 vacía (coherente: sólo aplica a filas con problema de completitud histórica) |
| PROGRAMATICA | Columna homónima en V3 | **COPIA_DIRECTA** | 100% idéntico (9.351 No / 112 Si / 43 vacío en ambos, ajustado por las 3 filas excluidas) |

**No se detectó ningún campo clasificable como `POSIBLE_PERDIDA_DE_MIGRACION`** entre los campos con datos reales en la V3: todo campo con dato en origen tiene el mismo dato (o una transformación explícita y verificable) en el destino. Los únicos campos vacíos en el destino ya estaban vacíos en el origen, **excepto** `TipoInventario`, `AplicaCantidad` y `RevisionMaestro`, que están vacíos en **ambos** por la misma causa (fórmula sin cachear), no por pérdida durante la migración — ver Sección 7 para el detalle y la vía de recuperación.

---

## 7. TipoInventario / AplicaCantidad — investigación de causa raíz

Se inspeccionó la V3 con `openpyxl` en dos modos: `data_only=False` (texto de la fórmula) y `data_only=True` (valor cacheado por la última vez que Excel recalculó y guardó el archivo).

1. **¿En la V3 eran fórmulas?** Sí, el 100% de las 4.384 filas de `TipoInventario`, `AplicaCantidad` y `RevisionMaestro` contienen una fórmula (no un valor estático).
2. **¿Qué resultados visibles producían?** Ninguno recuperable: el valor cacheado (`data_only=True`) es `None` en el 100% de las filas. El archivo `V3_AUDITADA` fue guardado (probablemente por una herramienta que no recalcula, como openpyxl, o con el cálculo de fórmulas desactivado) sin persistir el resultado calculado.
3. **¿Es posible recuperar esos resultados directamente de la V3?** No leyendo el archivo tal cual —el caché está vacío—, pero **sí es posible recalcularlos**, porque la lógica de la fórmula es 100% determinística y depende únicamente de columnas que sí tienen datos completos (`ElementoID`, `Medio`, `Ubicacion`, `Subcircuito`, `Descripcion`).
4. **¿La pérdida se produjo al convertir fórmulas a valores?** No exactamente: no hubo "conversión con pérdida", porque nunca existió un valor calculado que perder. La cadena completa (V3 → base final) heredó una columna que **nunca tuvo contenido materializado**, ni siquiera en el archivo de auditoría.
5. **¿Puede reconstruirse la clasificación de forma determinística?** Sí, con la fórmula literal extraída de la V3:
   - `TipoInventario = "" si ElementoID vacío`
   - `"Flexible gráfico"` si `Ubicacion|Descripcion|Subcircuito|CircuitoDashboard` contiene (mayúsculas) `CARRO`, `STOPPER`, `FLOORGRAPHIC`, `CUBRE` o `ALARMA`
   - si no, `"Digital"` cuando `Medio="Digital"`
   - si no, `"Físico estático"`
   - `AplicaCantidad = "SI"` si `TipoInventario="Flexible gráfico"`, si no `"NO"` (vacío si ElementoID vacío)
   - `RevisionMaestro`: `"Revisar: ElementoID duplicado"` si el ID se repite; `"Revisar: capacidad digital en 0"` si es Digital y `CapacidadSlotsReel=0`; `"Revisar: flexible no usa slots"` si es Flexible gráfico y `CapacidadSlotsReel≠0`; si no, `"OK"`.
6. **¿Qué regla exacta utilizaba la V3?** La descrita en el punto anterior, extraída literalmente de la celda `W2` de `BASE_MAESTRA_ELEMENTOS` (`=IF($G2="","",IF(OR(ISNUMBER(SEARCH("CARRO",...)),...),"Flexible gráfico",IF($C2="Digital","Digital","Físico estático")))`).
7. **¿Permite diferenciar Digital / Estático fijo / Estático flexible-por-cantidad?** Sí, exactamente esas tres categorías son el dominio de salida de la fórmula (`Digital`, `Físico estático`, `Flexible gráfico`), y `AplicaCantidad` marca cuáles de esas requieren gestión por cantidad en vez de por slot/tiempo.
8. **¿Son necesarias para los tableros de Power BI y HTML?** Sí — son la base de segmentación de inventario (digital vs. estático vs. flexible) que previsiblemente alimenta KPIs de fill-rate, disponibilidad y capacidad. Con las cuatro columnas vacías, **cualquier medida de Power BI que agrupe o filtre por tipo de inventario quedará en blanco o incorrecta** hasta que se recalculen.

**No se rellenó ningún dato en esta auditoría**, conforme a la instrucción. Se documenta la regla para que la recuperación, si se decide, sea un cálculo determinístico y auditable, no una interpretación manual.

---

## 8. CapacidadSlotsReel / SegundosDia

Auditoría de tipo de dato, fila por fila, en ambos archivos (`openpyxl`, `data_only=True`):

| Métrica | V3 (BASE_MAESTRA_ELEMENTOS) | Final (MAESTRO_ELEMENTOS) |
|---|---|---|
| Almacenado como `int` | 2.801 | 2.755 |
| Almacenado como texto numérico (`"20"` / `"100800"`) | 1.583 | 1.583 |
| Vacío | 5.615 (filas de relleno del rango físico, más allá de los datos reales) | 0 |
| No convertible a número | 0 | 0 |
| Mínimo / Máximo (CapacidadSlotsReel) | 0 / 40 | 0 / 40 |
| Mínimo / Máximo (SegundosDia) | 0 / 100.800 | 0 / 100.800 |
| Valores de texto distintos | `CapacidadSlotsReel`: únicamente `"20"` (1.583 veces); `SegundosDia`: únicamente `"100800"` (1.583 veces) | Idéntico |

**Comparación valor a valor por ElementoID (4.338 pares comparados):**
- Diferencias de **tipo solamente** (mismo valor numérico, distinto tipo Python): **0**
- Diferencias de **valor real**: **0**

**Conclusión: "20" (texto) y 20 (entero), así como "100800" y 100800, son única y exclusivamente diferencias de tipo de almacenamiento — nunca de valor.** Además, la mezcla int/texto **ya existía idéntica en la V3 auditada**; la migración no la introdujo ni la agravó, solo la heredó sin alterarla (2.755 + 1.583 = 4.338, exactamente el total de filas finales).

**Regla de conversión segura propuesta (no ejecutada):** `pd.to_numeric(columna, errors="raise")` o equivalente `CInt`/`VALUE()` en Power Query, aplicado en la capa de carga de Power BI (no sobre el Excel origen), ya que no hay valores no convertibles ni pérdida de precisión esperada (todos los 1.583 valores de texto son enteros simples).

---

## 9. Auditoría completa de PENDIENTES

`OCU26_PENDIENTES_VALIDACION.xlsx`, hoja `PENDIENTES`, tabla `tblPendientes`, **157 filas** de datos, 13 columnas: `TipoRegistro, HojaOrigen, FilaOrigen, ElementoID, IDCampaña, Campaña, TipoPendiente, CamposAfectados, ValorActual, Motivo, Impacto, AccionManualRequerida, EstadoResolucion`.

| TipoRegistro | TipoPendiente | EstadoResolucion | Filas |
|---|---|---|---|
| ELEMENTO | DUPLICADO_IDENTICO_CONSOLIDADO | CONSOLIDADO | 5 |
| ELEMENTO | ID_DUPLICADO_CONFLICTIVO | PENDIENTE | 41 |
| ELEMENTO | PENDIENTE_CAPACIDAD | PENDIENTE | 10 |
| CAMPANA | PENDIENTE_COLISION | PENDIENTE | 7 |
| CAMPANA | PENDIENTE_DUPLICADO | PENDIENTE | 11 |
| CAMPANA | PENDIENTE_HISTORICO | PENDIENTE | 80 |
| CAMPANA | PENDIENTE_ID_CONFLICTIVO | PENDIENTE | 3 |
| **Total** | | | **157** |

- **HojaOrigen**: únicamente `BASE_MAESTRA_ELEMENTOS` y `BASE_CAMPAÑAS` (consistente con las dos hojas de datos reales de la V3).
- **FilaOrigen**: el 100% de los valores corresponde a una fila física real de la V3 (0 huérfanos, verificado por intersección de conjuntos). 153 valores de `FilaOrigen` son únicos; 2 valores (551, 552) aparecen dos veces porque esas filas físicas acumulan dos motivos distintos (duplicado idéntico + capacidad en cero) — explicado en la Sección 4, no es un error.
- **ElementoID / IDCampaña / Campaña**: presentes cuando corresponde al tipo de registro (ELEMENTO usa ElementoID; CAMPANA usa IDCampaña/Campaña); consistentes con los valores reales encontrados en la V3 para esas mismas filas.
- **Cruce de contradicción solicitado** ("¿alguna fila aparece simultáneamente como excluida de la base operativa Y migrada como OK sin explicación?"): **No se encontró ningún caso.** Se verificó que:
  - Las 41 filas `ID_DUPLICADO_CONFLICTIVO` (maestro) **no** están en `MAESTRO_ELEMENTOS` final.
  - Las 3 filas `PENDIENTE_ID_CONFLICTIVO` (campañas) **no** están en `CAMPANAS` final.
  - Todas las demás filas de `PENDIENTES` **sí** están en la base final, pero con `EstadoValidacion`/estado coherente con el motivo documentado (no dicen "OK").

---

## 10. Campos que parecen haberse perdido

Sólo un grupo de campos entra en esta categoría, y con causa raíz identificada (no es pérdida silenciosa):

| Campo | Hoja | ¿Perdido? | Explicación |
|---|---|---|---|
| TipoInventario | MAESTRO_ELEMENTOS | Vacío en ambos lados | Fórmula sin caché en la V3 (Sección 7); recuperable determinísticamente |
| AplicaCantidad | MAESTRO_ELEMENTOS | Vacío en ambos lados | Ídem |
| RevisionMaestro | MAESTRO_ELEMENTOS | Vacío en ambos lados | Ídem; además es un campo de control interno de la V3, cuestionable si debe migrar tal cual a la base operativa (ver pregunta 7) |

No se identificó ningún otro campo con datos reales en la V3 que no haya llegado íntegro a la base final.

---

## 11. Campos realmente vacíos en el origen (no son pérdida)

| Campo | Hoja | Vacío en V3 | Vacío en Final | Interpretación |
|---|---|---|---|---|
| Proveedor | MAESTRO_ELEMENTOS | 4.384/4.384 | 4.338/4.338 | Nunca tuvo dato; reservado |
| CantidadUnidades | CAMPANAS | 9.506/9.506 | 9.503/9.503 | Nunca tuvo dato; reservado para cargas futuras basadas en cantidad |
| TipoExclusividad / EsExclusividad | CAMPANAS | 9.506/9.506 | 9.503/9.503 | Nunca tuvo dato en la V3 |
| HoraInicio / HoraFin | CAMPANAS | 9.503/9.506 vacío (sólo 3 con dato) | 9.500/9.503 vacío (mismas 3 con dato) | Casi enteramente vacío ya en origen; se conservaron exactamente las 3 filas con dato real |
| FechaHoraCarga | CAMPANAS | No existe en V3 | 9.503/9.503 vacío | Campo nuevo de control operativo, reservado para el proceso de carga futuro (formulario), no aplica a la migración histórica |

---

## 12. Controles dobles realizados (Fase 10)

Para cada punto exigido se aplicaron dos métodos genuinamente independientes: **Método A** = pandas (`merge`, `groupby`, `crosstab`, comparación vectorizada); **Método B** = Python puro con `openpyxl` + `dict`/`Counter`/`set` (sin pandas, sin `merge`, sin `groupby`).

| Control | Método A (pandas) | Método B (dict/Counter/set) | Coincide |
|---|---|---|---|
| Reconciliación del maestro | 4.384 = 4.334 (directa, ID único) + 4 (representante de grupo idéntico) + 5 (consolidada, no migra) + 41 (conflictiva, excluida) | 4.384 = 4.330 (en final, sin ningún flag) + 13 (en final CON flag de `PENDIENTES`, incluye representantes y filas con doble motivo) + 41 (sólo en `PENDIENTES`, fuera del final) + 0 (sin destino) | ✔ Ambos métodos coinciden en lo esencial: 4.343 filas de la V3 tienen su ElementoID en el final (4.330+13) y 41 quedan fuera; la única diferencia es de agrupación interna (Método A separa "representante" de "consolidada"; Método B los junta bajo "en final + en pendientes"), no de resultado — 0 sin destino en ambos |
| Reconciliación de campañas | 9.506 = 9.405+80+11+7+3 | 9.506 = 9.405+98(en final y en pendientes)+3(sólo pendiente)+0(sin destino), con 98=80+11+7 | ✔ Exacto en ambos métodos |
| ElementoID duplicados en V3 | `value_counts()>1` → 22 IDs / 50 filas | `Counter()` → 22 IDs / 50 filas | ✔ Exacto |
| ElementoID duplicados en base final | `value_counts()>1` → 0 | `Counter()` → 0 | ✔ Exacto |
| ElementoID huérfanos (en CAMPANAS sin maestro) | `isin()` sobre sets → 0 | Resta de `set()` → 0 | ✔ Exacto |
| Conteo de filas (las 5 hojas/tablas relevantes) | `len(df)` / `ws.max_row` con `dropna(how="all")` | Conteo manual fila a fila con `openpyxl`, cortando en la primera fila 100% vacía | ✔ Exacto en las 5 tablas (4.384, 4.338, 9.506, 9.503, 157) |
| Comparación de datos migrados (spot-check campo por campo) | `merge` + comparación vectorizada normalizada, 26 campos de maestro × 4.338 filas → 0 diferencias | Comparación manual campo por campo sobre 5 filas elegidas con `random.seed(42)` (sin pandas) → 0 diferencias | ✔ Exacto |

---

## 13. Errores

**Ninguno.** No se encontró ninguna fila sin destino, ningún ElementoID duplicado en la base final, ningún huérfano, ninguna diferencia de valor real (no de tipo) en los campos comparables, y ninguna contradicción entre `PENDIENTES` y el estado de la base final.

---

## 14. Advertencias

1. **`TipoInventario`, `AplicaCantidad`, `RevisionMaestro`, `Proveedor` están completamente vacíos** en `MAESTRO_ELEMENTOS` (4.338/4.338 filas). No bloquean el modelado general, pero **sí bloquean cualquier medida de Power BI que dependa de la clasificación Digital/Estático fijo/Flexible**.
2. Dos filas de `PENDIENTES` (551, 552) llevan doble motivo simultáneo; es válido, pero puede confundir a quien cuente filas de `PENDIENTES` de forma ingenua (157 filas ≠ 155 eventos físicos distintos).
3. `RevisionMaestro` es, por diseño de la V3, una columna de control interno de auditoría (marca "Revisar: …" u "OK"); antes de migrarla a la base operativa conviene decidir si tiene valor para el usuario final de Power BI o si debe quedar sólo como artefacto de auditoría (ver pregunta 7).
4. La mezcla de tipos (`int`/texto) en `CapacidadSlotsReel` y `SegundosDia` (1.583 de 4.338 filas) es preexistente en la V3 y no se corrigió en la migración; si Power BI/Power Query no normaliza el tipo de columna, puede producir agregaciones incorrectas (silenciosamente) al mezclar texto y número en la misma columna.

---

## 15. Recomendaciones

1. Antes de construir medidas de Power BI basadas en tipo de inventario: **recalcular `TipoInventario`, `AplicaCantidad` y `RevisionMaestro`** aplicando la regla determinística documentada en la Sección 7, en un paso de transformación (Power Query o script Python), no editando el Excel manualmente.
2. Normalizar el tipo de dato de `CapacidadSlotsReel` y `SegundosDia` a numérico en la capa de carga (Power Query `Table.TransformColumnTypes` o `pd.to_numeric`), no en el Excel fuente.
3. Mantener `PENDIENTES` como tabla de auditoría viva: las 41 filas `ID_DUPLICADO_CONFLICTIVO` (maestro) y las 3 `PENDIENTE_ID_CONFLICTIVO` (campañas) representan datos reales de negocio actualmente **fuera** de la base operativa; deben resolverse manualmente antes de considerarlos completos, aunque no bloquean el uso de lo ya migrado.
4. Documentar en el modelo de Power BI que `EstadoValidacion="OK"` es la condición para considerar una fila de `CAMPANAS` "limpia"; las filas con otro estado están migradas pero requieren criterio del usuario antes de tratarlas como definitivas en reportes ejecutivos.
5. Evaluar si `Proveedor` (maestro) debe poblarse desde otra fuente (no existe en la V3 en absoluto) antes de construir cualquier medida que dependa de proveedor de elemento.

---

## 16. Tabla final de control

| Control | Fuente | Destino | Diferencia | Explicación | Estado |
|---|---|---|---|---|---|
| Filas BASE_MAESTRA_ELEMENTOS → MAESTRO_ELEMENTOS + PENDIENTES | 4.384 | 4.338 + 46 (41 excluidas + 5 consolidadas) | 0 (cierra exacto) | Reconciliado por ElementoID + FilaOrigen, doble método | OK |
| Filas BASE_CAMPAÑAS → CAMPANAS + PENDIENTES | 9.506 | 9.503 + 3 (excluidas) | 0 (cierra exacto) | Reconciliado por FilaOrigen + clave lógica, doble método | OK |
| ElementoID duplicados en base final | 0 esperados | 0 encontrados | 0 | Verificado por `Counter` y `value_counts` | OK |
| ElementoID huérfanos en CAMPANAS | 0 esperados | 0 encontrados | 0 | 2.935 ElementoID distintos, todos con maestro | OK |
| Campos copiados directos (maestro, 22 campos) | V3 | Final | 0 | 100% idénticos en 4.338 filas | OK |
| TipoInventario / AplicaCantidad / RevisionMaestro | Fórmula sin caché en V3 | Vacío en final | 100% vacío en ambos | Recuperable de forma determinística; no migrado aún | REQUIERE_DECISION |
| Proveedor (maestro) | Vacío en V3 | Vacío en final | 0 (coherente) | Nunca tuvo dato | OK |
| CapacidadSlotsReel / SegundosDia | Mixto int/texto en V3 | Mismo mixto en final | 0 valor, sólo tipo | Preexistente, sólo diferencia de tipo de almacenamiento | REQUIERE_DECISION (normalizar tipo antes de Power BI) |
| FechaFin (formato + INDET) | V3 datetime/"INDET" | Final fecha/"FechaIndefinida" | 0 valor real | Transformación documentada y verificada | OK |
| ModalidadPauta, TipoCargaDeclarado, EstadoValidacion, ClaveNegocio, CargaID | No existen con datos en V3 | Derivados/nuevos en final | N/A | Campos nuevos de control, reglas verificadas sin excepciones | OK |
| PENDIENTES vs base final (contradicciones) | 157 filas | — | 0 contradicciones | Ninguna fila "OK" y "excluida" simultáneamente | OK |

---

## 17. Preguntas finales obligatorias

1. **¿Se perdió alguna fila del maestro?** No. Las 4.384 filas de la V3 se reconcilian exactamente contra 4.338 filas finales + 46 filas documentadas en `PENDIENTES` (41 excluidas por conflicto, 5 consolidadas por ser copias idénticas).
2. **¿Se perdió alguna activación de campañas?** No. Las 9.506 filas de la V3 se reconcilian exactamente contra 9.503 filas finales + 3 filas excluidas y documentadas en `PENDIENTES`.
3. **¿Todas las diferencias están explicadas por pendientes, consolidaciones o transformaciones?** Sí, el 100% de las filas de la V3 (tanto maestro como campañas) tiene un destino verificable: base final, `PENDIENTES`, o ambos.
4. **¿Hay registros sin destino?** No, 0 en ambas hojas, confirmado por dos métodos independientes.
5. **¿TipoInventario debe recuperarse?** Sí, recomendado antes de construir medidas de Power BI que segmenten por tipo de inventario (Digital / Estático / Flexible), usando la fórmula determinística documentada en la Sección 7.
6. **¿AplicaCantidad debe recuperarse?** Sí, por la misma razón y con la misma regla (depende de `TipoInventario`).
7. **¿RevisionMaestro sigue siendo necesario?** Depende del uso previsto: es un campo de control de calidad interno de la V3 (marca duplicados/inconsistencias), no un atributo de negocio. Se recomienda decidir explícitamente si migra a la base operativa o si su función ya está cubierta por `PENDIENTES` en la nueva arquitectura.
8. **¿Proveedor del maestro debe recuperarse?** No puede "recuperarse" porque nunca existió en la V3; si se necesita, debe conseguirse de otra fuente de datos.
9. **¿CantidadUnidades perdió datos?** No. Estaba 100% vacío en la V3 (0/9.506) y sigue 100% vacío en la base final (0/9.503); es un campo reservado, no una pérdida.
10. **¿TipoExclusividad perdió datos?** No. La columna equivalente de origen (`EsExclusividad`) también estaba 100% vacía en la V3.
11. **¿HoraInicio/HoraFin perdieron datos?** No. La V3 sólo tenía 3 filas con dato real de 9.506; la base final conserva exactamente esas 3.
12. **¿PROGRAMATICA perdió datos?** No. Comparación campo por campo: 100% idéntico (ajustado por las 3 filas excluidas).
13. **¿CANJE perdió datos?** No. 100% idéntico tras ajustar por las 3 filas excluidas.
14. **¿IDCampaña perdió datos?** No. 26 vacíos en V3, 26 vacíos en final — mismos 26, verificado por `FilaOrigen`.
15. **¿Cliente/Agencia perdieron datos?** No. Tras normalizar la comparación (NaN=NaN), 0 diferencias reales; los ~35% de vacíos en ambos campos ya eran vacíos en la V3.
16. **¿Las diferencias de tipo en CapacidadSlotsReel son solamente de almacenamiento?** Sí, confirmado: 0 diferencias de valor real en 4.338 comparaciones, y la mezcla int/texto (2.755/1.583) es idéntica en V3 y en final.
17. **¿Las diferencias de tipo en SegundosDia son solamente de almacenamiento?** Sí, mismo resultado que el punto anterior.
18. **¿La base plana representa fielmente la información útil de la V3?** Sí, con la salvedad documentada de `TipoInventario`/`AplicaCantidad`/`RevisionMaestro`/`Proveedor` (vacíos en ambos lados, no por la migración sino por el estado del archivo de auditoría).
19. **¿OCU26_BASE_DATOS.xlsx está lista para desarrollar las medidas de Power BI?** Sí, **para todas las medidas que no dependan de tipo de inventario, capacidad/segundos correctamente tipados o proveedor del maestro**. Ver punto 20.
20. **¿Qué medidas NO deberían construirse todavía?** Cualquier medida que segmente por `TipoInventario`/`AplicaCantidad` (ej. "fill rate por tipo de inventario", "digital vs. estático"), cualquier medida de capacidad/fill-rate que sume `CapacidadSlotsReel`/`SegundosDia` sin normalizar antes el tipo de columna, y cualquier medida que use `Proveedor` del maestro.
21. **¿Qué correcciones manuales siguen pendientes?** Resolver los 41 ElementoID conflictivos del maestro y las 3 activaciones de campaña ligadas a ellos; revisar las 98 filas de campaña migradas-con-flag (`PENDIENTE_HISTORICO`/`PENDIENTE_DUPLICADO`/`PENDIENTE_COLISION`); revisar los 10 elementos digitales con capacidad en cero.
22. **¿Es necesario modificar la base antes de continuar con Power BI?** Sí, se recomienda un paso de transformación (no edición manual del Excel) que: (a) recalcule `TipoInventario`/`AplicaCantidad` con la regla documentada, y (b) normalice el tipo numérico de `CapacidadSlotsReel`/`SegundosDia`. Ninguna otra modificación es indispensable para empezar a modelar el resto de las medidas.

---

## Veredicto final

**B. APROBADA CON AJUSTES DE DATOS PENDIENTES QUE NO BLOQUEAN EL MODELADO.**

La reconciliación de filas cierra de forma exacta y verificada por doble método tanto para el maestro como para las campañas, sin pérdidas, sin huérfanos y sin contradicciones en `PENDIENTES`. El único ajuste pendiente con impacto real es la recuperación determinística de `TipoInventario`/`AplicaCantidad`/`RevisionMaestro` y la normalización de tipo de `CapacidadSlotsReel`/`SegundosDia`, ambos con regla de reconstrucción ya documentada en este informe. Esto permite avanzar con el modelado de Power BI y el desarrollo HTML para todas las medidas que no dependan de esos campos, mientras se decide y ejecuta (en un paso de transformación, no en el Excel origen) la recuperación de los campos señalados.

---

*Fin del informe. No se modificó, corrigió, completó ni eliminó ningún dato en `input/OCU26_BASE_DATOS.xlsx`, `audit_sources/Base_ocupacion_26_FINAL_CON_YPF_AUDITADA_V3.xlsx` ni `audit_sources/OCU26_PENDIENTES_VALIDACION.xlsx` durante esta auditoría. No se realizó ningún commit, push ni pull request.*
