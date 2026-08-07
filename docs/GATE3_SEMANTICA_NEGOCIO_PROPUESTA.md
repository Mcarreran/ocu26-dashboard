# Gate 3A — Relevamiento y definición semántica del negocio OCU26

**Fecha:** 2026-08-07 (revisión de corrección: 2026-08-07)
**Alcance:** Solo lectura. Fuente única de datos: `transform_data()` (`scripts/transform_data.py`) sobre `input/OCU26_BASE_DATOS.xlsx`. No se modificó `scripts/validate_input.py`, `scripts/transform_data.py`, ningún test, `input/`, `audit_sources/` ni ningún HTML/JS existente (no existen en este repositorio, ver Sección 38). Único archivo creado/modificado: este documento.

**Nota de corrección (2026-08-07, ronda 1):** esta revisión incorpora decisiones de negocio confirmadas por Brand Plus sobre las preguntas pendientes de la versión anterior de este documento (London Supply, YPF, Puentes LED, performance core jerárquico, APSA, Cencomedia, MAB, Pilar, Pantalla LED Córdoba, ocupación estática inclusiva, campañas Reservadas, 72.000 segundos). No se repitió el relevamiento de datos; se corrigieron únicamente las secciones afectadas por estas decisiones, manteniendo la separación entre hecho observado, regla confirmada, regla legacy, propuesta y pregunta abierta.

**Nota de corrección (2026-08-07, ronda 3 — final antes de commit; la ronda 2 fue la microcorrección de London Supply/Pilar Frontlight/MAB en `CoberturaCatalogo` y `PortfolioTier`, Sección 18.2):** se separó `CoberturaCatalogo` (¿conocemos el universo comercial?) de una nueva dimensión `CompletitudMaestro` (¿ya cargamos en la base todo ese universo conocido?) — AA2000 es el caso que motiva la separación: universo conocido (`COMPLETO`) pero carga incompleta (`CompletitudMaestro = PARCIAL`, faltan Mendoza/Córdoba). Además se corrigió `CertezaDato` para que "elemento con campaña = CONFIRMADO" **no sea una regla global**: es específica de YPF; en los circuitos de catálogo cerrado/gobernado (Cencosud, Remeros, Pantallas LED, Pilar Frontlight, AA2000) un elemento cargado y validado es `CONFIRMADO` aunque no tenga campañas — 0 campañas significa "sin actividad comercial registrada", no "elemento no confirmado".

**Convención de clasificación** (obligatoria, no se mezclan categorías):
- **HECHO OBSERVADO EN DATOS** — verificado ejecutando `transform_data()` sobre el input productivo.
- **REGLA CONFIRMADA POR BRAND PLUS** — enunciada explícitamente en el prompt de Gate 3A.
- **REGLA LEGACY / HISTÓRICA** — proviene de `audit_sources/INFORME_RECONCILIACION_MIGRACION.md` o de convenciones heredadas de la V3, sin haber sido re-confirmada en este prompt.
- **INFERENCIA / PROPUESTA** — razonamiento propio a partir de lo anterior, no un hecho ni una regla confirmada.
- **PREGUNTA ABIERTA** — no puede resolverse con los datos ni con este prompt; requiere respuesta de Brand Plus.

---

## 1. Resumen ejecutivo

- El repositorio **no contiene ningún HTML, JavaScript ni JSON de dashboard** (ver Sección 38). Todo el análisis de "vistas actuales" que pedía el prompt no pudo hacerse porque no hay vistas en el repositorio: no se inventó ninguna, no se buscó fuera del repositorio.
- Gate 1 (`validate_input.py`) y Gate 2 (`transform_data.py`) están intactos, versionados y son la única fuente de datos usada aquí. El input pasa Gate 1 como `VALID_WITH_WARNINGS` (0 errores, solo advertencias de completitud). **HECHO OBSERVADO EN DATOS**.
- `MAESTRO_ELEMENTOS` tiene 4.338 filas, 26 columnas; `CAMPANAS` tiene 9.503 filas, 29 columnas (sin `_row`); `PARAMETROS` tiene 23 filas. **HECHO OBSERVADO EN DATOS**.
- `CircuitoDashboard` tiene **31 valores distintos**, y ninguno se llama literalmente "Cencomedia" ni "APSA" como circuito de primer nivel: APSA vive *dentro* de `CircuitoDashboard = "Shoppings Estático"` (como `Subcircuito = "APSA"`), y lo que el prompt llama "Cencomedia" aparece como **22 `CircuitoDashboard` individuales** (`Jumbo …` ×20, `Disco …` ×2), sin ninguna etiqueta agrupadora. Esto es una diferencia estructural real entre el vocabulario de negocio del prompt y el vocabulario literal del Excel, documentada en la Sección 39 (mapeo de circuitos). **HECHO OBSERVADO EN DATOS**.
- Los circuitos cerrados confirmados por Brand Plus (Cencosud, Remeros, Pantallas LED, AA2000) sí son identificables en los datos, pero **Remeros y APSA comparten el mismo `CircuitoDashboard` que Cencosud** (`Shoppings Estático` / `Shoppings Digital`); solo se distinguen por `Subcircuito`. Esto es crítico para Gate 3B: filtrar por `CircuitoDashboard = "Shoppings Estático"` sin excluir `Subcircuito = "APSA"` mezclaría un circuito cerrado con uno legacy. **HECHO OBSERVADO EN DATOS**.
- Actualmente, **London Supply, APSA, AA2000 y las 22 ubicaciones de "Cencomedia" (Jumbo/Disco) tienen 0 campañas asociadas** en `CAMPANAS`. Toda la actividad comercial registrada hoy se concentra en YPF Digital (2.651 de 2.699 elementos con actividad), Shoppings Estático/Digital (273 de 650 elementos) y, marginalmente, Pantalla LED (11 de 11) y — sorprendentemente — **YPF Estático tiene 0 campañas** a pesar de tener 383 elementos cargados. **HECHO OBSERVADO EN DATOS**.
- No existe en los datos ningún valor `SegundosDia = 72.000`. Los valores reales observados son `{0, 50.400, 75.000, 100.800}`. La referencia de "72.000 segundos" del prompt es un ancla comercial conceptual, no un valor presente en la base. **HECHO OBSERVADO EN DATOS**.
- **[Resuelto por Brand Plus]** `MAB` y `Pilar` ya no son preguntas abiertas: `MAB` se trata semánticamente como inventario abierto/flexible/complementario (aunque su `TipoCatalogo` en el Excel siga siendo "Cerrado"); `CircuitoDashboard = "Pilar"` es la contracara **estática** (Frontlight) de la Pantalla LED Pilar (**Digital**) — ambos forman parte del performance core, como dos elementos/medios distintos de un mismo sitio. **REGLA CONFIRMADA POR BRAND PLUS.**
- **[Resuelto por Brand Plus]** Pantalla LED Córdoba queda confirmada como una de las **11 ubicaciones reales** del circuito Pantalla LED — no hay contradicción "10 vs. 11"; la lista de 10 ubicaciones del prompt original estaba incompleta. **REGLA CONFIRMADA POR BRAND PLUS.**
- **[Resuelto por Brand Plus]** Puente LED tiene una capacidad comercial confirmada de **13 espacios/campañas**; el dato legado cargado en el maestro (`CapacidadSlotsReel = 10`) se documenta como **inconsistencia entre dato legacy y regla comercial actual**, se mantiene sin modificar para trazabilidad, y Gate 3B deberá aplicar `slots_comerciales = 13` vía perfil de configuración. **REGLA CONFIRMADA POR BRAND PLUS + REGLA LEGACY / HISTÓRICA (el valor 10 del Excel).**
- **[Resuelto por Brand Plus]** APSA queda excluida por defecto de **todos** los universos operativos estándar (no solo del performance core): conteo general, inventario general, ocupación/disponibilidad general y KPIs ejecutivos. Sigue disponible solo bajo consulta explícita de histórico/legacy. **REGLA CONFIRMADA POR BRAND PLUS.**
- Este documento **no implementa nada**: no crea `semantic_model.py`, ni YAML, ni HTML, ni motor de métricas. Es exclusivamente el relevamiento y la propuesta semántica para aprobación de Brand Plus antes de Gate 3B.

---

## 2. Arquitectura Gate 1 → Gate 2 → Gate 3

```
input/OCU26_BASE_DATOS.xlsx  (única base operativa)
        ↓
Gate 1 — validate_input.py   (estructura, dominios, integridad referencial; 0 escritura)
        ↓
Gate 2 — transform_data.py   (reconstruye TipoInventario/AplicaCantidad, normaliza tipos numéricos,
                               100% en memoria, verifica SHA-256 antes/después)
        ↓
Gate 3 — (este documento es 3A: semántica; 3B será el motor de métricas)
        ↓
Múltiples vistas (HTML, Power BI, YPF, futuros tableros) — NO EXISTEN TODAVÍA EN ESTE REPO
```

Puntos verificados en el código (no reimplementados aquí, solo citados):
- `transform_data()` **llama a** `validate_input()` como gate obligatorio (`transform_data.py:237`) y aborta con `TransformError` si el resultado es `INVALID`.
- `transform_data()` es estrictamente de lectura: nunca usa `Workbook.save()`; verifica `sha256` del archivo fuente antes y después (`transform_data.py:328-333`).
- El único acoplamiento a `audit_sources/` es **documental, no de runtime**: la regla de `TipoInventario` fue *reproducida* como constante literal (`TIPO_INVENTARIO_KEYWORDS`, `transform_data.py:45`) a partir de la fórmula documentada en el informe de reconciliación; el archivo de auditoría no se lee en ningún punto del pipeline. **HECHO OBSERVADO EN DATOS** (leído directamente del código).

---

## 3. Principios semánticos

**REGLA CONFIRMADA POR BRAND PLUS** (reafirmadas del prompt, sin modificación):
1. Una única lógica semántica (Gate 3) debe alimentar todas las vistas; ninguna vista debe tener lógica de negocio propia.
2. El objetivo es libertad total de cruces: dimensiones + métricas + filtros + agrupaciones + reglas semánticas, no funciones por circuito.
3. `TipoCatalogo` es una clasificación operativa del Excel, no la única fuente de verdad semántica. Se separan explícitamente: `CoberturaCatalogo`, `CertezaDato`, `ModoDisponibilidad`, `PortfolioTier`/`IncluyePerformanceCore` (Secciones 15-18).
4. `MAESTRO_ELEMENTOS` es un catálogo vivo: cualquier fila nueva válida debe fluir por Gate 1→2→3→métricas→vistas sin rehacer código. Ninguna cantidad actual (4.338 elementos, 9.503 campañas, 31 `CircuitoDashboard`, 440 APIE de YPF, etc.) debe hardcodearse en Gate 3B.
5. Circuito cerrado ≠ dato confiable ≠ debe entrar en el performance core. Son tres preguntas independientes (Secciones 15-18).

---

## 4. Inventario cerrado

**REGLA CONFIRMADA POR BRAND PLUS:** un circuito cerrado es aquel donde Brand Plus conoce razonablemente el universo total y tiene disponibilidad total/libre para comercializarlo (propiedad en digital; alquiler con responsabilidad de mantenimiento en estático).

**HECHO OBSERVADO EN DATOS — mapeo real de `TipoCatalogo` a nivel `CircuitoDashboard`:**

| CircuitoDashboard | TipoCatalogo=Cerrado | TipoCatalogo=Abierto |
|---|---|---|
| YPF Digital | 0 | 2.699 |
| YPF Estático | 0 | 383 |
| Shoppings Estático | 381 (Cencosud+Remeros) | 201 (APSA) |
| Shoppings Digital | 68 (Cencosud+Remeros) | 0 |
| London Supply | 0 | 441 |
| AA2000 | 52 | 0 |
| MAB | 13 | 0 |
| Pantalla Led | 11 | 0 |
| Pilar | 1 | 0 |
| Jumbo × 20, Disco × 2 (22 valores) | 0 | 88 (4 cada uno) |

**Conclusión:** `TipoCatalogo = Cerrado` en los datos actuales coincide con Cencosud+Remeros, Pantalla LED (incluida Córdoba, ver Sección 7), AA2000 y también con `MAB` y `Pilar`. `TipoCatalogo = Abierto` coincide con YPF, London Supply, APSA y Cencomedia (Jumbo/Disco). **HECHO OBSERVADO EN DATOS.**

**[Resuelto por Brand Plus]** `TipoCatalogo` **no** debe leerse literalmente para `MAB` ni para `Pilar`:
- `Pilar` (1 elemento, Frontlight, Estático) está correctamente en `Cerrado` — es la contracara estática de la Pantalla LED Pilar y forma parte del performance core. **REGLA CONFIRMADA POR BRAND PLUS** (Sección 7).
- `MAB` (13 elementos, panel móvil/flexible, Estático) tiene `TipoCatalogo = Cerrado` en el Excel, pero **la regla de negocio actual lo trata como abierto/flexible/complementario**, porque es un soporte móvil que puede reubicarse entre circuitos — el nombre actual no representa un universo geográfico fijo. Esto es un ejemplo directo de por qué `TipoCatalogo` no puede usarse como única fuente de verdad semántica (Sección 14, 16). **REGLA CONFIRMADA POR BRAND PLUS.**

---

## 5. Shoppings Cencosud

**HECHO OBSERVADO EN DATOS:** No existe un `CircuitoDashboard` llamado "Shoppings Cencosud". Los shoppings de Cencosud viven dentro de `CircuitoDashboard ∈ {"Shoppings Estático", "Shoppings Digital"}`, identificados por `Subcircuito = "CENCOSUD"`, y se diferencian entre sí por `Ubicacion`.

**Mapeo `Ubicacion` → nombre comercial** (confirmado vía texto libre en `Descripcion`, no vía valor directo — la columna `Ubicacion` usa códigos abreviados):

| Ubicacion (código real) | Nombre comercial evidenciado | Evidencia | Clasificación |
|---|---|---|---|
| UNICENTER | Unicenter | Descripcion contiene "Unicenter" | MAPEO CONFIRMADO |
| P.PILAR | Palmas del Pilar | Descripcion: "Palmas del Pilar - Totem 1..5" | MAPEO CONFIRMADO |
| P.OESTE | Plaza Oeste | Descripcion: "Plaza Oeste - Totem..." | MAPEO CONFIRMADO |
| P.PALERMO | Portal Palermo | Descripcion: "Frontlight … Rampa Portal Palermo" | MAPEO CONFIRMADO |
| P.LOMAS | Portal Lomas | Descripcion: "Portal Lomas - Totem 1..3" | MAPEO CONFIRMADO |
| P.ESCOBAR | Portal Escobar | Descripcion: "Portal Escobar - Totem 1..2" | MAPEO CONFIRMADO |
| F.QUILMES | Factory Quilmes | Descripcion: "Factory Quilmes - Totem 1..3" | MAPEO CONFIRMADO |
| F.P.BROWN | Factory Parque Brown (~"Factory Brown") | Descripcion: "Factory Parque Brown - Totem 1" | MAPEO PROBABLE (nombre exacto lleva "Parque" que el prompt no incluyó) |
| F.SAN MARTIN / F.SANMARTIN | Factory San Martín | Descripcion: "Factory San Martin - Totem 1" | MAPEO CONFIRMADO — **pero con inconsistencia de escritura**: 4 filas estáticas usan `Subcircuito`/`Ubicacion` = "F.SAN MARTIN" (con espacio) y 1 fila digital usa "F.SANMARTIN" (sin espacio). No corregido (fuera de alcance de Gate 3A). |
| P.ROSARIO | Portal Rosario | Sin texto explícito en muestra revisada | MAPEO PROBABLE |
| P.TUCUMAN | Portal Tucumán | Sin texto explícito en muestra revisada | MAPEO PROBABLE |
| P.SANTIAGO | Portal Santiago | Sin texto explícito en muestra revisada | MAPEO PROBABLE |
| P.SALTA | Portal Salta | Descripcion: dirección real de Salta capital | MAPEO PROBABLE |
| P.TRELEW | Portal Trelew | Sin texto explícito en muestra revisada | MAPEO PROBABLE |
| P.LOSANDES | Portal Los Andes | Sin texto explícito en muestra revisada | MAPEO PROBABLE |
| P.PATAGONIA | Portal Patagonia | Sin texto explícito en muestra revisada | MAPEO PROBABLE |

**Total: 16 nombres comerciales distintos bajo `Subcircuito = "CENCOSUD"`** (contando F.SAN MARTIN/F.SANMARTIN como uno solo) — coincide exactamente con los 16 shoppings listados en el prompt. **HECHO OBSERVADO EN DATOS.**

**Sobre "Plaza Oeste mencionado dos veces":** en los datos actuales existe **un único** valor de `Ubicacion` para Plaza Oeste (`P.OESTE`), sin duplicación. La advertencia del prompt no se materializa como una duplicación de dato en el estado actual del maestro. **HECHO OBSERVADO EN DATOS.**

**Volumen:** Shoppings Estático (Cencosud+Remeros) = 381 elementos; Shoppings Digital (Cencosud+Remeros) = 68 elementos. Unicenter es, por lejos, la ubicación con más elementos (97 estáticos + 35 digitales = 132).

**Formatos digitales identificados dentro de Cencosud** (ver también Secciones 24-29): Totems (20 seg. slots=20/100.800), Triedros (6 elementos, slots=10/50.400) y Puentes LED (5 elementos, slots=10/50.400), todos con `TipoInstalacion = "No aplica"` y `ModalidadPauta` no derivable desde `MAESTRO_ELEMENTOS` (esa columna vive en `CAMPANAS`).

---

## 6. Remeros

**HECHO OBSERVADO EN DATOS:** Remeros **no tiene su propio `CircuitoDashboard`**. Vive dentro de `CircuitoDashboard ∈ {"Shoppings Estático", "Shoppings Digital"}` con `Subcircuito = "REMEROS"`, `Ubicacion = "REMEROS"` — es decir, en el mismo `CircuitoDashboard` que Cencosud, aunque el prompt aclara que "no pertenece a Cencosud".

- Shoppings Estático → REMEROS: 27 elementos.
- Shoppings Digital → REMEROS: 6 elementos.
- `TipoCatalogo` = Cerrado en el 100% de sus filas (agrupado junto con Cencosud bajo el mismo valor de `TipoCatalogo`).
- Adicionalmente, Remeros tiene un elemento digital propio en el circuito `Pantalla Led` (Subcircuito=PLED, Ubicacion=REMEROS, capacidad 40/75.000) — es decir, **Remeros aparece dos veces en la taxonomía real: una vez como "shopping" (dentro de Shoppings Estático/Digital) y otra vez como una de las 11 pantallas LED**. Esto es coherente con el prompt (Remeros tiene inventario propio de shopping *y* está en la lista de pantallas LED), pero implica que **Gate 3B no puede usar `CircuitoDashboard` solo como filtro para "todo lo de Remeros"**; necesita combinar `CircuitoDashboard` + `Subcircuito`/`Ubicacion`. **INFERENCIA / PROPUESTA** (consecuencia directa de los dos hechos anteriores).

**PREGUNTA ABIERTA:** ¿debe la métrica "Remeros" consolidar shopping + pantalla LED de Remeros, o deben tratarse como dos líneas de negocio separadas aunque compartan geografía?

---

## 7. Pantallas LED

**HECHO OBSERVADO EN DATOS:** `CircuitoDashboard = "Pantalla Led"`, `Subcircuito = "PLED"`, **11 elementos**, 100% Digital, 100% `TipoCatalogo = Cerrado`, capacidad uniforme `CapacidadSlotsReel = 40` / `SegundosDia = 75.000`. Es el único circuito con 100% de sus elementos con al menos una campaña asociada.

| ElementoID | Ubicacion (real) | Código interno (Descripcion) |
|---|---|---|
| C1 - CAB | CABILDO | PCAB-3600seg |
| C2 - JBJ | JUAN B JUSTO | PJBJ-3600seg |
| C3 - REM | REMEROS | PREM-3600seg |
| C4 - 9DJ | 9DEJULIO | P9J-1800seg |
| C5 - OLA | OLAZABAL | PTRI-3600seg |
| C6 - PIL | PILAR | PPIL-3600seg |
| C7 - AVE | AVELLANEDA | PAV-3600seg |
| C8 - PAR | CERRITO | PPAR-3600seg |
| C9 - CHA | CORRIENTES Y STOS. DUMONT | PCHACA-3600seg |
| C10 - COR | CORDOBA | PCOR-3600seg |
| C11 - PAN | PANAMERICANA | PANA-3600seg |

**Comparación contra la lista del prompt original** (9 de Julio, Remeros, Pilar, Cabildo y Juramento, Cerrito, Avellaneda, Chacarita, Olazábal, Juan B. Justo, Panamericana = 10 ubicaciones):

**[Resuelto por Brand Plus]** La lista correcta y confirmada de Pantalla LED es de **11 ubicaciones**, incluyendo `CORDOBA` (C10). No existe contradicción "10 vs. 11": la lista original de 10 ubicaciones estaba incompleta y `CORDOBA` forma parte del circuito, con la misma naturaleza `Cerrado` / `CORE` / `Digital` que las otras 10. **REGLA CONFIRMADA POR BRAND PLUS.**

- `CORRIENTES Y STOS. DUMONT` mapea probablemente a "Chacarita" (el código interno es `PCHACA`). MAPEO PROBABLE.
- `CERRITO` (C8) tiene código interno `PPAR`, que no corresponde a "Cerrito" en ninguna lectura obvia — sugiere que el código fue asignado a una ubicación anterior y la pantalla fue reubicada o renombrada sin actualizar el código. **PREGUNTA ABIERTA** (deseable, no bloqueante — ver Sección 43), no se corrige.
- `OLAZABAL` (C5) tiene código interno `PTRI`, con el mismo patrón de posible desalineación histórica código↔ubicación actual. **PREGUNTA ABIERTA** (deseable, no bloqueante).
- "Cabildo y Juramento" del prompt se registra simplemente como `CABILDO`. MAPEO PROBABLE (coherente, solo abreviado).

**[Resuelto por Brand Plus — vínculo con Pilar Frontlight]** La ubicación `PILAR` (C6, Digital, Pantalla LED) tiene una contracara **estática** confirmada: el elemento `PIL- PLED FRONT-1` (`CircuitoDashboard = "Pilar"`, único elemento de ese valor, `Descripcion = "Frontlight - Dorso Pantalla Led Pilar - Panamericana KM 50"`, Estático, `TipoCatalogo = Cerrado`, ver tabla de la Sección 4). Ambos elementos representan el mismo **sitio** ("Pilar") con dos **medios** distintos (Digital y Estático), cada uno con sus propias reglas de ocupación — no son un error ni un circuito accidental. Ambos forman parte del performance core (Sección 18.2). Esto refuerza la necesidad de distinguir conceptualmente **sitio/ubicación** de **elemento/medio** en el modelo de dimensiones (Sección 35). **REGLA CONFIRMADA POR BRAND PLUS.**

**Conclusión:** el circuito Pantalla LED es cerrado, íntegramente comercializado, con capacidad uniforme, **11 ubicaciones confirmadas** — el candidato más simple para un primer piloto de métricas de ocupación digital en Gate 3B. Solo quedan como preguntas deseables (no bloqueantes) los códigos internos desalineados (`PPAR`, `PTRI`).

---

## 8. AA2000

**REGLA CONFIRMADA POR BRAND PLUS:** AA2000 es circuito cerrado (Digital + Estático).

**HECHO OBSERVADO EN DATOS:** `CircuitoDashboard = "AA2000"`, 52 elementos, 100% `TipoCatalogo = Cerrado`. `Subcircuito` distintos: `EZEIZA` (35, incluye 10 elementos digitales `TRIPSTORE`/Totem), `AEROPARQUE` (17, incluye 3 elementos digitales `TRIPSTORE`).

**Verificación específica Mendoza/Córdoba (pedida explícitamente por el prompt):**
- **0 filas** con `Ciudad = "MENDOZA"` bajo `CircuitoDashboard = "AA2000"`.
- **0 filas** con `Ciudad = "CORDOBA"` bajo `CircuitoDashboard = "AA2000"` como aeropuerto — la única fila que menciona "Cordoba" es `ElementoID = "COREX022"`, `Subcircuito = "AEROPARQUE"`, `Ubicacion = "CORDOBA"`, `Ciudad = "CABA"`, `Descripcion = "CARTEL RUTERO INGRESO AEROPUERTO"` — es decir, es un **cartel físicamente ubicado en Aeroparque (Buenos Aires)** que hace referencia direccional a Córdoba, no un elemento en el Aeropuerto de Córdoba. **HECHO OBSERVADO EN DATOS.**
- Conclusión: **Ezeiza y Aeroparque tienen ElementoID cargados; Mendoza y Córdoba tienen 0 ElementoID cargados hoy.** Esto confirma la distinción que pide el prompt: la ausencia de registros en Mendoza/Córdoba es una cuestión de cobertura del catálogo (`CoberturaCatalogo`, Sección 15), **no** una redefinición de AA2000 como circuito abierto. AA2000 sigue siendo `TipoCatalogo = Cerrado` en el 100% de sus filas reales; simplemente el universo cargado hoy no incluye esos dos aeropuertos. **HECHO OBSERVADO EN DATOS.**

**Actividad comercial:** **0 campañas** asociadas a ningún elemento de AA2000 en la base actual (ver Sección 35 para la lista completa de circuitos sin campañas). Esto no cambia su clasificación como cerrado ni como parte del portfolio core (Sección 18), pero sí es relevante para cualquier métrica de ocupación/fill-rate: hoy el numerador de "elementos ocupados" en AA2000 es 0 en toda la ventana de datos cargada. **HECHO OBSERVADO EN DATOS.**

**Capacidad digital:** los 10 Totems Tripstore (`AEP-TS-*`, `EZE-TS-*`) tienen `CapacidadSlotsReel=20`/`SegundosDia=100.800`; 2 elementos digitales (`EZEPAW005`, `EZEPAW011`, sin `Descripcion`) tienen `CapacidadSlotsReel=0` — coinciden con las filas que la auditoría histórica (Sección 4 de `INFORME_RECONCILIACION_MIGRACION.md`) marcó como duplicado consolidado / `PENDIENTE_CAPACIDAD`. **HECHO OBSERVADO EN DATOS + REGLA LEGACY** (el origen de esas dos filas está documentado en la auditoría de migración, no se re-verifica aquí).

---

## 9. Inventario abierto/progresivo

**REGLA CONFIRMADA POR BRAND PLUS:** un circuito abierto no implica que sus datos sean incorrectos; implica que no debe asumirse que el maestro contiene el universo comercial completo. Cobertura del catálogo y modo de disponibilidad son conceptos distintos.

**HECHO OBSERVADO EN DATOS:** bajo esta regla caen, en los datos actuales: YPF Digital (2.699), YPF Estático (383), London Supply (441), APSA (201, dentro de Shoppings Estático) y las 22 ubicaciones Jumbo/Disco (88). Total: 3.812 de 4.338 filas del maestro (87,9%) tienen `TipoCatalogo = Abierto`. Es decir: **la mayoría del maestro es, por diseño, de cobertura no garantizada** — un dato importante para calibrar expectativas sobre cualquier "% de ocupación total" agregado sin segmentar por circuito (Sección 42).

---

## 10. Cencomedia

**HECHO OBSERVADO EN DATOS — hallazgo estructural clave:** no existe ningún valor de `CircuitoDashboard`, `Subcircuito` ni `TipoCatalogo` llamado "Cencomedia". Lo que el prompt describe como "Cencomedia / Supermercados" corresponde, en los datos actuales, a **22 `CircuitoDashboard` individuales**, todos con nombre de cadena/sucursal:

- 20 ubicaciones **Jumbo**: Acoyte, Almagro, Escalada, Escobar, Juan B. Justo, La Palmera, Las Heras, Lomas, Madero Harbour, Martinez, Moron, Nordelta, Pacheco Novo, Palermo, Paseo Del Pilar, Pilar, Quilmes, Quilmes Ii, San Martin, Tronador.
- 2 ubicaciones **Disco**: Pinamar, Gesell.

Cada una tiene **exactamente 4 elementos**: `Cubre Alarmas`, `Stoppers`, `Floorgraphics`, `Carros` — con `ElementoID` en formato `{código}_{Formato}` (ej. `JAC_Cubre Alarmas`, `PI_Carros`). Total: **88 elementos**, que coinciden exactamente con el 100% del universo `TipoInventario = "Flexible gráfico"` del maestro (2.803 Digital + 1.447 Físico estático + **88 Flexible gráfico** = 4.338). **HECHO OBSERVADO EN DATOS.**

Todas las filas comparten: `Medio = Estático`, `AplicaCantidad = SI`, `TipoCatalogo = Abierto`, `CapacidadSlotsReel = 0`/`SegundosDia = 0` (no aplica capacidad de reel a este tipo de inventario), `Ubicacion = "AMBA"` (para las 20 Jumbo) o el nombre de la localidad (`PINAMAR`, `GESELL` para Disco). **HECHO OBSERVADO EN DATOS.**

**Contradicción con el prompt a documentar (no corregir):** el prompt describe Cencomedia como un universo donde "Brand Plus tampoco tiene certeza total de cuántos supermercados forman el universo… el maestro crece a medida que se cotiza/solicita/instala" y menciona "campañas iniciales… luego una discontinuidad". En los datos actuales:

- **0 campañas** están asociadas a ninguno de los 88 elementos Flexible gráfico (Jumbo/Disco). No hay evidencia en `CAMPANAS` de las "campañas iniciales" que menciona el prompt — pueden haber existido y no estar en esta base, o la actividad descripta es anterior a este período de carga. **PREGUNTA ABIERTA** (no se puede resolver con los datos disponibles).
- `CantidadUnidades` en `CAMPANAS` está **100% vacío** (0 de 9.503 filas) en toda la base, no solo para Cencomedia — es un campo reservado sin uso real todavía (coincide con lo documentado en la Sección 11 del informe de reconciliación). **HECHO OBSERVADO EN DATOS + REGLA LEGACY.**
- El maestro actual solo modela 22 ubicaciones "conocidas" (todas Jumbo/Disco, ambas cadenas del grupo Cencosud) — no hay en el maestro ninguna otra cadena de supermercados (Carrefour, Coto, Día, Vea, etc.) que el prompt sugiere podrían formar parte del universo potencial. **HECHO OBSERVADO EN DATOS.**

**Clasificación confirmada:** `PortfolioTier = COMPLEMENTARIO`, `IncluyePerformanceCore = NO` **por ahora** (Sección 18.2). **REGLA CONFIRMADA POR BRAND PLUS:** Cencomedia queda fuera del performance core por ahora, pero debe quedar completamente preparada para incorporarse en el futuro mediante configuración (`IncluyePerformanceCore = NO → SI`) sin modificar el motor. A diferencia de APSA, Cencomedia **sí** puede seguir disponible en el universo operativo general (Sección 18.3, Universo A) si la consulta la solicita explícitamente — no queda oculta por defecto del conteo general, solo fuera del performance core. No se eliminan sus 88 elementos.

---

## 11. London Supply

**REGLA CONFIRMADA POR BRAND PLUS:** Brand Plus **sí conoce** la cantidad/universo informado de elementos de London Supply — no debe clasificarse `CoberturaCatalogo = DESCONOCIDO`. Se distingue expresamente **conocimiento del universo/catálogo** (A) de **modo de disponibilidad** (B): London Supply es (A) universo/cantidad **conocida**, pero (B) disponibilidad por **CONSULTA**, con muy baja actividad comercial, fuera del performance core y fuera de los KPIs comerciales principales por defecto. No debe eliminarse del maestro; sigue disponible para análisis específicos. Esto corrige la versión anterior de este documento, que proponía `CoberturaCatalogo = DESCONOCIDO` para London Supply solo por tener `ModoDisponibilidad = CONSULTA` — ambas dimensiones son independientes (ver Sección 16).

**HECHO OBSERVADO EN DATOS:** `CircuitoDashboard = "London Supply"`, 441 elementos (10,2% del maestro completo), 100% `TipoCatalogo = Abierto`. `Medio`: 428 Estático / 13 Digital. `Subcircuito` → `Ciudad`: `USH`→USHUAIA (171), `FTE`→CALAFATE (141), `REL`→TRELEW (129).

**Actividad comercial:** **0 campañas** asociadas a los 441 elementos en toda la base — confirma cuantitativamente la descripción del prompt ("muy baja rotación, casi no se mueve comercialmente"): en esta ventana de datos, la rotación observada es exactamente cero, no solo baja. **HECHO OBSERVADO EN DATOS.**

**Proporción relativa al maestro:** 441/4.338 = 10,2% de todos los elementos del maestro, con 0% de actividad comercial. Si se incluyera en un denominador de ocupación global, aportaría 441 "elementos disponibles" sin ninguna contraparte de "elementos ocupados" — bajando cualquier % de ocupación agregada en aproximadamente 10 puntos porcentuales de forma artificial, exactamente el riesgo que describe el prompt. **INFERENCIA / PROPUESTA** (cálculo directo a partir de los hechos anteriores, no una métrica ya implementada).

**Nota de calidad de dato:** 6 de los 13 elementos digitales (`PAN1-USH`…`PAN6-USH`, Descripcion="Pantalla Digital") tienen `CapacidadSlotsReel = 0`. **HECHO OBSERVADO EN DATOS.**

---

## 12. APSA

**REGLA CONFIRMADA POR BRAND PLUS (decisión ampliada respecto de la versión anterior de este documento):** APSA queda **fuera de todos los universos operativos estándar actuales**, no solo del performance core. Esto incluye por defecto: performance core, conteo general de elementos, inventario general mostrado, ocupación general, disponibilidad general, KPIs ejecutivos y dashboards comerciales estándar. **No se borra APSA del Excel ni se elimina de Gate 1/Gate 2** — se mantiene únicamente como `PortfolioTier = LEGACY`, visible solo si el usuario solicita explícitamente incluir histórico/legacy (`IncluyePerformanceCore = NO`, `IncluyeConteoGeneral = NO`, `VisiblePorDefecto = NO`; nombres técnicos definitivos a decidir en Gate 3B — ver Sección 18). **APSA no debe contaminar ningún KPI o conteo operativo estándar actual.**

**HECHO OBSERVADO EN DATOS:** APSA **no es un `CircuitoDashboard`**; es `Subcircuito = "APSA"` dentro de `CircuitoDashboard = "Shoppings Estático"` (201 elementos, 100% `TipoCatalogo = Abierto` — nótese que esto contrasta con Cencosud/Remeros, que están en `Cerrado` dentro del mismo `CircuitoDashboard`). Distribución por `Ubicacion` (ciudad): CABA (132), NEUQUEN (14), MENDOZA (14), AVELLANEDA (12), CORDOBA (11), ROSARIO (8), SALTA (8), GBA (2).

**Actividad comercial:** **0 campañas** asociadas a los 201 elementos APSA — confirma cuantitativamente "sin acuerdo comercial estable actual, uso excepcional". **HECHO OBSERVADO EN DATOS.**

**Implicancia para Gate 3B:** cualquier filtro que use solo `CircuitoDashboard = "Shoppings Estático"` para "Cencosud" incluiría automáticamente estos 201 elementos legacy sin actividad — y ahora, por decisión confirmada, también los excluiría por defecto del universo operativo general, no solo del performance core. El filtro correcto para el circuito comercial "Cencosud" es `CircuitoDashboard IN ("Shoppings Estático","Shoppings Digital") AND Subcircuito = "CENCOSUD"` (y para Remeros, `Subcircuito = "REMEROS"`); el filtro para excluir APSA del universo general es `Subcircuito ≠ "APSA"` salvo consulta explícita de histórico/legacy. **PROPUESTA** de implementación de la regla confirmada — la mecánica de filtrado exacta se define en Gate 3B (resolución jerárquica de metadata, Sección 18).

---

## 13. YPF

**REGLA CONFIRMADA POR BRAND PLUS:** `TipoCatalogo` permanece "Abierto" por compatibilidad histórica; semánticamente se analiza como `CIRCUITO = YPF`, distinguiendo Digital/Estático, sin depender únicamente de `TipoCatalogo`.

**HECHO OBSERVADO EN DATOS:** `CircuitoDashboard ∈ {"YPF Digital","YPF Estático"}`, 3.082 elementos totales (2.699 + 383), 100% `TipoCatalogo = Abierto` en ambos. Es el circuito más grande del maestro (71% de todas las filas).

### 13.1 Estaciones — APIE

`Subcircuito` es usado como el campo APIE: **440 valores distintos** dentro de YPF. Distribución de elementos por APIE: media 7,0, mediana 6, mínimo 1, máximo 25 (ver Sección 14 para el detalle completo). **HECHO OBSERVADO EN DATOS.**

### 13.2 ElementoID — convención APIE + FORMATO + CORRELATIVO

Confirmado en los datos: el segundo segmento de `ElementoID` (separado por `" - "`) es el token de formato. Ejemplos reales observados: `1047 - MB - 1`, `1142 - TT - 1`, `1142 - TT - 2` (repetición de CORRELATIVO dentro del mismo APIE). **HECHO OBSERVADO EN DATOS.**

| Token | Descripcion (texto real) | Medio | Cantidad |
|---|---|---|---|
| TT | Mueble Torre | Digital | 1.661 |
| PPUNTER | Puntera | Digital | 749 |
| MB | Menu Board | Digital | 289 |
| FB | Mupi | Estático | 383 |

Suma Digital (1.661+749+289=2.699) = exactamente `YPF Digital`. FB=383 = exactamente `YPF Estático`. **No hay ningún otro formato estático hoy** — el prompt menciona "Estático — Mupi/Fotobox", y en efecto el 100% del estático YPF es `Descripcion = "Mupi"`, `TipoInstalacion = "Backlight"`. **HECHO OBSERVADO EN DATOS.**

### 13.3 Ciudades

224 ciudades distintas en todo el maestro; dentro de YPF específicamente, la distribución está fuertemente diversificada geográficamente (top: CABA 265, ENTRE RIOS 149, CORDOBA 135, MAR DEL PLATA 104, MISIONES 103…), consistente con una red de estaciones de servicio a lo largo de rutas nacionales, no concentrada en un área metropolitana. **HECHO OBSERVADO EN DATOS.**

### 13.4 Capacidad digital

**100% de YPF Digital** tiene `CapacidadSlotsReel = 20` y `SegundosDia = 100.800`, sin ninguna excepción/override observada hoy. **HECHO OBSERVADO EN DATOS.**

### 13.5 Actividad comercial y certeza — reglas confirmadas

**REGLA CONFIRMADA POR BRAND PLUS (certeza de elementos YPF):** si un elemento YPF tiene una **campaña asociada**, se considera **efectivamente comercializado/montado** y se trata como elemento **CONFIRMADO**, sin exigir como condición adicional foto, relevamiento posterior ni verificación manual (esos procesos pueden existir operativamente, pero no son condición de Gate 3 para la confirmación). Un elemento YPF **sin** campaña es "registrado/conocido", pero no necesariamente confirmado físicamente. Esta regla permite avanzar sin revisar manualmente miles de elementos, y es compatible con que la arquitectura siga permitiendo sumar nuevas estaciones/elementos, quitar elementos si se comprueba que no existen, y modificar el catálogo conocido sin hardcodear los conteos actuales.

- YPF Digital: **2.651 de 2.699 elementos (98,2%) tienen al menos una campaña asociada → CONFIRMADOS** bajo la regla anterior.
- **YPF Estático: 0 de 383 elementos tienen ninguna campaña asociada → registrados/conocidos, no confirmados.** **HECHO OBSERVADO EN DATOS.**

**REGLA CONFIRMADA POR BRAND PLUS (cómo describir este hallazgo):** no debe interpretarse "383 elementos YPF Estático + 0 campañas" como **"0% de ocupación total YPF"**. La descripción correcta es: **"383 elementos estáticos registrados actualmente sin actividad de campaña registrada en esta base"**. YPF general puede mostrarse en dos niveles (Sección 13.7): a nivel general (estaciones registradas, elementos registrados, elementos confirmados, campañas, actividad general) y a nivel desagregado (Digital vs. Estático, formato, APIE/estación, ciudad/zona, ElementoID). Es válido mostrar un número general de "elementos registrados" de YPF y luego abrirlo por Digital/Estático — pero **no debe generarse un único porcentaje de ocupación YPF mezclando Digital y Estático**, porque sus reglas de ocupación son diferentes.

### 13.7 Libertad de cruces (dos niveles)

**REGLA CONFIRMADA POR BRAND PLUS:** YPF debe poder analizarse en (A) nivel general — estaciones registradas, elementos registrados, elementos confirmados, campañas, actividad general — y (B) nivel desagregado como mínimo por Digital vs. Estático, formato (Menu Board / Torre / Puntera / Mupi-Fotobox), APIE/estación, ciudad/zona y `ElementoID`. Esto no cambia ningún hecho ya observado en la Sección 13 (APIE=440, formatos TT/PPUNTER/MB/FB, etc.); formaliza que ambos niveles deben quedar disponibles como cruces en Gate 3B sin funciones dedicadas por nivel.

### 13.6 Confiabilidad

**REGLA CONFIRMADA POR BRAND PLUS:** la base de estaciones YPF no es 100% confiable todavía; direcciones contrastadas contra Google Maps pero sin certeza física/operativa total. Este documento **no verificó ninguna dirección en internet**, conforme a la instrucción explícita de no hacerlo en Gate 3A.

---

## 14. TipoCatalogo vs. semántica

**REGLA CONFIRMADA POR BRAND PLUS:** no usar `TipoCatalogo` como única fuente de verdad; YPF puede seguir siendo "Abierto" en `TipoCatalogo` mientras semánticamente tiene reglas propias (Circuito=YPF, Portfolio=crecimiento, Certeza=progresiva).

**HECHO OBSERVADO EN DATOS:** de los 4.338 elementos del maestro, `TipoCatalogo` coincide con la clasificación cerrado/abierto real en el 100% de los casos **excepto** en `MAB` (`TipoCatalogo = Cerrado` en el dato, pero tratado semánticamente como abierto/flexible por decisión confirmada de Brand Plus, Sección 4/18) y en la necesidad de tratar YPF con una capa semántica adicional (Sección 13.5-13.7) a pesar de que su `TipoCatalogo` sea correcto tal cual está. `Pilar` **sí** coincide correctamente con `Cerrado` (Sección 4/7). Es decir: **el problema no es que `TipoCatalogo` esté mal**, es que por sí solo no expresa "certeza progresiva", "portfolio en crecimiento", la distinción Digital/Estático dentro de YPF, ni el carácter flexible/reubicable de un soporte como MAB. Debe complementarse, no reemplazarse. **REGLA CONFIRMADA POR BRAND PLUS** (MAB, Sección 4/16) + **INFERENCIA / PROPUESTA** (el resto del razonamiento).

---

## 15. CoberturaCatalogo (propuesta)

**Pregunta que responde:** ¿conocemos razonablemente el universo total de soportes de este circuito?

**Propuesta de taxonomía** (INFERENCIA / PROPUESTA, no fijar sin validación de Brand Plus):

| Valor propuesto | Definición | Circuitos candidatos (según hechos observados) |
|---|---|---|
| `COMPLETO` | Brand Plus tiene certeza razonable de que **conoce** el universo comercial total de este circuito (independientemente de si ya está 100% cargado en el maestro, ver `CompletitudMaestro` más abajo) | Shoppings Cencosud+Remeros, Pantalla LED, **Pilar Frontlight**, **London Supply**, **AA2000** |
| `DESCONOCIDO` | No hay forma de estimar el universo total; el maestro crece por demanda comercial, no por relevamiento previo, o el soporte es móvil/reubicable | YPF, APSA, Cencomedia (Jumbo/Disco), **MAB** |

**[Resuelto por Brand Plus]** London Supply se reclasifica de `DESCONOCIDO` (versión anterior de este documento) a **`COMPLETO`**: Brand Plus **sí conoce** la cantidad/universo informado de elementos de London Supply. La versión anterior confundía esta dimensión con `ModoDisponibilidad = CONSULTA` (Sección 17) — son dos preguntas independientes (Sección 16): "¿conozco el universo?" (sí, `COMPLETO`) y "¿puedo calcular disponibilidad automáticamente?" (no, `CONSULTA`). **REGLA CONFIRMADA POR BRAND PLUS.**

**[Resuelto por Brand Plus]** `Pilar Frontlight` se agrega explícitamente como `CoberturaCatalogo = COMPLETO`: existe, se conoce el elemento, es la contracara estática confirmada de Pantalla LED Pilar y forma parte del core business (Sección 18.2). No debe confundirse con la cara Digital (Pantalla LED Pilar, mismo sitio, `Medio` distinto).

**[Resuelto por Brand Plus]** `MAB` se agrega explícitamente como `CoberturaCatalogo = DESCONOCIDO` — no por falta de relevamiento comercial como YPF/Cencomedia, sino porque es un soporte **móvil/flexible** que puede colocarse en distintos circuitos: su `CircuitoDashboard`/`Ubicacion` actual en el maestro **no debe interpretarse como un universo o sitio permanente**, ya que el soporte físico puede reubicarse. Para no crear una categoría nueva en la taxonomía, se reutiliza `DESCONOCIDO` con esta salvedad documentada.

**[Resuelto por Brand Plus — AA2000 pasa de `PARCIAL` a `COMPLETO`, y la pregunta abierta anterior queda eliminada]** AA2000 es un circuito **cerrado cuyo universo comercial es conocido**: `CoberturaCatalogo = COMPLETO`. El hecho de que Mendoza y Córdoba todavía tengan 0 `ElementoID` cargados **no** significa que el universo sea desconocido o parcialmente conocido — significa que **el maestro actual todavía no tiene cargado todo el universo que Brand Plus ya conoce**. Esa es una pregunta distinta, resuelta en la Sección 15.1 con la nueva dimensión `CompletitudMaestro`. **REGLA CONFIRMADA POR BRAND PLUS.**

### 15.1 CompletitudMaestro (nueva dimensión propuesta)

**Pregunta que responde:** ¿la base/maestro actual contiene todos los elementos que Brand Plus ya sabe que deberían estar cargados? — distinta de `CoberturaCatalogo` (¿conocemos el universo comercial?) y de `CertezaDato` (¿confiamos en cada registro concreto?, Sección 16).

**Propuesta de taxonomía** (INFERENCIA / PROPUESTA — nombre técnico definitivo a confirmar en Gate 3B):

| Valor propuesto | Definición | Circuitos candidatos |
|---|---|---|
| `COMPLETO` | Todo el universo conocido (`CoberturaCatalogo=COMPLETO`) ya está cargado en `MAESTRO_ELEMENTOS` | Shoppings Cencosud+Remeros, Pantalla LED, Pilar Frontlight, London Supply |
| `PARCIAL` | Se conoce el universo comercial, pero el maestro todavía no tiene cargados todos los elementos/ubicaciones conocidos | **AA2000** (Ezeiza/Aeroparque cargados; Mendoza/Córdoba conocidos comercialmente pero con 0 `ElementoID` cargados hoy) |

`CompletitudMaestro` solo tiene sentido cuando `CoberturaCatalogo = COMPLETO` — si el universo mismo es `DESCONOCIDO` (YPF, APSA, Cencomedia, MAB), la pregunta "¿ya cargamos todo lo que conocemos?" no aplica del mismo modo, porque no hay un universo cerrado contra el cual medir completitud (`NO_APLICA` o `DESCONOCIDA`, a definir en Gate 3B). **No se modifica el Excel** para reflejar esta dimensión en Gate 3A; es puramente conceptual, para incorporarse como metadata en Gate 3B (Sección 40).

Con esta separación, AA2000 queda:

```
CoberturaCatalogo     = COMPLETO
CompletitudMaestro    = PARCIAL   (Ezeiza/Aeroparque cargados; Mendoza/Córdoba no)
PortfolioTier         = CORE
IncluyePerformanceCore = SI
ModoDisponibilidad    = MIXTO    (calculable donde está cargado; consulta donde falta cargar)
```

---

## 16. CertezaDato (propuesta — corregida: la certeza se resuelve por semántica de circuito, no con una regla global)

**Pregunta que responde:** ¿qué tan confiable es el registro concreto (independiente de si se conoce el universo completo o si ya está todo cargado)?

**[Corrección confirmada por Brand Plus]** La versión anterior de este documento generalizaba "elemento con campaña = `CONFIRMADO`" como si fuera una regla única para todo el maestro. **Esa condición fue confirmada específicamente para YPF** (Sección 13.5), no como regla global. `CertezaDato` debe resolverse según la semántica de cada circuito, con al menos cuatro casos distintos:

**A. Circuitos cerrados / catálogo gobernado** (Cencosud, Remeros, Pantallas LED, Pilar Frontlight, AA2000): un elemento cargado dentro del catálogo cerrado y validado por Gate 1/2 puede considerarse `CONFIRMADO` **aunque todavía no tenga ninguna campaña**. La ausencia de campañas en estos circuitos significa **"sin actividad comercial registrada"**, no **"elemento no confirmado"** — son dos afirmaciones distintas (Actividad comercial es su propia dimensión, ver más abajo).

**B. YPF** (regla ya confirmada, sin cambios — Sección 13.5): `ElementoID` con ≥1 campaña asociada → `CONFIRMADO`. `ElementoID` sin campaña → `REGISTRADO` / `NO CONFIRMADO` (nombre técnico definitivo a resolver en Gate 3B). Esta regla es específica de YPF porque Brand Plus no tiene certeza física/operativa total de las estaciones (Sección 13.6) y usa la campaña ejecutada como evidencia sustituta — no aplica igual a un catálogo ya gobernado y cerrado como Cencosud o Pantalla LED, donde la certeza viene de pertenecer al catálogo validado, no de la actividad.

**C. London Supply**: Brand Plus conoce el universo informado (`CoberturaCatalogo=COMPLETO`, Sección 15). **La ausencia de campañas no debe usarse como prueba de incertidumbre del elemento.** Se separan explícitamente tres cosas que la versión anterior mezclaba parcialmente: certeza del registro (`CertezaDato`), actividad comercial (campañas) y disponibilidad por consulta (`ModoDisponibilidad=CONSULTA`, Sección 17). Con esa separación, los elementos de London Supply pueden tratarse como `CONFIRMADO` en cuanto a registro (Brand Plus conoce y confía en el dato informado), con `Actividad comercial = 0 campañas` como un hecho aparte, no como evidencia de duda sobre el registro.

**D. Cencomedia**: se mantiene la lógica progresiva ya descripta (Sección 10) — no debe asumirse que la presencia de una fila representa necesariamente una instalación física permanente. Su certeza **puede requerir reglas propias futuras** que este documento no fija; se deja explícitamente como **PREGUNTA ABIERTA para Gate 3B**, no se le asigna hoy un valor cerrado de `CertezaDato`.

**`PROVISORIO` / `REQUIERE_REVISION`** (sin cambios de fondo, pero con la salvedad explícita que pedía la corrección): reservado para registros con **problemas técnicos concretos** — capacidad no cargada (`CapacidadSlotsReel=0`, 8 elementos: 2 AA2000 + 6 London Supply), `EstadoValidacion` pendiente en `CAMPANAS` (98 filas), etc. **Nunca debe asignarse automáticamente solo por "0 campañas"** — eso es un hecho de actividad comercial, no un problema técnico del registro.

**Tabla resumen** (INFERENCIA / PROPUESTA, con las reglas A-D arriba como REGLA CONFIRMADA POR BRAND PLUS donde se indica):

| Circuito | `CertezaDato` de un elemento cargado y validado | Base de la certeza |
|---|---|---|
| Cencosud, Remeros, Pantallas LED, Pilar Frontlight, AA2000 | `CONFIRMADO` (con o sin campaña) | Pertenencia a catálogo cerrado/gobernado y validado por Gate 1/2 |
| YPF, con ≥1 campaña | `CONFIRMADO` | Campaña ejecutada (evidencia sustituta de confiabilidad física, Sección 13.5/13.6) |
| YPF, sin campaña | `REGISTRADO` / `NO CONFIRMADO` | Solo carga en el maestro, sin evidencia de montaje |
| London Supply | `CONFIRMADO` (registro), independiente de actividad | Universo informado y conocido por Brand Plus (`CoberturaCatalogo=COMPLETO`) |
| Cencomedia | **PREGUNTA ABIERTA** — no asignado en Gate 3A | Requiere regla propia futura (Sección 10) |
| Cualquier circuito, con problema técnico concreto | `PROVISORIO` / `REQUIERE_REVISION` | Capacidad no cargada, `EstadoValidacion` pendiente, etc. — nunca por sola ausencia de campaña |

Esta dimensión es **ortogonal** a `CoberturaCatalogo` y a `CompletitudMaestro` (Sección 15/15.1), y **distinta de la actividad comercial** (¿tiene campañas?, una séptima dimensión aparte, ver Sección 16.1): un elemento AA2000 puede ser `CONFIRMADO` (catálogo cerrado validado) con `Actividad comercial = 0 campañas` simultáneamente — ambos hechos coexisten sin contradicción. **REGLA CONFIRMADA POR BRAND PLUS.**

### 16.1 Resumen de independencia entre las siete dimensiones semánticas

**REGLA CONFIRMADA POR BRAND PLUS** — deben quedar explícitamente separadas, sin equivalencias implícitas entre ellas:

| Dimensión | Pregunta que responde |
|---|---|
| `CoberturaCatalogo` (Sección 15) | ¿Conocemos el universo comercial? |
| `CompletitudMaestro` (Sección 15.1) | ¿Ya cargamos en la base todo el universo que conocemos? |
| `CertezaDato` (Sección 16) | ¿Confiamos en este registro/elemento concreto? |
| `ModoDisponibilidad` (Sección 17) | ¿Podemos saber su disponibilidad desde el sistema, o requiere consulta? |
| `PortfolioTier` (Sección 18) | ¿Qué papel tiene comercialmente? |
| `IncluyePerformanceCore` (Sección 18) | ¿Entra en los KPIs core? |
| Actividad comercial | ¿Tiene campañas registradas? |

**Ejemplos que demuestran la independencia (los cuatro de la corrección anterior, más AA2000 actualizado):**
- **London Supply** demuestra que universo conocido ≠ disponibilidad calculable, y que certeza de registro ≠ actividad comercial (`CoberturaCatalogo=COMPLETO`, `CertezaDato=CONFIRMADO`, `ModoDisponibilidad=CONSULTA`, `Actividad=0 campañas`, todo simultáneo y sin contradicción).
- **AA2000** demuestra que universo comercial conocido ≠ maestro completamente cargado (`CoberturaCatalogo=COMPLETO`, `CompletitudMaestro=PARCIAL` — Mendoza/Córdoba conocidos pero no cargados), y que catálogo cerrado ≠ requiere campaña para confirmarse (`CertezaDato=CONFIRMADO` con `Actividad=0 campañas`).
- **YPF** demuestra que elemento confirmado ≠ universo total conocido, y que la regla de confirmación por campaña es **específica de YPF**, no generalizable (`CertezaDato=CONFIRMADO` por campaña, mientras `CoberturaCatalogo=DESCONOCIDO` a nivel de circuito).
- **MAB** demuestra que `TipoCatalogo` histórico del Excel ≠ semántica comercial actual (`TipoCatalogo=Cerrado` en el dato, pero tratado como abierto/flexible en la regla de negocio, Sección 4/18).

---

## 17. ModoDisponibilidad (propuesta)

**Pregunta que responde:** ¿la disponibilidad se calcula directamente desde el sistema, o requiere consulta/confirmación manual?

**Propuesta de taxonomía** (INFERENCIA / PROPUESTA):

| Valor propuesto | Definición | Ejemplo |
|---|---|---|
| `CALCULABLE` | Capacidad y calendario están en el maestro/campañas de forma suficiente para calcular disponibilidad sin intervención humana | Pantalla LED (capacidad uniforme 40/75.000, 100% con campañas registradas) |
| `CONSULTA` | La disponibilidad real depende de información que no está en el sistema (relevamiento físico, confirmación de estación) o de un proceso comercial de consulta | YPF (confiabilidad de estaciones aún no total, Sección 13.6); **London Supply** |
| `MIXTO` | Parte calculable, parte requiere consulta | AA2000 (Ezeiza/Aeroparque con datos completos; Mendoza/Córdoba sin ningún dato) |

**[Resuelto por Brand Plus]** London Supply es `ModoDisponibilidad = CONSULTA` (no `CALCULABLE`), aunque tenga `CoberturaCatalogo = COMPLETO` (Sección 15) y capacidad numérica cargada en el maestro — la capacidad cargada no es evidencia de que el proceso comercial sea autoservicio. Esta es precisamente la separación que Brand Plus pidió formalizar: **conocer el universo (A) es independiente de poder calcular la disponibilidad automáticamente (B)**. **REGLA CONFIRMADA POR BRAND PLUS.**

**Nota (corrección ronda 3):** el `MIXTO` de AA2000 es consecuencia de `CompletitudMaestro = PARCIAL` (Sección 15.1), no de `CoberturaCatalogo` — Brand Plus conoce el universo completo de AA2000 (`COMPLETO`); lo que está mixto es qué porción de ese universo ya tiene datos cargados para calcular disponibilidad automáticamente (Ezeiza/Aeroparque) frente a la porción que aún no está cargada (Mendoza/Córdoba) y por tanto requiere consulta.

---

## 18. Portfolio / performance core (propuesta)

**REGLA CONFIRMADA POR BRAND PLUS:** performance core actual NO debe incluir automáticamente todo `MAESTRO_ELEMENTOS`; London Supply, APSA, Cencomedia y MAB deben quedar fuera por defecto (sin eliminarlos), y Cencomedia debe poder incorporarse en el futuro con un cambio de configuración (`IncluyePerformanceCore = NO → SI`) sin tocar Gates 1/2/motor/vistas.

### 18.1 Resolución jerárquica de metadata (resuelve la pregunta bloqueante de granularidad)

**[Resuelto por Brand Plus]** `IncluyePerformanceCore` (y el resto de la metadata semántica: `PortfolioTier`, `CoberturaCatalogo`, `ModoDisponibilidad`, etc.) **no depende únicamente de `CircuitoDashboard`**. Se resuelve con una jerarquía conceptual de configuración, de más a menos específico:

```
1. override por ElementoID
2. regla por CircuitoDashboard + Subcircuito
3. regla por CircuitoDashboard
4. default semántico
```

Esto es exactamente lo que exige el caso `"Shoppings Estático"`, que contiene simultáneamente `CENCOSUD`, `REMEROS` y `APSA` con comportamientos comerciales distintos (nivel 2: `CircuitoDashboard + Subcircuito`), y lo que exige el caso de los 2 elementos AA2000 con `CapacidadSlotsReel=0` si en el futuro necesitaran un tratamiento distinto al resto de AA2000 (nivel 1: override por `ElementoID`). Una clasificación comercial se cambia editando configuración en el nivel correspondiente, sin modificar el motor. **REGLA CONFIRMADA POR BRAND PLUS.**

### 18.2 Taxonomía `PortfolioTier` (actualizada — microcorrección)

**[Resuelto por Brand Plus]** London Supply **no es `LEGACY`**: es inventario comercial **vigente** (a diferencia de APSA, que sí es un circuito histórico sin acuerdo comercial estable). London Supply tiene poca rotación y queda fuera del performance core, pero se reclasifica a `COMPLEMENTARIO`. **APSA queda como el único caso `LEGACY`** de todo el maestro.

| PortfolioTier propuesto | `IncluyePerformanceCore` propuesto | `CoberturaCatalogo` | Circuitos (nivel de resolución jerárquica) |
|---|---|---|---|
| `CORE` | SI | — (ver detalle por circuito) | Shoppings Cencosud (`CircuitoDashboard + Subcircuito = Shoppings Estático/Digital + CENCOSUD`), Remeros (`... + REMEROS`), Pantalla LED (11 ubicaciones, incluida Córdoba), **Pilar Frontlight** (`CircuitoDashboard=Pilar`, Medio=Estático, contracara estática de Pantalla LED Pilar — no confundir con la cara Digital), **AA2000** (`CoberturaCatalogo=COMPLETO`, `CompletitudMaestro=PARCIAL` — Sección 15.1, sin que esto lo excluya del core), YPF (según reglas de certeza/ocupación de la Sección 13) |
| `COMPLEMENTARIO` | NO (Cencomedia preparado para pasar a SI vía configuración; **London Supply y MAB no cambian de tier**, solo quedan fuera del core) | `COMPLETO` para **London Supply**; `DESCONOCIDO` para Cencomedia y **MAB** | Cencomedia (Jumbo/Disco), **London Supply**, **MAB** |
| `LEGACY` | NO — y además excluido del universo operativo general, no solo del core (Sección 12) | `DESCONOCIDO` (universo históricamente incierto) | **APSA únicamente** (`CircuitoDashboard=Shoppings Estático/Digital + Subcircuito=APSA`) |

**Cambios de esta microcorrección respecto de la versión anterior de este documento:**
- **London Supply** pasa de `LEGACY` a **`COMPLEMENTARIO`**: sigue fuera del performance core (`IncluyePerformanceCore=NO`) por baja actividad comercial y `ModoDisponibilidad=CONSULTA`, pero su `CoberturaCatalogo` ya era `COMPLETO` (Sección 15) — Brand Plus conoce su universo, y no es un circuito histórico abandonado. Sigue disponible en el universo operativo general (Sección 18.3).
- **Pilar Frontlight** queda explícitamente `CoberturaCatalogo = COMPLETO`, `PortfolioTier = CORE`, `IncluyePerformanceCore = SI`, `Medio = Estático` — se conoce el elemento, es la contracara estática confirmada de Pantalla LED Pilar (Digital) y forma parte del core business; no deben confundirse ambas caras del mismo sitio.
- **MAB** queda explícitamente `CoberturaCatalogo = DESCONOCIDO` (para no crear una categoría nueva en la taxonomía): es un soporte móvil/flexible que puede colocarse en distintos circuitos, se comercializa poco, se trata semánticamente como abierto/flexible y queda fuera del performance core (`PortfolioTier = COMPLEMENTARIO`, `IncluyePerformanceCore = NO`). Su ubicación/circuito actual en el maestro **no debe interpretarse como permanente**, porque el soporte físico puede moverse — a diferencia de Cencosud, Remeros, Pantalla LED o YPF, donde `CircuitoDashboard`/`Subcircuito`/`Ubicacion` sí representan un sitio fijo.
- **APSA** queda como el **único** caso `LEGACY` de la taxonomía; su exclusión sigue siendo la más amplia (no solo `IncluyePerformanceCore=NO`, sino también fuera del conteo/inventario/ocupación/disponibilidad general por defecto, Sección 12).

Esta tabla sigue siendo **INFERENCIA / PROPUESTA** en cuanto a los nombres técnicos exactos (`PortfolioTier`, `CORE`, `COMPLEMENTARIO`, `LEGACY` pueden decidirse distinto en Gate 3B), pero las **decisiones de negocio subyacentes** (qué queda dentro/fuera de qué universo, y que APSA es el único `LEGACY`) son **REGLA CONFIRMADA POR BRAND PLUS**.

### 18.3 Universos de reporte (propuesta)

**PROPUESTA**, a partir de las decisiones confirmadas de esta corrección — se proponen al menos tres universos distintos de reporte, todos derivables de la misma fuente (`transform_data()`) y de la misma capa de metadata semántica, sin funciones dedicadas por universo:

| Universo | Contenido | Regla de exclusión |
|---|---|---|
| **A. Universo operativo general** | Todos los elementos relevantes actuales, **incluyendo London Supply, Cencomedia y MAB por defecto** (sin necesidad de consulta explícita — son `COMPLEMENTARIO`, no `LEGACY`) | Excluye **APSA** por defecto (único circuito oculto incluso del conteo general) |
| **B. Performance core** | Shoppings Cencosud, Remeros, Pantallas LED, Pilar Frontlight, AA2000, YPF (según reglas de certeza/ocupación de la Sección 13) | Excluye APSA, London Supply, Cencomedia (por ahora), MAB |
| **C. Universo completo / histórico** | Todo lo anterior + APSA + cualquier otro circuito legacy que se agregue en el futuro | Sin exclusiones — requiere consulta explícita del usuario |

**Regla confirmada asociada (corregida):** **London Supply pertenece al universo operativo general (A) por defecto**, aunque no al performance core (B) — es inventario comercial vigente, no histórico. Cencomedia y MAB también permanecen disponibles por defecto en el universo A. **Únicamente APSA permanece oculto/excluido por defecto incluso del conteo general (A)**, y solo aparece si el usuario pide explícitamente el universo histórico/legacy (C). **REGLA CONFIRMADA POR BRAND PLUS.**

---

## 19. Ocupación estática

**REGLA CONFIRMADA POR BRAND PLUS:** la ocupación estática utiliza **días calendario reales e INCLUSIVOS**. `FechaInicio = FechaFin` → **1 día ocupado** (no 0). La fórmula conceptual de días ocupados es `(FechaFin - FechaInicio) + 1 día` cuando ambas fechas existen. No se asume meses de 30 días; la métrica de ocupación por período deberá calcular el solapamiento **inclusivo** entre `[FechaInicio, FechaFin]` y `[InicioPeriodo, FinPeriodo]`.

**HECHO OBSERVADO EN DATOS (diferencia bruta de fechas, antes de aplicar la regla de negocio):** `FechaFin - FechaInicio` sobre 9.443 filas con ambas fechas completas tiene mínimo 0 (diferencia bruta), máximo **1.095** (3 años exactos: `PPIL-FR-1` / "YPF PIER 3", 2024-01-01 a 2026-12-31), media 92, mediana 36, percentil 75 = 194. **Ninguna diferencia es un múltiplo limpio de 30**, confirmando empíricamente que no existe la convención "1 campaña = 1 mes". **HECHO OBSERVADO EN DATOS.**

**[Corrección de interpretación — resuelto por Brand Plus]** Las **17 filas con `FechaInicio == FechaFin`** (diferencia bruta = 0) **no son campañas de "0 días"**: bajo la regla de ocupación confirmada, son campañas de **1 día ocupado**. El dato bruto de diferencia de fechas puede ser 0, pero el significado de negocio, aplicando `+1` inclusivo, es 1 día. Aplicando la regla inclusiva a las 9.443 filas: duración de negocio = diferencia bruta + 1, es decir, mínimo **1 día**, máximo **1.096 días**, con la misma distribución desplazada en +1 respecto de la diferencia bruta reportada arriba.

El rango temporal completo de la base es `FechaInicio` desde 2024-01-01 hasta 2026-08-04, y `FechaFin` hasta 2027-03-31 — abarca pasado, presente y futuro reservado. **HECHO OBSERVADO EN DATOS.**

Conceptos a definir en Gate 3B (**no implementados aquí**, solo nombrados como pide el prompt): `dias_periodo`, `dias_ocupados`, `elemento_dias_disponibles`, `elemento_dias_ocupados`, `ocupacion_calendario_pct`. Su cálculo deberá aplicar la regla inclusiva confirmada arriba sobre el solapamiento real de intervalos `[FechaInicio, FechaFin]` contra el período de consulta, con tratamiento explícito de `FechaIndefinida = "Si"` (5 filas en la base actual, ver Sección 31) y de `Estado = "Reservada"` como bloqueo de disponibilidad futura (Sección 32).

---

## 20. Ocupación calendario digital

**REGLA CONFIRMADA POR BRAND PLUS:** la ocupación digital no es una sola métrica; separa como mínimo (1) ocupación calendario, (2) capacidad comercial, (3) slots/reel, (4) exclusividad. "Elemento con una campaña" ≠ "elemento sin disponibilidad".

**HECHO OBSERVADO EN DATOS que sostiene esta regla:** un elemento digital YPF con `CapacidadSlotsReel = 20` puede tener una campaña activa ocupando 2 salidas (`SalidasVendidas = 2`, la moda absoluta: 9.140 de 9.141 filas con `ModalidadPauta = "Slot / Reel normal"`) — es decir, en el escenario típico, **una campaña ocupa una fracción pequeña de la capacidad del elemento** (2 de hasta 20 slots nominales), no el 100%. Confirma en datos reales por qué "tiene una campaña" no puede leerse como "está lleno". **HECHO OBSERVADO EN DATOS.**

---

## 21. Capacidad digital

**HECHO OBSERVADO EN DATOS — matriz real de `CapacidadSlotsReel` / `SegundosDia` por circuito** (elementos Digital):

| CircuitoDashboard / perfil | CapacidadSlotsReel | SegundosDia | n |
|---|---|---|---|
| YPF Digital (Torre/Puntera/Menu Board) | 20 | 100.800 | 2.699 |
| Shoppings Digital — Totems | 20 | 100.800 | 40 |
| Shoppings Digital — Puentes LED + Triedros | 10 | 50.400 | 28 |
| AA2000 — Totems Tripstore | 20 | 100.800 | 10 |
| AA2000 — 2 elementos sin capacidad cargada | 0 | 0 | 2 |
| Pantalla Led | 40 | 75.000 | 11 |
| London Supply — Digital con capacidad | 20 | 100.800 | 7 |
| London Supply — Digital sin capacidad (Pantallas Ushuaia) | 0 | 0 | 6 |

No existe hoy ningún override por `ElementoID` individual dentro de un mismo perfil (ej. todos los Totems tienen exactamente 20/100.800, sin excepciones) — la variación es por **perfil de tipo de soporte** (Totem vs. Puente/Triedro vs. Pantalla LED), no por instancia. **HECHO OBSERVADO EN DATOS.**

---

## 22. SegundosDia

**REGLA CONFIRMADA POR BRAND PLUS:** **72.000 segundos es la referencia/máximo comercial normalizado** que Brand Plus quiere utilizar conceptualmente para la capacidad digital donde corresponda — es la **capacidad comercial efectiva** a aplicar hacia adelante, no un valor a derivar de lo ya cargado. Los valores actualmente observados en el Excel (`0`, `50.400`, `75.000`, `100.800`) son **datos legacy/históricos** y deben mantenerse para trazabilidad; **no se modifican en Gate 3A** ni se reinterpretan como tiempo físico exacto de encendido.

**HECHO OBSERVADO EN DATOS:** los valores reales de `SegundosDia` en el maestro son exactamente `{0, 50.400, 75.000, 100.800}` — **72.000 no aparece en ningún registro del Excel**; es un valor de referencia comercial a introducir en la capa semántica de Gate 3B, no un dato ya presente. Todos los valores legacy son > 24h en segundos (86.400) salvo 50.400 (14h) — es decir, 75.000s (20,8h) y 100.800s (28h, físicamente imposible como tiempo real de un día) refuerzan la lectura de que `SegundosDia` (legacy) es una capacidad comercial normalizada histórica, no segundos físicos de encendido. **HECHO OBSERVADO EN DATOS.**

**PROPUESTA para Gate 3B (no implementada en Gate 3A):** separar explícitamente (A) **valor fuente/legacy** = `SegundosDia` tal cual está en el Excel (preservado sin borrar, para trazabilidad) de (B) **capacidad comercial efectiva** = proveniente de un perfil/configuración semántica, con `72.000` como default/máximo comercial normalizado donde Brand Plus lo confirme, más soporte de `default + override` por perfil/`ElementoID`, sin borrar nunca el dato original de (A).

**PREGUNTA IMPORTANTE (deseable, no bloqueante):** el origen histórico exacto del cálculo de 50.400/75.000/100.800 (no son múltiplos limpios entre sí: 100.800/50.400=2, pero 75.000 no es múltiplo de ninguno de los otros dos) queda documentado sin interpretar, solo si se considera útil investigarlo — no bloquea Gate 3B porque la regla comercial efectiva (72.000) ya está confirmada independientemente del legacy.

---

## 23. Slots y reels

**HECHO OBSERVADO EN DATOS:** `CapacidadSlotsReel` toma solo los valores `{0, 10, 20, 40}` en el maestro completo (mínimo 0, máximo 40, ver stats de Gate 2). No hay un valor único "20" aplicado a todo `Medio=Digital`; varía por perfil (Sección 21). Esto confirma en datos la regla del prompt de que la arquitectura NO debe asumir capacidad idéntica para todo lo digital, y que ya existen al menos 3 perfiles reales de capacidad (20, 10, 40) más el caso "0" (pendiente de carga). **HECHO OBSERVADO EN DATOS.**

**PREGUNTA IMPORTANTE / relación con `SegundosDia`:** no hay una relación aritmética simple entre `CapacidadSlotsReel` y `SegundosDia` observable en los datos (ej. Totems: 20 slots / 100.800s → 5.040s por slot; Puentes/Triedros: 10 slots / 50.400s → también 5.040s por slot; pero Pantalla LED: 40 slots / 75.000s → 1.875s por slot, una proporción distinta). Es decir, **la relación segundos-por-slot no es constante entre perfiles** (5.040 para Totems/Puentes/Triedros, pero 1.875 para Pantalla LED). Esto es una **PREGUNTA IMPORTANTE PARA GATE 3B**: ¿`SegundosDia` y `CapacidadSlotsReel` son independientes por diseño (cada uno mide algo distinto — total de loops/día vs. slots por loop), o debería existir una relación fija que hoy está rota en Pantalla LED? No se resuelve aquí.

---

## 24. Spots / SalidasVendidas

**REGLA CONFIRMADA POR BRAND PLUS (ya cerrada, no es pregunta):** `SEGUNDOS VENDIDOS = SALIDAS × 1.800`.

**HECHO OBSERVADO EN DATOS:** `DuracionSpotSeg = 10` en 9.141 de 9.503 filas (el resto, 362, es `NaN` y corresponde exactamente a las filas con `ModalidadPauta = "No aplica"` — coincidencia exacta 362=362). `SalidasVendidas = 2` en 9.140 de esas 9.141 filas; **una sola fila tiene `SalidasVendidas = 4`**. No se observó ningún otro valor. **HECHO OBSERVADO EN DATOS.**

Con la regla confirmada, esto implica que el 99,99% de las campañas actuales representan `2 × 1.800 = 3.600` segundos vendidos por día, y una única campaña representa `4 × 1.800 = 7.200` segundos vendidos. Se documenta cómo debería integrarse esta fórmula en Gate 3B (no se implementa): `segundos_vendidos = SalidasVendidas × 1.800`, combinable con `DuracionSpotSeg` para el caso de piezas de duración distinta a 10 segundos (Sección 25) y comparable contra `SegundosDia` del elemento para obtener una fracción de capacidad vendida.

---

## 25. Videos de duración distinta a 10 segundos

**HECHO OBSERVADO EN DATOS:** el 100% de las filas con `DuracionSpotSeg` no vacío tiene el valor `10.0`. **No existe en la base actual ningún spot con duración distinta a 10 segundos.** Por lo tanto, no hay evidencia empírica disponible para inferir la regla de escalado proporcional que el prompt pide evaluar. **HECHO OBSERVADO EN DATOS.**

**PREGUNTA IMPORTANTE PARA GATE 3B** (tal como anticipa el prompt): cuando existan piezas de duración distinta a 10 segundos, ¿el consumo de segundos escala linealmente con `DuracionSpotSeg` (ej. un spot de 20 segundos con 2 salidas consumiría `2 × 1.800 × (20/10) = 7.200` segundos) o existe una tabla de equivalencias distinta? No puede confirmarse ni refutarse con los datos actuales porque no hay ningún caso real que lo evidencie.

---

## 26. Exclusividades

**HECHO OBSERVADO EN DATOS:** `TipoExclusividad` está **vacío en el 100% de las 9.503 filas de CAMPANAS**, a pesar de que `PARAMETROS` define explícitamente los valores `"Día completo"` y `"Bloque horario"` como vocabulario válido. `HoraInicio`/`HoraFin` tienen dato real en solo **3 de 9.503 filas**. **No existe hoy ningún caso real de exclusividad registrado en la base**, ni de día completo ni de bloque horario. **HECHO OBSERVADO EN DATOS.**

Esto significa que toda la Sección de exclusividades del prompt (día completo vs. bloque horario, bloqueo de capacidad correspondiente) es, en el estado actual de los datos, un **diseño para un caso de uso que aún no tiene instancias reales que lo validen**. Se documenta la estructura de columnas disponible (`TipoExclusividad`, `HoraInicio`, `HoraFin`) para Gate 3B, sin construir la matemática, tal como indica el prompt.

---

## 27. Tótems

**HECHO OBSERVADO EN DATOS:** 25 elementos con "Totem" en su descripción (excluyendo el caso estático "Totems Mall x 7", que es un ítem físico distinto, ver Sección 31): 3 en AA2000 (Aeroparque) + 7 en AA2000 (Ezeiza) + 15 en Shoppings Digital (Factory Parque Brown ×3, Factory Quilmes ×3, Factory San Martín ×1, Portal Escobar ×2, Palmas del Pilar ×5, Portal Lomas ×3, Plaza Oeste ×2). Todos con `CapacidadSlotsReel = 20`, `SegundosDia = 100.800`, `TipoInstalacion = "No aplica"`. **100% de los tótems observados hoy usan el mismo perfil de capacidad (20/100.800)** — no se observó ningún tótem con capacidad reducida ni ningún caso de venta por exclusividad (ver Sección 26: 0 exclusividades registradas en toda la base). **HECHO OBSERVADO EN DATOS.**

La "intención de normalizar en 20 como estándar" que confirma el prompt ya coincide al 100% con lo cargado hoy — no hay evidencia de tótems con capacidad distinta a 20 que requieran migración.

---

## 28. Puentes LED de shopping

**HECHO OBSERVADO EN DATOS:** 5 elementos con "Puente" en su descripción, todos en Unicenter (`UNI-PUENTELED-2` a `UNI-PUENTELED-6`), `CapacidadSlotsReel = 10`, `SegundosDia = 50.400`, `Medio = Digital`.

**[Resuelto por Brand Plus — inconsistencia documentada, ya no es pregunta abierta]** La capacidad comercial de los Puentes LED es de **13 espacios/campañas** (**REGLA CONFIRMADA POR BRAND PLUS**). El dato cargado en el maestro es **`CapacidadSlotsReel = 10`** (**REGLA LEGACY / HISTÓRICA** — dato de origen, no se modifica en Gate 3A ni en el Excel). Esto queda documentado como **INCONSISTENCIA ENTRE DATO LEGACY Y REGLA COMERCIAL ACTUAL**, no como pregunta abierta: Gate 3B deberá aplicar un perfil de configuración `PUENTE_LED` con `slots_comerciales = 13` como capacidad comercial efectiva, **manteniendo el valor original `CapacidadSlotsReel = 10` del Excel disponible para trazabilidad** (mismo patrón que la separación valor-legacy/capacidad-efectiva de `SegundosDia`, Sección 22). No se implementa ese perfil en Gate 3A.

---

## 29. Triedros

**HECHO OBSERVADO EN DATOS:** 6 elementos con "Triedro" en su descripción, todos en Shoppings Digital: Palmas del Pilar (1), Plaza Oeste (1), Unicenter (4: Octógono Chico, Garganta Central, Octógono Grande Pasillo Parking, Octógono Grande Pasillo Nivel 3). `CapacidadSlotsReel = 10`, `SegundosDia = 50.400`, `TipoInstalacion = "No aplica"` en el 100% de los casos.

**HECHO OBSERVADO EN DATOS:** ningún triedro tiene registrado un modo "exclusividad completa" distinto de los demás elementos digitales (0 exclusividades en toda la base, Sección 26). Es decir, hoy en los datos **todos los triedros están modelados con el mismo esquema de slots/reel que un puente LED** (misma capacidad 10/50.400), sin distinción de modo de venta. Esto es consistente con la instrucción del prompt de no fijar "TRIEDRO = siempre exclusivo" — en los datos actuales, de hecho, no hay ningún indicio de venta por exclusividad para ningún triedro. **HECHO OBSERVADO EN DATOS.**

---

## 30. Pantallas LED (capacidad — detalle)

Ver Sección 7 para el detalle completo de ubicaciones. En términos de capacidad: **100% uniforme** — `CapacidadSlotsReel = 40`, `SegundosDia = 75.000`, sin ningún override por `ElementoID` observado. Es, junto con YPF Digital, el circuito digital con menor variabilidad interna de capacidad. **HECHO OBSERVADO EN DATOS.**

---

## 31. Inventario flexible / cantidades

**HECHO OBSERVADO EN DATOS:**
- `AplicaCantidad = "SI"` coincide exactamente con `TipoInventario = "Flexible gráfico"` (88 filas, las 88 de Jumbo/Disco) — es una relación 1:1 determinística ya calculada por Gate 2 (`derive_aplica_cantidad`, `transform_data.py:118`), no algo a re-derivar en Gate 3.
- Sin embargo, `CantidadUnidades` en `CAMPANAS` está **100% vacío** (0/9.503) — no existe, en la base actual, ningún registro de cuántas unidades físicas se vendieron/instalaron para ninguna campaña de Cencomedia (ni de ningún otro circuito). **HECHO OBSERVADO EN DATOS.**
- Un mismo `ElementoID` puede representar más de una unidad física: el caso `"Totems Mall"` (`Shoppings Estático`, Portal Tucumán) tiene `Descripcion = "Totems Mall x 7"` — **un solo `ElementoID` agrupando 7 tótems físicos**, con `q` (cantidad) vacío/0 y `m2 = 0`. Esto confirma en datos reales, incluso fuera de Cencomedia, que **1 `ElementoID` no siempre equivale a 1 unidad física instalada** — coincide con la advertencia explícita del prompt. **HECHO OBSERVADO EN DATOS.**

**PREGUNTA IMPORTANTE PARA GATE 3B:** dado que `CantidadUnidades` nunca tiene dato hoy, ¿de dónde se espera que provenga el dato de "cantidad comprada/instalada" para Cencomedia cuando existan campañas reales — un campo nuevo en la carga de campañas, o inferencia desde otro lugar? No puede resolverse con la evidencia actual.

---

## 32. Campañas y temporalidad

**HECHO OBSERVADO EN DATOS:**

| Estado | Filas |
|---|---|
| Activa | 4.991 |
| Finalizada | 4.437 |
| Reservada | 25 |
| (vacío) | 50 |

`PARAMETROS` define también `"Cancelado"` como valor válido de `Estado`, pero **no hay ninguna fila con `Estado = "Cancelado"` en la base actual**. **HECHO OBSERVADO EN DATOS.**

`FechaIndefinida`: `No` (9.443), `Si` (5), vacío (55). `FechaFin` vacío en 60 filas; `FechaInicio` vacío en 55 filas. Las 5 filas `FechaIndefinida="Si"` tienen, por regla de Gate 1 (`validate_dates`, `validate_input.py:686`), `FechaFin` obligatoriamente vacío — esto es una **regla ya validada y forzada por Gate 1**, no una ambigüedad a resolver en Gate 3.

**No se encontró en el código ni en los datos ninguna lógica ya implementada de "Activa/Finalizada/Reservada/Futura" derivada de fechas** — `Estado` es un campo cargado directamente en `CAMPANAS`, no calculado a partir de `FechaInicio`/`FechaFin`/fecha actual por ningún script del repositorio (ni Gate 1 ni Gate 2 lo derivan). **HECHO OBSERVADO EN DATOS** (verificado leyendo el código: ni `validate_input.py` ni `transform_data.py` calculan `Estado`).

**[Resuelto por Brand Plus] `Estado = "Reservada"` bloquea disponibilidad futura:** una reserva representa capacidad comercial comprometida. Para consultas de disponibilidad futura ("¿puedo vender este elemento en estas fechas?"), tanto `Activa` como `Reservada` deben considerarse **ocupación/compromiso**, cada una según sus fechas correspondientes (25 filas con `Estado="Reservada"` en la base actual, Sección 32). Esto **no** implica que ambas deban contarse igual en todos los KPIs comerciales: Gate 3B deberá poder distinguir **actividad ejecutándose** (`Activa`) de **capacidad futura comprometida** (`Reservada`) como dos conceptos relacionados pero distintos — bloqueo de disponibilidad sí, pero no necesariamente el mismo tratamiento en cada métrica. **REGLA CONFIRMADA POR BRAND PLUS.**

**PREGUNTA IMPORTANTE (deseable, no bloqueante):** más allá de la regla de bloqueo ya confirmada arriba, ¿debe Gate 3B además cruzar `Estado` contra `FechaInicio`/`FechaFin`/fecha de hoy para detectar inconsistencias de carga (ej. `Estado="Activa"` con `FechaFin` en el pasado)? Esto es un control de calidad de dato adicional, no bloquea el diseño de bloqueo de disponibilidad ya resuelto.

---

## 33. Programática

**HECHO OBSERVADO EN DATOS:** `PROGRAMATICA`: `No` (9.348), `Si` (112), vacío (43). Cruce con `ModalidadPauta`:

| PROGRAMATICA | No aplica | Slot / Reel normal |
|---|---|---|
| No | 362 | 8.986 |
| Si | 0 | 112 |
| (vacío) | 0 | 43 |

**Toda campaña `PROGRAMATICA = "Si"` tiene `ModalidadPauta = "Slot / Reel normal"`; ninguna es `"No aplica"`.** No hay evidencia de que programática implique una regla de ocupación distinta a la de una campaña normal de slot/reel — se comporta, en los datos, igual que cualquier otra campaña de reel en términos de `ModalidadPauta`. **HECHO OBSERVADO EN DATOS**, consistente con la instrucción del prompt de no asumir una regla especial sin evidencia.

---

## 34. Canje

**HECHO OBSERVADO EN DATOS:** `CANJE`: `No` (9.188), vacío (302), `Si` (13). Es una dimensión comercial minoritaria (13 de 9.503 filas, 0,14%). No se encontró ninguna relación observable entre `CANJE = "Si"` y ninguna otra columna de capacidad/ocupación (no hay una `ModalidadPauta`, `CapacidadSlotsReel` ni `Estado` diferencial asociada a canje en los datos). **HECHO OBSERVADO EN DATOS**, consistente con la instrucción de no asumir que CANJE afecta capacidad/ocupación sin evidencia explícita.

---

## 35. Dimensiones centrales (catálogo propuesto)

**INFERENCIA / PROPUESTA** — catálogo de dimensiones para Gate 3B, organizadas por familia:

**Dimensiones de inventario** (fuente: `MAESTRO_ELEMENTOS`): `ElementoID`, `Ciudad` (224 valores), `Medio` (Digital/Estático), `CircuitoDashboard` (31 valores), `Subcircuito` (459 valores — usado con semántica distinta según circuito: APIE en YPF, cadena en Cencomedia/London Supply, agrupador comercial "CENCOSUD"/"APSA"/"REMEROS" en Shoppings), `Ubicacion`, `Nivel` (fuertemente sobrecargado semánticamente: Indoor/Outdoor en YPF, piso en shoppings, zona de terminal en AA2000, "Carros" en Cencomedia — no es una única dimensión limpia de "nivel de piso"), `TipoInventario` (Digital/Físico estático/Flexible gráfico), `TipoInstalacion` (29 valores, mezcla materiales físicos y "No aplica" para digital), `TipoCatalogo` (Abierto/Cerrado), `Material`, `AplicaCantidad`.

**Dimensiones comerciales** (fuente: `CAMPANAS`): `Cliente` (137 distintos, 3.322 vacíos), `Marca` (266 distintos), `Agencia` (7 distintos, mayoritariamente vacío: 3.490/9.503), `Proveedor` (6 distintos), `Campaña`, `IDCampaña`, `ModalidadPauta` (2 valores en uso: "Slot / Reel normal", "No aplica"), `PROGRAMATICA`, `CANJE`, `TipoExclusividad` (0 valores en uso hoy).

**Dimensiones temporales:** `FechaInicio`, `FechaFin`, `FechaIndefinida`, `HoraInicio`/`HoraFin` (3 filas con dato real en toda la base).

**Dimensiones semánticas / gobernanza (propuestas en este documento, no implementadas):** `CoberturaCatalogo`, **`CompletitudMaestro`** (nueva, Sección 15.1 — distinta de `CoberturaCatalogo`: AA2000 es `COMPLETO` en cobertura pero `PARCIAL` en completitud de carga), `CertezaDato` (resuelta por semántica de circuito, no con una regla global — Sección 16), `ModoDisponibilidad`, `PortfolioTier`, `IncluyePerformanceCore`, **Actividad comercial** (¿tiene campañas? — dimensión propia, independiente de `CertezaDato`, Sección 16.1), más `EstadoValidacion` (ya existente y calculado por Gate 1/2: OK=9.405, PENDIENTE_HISTORICO=80, PENDIENTE_DUPLICADO=11, PENDIENTE_COLISION=7).

---

## 36. Métricas legacy detectadas

**No aplica en este repositorio.** No existe ningún HTML, JavaScript, JSON de dashboard, ni ninguna fórmula de Power BI versionada en `C:\brand plus\ocu26-dashboard`. La única fórmula legacy documentada y verificable es la de `TipoInventario`/`AplicaCantidad`, ya **reimplementada correctamente en Gate 2** (`transform_data.py:98-122`), citada en `audit_sources/INFORME_RECONCILIACION_MIGRACION.md` Sección 7. No hay ninguna otra métrica de ocupación/disponibilidad/performance calculada en ningún archivo de este repositorio para YPF, London Supply, AA2000, Supermercados o APSA. **HECHO OBSERVADO EN DATOS** (relevamiento de archivos, Sección 38).

---

## 37. Métricas centrales propuestas (catálogo)

**INFERENCIA / PROPUESTA** — ninguna de estas métricas está implementada; se documenta su definición conceptual para aprobación antes de Gate 3B.

| Métrica | Definición | Numerador | Denominador | Aplica sin reservas | Aplica con advertencia / estado especial |
|---|---|---|---|---|---|
| `elementos_registrados` | Cantidad de `ElementoID` cargados en el maestro para un filtro dado | — | — | Todos los circuitos | — (es un conteo, no un %) |
| `elementos_confirmados` | Subconjunto de `elementos_registrados` con `CertezaDato = CONFIRMADO` — **la regla de resolución de `CertezaDato` es por circuito** (Sección 16), no una condición única de "tiene campaña" para todo el maestro | count(`CertezaDato=CONFIRMADO`) | — | Todos (con la resolución de `CertezaDato` correspondiente a cada circuito) | — |
| `inventario_total` | Universo comercializable/informado total de un circuito, según `CoberturaCatalogo` | — | — | Pantalla LED, Cencosud, Remeros, **Pilar Frontlight**, **London Supply**, **AA2000** (cobertura `COMPLETO`, Sección 15) | **NO_APLICA** en YPF/APSA/Cencomedia/MAB (cobertura `DESCONOCIDO`) — devolver `REQUIERE_CONFIRMACION`. Nota: `inventario_total` calculable (`CoberturaCatalogo=COMPLETO`) no implica que el maestro ya tenga todo ese universo **cargado** — ver `elementos_registrados` vs. `inventario_total` con `CompletitudMaestro=PARCIAL` en AA2000 (Sección 15.1), donde `elementos_registrados < inventario_total` conocido |
| `elementos_con_actividad` | `ElementoID` con ≥1 fila en `CAMPANAS` en el período | count distinct ElementoID en CAMPANAS | — | Todos | — |
| `elementos_ocupados_calendario` | `ElementoID` con campaña vigente en una fecha/rango dado | count con solapamiento de fecha | — | Todos (requiere Gate 3B) | — |
| `elementos_disponibles` | `elementos_registrados` − `elementos_ocupados_calendario` (solo si capacidad=1 campaña bloquea 100%, lo cual **no es cierto en digital**, Sección 20) | — | — | Estático (Cerrado) | Digital: **PARCIAL** — requiere considerar slots, no solo presencia de campaña |
| `ocupacion_calendario_pct` | `elementos_ocupados_calendario / inventario_total` | — | `inventario_total` | Pantalla LED, Cencosud/Remeros | **NO_APLICA** donde `inventario_total` sea `REQUIERE_CONFIRMACION` |
| `dias_periodo`, `dias_ocupados`, `elemento_dias_disponibles`, `elemento_dias_ocupados` | Base de cálculo de ocupación estática por solapamiento real de fechas (Sección 19) | — | — | Cencosud, Remeros, APSA (si se decide incluir), Cencomedia | — |
| `capacidad_digital_total`, `capacidad_digital_vendida`, `fill_rate_digital` | Basadas en `CapacidadSlotsReel`/`SegundosDia` × campañas vendidas (Sección 20-24) | segundos/slots vendidos | segundos/slots totales del perfil | YPF Digital, Pantalla LED, Shoppings Digital, AA2000 Totems | Elementos con `CapacidadSlotsReel=0` (8 casos, Sección 21): **REQUIERE_CONFIRMACION**, no calcular fill rate con denominador 0 |
| `slots_totales`, `slots_vendidos`, `slots_disponibles` | Directamente desde `CapacidadSlotsReel` y `SalidasVendidas` agregadas | — | — | Igual que fill_rate_digital | Igual que arriba |
| `campanas_activas`, `clientes_activos`, `marcas_activas` | Conteos simples sobre `CAMPANAS` filtrado por `Estado`/fecha | — | — | Todos | — |
| `unidades_estaticas_vendidas` | Suma de `CantidadUnidades` para elementos `AplicaCantidad=SI` | sum(`CantidadUnidades`) | — | **NO_APLICA hoy** — `CantidadUnidades` está 100% vacío en la base actual (Sección 31) | — |

---

## 38. HTML actuales y lógica a retirar

**Relevamiento del repositorio realizado con `find`/`Glob` sobre todo el árbol, excluyendo `.venv/` y `.git/`.**

**Resultado: no existe ningún archivo `.html`, `.js` ni `.json` de aplicación en este repositorio.** El repositorio contiene únicamente:

```
audit_sources/   (2 .xlsx + 1 .md de auditoría — no HTML/JS)
input/           (1 .xlsx — el input productivo)
scripts/         (validate_input.py, transform_data.py)
tests/           (test_validate_input.py, test_transform_data.py)
requirements.txt
```

**No se inventó ningún HTML ni se buscó fuera del repositorio, conforme a la instrucción explícita del prompt.** Por lo tanto, las Secciones que pedían "análisis de HTML/vistas actuales" (pregunta de negocio, KPIs, dimensiones, filtros, cruces, cálculos, lógica embebida, supuestos, métricas repetidas, qué migrar) **no tienen material sobre el cual aplicarse**: no hay lógica de negocio embebida en ningún HTML para retirar, porque no hay ningún HTML. Si Brand Plus tiene HTML/Power BI de OCU26 fuera de este repositorio (por ejemplo en otra carpeta, otro repo, o Power BI Service), **deben incorporarse al repositorio o señalarse su ubicación explícitamente** antes de poder auditarlos en una futura vuelta de Gate 3.

---

## 39. Cruces futuros

**INFERENCIA / PROPUESTA** — con las dimensiones y métricas de las Secciones 35 y 37, todos los cruces de ejemplo que pide el prompt son expresables como `(métrica, group_by[], filters[])` sin funciones dedicadas, siempre que la matriz de aplicabilidad (Sección 40) se respete:

| Cruce pedido | Expresión conceptual | Nota de aplicabilidad |
|---|---|---|
| YPF × Ciudad | `metric=elementos_con_actividad, group_by=[Ciudad], filters=[CircuitoDashboard IN (YPF Digital, YPF Estático)]` | Directo |
| YPF × Tipo de soporte | `group_by=[Medio o TipoInventario], filters=[CircuitoDashboard IN (YPF Digital, YPF Estático)]` | Directo — separar Digital/Estático es obligatorio (Sección 13.5) |
| Cliente × Ciudad × Medio | `group_by=[Cliente, Ciudad, Medio]` | Directo, sin filtro de circuito |
| Circuito × TipoInventario | `group_by=[CircuitoDashboard, TipoInventario]` | Directo |
| Marca × Mes | `group_by=[Marca, mes(FechaInicio)]` | Requiere definir "mes" de una campaña con fechas reales (Sección 19), no asumir 1 mes calendario |
| Agencia × Medio | `group_by=[Agencia, Medio]` | Directo — pero 3.490/9.503 filas no tienen `Agencia` |
| PROGRAMATICA × Circuito | `group_by=[CircuitoDashboard], filters=[PROGRAMATICA="Si"]` | Directo |
| CANJE × Cliente | `group_by=[Cliente], filters=[CANJE="Si"]` | Directo, pero solo 13 filas totales |
| Shopping × Mes × Medio | `group_by=[Ubicacion, mes, Medio], filters=[Subcircuito="CENCOSUD"]` | Requiere excluir explícitamente APSA/Remeros del mismo `CircuitoDashboard` |
| Circuito × ocupación calendario | `metric=ocupacion_calendario_pct, group_by=[CircuitoDashboard]` | **NO_APLICA** o `REQUIERE_CONFIRMACION` para circuitos con `CoberturaCatalogo≠COMPLETO` (Sección 15) |
| Circuito × fill rate digital | `metric=fill_rate_digital, group_by=[CircuitoDashboard], filters=[Medio="Digital"]` | Excluir los 8 elementos con capacidad 0 del denominador o marcarlos aparte |
| Cliente × circuito × período | `group_by=[Cliente, CircuitoDashboard, periodo]` | Directo |

Ninguno de estos cruces requiere una función `calcular_X()` dedicada si Gate 3B implementa el motor genérico descripto en la Sección 44.

---

## 40. Metadata semántica propuesta

**REGLA CONFIRMADA POR BRAND PLUS (preferencia arquitectónica inicial a evaluar, no a fijar sin análisis):** Excel = hechos operativos, configuración semántica externa = reglas de negocio, código = motor genérico.

**Evaluación de las tres alternativas** (INFERENCIA / PROPUESTA):

| Alternativa | Mantenibilidad | Auditoría | Escalabilidad ante nuevos circuitos | Riesgo |
|---|---|---|---|---|
| A. Agregar columnas al Excel (ej. `PortfolioTier` como columna de `MAESTRO_ELEMENTOS`) | Baja: cualquier cambio de regla de negocio requiere editar el Excel productivo, que Gate 1 valida estructuralmente (headers exactos, Sección de `validate_input.py`) — un cambio de columna rompe Gate 1 si no se actualiza el validador en paralelo | Alta: el dato queda versionado junto con el hecho operativo, pero mezcla hecho con regla | Baja: cada circuito nuevo requiere decidir manualmente el valor en cada fila | Alto: el Excel productivo no está pensado como superficie de configuración; el prompt mismo prohíbe tocar `input/` |
| B. Metadata/configuración externa (ej. `config/business_semantics.yaml`, **no creado en Gate 3A**) | Alta: cambiar `IncluyePerformanceCore` de Cencomedia es editar una línea de configuración, sin tocar Gates 1/2 ni el Excel | Alta: diff de configuración auditable en control de versiones, separado de los datos | Alta: reglas por `CircuitoDashboard`/`Subcircuito` se agregan como entradas nuevas, sin tocar código | Medio: requiere que Gate 3B sepa resolver "elemento no tiene entrada en config" con un default explícito, no un fallo silencioso |
| C. Reglas hardcodeadas en Python (ej. `if circuito == "London Supply": ...`) | Muy baja: exactamente el antipatrón `calcular_ypf()`/`calcular_cliente()` que el prompt prohíbe explícitamente | Baja: la regla de negocio queda enterrada en lógica de control de flujo | Muy baja: cada circuito nuevo exige una nueva rama de código y un despliegue | Alto: es la opción que el prompt pide evitar de forma explícita |

**Conclusión propuesta:** la preferencia inicial del prompt (B, con Excel para hechos operativos) es la más consistente con el principio de "libertad total de cruces" y "crecimiento del maestro sin rehacer dashboards". La alternativa A queda descartada porque el prompt prohíbe modificar `input/` y porque mezclaría regla de negocio con hecho operativo dentro de un archivo que Gate 1 valida estructuralmente de forma estricta. **No se crea ningún archivo de configuración en Gate 3A**, conforme a la instrucción.

---

## 41. Escalabilidad del maestro

**REGLA CONFIRMADA POR BRAND PLUS:** el maestro debe poder crecer indefinidamente sin rehacer dashboards ni métricas; ninguna cantidad actual debe hardcodearse.

**Verificación en el código actual (HECHO OBSERVADO EN DATOS, leído directamente):** `transform_data.py` y `validate_input.py` **no hardcodean ninguna cantidad de filas**. `read_excel_table()` (`transform_data.py:65`) lee el rango real de la tabla de Excel (`ws.tables[table_name]`, `range_boundaries(table.ref)`), no un número de filas fijo. Los únicos "números" hardcodeados en el código son **nombres de hojas/tablas/columnas y vocabularios de dominio** (ej. `TIPO_INVENTARIO_VALUES`, `MEDIO_VALUES`), no cantidades. Esto significa que agregar filas nuevas y válidas a `tblElementos` **ya fluye correctamente por Gate 1 y Gate 2 sin cambios de código**, tal como exige el prompt. **HECHO OBSERVADO EN DATOS.**

**Casos que sí exigirían modificación de configuración o código** (a diferencia de simplemente agregar filas):
- Un nuevo `CircuitoDashboard` que Brand Plus quiera incluir en `PortfolioTier=CORE` requiere una entrada nueva en la configuración semántica de Gate 3B (Sección 40), no código nuevo — **si** Gate 3B se construye como se propone.
- Un nuevo perfil de capacidad digital (ej. una pantalla con `CapacidadSlotsReel=60`) no requiere cambios si el motor de Gate 3B lee capacidad directamente de `MAESTRO_ELEMENTOS` en vez de mapear por nombre de perfil — es una decisión de diseño a confirmar en Gate 3B, no un hecho ya resuelto.
- Un nuevo valor de `Medio`, `TipoInventario` o `TipoCatalogo` **sí** requeriría actualizar los vocabularios cerrados de `validate_input.py` (`MEDIO_VALUES`, `TIPO_INVENTARIO_VALUES`, etc.) — eso es Gate 1, fuera del alcance de Gate 3, y es una decisión deliberada de Gate 1 (dominio cerrado y auditado), no una limitación a resolver en Gate 3.

---

## 42. Riesgos de interpretación

**INFERENCIA / PROPUESTA**, a partir de los hechos observados en este documento:

1. **Riesgo de denominador mezclado:** `CircuitoDashboard = "Shoppings Estático"` mezcla Cencosud+Remeros (Cerrado, con actividad) y APSA (Abierto, 0 actividad) en el mismo valor. Cualquier métrica que agrupe solo por `CircuitoDashboard` sin descomponer por `Subcircuito` heredará automáticamente este riesgo — es el mismo riesgo que el prompt describe para London Supply, pero aplicado a un circuito ("Shoppings Estático") que a primera vista parece homogéneo y no lo es.
2. **Riesgo de "Cencomedia invisible":** como no existe ningún valor de dato que agrupe las 22 ubicaciones Jumbo/Disco, cualquier consulta futura por "Cencomedia" requiere una lista explícita de 22 valores de `CircuitoDashboard` (o un patrón `startswith("Jumbo")/startswith("Disco")`, que es fragilísimo si Brand Plus agrega una cadena nueva con otro nombre de marca). Esto refuerza la necesidad de la capa de metadata semántica (Sección 40) antes de escalar Cencomedia.
3. **Riesgo de "0 campañas" leído como "0 disponibilidad":** London Supply, APSA, AA2000 y Cencomedia tienen hoy 0 campañas. Sin la dimensión `CertezaDato`/`CoberturaCatalogo`, un dashboard ingenuo podría reportar "0% de ocupación" en estos circuitos, cuando la interpretación correcta (por regla confirmada del prompt) es "sin actividad registrada en este período" o "fuera del performance core", no "vacío/malo".
4. **Riesgo de doble conteo geográfico de Remeros** (Sección 6): Remeros aparece en Shoppings Estático/Digital y en Pantalla LED. Sumar "elementos totales de Remeros" sin desambiguar estas dos fuentes duplicaría parcialmente el conteo si se buscara solo por texto "Remeros" en `Ubicacion`/`Ciudad` sin considerar `CircuitoDashboard`.
5. **Riesgo de capacidad-cero como disponibilidad real:** los 8 elementos digitales con `CapacidadSlotsReel=0` (Sección 21) no significan "0% disponible" ni "100% disponible" — significan "capacidad no cargada"; deben excluirse de cualquier fill-rate o marcarse `REQUIERE_CONFIRMACION`, nunca promediarse como si fueran 0 de capacidad real.
6. **Riesgo de confundir `CoberturaCatalogo` con `CompletitudMaestro`** (Sección 15/15.1): AA2000 es el ejemplo directo — "conocemos el universo" (`COMPLETO`) no equivale a "ya está todo cargado" (`PARCIAL`, faltan Mendoza/Córdoba). Un dashboard que trate ambas preguntas como una sola podría concluir erróneamente que AA2000 tiene cobertura incompleta cuando en realidad Brand Plus sí conoce su universo comercial completo; el problema es de carga, no de conocimiento.
7. **Riesgo de generalizar la regla de `CertezaDato` de YPF a todo el maestro** (Sección 16): "elemento con campaña = confirmado" es una regla **específica de YPF**, no un principio general. Aplicarla sin distinción a Cencosud, Remeros, Pantallas LED, Pilar Frontlight o AA2000 marcaría como "no confirmados" miles de elementos de catálogos cerrados y ya validados simplemente por no tener actividad comercial en la ventana de datos actual — exactamente el error que esta corrección busca prevenir.

---

## 43. Preguntas pendientes

**Revisada en la corrección de 2026-08-07.** Se eliminaron de esta sección todas las preguntas ya resueltas por decisión confirmada de Brand Plus (granularidad de `IncluyePerformanceCore`, YPF Digital vs. Estático a nivel conceptual, MAB, Pilar, Pantalla LED Córdoba/10 vs. 11, Puente LED 10 vs. 13, ocupación estática mismo día, `Reservada` y bloqueo futuro, 72.000 segundos como referencia comercial, certeza YPF cuando existe campaña, London Supply `ModoDisponibilidad`). Esas decisiones están incorporadas en las Secciones 4, 7, 11, 12, 13.5-13.7, 15, 17, 18, 19, 22, 28 y 32. Solo se mantienen a continuación las preguntas que **siguen genuinamente sin resolver**.

### BLOQUEANTE PARA GATE 3B

Ninguna pregunta bloqueante permanece abierta tras esta corrección. Todas las que impedían fijar el diseño de metadata semántica (granularidad de `IncluyePerformanceCore`, capacidad de Puentes LED, conteo de Pantalla LED, clasificación de MAB/Pilar) quedaron resueltas.

### IMPORTANTE

- **[Videos >10 segundos]** No hay ningún caso real en la base hoy; la fórmula de escalado de segundos para piezas de duración distinta a 10 segundos queda sin poder validarse empíricamente (Sección 25). Se mantiene como pregunta IMPORTANTE, tal como el prompt de corrección permite explícitamente.
- **[Exclusividades — matemática]** 0 casos reales de exclusividad en la base hoy; la matemática exacta de bloqueo de capacidad por franja horaria/día completo queda sin poder validarse empíricamente cuando existan casos reales (Sección 26).
- **[Ocupación digital / slots-reel]** Relación exacta entre `CapacidadSlotsReel` y `SegundosDia` (legacy) — no es una proporción constante entre perfiles (5.040 seg/slot en Totems/Puentes/Triedros vs. 1.875 seg/slot en Pantalla LED, Sección 23). No bloquea Gate 3B porque la capacidad comercial efectiva (72.000, Sección 22) se define independientemente del legacy, pero conviene entender el origen antes de fijar los perfiles de configuración.
- **[Segundos — origen histórico]** Origen histórico del cálculo de 50.400/75.000/100.800 en el Excel legacy — no son múltiplos limpios entre sí (Sección 22). Se mantiene como pregunta IMPORTANTE solo si Brand Plus considera útil investigarlo; no bloquea Gate 3B.
- **[Cencomedia]** Dado que hay 0 campañas y `CantidadUnidades` 100% vacío en toda la base, ¿de dónde provendrá el dato real de "cantidad vendida" cuando arranque la comercialización — un campo nuevo en la carga de campañas, u otra fuente? (Sección 31)
- **[Campañas / temporalidad — control de calidad]** ¿Debe Gate 3B, además de aplicar la regla de bloqueo ya confirmada para `Reservada` (Sección 32), cruzar `Estado` contra fechas para detectar inconsistencias de carga (ej. `Estado="Activa"` con `FechaFin` en el pasado)? Es un control de calidad adicional, no bloquea el diseño ya resuelto.

### DESEABLE

- Confirmar los 8 mapeos "MAPEO PROBABLE" de Shoppings Cencosud (P.ROSARIO, P.TUCUMAN, P.SANTIAGO, P.SALTA, P.TRELEW, P.LOSANDES, P.PATAGONIA, F.P.BROWN) contra el nombre comercial exacto (Sección 5).
- Corregir (fuera de Gate 3A) la inconsistencia de escritura `F.SAN MARTIN` vs. `F.SANMARTIN` — no es bloqueante, solo higiene de dato.
- Aclarar el origen de los códigos internos `PPAR` (→ Cerrito) y `PTRI` (→ Olazábal) en Pantalla LED, que no coinciden con el nombre de ubicación actual (Sección 7).
- **[Resuelto — eliminada de esta lista]** La pregunta "¿AA2000 es `PARCIAL` o `COMPLETO` en `CoberturaCatalogo`?" queda resuelta: `CoberturaCatalogo = COMPLETO` (el universo comercial es conocido) y `CompletitudMaestro = PARCIAL` (Mendoza/Córdoba aún no cargados) son dos dimensiones separadas, no una disyuntiva (Sección 15/15.1).
- **[Nueva, no bloqueante]** ¿Cuál es la regla de negocio definitiva para `CertezaDato` de Cencomedia (Sección 16)? Se dejó explícitamente sin asignar en Gate 3A porque requiere una regla propia futura, no derivable de los datos actuales.

---

## 44. Arquitectura propuesta Gate 3B (opinión técnica, no implementada)

**INFERENCIA / PROPUESTA — no se programa nada de esto en Gate 3A.** Actualizada tras la corrección: no quedan preguntas bloqueantes (Sección 43), por lo que esta arquitectura ya puede tomarse como base de diseño para Gate 3B.

La arquitectura recomendada tiene tres capas:

```
FUENTE DE HECHOS: transform_data()
        +
METADATA SEMÁNTICA CONFIGURABLE (resolución jerárquica)
        +
MOTOR GENÉRICO DE MÉTRICAS
```

1. **Fuente de hechos:** `transform_data()` sigue siendo la única fuente de datos de Gate 3B, sin reimplementar nada de Gate 1/Gate 2.
2. **Metadata semántica configurable**, separada del Excel y del código (alternativa B de la Sección 40), con al menos cinco tablas: `circuito_semantica` (→ `CoberturaCatalogo`, **`CompletitudMaestro`** (Sección 15.1 — AA2000 es el caso que exige esta columna: `COMPLETO`/`PARCIAL` independiente de `CoberturaCatalogo`), `ModoDisponibilidad`, `PortfolioTier`, `IncluyePerformanceCore`/`IncluyeConteoGeneral`/`VisiblePorDefecto`, Sección 12/18), `certeza_dato_regla` (por circuito, **no una condición global** — ej. YPF: "campaña→CONFIRMADO"; Cencosud/Remeros/Pantallas LED/Pilar Frontlight/AA2000: "cargado y validado→CONFIRMADO"; London Supply: "informado y conocido→CONFIRMADO"; Cencomedia: sin regla fijada, Sección 16), `perfil_capacidad_digital` (perfil → slots_comerciales/segundos_comerciales default + override, ej. `PUENTE_LED: slots_comerciales=13`, Sección 22/28), `dimension_catalogo` y `metrica_catalogo` (Sección 37).
3. **Resolución jerárquica de metadata (REGLA CONFIRMADA POR BRAND PLUS, Sección 18.1)** — toda consulta a la metadata semántica se resuelve en este orden, de más a menos específico:
   ```
   1. override por ElementoID
   2. regla por CircuitoDashboard + Subcircuito
   3. regla por CircuitoDashboard
   4. default semántico
   ```
   Esto es lo que permite separar `CENCOSUD`/`REMEROS`/`APSA` dentro de `"Shoppings Estático"` sin bifurcar el motor por circuito.
4. **Motor de consulta genérico** con la forma `resolver(metric, group_by[], filters[])` (Sección 39), que resuelve dimensiones sin funciones por circuito, consulta la matriz de aplicabilidad antes de calcular (devolviendo `NO_APLICA`/`REQUIERE_CONFIRMACION` en vez de un número engañoso) y nunca hardcodea un `CircuitoDashboard`, `ElementoID` ni cantidad de filas.
5. **Separación estricta Digital/Estático** en cualquier métrica de ocupación, incluso dentro de un mismo circuito de negocio o de un mismo sitio (YPF, Sección 13.5-13.7; Pilar Digital/Estático, Sección 7/9).
6. **Tratamiento explícito de "capacidad no cargada" (`CapacidadSlotsReel=0`) como estado propio**, no como 0% ni 100% de disponibilidad.
7. **Separación valor-legacy / capacidad-efectiva** para `SegundosDia` y `CapacidadSlotsReel` (Sección 22/28): el motor lee el perfil comercial configurado (ej. `72.000` segundos, `13` slots en Puente LED) sin borrar ni sobrescribir el valor original del Excel, que permanece disponible para trazabilidad.

**Cambios futuros que la metadata debe absorber sin reescribir métricas ni dashboards** (lista confirmada, no exhaustiva): Cencomedia entra/sale del core; un circuito cambia de `PortfolioTier`; cambia la capacidad de un perfil (ej. Puente LED); aparece un nuevo `ElementoID`; aparece una nueva estación YPF; cambia la `CertezaDato` de un elemento. Todos estos casos se resuelven editando la metadata en el nivel jerárquico correspondiente (punto 3), no modificando Gates 1/2, el motor ni ninguna vista — reafirmando el principio rector: **una lógica central → múltiples cruces y vistas.**

---

## Validación final

**Actualizada tras la corrección de 2026-08-07.** Esta corrección fue exclusivamente editorial sobre el documento: no se volvió a ejecutar `transform_data()` ni se relevaron datos nuevos, conforme a la instrucción de no repetir el relevamiento.

- Gate 1 (`scripts/validate_input.py`): **intacto**, no modificado en esta corrección.
- Gate 2 (`scripts/transform_data.py`): **intacto**, no modificado en esta corrección.
- `input/OCU26_BASE_DATOS.xlsx`: **intacto**, no modificado en esta corrección.
- `audit_sources/`, `tests/`, `requirements.txt`, `.gitignore`: **intactos**, no modificados en esta corrección.
- No existen HTML/JS en el repositorio (Sección 38); no hay nada adicional que verificar en esa categoría.
- Único archivo modificado en esta corrección: `docs/GATE3_SEMANTICA_NEGOCIO_PROPUESTA.md` (no se creó ningún archivo nuevo).
- No se ejecutó ningún `git add`, `git commit`, `git push` ni se abrió ningún PR.

```
git status --short
git diff --stat
```

se ejecutan y se reportan tal cual en la respuesta final de esta sesión.
