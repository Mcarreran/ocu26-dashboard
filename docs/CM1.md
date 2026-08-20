# CONTEXTO MAESTRO DEL PROYECTO

**Proyecto:** BRAND PLUS · OCU26 · Sistema de inteligencia comercial, ocupación y dashboards  
**Estado documentado:** 10 de agosto de 2026, aproximadamente 16:14 ART  
**Repositorio principal:** `C:\brand plus\ocu26-dashboard`  
**Branch vigente:** `main`  
**Fuente local vigente:** `input/OCU26_BASE_DATOS.xlsx`  
**Estado de entrega:** las seis TVs productivas existen y fueron mostradas juntas en su estado final de trabajo; TV5 quedó visualmente aprobada después de reemplazar el falso fondo cartográfico por geometría real de Natural Earth.  
**Prioridad inmediata después de esta migración:** preservar el estado actual, hacer futuras ediciones estéticas TV por TV en su sesión de Claude Code correspondiente, no reauditar lógica durante retoques visuales, y luego ejecutar la Etapa 2 de reconciliaciones/microauditorías ya documentada. Power BI sigue pausado.  
**Finalidad:** que un nuevo ChatGPT/Claude Code pueda continuar desde el estado exacto actual sin releer este chat.

> **REGLA DE VIGENCIA ABSOLUTA — 10/8/2026 ~16:14 ART:** las secciones 0–45 de esta consolidación son la capa vigente. El anexo histórico conserva íntegramente el CM1 anterior (~08:22 ART) y todas sus capas previas para trazabilidad. Cuando exista contradicción, prevalece la decisión posterior documentada en estas secciones actuales. No usar estados viejos de TV4/TV5/TV6 ni numeraciones históricas si contradicen la arquitectura final vigente.

---

## 0. Cómo utilizar este documento

Este archivo es la migración acumulada del proyecto OCU26 desde varios chats de ChatGPT y sesiones de Claude/Claude Code. Debe leerse como fuente de verdad inicial antes de modificar código, datos, dashboards, publicación o arquitectura.

Reglas de uso:

1. Leer primero estas secciones 0–45.
2. Usar el **Anexo histórico** solamente para recuperar trazabilidad, detalle o decisiones previas no repetidas aquí.
3. Si una regla histórica contradice una regla actual, usar la actual.
4. No reconstruir componentes ya funcionales por preferencia técnica.
5. No tocar Gates, Excel, semántica central, otras TVs, Git o deploy durante una edición visual salvo necesidad real y autorización.
6. Para futuras ediciones estéticas, volver a la **sesión de Claude Code correspondiente a cada TV** porque conserva mejor el contexto específico y reduce consumo de tokens.
7. En esas sesiones, usar como encuadre: `solo ajustes estéticos; lógica y métricas congeladas salvo error evidente`.
8. Los datos desconocidos no se convierten en cero: conservar `MetricStatus`.
9. Toda cifra de referencia legacy en HTML de auditoría es visual, no fuente de verdad.
10. Antes de publicar, revisar Git, tests pertinentes y compatibilidad 1920×1080/webOS.

---

## 1. Objetivo general del proyecto

### Qué estamos construyendo

Un sistema único de inteligencia comercial para Brand Plus que toma una base Excel validada, aplica lógica central de negocio y produce múltiples salidas consistentes:

- dashboards HTML para 6 TVs;
- data mart/outputs reutilizables;
- Power BI como capa de validación/análisis posterior;
- futura alimentación desde SharePoint/Microsoft 365;
- futura publicación/automatización sin depender de una PC encendida.

### Por qué

El problema original era tener información comercial/ocupación distribuida entre bases, fórmulas, lógicas manuales y tableros que podían divergir. La solución busca:

- una sola base validada;
- una sola lógica de negocio;
- múltiples vistas sin duplicar cálculos;
- operación simple;
- capacidad de reemplazar el Excel por una versión con igual estructura;
- dashboards aptos para TVs/CMS.

### Para quién

- Brand Plus;
- equipo de Producto/Gerencia;
- usuarios internos que necesitan lectura ejecutiva del inventario, ocupación, pipeline y demanda;
- presentación en pantallas 1920×1080.

### Resultado final actual

Seis HTML productivos:

1. TV1 — **Visión general del negocio**
2. TV2 — **Core Comercial Digital**
3. TV3 — **Core Comercial Estático**
4. TV4 — **Pipeline Comercial**
5. TV5 — **Pulso comercial YPF**
6. TV6 — **Así se comporta la demanda**

### Entregables acumulados

- pipeline/Gates;
- modelo semántico y MetricsEngine;
- outputs/data mart;
- builders + payloads + templates + HTML TV1–TV6;
- tests específicos por dashboard;
- referencias visuales read-only;
- assets cartográficos locales de Argentina para TV5;
- CM1;
- guía PDF final simple de las seis TVs: `GUIA_FINAL_TABLEROS_OCU26.pdf` (generada en ChatGPT; no asumir que está dentro del repo hasta comprobarlo);
- infraestructura Git/GitHub y plan de publicación;
- documentación de Power BI y SharePoint para etapas posteriores.

---

## 2. Resumen de la evolución del proyecto

### Etapa inicial

Se partió de una base de ocupación/elementos/campañas con fórmulas y de la necesidad de construir Power BI y tableros para Brand Plus.

### Primeros cambios

Se decidió abandonar la idea de que cada salida tuviera su propia lógica. Se construyó una arquitectura central:

`Excel → validación → transformación → semántica/MetricsEngine → outputs → salidas`

Se desarrollaron Gates 1–4B, se validó la base y se generó un data mart.

### Prioridad Power BI → HTML

Power BI quedó especificado y parcialmente trabajado, pero apareció el error:

`El argumento 'dataType' no puede ser nulo. Nombre del parámetro: dataType`

Por urgencia de presentación y restricciones de licencia, Power BI se pausó y los HTML para TV pasaron a prioridad.

### Cambio de arquitectura de TVs

La numeración cambió varias veces. La **numeración final vigente** es:

- TV1 General
- TV2 Digital
- TV3 Estático
- TV4 Pipeline
- TV5 YPF
- TV6 Demanda

Cualquier historial que diga TV4=YPF o TV5=Demanda queda reemplazado.

### TV1–TV3

Se consolidó un sistema visual común:

- header Brand Plus;
- tarjetas superiores;
- paneles de análisis;
- footer `LECTURA | PUNTO POSITIVO | A ATENDER`;
- comparación con mes anterior;
- uso de YTD donde aporta;
- entero principal + porcentaje secundario cuando existe denominador interpretable.

### TV4 Pipeline

Se construyó después de TV3 y quedó cerrado para entrega inmediata. No usa comparación mensual como eje principal: trabaja con estado al corte y ventanas de 30 días.

### Cambio de orden TV6 antes que TV5

Por decisión operativa, después de TV4 se implementó TV6 Demanda antes que TV5 YPF.

### TV6 Demanda

La primera lectura mensual de la parte inferior no mostraba suficiente profundidad de agencias/programática. Se adoptó modelo híbrido:

- tarjetas superiores = julio 2026;
- cuerpo inferior = acumulado Ene–Jul 2026;
- ranking dinámico marcas/agencias/programática;
- matriz por circuito acumulada.

La captura final ya muestra valores distintos de la primera preview; los valores vigentes están en la sección 23.

### TV5 YPF

Se construyó con KPIs de catálogo/actividad y un mapa territorial.

Primer intento de mapa:
- puntos lat/lon reales;
- fondo/silueta aproximado manualmente;
- visualmente no se reconocía Argentina/AMBA/CABA;
- **DESCARTADO**.

Solución final:
- descarga puntual autorizada de Natural Earth 1:10m;
- Admin 0 real de Argentina + Admin 1 provincias;
- filtrado local;
- assets GeoJSON dentro del repo;
- sin dependencia de internet en runtime;
- puntos YPF sobre cartografía real;
- usuario aprobó la versión final.

### Cierre de las seis TVs

El 10/8 se mostraron juntas las seis capturas finales. Después se generó un PDF explicativo simple y se acordó que los retoques estéticos futuros se harán TV por TV en sus sesiones correspondientes, sin reauditar lógica ni gastar tokens innecesariamente.

---

## 3. Estado actual exacto

### Terminado

- Gate 1 validación.
- Gate 2 transformación.
- Gate 3A/3B semántica/MetricsEngine.
- Gate 4A/4B outputs/data mart.
- TV1 productiva.
- TV2 productiva.
- TV3 productiva.
- TV4 productiva.
- TV5 productiva.
- TV6 productiva.
- mapa TV5 con cartografía real local.
- guía PDF ejecutiva/glosario de TVs generada.
- referencias visuales y patrón de branding consolidados.

### Funcionando

- `./.venv/Scripts/python.exe`.
- builders TV1–TV6.
- payloads `output/tvN_data.json`.
- templates y HTML raíz.
- pipeline central.
- assets cartográficos TV5 offline.
- visualización 1920×1080 sin necesidad de red externa para el mapa TV5.

### Validado

- Excel fuente con SHA esperado histórico:
  `2f165e12ad90c2f05963367a8e1717dcd1006e9ce669f976afa6e57470aca2cd`
- snapshot base previo documentado:
  - Maestro: 4.338 filas
  - Campañas: 9.503 filas
  - Parámetros: 23
- Gates 1–4B completados.
- baseline previo: 199 tests.
- Gate4B: 8/8.
- 5 tablas Parquet + manifest.
- TV1: 39/39 tests específicos confirmados; en una validación se mostró suite completa 238/238.
- TV2: 37/37 confirmados.
- TV4: 35/35 confirmados.
- TV5: último resultado confirmado **34/34 antes de los cambios geográficos finales**.
- TV3: preview funcional aceptada; resultado textual final de pytest no quedó capturado → `A VALIDAR`.
- TV6: versión final visual aceptada; resultado textual final de tests híbridos no quedó capturado → `A VALIDAR`.

### En desarrollo

No hay una TV nueva en desarrollo. Lo próximo es refinamiento/Etapa 2, no reconstrucción.

### Pendiente

- retoques estéticos puntuales TV por TV;
- Etapa 2 de reconciliación/microauditorías;
- tests específicos finales de TV5 después de integración geográfica si se desea evidencia formal;
- confirmación de tests finales TV6;
- confirmación documental de tests TV3;
- `git status`/`git diff` final;
- commit/push solo con aprobación;
- deploy/publicación;
- prueba física en TV/webOS;
- Power BI;
- SharePoint/automatización.

### Bloqueado

- Power BI por error `dataType` hasta retomarlo.
- No hay bloqueo para los HTML.

### A validar

1. TV1: coherencia semántica del porcentaje YPF `61,6% del total` frente al patrón de “ocupación sobre universo propio”; no cambiar sin decisión explícita.
2. TV1 vs TV3: denominador estático `29,4%` en TV1 vs `28,2%` en TV3; reconciliar en Etapa 2.
3. TV2: Shoppings Digital ~68 elegibles actuales vs ~61 recordados.
4. TV2: clasificación de Remeros.
5. TV3: `Otro` = 109/119 en soportes, clasificación demasiado amplia.
6. TV3: semántica exacta del histórico visible y denominadores.
7. TV3: resultado textual de pytest.
8. TV5: tests específicos posteriores al cambio de GeoJSON.
9. TV5: cobertura geográfica 79/305 (25,9%) es parcial.
10. TV5: integración futura de APIE/ID real en capa central.
11. TV6: resultado textual final de tests.
12. TV6: 377 activaciones con agencia `A Confirmar`.
13. TV6: 112 activaciones programáticas YTD sin agencia imputada.
14. reconciliación transversal TV1–TV6.
15. Git status/diff exacto actual.
16. deploy final y prueba física.
17. Power BI `dataType`.

---

## 4. Decisiones vigentes

### Decisión vigente: una sola lógica central
**Origen / contexto:** evitar divergencias entre HTML y Power BI.  
**Por qué:** consistencia y mantenibilidad.  
**Reemplaza:** cálculos independientes por salida.  
**Consecuencia:** HTML consume payload; no se convierte en motor de negocio.  
**Estado:** confirmada.

### Decisión vigente: Excel fuente protegido
**Origen:** Gates.  
**Consecuencia:** no editar `input/OCU26_BASE_DATOS.xlsx` desde builders.  
**Estado:** confirmada.

### Decisión vigente: APSA y London Supply excluidos
**Origen:** redefinición de universo.  
**Consecuencia:** fuera de numeradores, denominadores, rankings, composición y payloads estándar.  
**Estado:** confirmada.

### Decisión vigente: Cencomedia no se fuerza dentro del Core
**Origen:** inventario flexible.  
**Consecuencia:** puede aparecer en catálogo ampliado/estático cuando corresponda, pero no inventar capacidad.  
**Estado:** confirmada.

### Decisión vigente: YPF no tiene fill rate total
**Origen:** no se conoce la capacidad real de slots del CMS YPF.  
**Consecuencia:** medir estaciones, campañas, activaciones, elementos/formato; no fabricar fill.  
**Estado:** confirmada.

### Decisión vigente: TV4 = Pipeline; TV5 = YPF; TV6 = Demanda
**Reemplaza:** roadmaps antiguos que asignaban TV4 a YPF.  
**Estado:** confirmada.

### Decisión vigente: comparación `vs` = mes anterior
**Aplicación:** TV1–TV3 y TV5 cuando tiene sentido.  
**Excepción:** TV4 usa corte/ventana temporal; no necesita forzar junio.  
**Estado:** confirmada.

### Decisión vigente: YTD
`YTD` = referencia del indicador “year to date”, es decir, lo acumulado/nivel calculado en lo que va del año hasta el mes vigente. En julio, Ene–Jul 2026.  
No asumir una fórmula distinta de la codificada por el builder sin revisar el cálculo específico.  
**Estado:** confirmada.

### Decisión vigente: `pp`
`pp` = puntos porcentuales.  
Ejemplo aceptado por el usuario: 30,5% → 29,4% = `-1,1 pp`, no `-1,1%`.  
**Estado:** confirmada.

### Decisión vigente: tarjetas
Cuando existe denominador interpretable:
- entero grande = cantidad absoluta;
- porcentaje secundario = proporción sobre universo/capacidad;
- línea inferior = comparación con mes anterior;
- mejora verde, deterioro rojo, neutro gris;
- disponibilidad invierte la lectura comercial: más disponible puede ser peor utilización.
**Estado:** confirmada.

### Decisión vigente: bottom insights
Todas las TVs siguen, cuando aplica:
`LECTURA | PUNTO POSITIVO | A ATENDER`.  
No inventar causas.  
**Estado:** confirmada.

### Decisión vigente: ediciones estéticas
Volver a la sesión de Claude Code de cada TV.  
No reauditar datos ni lógica por un cambio visual.  
No tocar Gates u otras TVs.  
**Estado:** confirmada.

### Decisión vigente: eficiencia Claude Code
- prompt maestro consolidado;
- una auditoría consolidada;
- evitar microauditorías;
- evitar full suite salvo necesidad;
- permisos de una vez;
- no `git add .`, commit, push, deploy o dependencias nuevas sin aprobación.
**Estado:** confirmada.

### Decisión vigente: mapa TV5
Natural Earth 1:10m local/offline; no siluetas aproximadas ni geocodificación inventada.  
**Estado:** confirmada.

---

## 5. Historial de cambios de decisiones

### Tema: numeración TVs
**Inicial:** varias configuraciones, incluyendo TV4=YPF.  
**Intermedio:** Pipeline fue desplazado conceptualmente.  
**Final:** TV4 Pipeline, TV5 YPF, TV6 Demanda.  
**Motivo:** orden de lectura y decisión operativa.  
**Regla actual:** usar numeración final.

### Tema: Power BI
**Inicial:** principal superficie de validación/publicación.  
**Intermedio:** error `dataType`, límites de Free.  
**Final:** pausado; HTML productivo es prioridad.  
**Regla actual:** retomar después de TVs.

### Tema: TV6 temporalidad
**Inicial:** lectura principalmente mensual.  
**Cambio:** pocas agencias/programática visibles.  
**Final:** top mensual + cuerpo YTD Ene–Jul.  
**Regla actual:** enfoque híbrido.

### Tema: TV5 mapa
**Inicial:** ranking/territorio sin coords completas.  
**Cambio:** GEO auxiliar con lat/lon permitió scatter.  
**Error:** silueta manual aproximada.  
**Final:** Natural Earth real, offline, con provincias.  
**Regla actual:** no volver a cartografía dibujada a mano.

### Tema: YPF geografía
**Inicial:** fuente principal sin lat/lon/APIE útil para mapa.  
**Cambio:** aparece `BASE TDT - ELEMENTOS - JUNIO26.xlsx`, incorporada como `input_aux/YPF_GEO_COORDENADAS.xlsx.xlsx`.  
**Final:** join auxiliar builder-only para TV5; no alterar Gates.  
**Regla actual:** no inventar coordenadas; mostrar cobertura parcial.

### Tema: YPF porcentaje TV1
**Actual visible:** 305 estaciones y `61,6% del total`, que corresponde a participación dentro de las unidades activas.  
**Discusión posterior:** el usuario entendió como patrón deseado “entero activo + % ocupación de universo propio”; TV5 muestra 305/451 = 67,6%.  
**Final actual:** TV1 no fue cambiado.  
**Regla:** `A VALIDAR` antes de tocarlo.

### Tema: ediciones visuales
**Inicial:** sesiones largas podían acumular contexto.  
**Final:** usar una sesión por TV para estética, congelando lógica.  
**Motivo:** ahorro de tokens y menor riesgo.

---

## 6. Reglas de negocio

### Identidad de campaña
`IDCampaña` identifica una campaña concreta y puede repetirse en varias filas porque una campaña puede estar presente en múltiples elementos/activaciones.

### Activación
Para TV5 YPF, activación = combinación distinta `(IDCampaña, ElementoID)`.

### Elemento activo
`ElementoID` distinto con campaña en el período.

### Estación YPF activa
Estación distinta con al menos una campaña válida en el período.

### Campañas únicas
`COUNT DISTINCT IDCampaña`.

### No confundir
- campañas únicas;
- activaciones;
- elementos activos;
- estaciones activas;
- slots;
- capacidad.

Son granos diferentes.

### Ocupación por calendario
`elementos con campaña / elementos elegibles`.

### Fill rate
`slots ocupados / capacidad de slots elegible`.

YPF queda fuera del fill rate total.

### Disponibilidad
Según dashboard:
`elegibles - ocupados`, o capacidad libre; interpretar con la métrica canónica del payload.

### Exclusiones
APSA y London Supply: fuera hasta nuevo aviso.

### Cencomedia
No inventar capacidad/denominadores. Puede formar parte del catálogo ampliado, no necesariamente Core.

### YPF catálogo
TV5:
- 451 estaciones;
- 3.082 elementos;
- catálogo actual sin histórico equivalente confirmado;
- no inventar comparación mensual de catálogo.

### TV5 estaciones activas
Julio: 305/451 = 67,6%.  
Junio: 263/451 = 58,3%.  
Delta: +42 estaciones y +9,3 pp.

### TV5 campañas/activaciones
Julio:
- 9 campañas únicas;
- 7.681 activaciones;
- 853,4 activaciones/campaña.

Junio:
- 4 campañas;
- 4.286 activaciones.

### TV5 elementos activos
Julio: 2.651.  
Junio: 2.371.  
Delta: +280.

### TV5 formato líder
Se define por **campañas distintas**, no activaciones.  
Julio: Punteras, 9 campañas.  
Junio: Torres, 4.

### TV5 histórico por formato
May–Jul:
- Menu Board: 0, 103, 379
- Torres: 1.560, 2.878, 4.760
- Punteras: 708, 1.305, 2.542
Sin actividad registrada Ene–Abr.

### TV5 geografía
- solo puntos con coordenadas válidas;
- no geocodificación externa;
- no posiciones inventadas;
- cobertura actual: 79 de 305 estaciones activas = 25,9%;
- intensidad/tamaño = activaciones acumuladas Ene–Jul;
- join auxiliar por prefijo/identificador disponible, no integración canónica todavía.

### TV6 Programática
Usar únicamente `PROGRAMATICA == 'Si'`.  
No inferir agencia por otros campos.  
Los pendientes sin agencia se conservan.

### MetricStatus
Valores esperados:
- `OK`
- `PARTIAL`
- `NO_APLICA`
- `REQUIERE_CONFIRMACION`

Regla absoluta: unknown ≠ 0.

---

## 7. Arquitectura actual

```text
input/OCU26_BASE_DATOS.xlsx
        ↓
Gate 1 · validate_input.py
        ↓
Gate 2 · transform_data.py
        ↓
Gate 3 · semantic_model.py + metrics_engine.py
        ↓
Gate 4 · export_data.py / data mart
        ↓
Builders específicos TV1–TV6
        ↓
output/tvN_data.json
        ↓
scripts/templates/tvN_template.html
        ↓
tvN.html
        ↓
hosting público / CMS / TV
```

### Capa auxiliar TV5 mapa

```text
input_aux/YPF_GEO_COORDENADAS.xlsx.xlsx
        ↓
join SOLO dentro de build_tv5_dashboard.py
        ↓
coordenadas válidas de estaciones
        +
scripts/templates/assets/argentina_pais.geojson
scripts/templates/assets/argentina_provincias.geojson
        ↓
mapa local/offline TV5
```

Esta capa geográfica **no modifica Gates** y no debe convertirse en lógica central sin diseño de Etapa 2.

### Principio técnico

Los HTML no deben contener lógica canónica de negocio. El builder prepara el payload; el template representa.

### Target

- 1920×1080;
- horizontal;
- LG 43SM5KB-BD;
- webOS ~2.0;
- sin scroll/overflow;
- JavaScript/CSS estable;
- mapa TV5 sin dependencia de red externa en runtime.

---

## 8. Arquitecturas anteriores o descartadas

### Power BI como salida principal inmediata
Descartado temporalmente por error y licencias. Sigue válido como capa posterior.

### Power BI “Publicar en web” como único mecanismo
Condicionado por tenant/licencias. No depender de él.

### GitHub Pages como plan inicial exclusivo
Se consideró como Plan B gratuito. El proyecto luego manejó Netlify como objetivo/entorno conocido. No asumir deploy final sin comprobar estado.

### Office Scripts + muchas automatizaciones
No deseado en Etapa 1. Se prefirió flujo simple M365.

### SharePoint Lists desde el inicio
Postergado a Etapa 2.

### Lógica en HTML
Descartada.

### Usar `window.OCU_DATA` legacy
Descartado como fuente de datos; solo referencias visuales.

### Mapa YPF con silueta dibujada a mano
Descartado explícitamente por el usuario.

### Tiles online obligatorios
No usados en solución final TV5; se eligió GeoJSON local.

---

## 9. Bases de datos

### `input/OCU26_BASE_DATOS.xlsx`
**Función:** fuente maestra vigente.  
**Estado:** protegida.  
**SHA esperado histórico:** `2f165e12ad90c2f05963367a8e1717dcd1006e9ce669f976afa6e57470aca2cd`  
**Hojas principales conocidas:** maestro de elementos, campañas, parámetros.  
**Snapshot documentado:** 4.338 / 9.503 / 23.  
**Regla:** no modificar desde dashboards.

### `BPBI_OCU26_V2_REPARADO_DESPLEGABLES_OK.xlsx`
Archivo histórico de modelo Power BI/base de ocupación. Conservar como antecedente.

### `Base_ocupacion_26__4-8.xlsx`
Base histórica mencionada. No asumir vigente.

### `BASE TDT - ELEMENTOS - JUNIO26.xlsx`
Archivo auxiliar aportado por el usuario para YPF GEO.  
En repo se trabajó como:

`input_aux/YPF_GEO_COORDENADAS.xlsx.xlsx`

Campos confirmados:
- `CODIGO`
- `LATITUD`
- `LONGITUD`
- `CIUDAD*`
- `DIRECCION*`

Puede contener APIE/IDs. La integración real a la clave central queda para Etapa 2.

### Datos cartográficos Natural Earth
No son base comercial. Son assets geográficos reales:
- Admin 0 1:10m;
- Admin 1 1:10m;
- filtrados a Argentina;
- guardados en GeoJSON local.

---

## 10. Tablas, hojas y columnas

### Excel OCU26

Nombres exactos conocidos por la evolución del proyecto:
- `BASE_MAESTRA_ELEMENTOS` / maestro equivalente vigente en OCU26;
- `BASE_CAMPAÑAS`;
- `CONTROL_DISPONIBILIDAD` en modelos anteriores;
- hojas actuales de maestro/campañas/parámetros según la versión OCU26.

No renombrar sin inspección.

Columnas/variables críticas documentadas:
- `ElementoID`
- `IDCampaña`
- `CargaID`
- `ClaveNegocio`
- `PortfolioTier`
- `CircuitoNegocio`
- `PROGRAMATICA`
- `Cliente`
- `Marca`
- `Agencia`
- `Subcircuito`
- `Ciudad`
- `Ubicacion`

### GEO auxiliar

- `CODIGO`
- `LATITUD`
- `LONGITUD`
- `CIUDAD*`
- `DIRECCION*`

No normalizar nombres en documentación.

---

## 11. Identificadores

### `IDCampaña`
Campaña comercial/publicitaria única. Puede aparecer repetida por múltiples elementos.

### `ElementoID`
Elemento comercial. Es el grano para “elementos activos” y parte de la clave de activación TV5.

### `(IDCampaña, ElementoID)`
Clave compuesta usada para contar activaciones YPF distintas.

### `StationKey_TV1`
Clave/surrogate de estación usada cuando la fuente principal no aportaba APIE real utilizable. Cuenta estaciones distintas.

### `CODIGO` GEO
Código del auxiliar. Se exploró/uso el prefijo numérico para reconciliar estaciones con OCU26.

### APIE
Identificador de estación YPF deseado como clave real futura. La nueva base auxiliar parece contener IDs/APIE, pero la migración canónica no está hecha.

### `CargaID`, `ClaveNegocio`
Identificadores relevantes del pipeline; no modificar sin validación.

---

## 12. Archivos del proyecto

| Archivo | Tipo | Función | Estado | Vigencia | Dependencias | Observaciones |
|---|---|---|---|---|---|---|
| `input/OCU26_BASE_DATOS.xlsx` | Excel | fuente principal | protegido | vigente | Gates | no editar |
| `config/business_semantics.json` | JSON | reglas semánticas | protegido | vigente | Gate3 | tocar solo por bloqueo real |
| `scripts/validate_input.py` | Python | Gate1 | protegido | vigente | Excel | no tocar por estética |
| `scripts/transform_data.py` | Python | Gate2 | protegido | vigente | Gate1 | idem |
| `scripts/semantic_model.py` | Python | Gate3 | protegido | vigente | transform | idem |
| `scripts/metrics_engine.py` | Python | métricas | protegido | vigente | semántica | idem |
| `scripts/export_data.py` | Python | Gate4 | protegido | vigente | MetricsEngine | idem |
| `scripts/build_tv1_dashboard.py` | Python | builder TV1 | productivo | vigente | pipeline | cerrado funcionalmente |
| `scripts/build_tv2_dashboard.py` | Python | builder TV2 | productivo | vigente | pipeline | cerrado |
| `scripts/build_tv3_dashboard.py` | Python | builder TV3 | productivo | vigente | pipeline | cerrado visual |
| `scripts/build_tv4_dashboard.py` | Python | builder TV4 | productivo | vigente | pipeline | 35/35 tests confirmados |
| `scripts/build_tv5_dashboard.py` | Python | builder TV5 | productivo | vigente | pipeline + GEO aux | mapa real integrado |
| `scripts/build_tv6_dashboard.py` | Python | builder TV6 | productivo | vigente | pipeline | híbrido mensual/YTD |
| `scripts/templates/tv1_template.html` | HTML | template TV1 | productivo | vigente | payload TV1 | |
| `scripts/templates/tv2_template.html` | HTML | template TV2 | productivo | vigente | payload TV2 | |
| `scripts/templates/tv3_template.html` | HTML | template TV3 | productivo | vigente | payload TV3 | |
| `scripts/templates/tv4_template.html` | HTML | template TV4 | productivo | vigente | payload TV4 | |
| `scripts/templates/tv5_template.html` | HTML | template TV5 | productivo | vigente | payload TV5 + GeoJSON | |
| `scripts/templates/tv6_template.html` | HTML | template TV6 | productivo | vigente | payload TV6 | |
| `scripts/templates/assets/argentina_provincias.geojson` | GeoJSON | provincias reales | productivo | vigente TV5 | Natural Earth 1:10m | local/offline |
| `scripts/templates/assets/argentina_pais.geojson` | GeoJSON | contorno real Argentina | productivo | vigente TV5 | Natural Earth 1:10m | local/offline |
| `input_aux/YPF_GEO_COORDENADAS.xlsx.xlsx` | Excel | lat/lon YPF | auxiliar | vigente solo TV5 | builder TV5 | double extension real usada |
| `output/tv1_data.json` … `output/tv6_data.json` | JSON | payloads | productivo | vigente | builders | no editar manualmente |
| `tv1.html` … `tv6.html` | HTML | salidas TV | productivo | vigente | templates/payload | abrir estos, no template raw |
| `tests/test_build_tv1_dashboard.py` … `test_build_tv6_dashboard.py` | Python | tests | vigentes | por TV | builders | correr focalizados |
| `audit_sources/TV*_REFERENCE...` | HTML | referencias visuales | read-only | histórica/visual | ninguna | no usar datos legacy |
| `GUIA_FINAL_TABLEROS_OCU26.pdf` | PDF | explicación simple/glosario | generado en ChatGPT | vigente como documento | capturas finales | comprobar ubicación antes de asumir repo |
| `docs/CM1.md` | Markdown | contexto maestro repo | debe actualizarse con este archivo | vigente cuando se copie | proyecto | fuente de verdad |

---

## 13. Estructura de carpetas

```text
C:\brand plus\ocu26-dashboard\
├── input\
│   └── OCU26_BASE_DATOS.xlsx
├── input_aux\
│   └── YPF_GEO_COORDENADAS.xlsx.xlsx
├── config\
│   └── business_semantics.json
├── scripts\
│   ├── validate_input.py
│   ├── transform_data.py
│   ├── semantic_model.py
│   ├── metrics_engine.py
│   ├── export_data.py
│   ├── build_tv1_dashboard.py
│   ├── build_tv2_dashboard.py
│   ├── build_tv3_dashboard.py
│   ├── build_tv4_dashboard.py
│   ├── build_tv5_dashboard.py
│   ├── build_tv6_dashboard.py
│   └── templates\
│       ├── tv1_template.html
│       ├── tv2_template.html
│       ├── tv3_template.html
│       ├── tv4_template.html
│       ├── tv5_template.html
│       ├── tv6_template.html
│       └── assets\
│           ├── argentina_pais.geojson
│           └── argentina_provincias.geojson
├── output\
│   ├── tv1_data.json
│   ├── tv2_data.json
│   ├── tv3_data.json
│   ├── tv4_data.json
│   ├── tv5_data.json
│   └── tv6_data.json
├── tests\
│   ├── test_build_tv1_dashboard.py
│   ├── test_build_tv2_dashboard.py
│   ├── test_build_tv3_dashboard.py
│   ├── test_build_tv4_dashboard.py
│   ├── test_build_tv5_dashboard.py
│   └── test_build_tv6_dashboard.py
├── audit_sources\
│   └── referencias HTML read-only
├── powerbi\
├── docs\
│   └── CM1.md
├── tv1.html
├── tv2.html
├── tv3.html
├── tv4.html
├── tv5.html
└── tv6.html
```

La estructura exacta completa del repo puede incluir más archivos históricos, Parquet, manifests y documentación. No borrarlos por no aparecer en este árbol.

---

## 14. Código y scripts

### `validate_input.py`
Valida integridad/estructura del Excel. Protegido.

### `transform_data.py`
Normaliza/transforma datos. Protegido.

### `semantic_model.py`
Clasifica negocio y crea capa semántica. Protegido.

### `metrics_engine.py`
Centraliza métricas. Protegido.

### `export_data.py`
Produce outputs/data mart. Protegido.

### Builders TV1–TV6
Cada builder:
1. consume pipeline real;
2. arma payload específico;
3. escribe `output/tvN_data.json`;
4. renderiza template;
5. genera `tvN.html`.

### Particularidad TV5
`build_tv5_dashboard.py` además:
- lee GEO auxiliar;
- relaciona estaciones;
- calcula cobertura;
- incorpora puntos;
- usa GeoJSON local;
- no debe geocodificar faltantes;
- no debe modificar pipeline central.

### Regla de preview
Abrir `tvN.html` raíz. Abrir el template directamente puede mostrar `S/D`/datos vacíos y no prueba un error de datos.

---

## 15. Fórmulas y cálculos

### Puntos porcentuales

`delta_pp = porcentaje_actual - porcentaje_mes_anterior`

Ejemplo:
`29,4 - 30,5 = -1,1 pp`

### TV1 Core
Core = 964 unidades comerciales según composición vigente.

### TV1 unidades con campaña
Julio: 495.  
`495 / 964 = 51,3%`

### TV1 YPF
305 estaciones activas.  
Visible: `61,6% del total` = participación dentro de 495 unidades con campaña.  
TV5: `305 / 451 = 67,6%` = penetración del catálogo YPF.  
Esta diferencia de semántica queda `A VALIDAR`, no corregir automáticamente.

### TV1 estático
119 activos; 29,4% visible; junio 30,5%; -1,1 pp; YTD 28,6%.

### TV1 digital calendario
71 activos; 78,0%; junio 79,1%; -1,1 pp.

### TV1 fill
321 slots; 20,4%; junio 27,3%; -6,9 pp.

### TV2
- calendario: 71 / 78,0%;
- fill: 321 / 20,4%;
- LED: 11 / 100,0%;
- Shoppings Digital: 60 / 88,2%;
- slots disponibles: 1.254 / 79,6%.

### TV3
- Core Estático: 119/422 = 28,2%;
- Shoppings Estático: 119 / 31,2%;
- AA2000 Estático: 0/40 = 0,0%;
- shoppings con actividad: 7/17 = 41,2%;
- disponibles: 303 / 71,8%.

### TV4
- 77 campañas únicas activas;
- 421 activaciones en curso;
- 4 reservas futuras;
- 15 activaciones de reservas;
- 4 campañas inician en 30 días;
- 9 finalizan en 30 días;
- pipeline en Core: 77/82 = 93,9%.

### TV5
Ya documentado en sección 6.

### TV6
- top KPIs = julio;
- cuerpo = Ene–Jul;
- concentración Top5 = activaciones de Top5 marcas / activaciones totales del scope;
- final visible: 77,6% sobre total 8.203 activaciones.

---

## 16. Validaciones

### Integridad de archivos
SHA Excel histórico documentado y protegido.

### Estructura
Gates completos.

### Columnas
Validaciones ejecutadas en Gates; nuevas columnas GEO solo en auxiliar.

### IDs
`IDCampaña`, `ElementoID`, claves semánticas revisadas en pipeline.

### Duplicados
Repetición de `IDCampaña` por múltiples elementos es válida.

### Relaciones
Main pipeline estable. GEO auxiliar es relación local no canónica.

### Reglas de negocio
APSA/London exclusiones; YPF sin fill; MetricStatus; mes anterior.

### Fórmulas
Tests por TV verifican payloads; pendientes específicos arriba.

### Outputs
Seis HTML generados.

### Visualizaciones
Seis capturas finales revisadas. TV5 mapa real aprobado.

### Publicación
No deploy final documentado en esta migración.

### Automatizaciones
Planificadas; no activar sin cerrar publicación.

---

## 17. Tests

| Test | Qué valida | Resultado actual |
|---|---|---|
| Gates/core suite histórica | pipeline base | 199 pass documentado |
| Gate4B | export/data mart | 8/8 |
| TV1 | builder/payload/HTML/exclusiones | 39/39 PASS |
| Suite tras TV1 | regresión amplia | 238/238 PASS en evidencia visual |
| TV2 | builder digital | 37/37 PASS |
| TV3 | builder estático | `A VALIDAR` resultado textual |
| TV4 | Pipeline | 35/35 PASS |
| TV5 | YPF antes de GEO final | 34/34 PASS |
| TV5 después de GEO final | integración final | `A VALIDAR` |
| TV6 versión final híbrida | Demanda | `A VALIDAR` resultado textual |

Regla: no correr full suite cada vez. Para ediciones puramente estéticas, evitar tests innecesarios. Para cambios materiales de builder/lógica, correr test específico una vez al final.

---

## 18. Herramientas utilizadas

### ChatGPT
Definición funcional, revisión, prompts, migración, documentación, QA conceptual.

### Claude / Claude Code
Auditorías, código, builders, tests, previews y modificaciones de repo.

### Excel
Fuente operativa.

### Python
Gates/builders/tests.

### HTML/CSS/JavaScript
Dashboards TV.

### Power BI Desktop
Modelo/validación posterior; pausado.

### Git/GitHub/GitHub Desktop
Versionado.

### GitHub Actions
Plan de automatización futura.

### SharePoint/M365
Futura operación online.

### Netlify
Entorno/objetivo de hosting conocido; no asumir deploy final.

### CMS
Consumo en TVs.

### Natural Earth
Cartografía real offline TV5.

---

## 19. Restricciones de herramientas y licencias

- Power BI Free no permite compartir libremente a terceros como Pro.
- `Publicar en web` depende del tenant.
- No pagar licencias adicionales en esta etapa.
- No depender de PC local encendida para solución final.
- Datos pueden ser públicos; privacidad no es restricción.
- CMS/TVs deben abrir HTML/URL estable.
- webOS antiguo: evitar tecnología innecesariamente pesada.
- No incorporar dependencia cartográfica online obligatoria.
- Claude Code: ahorro de tokens prioritario.
- No instalar dependencias nuevas sin necesidad/autorización.

---

## 20. Git / GitHub

**Repo local:** `C:\brand plus\ocu26-dashboard`  
**Branch:** `main`

Histórico:
- GitHub Desktop instalado/autenticado;
- commits/push de Gates realizados en etapas previas;
- dashboard work posterior no debe asumirse committed.

Regla vigente:
- antes de commit: `git status` + `git diff`;
- revisar archivos protegidos;
- no `git add .`;
- no commit/push/deploy sin aprobación.

**Estado exacto Git actual:** `A VALIDAR`.

Hosting inicialmente contemplado:
- GitHub Pages gratuito.

Objetivo/URL conocido posteriormente:
- `https://digitalcore-brandplus.netlify.app/`

No afirmar que la versión final de las 6 TVs esté desplegada allí sin comprobarlo.

---

## 21. SharePoint / Microsoft 365

Sitio propio creado.

### Etapa 1 definida
Flujo simple:
- una persona carga;
- luego hasta 3–4;
- simultaneidad rara, máximo 2;
- Excel plano;
- Forms/Power Automate mínimo;
- sin complejidad innecesaria.

### Etapa 2
- SharePoint Lists;
- mayor robustez;
- automatización;
- conexión Power BI/outputs.

No rehacer Etapa 1 cuando se pase a Etapa 2; evolucionar sobre la estructura existente.

---

## 22. Power BI

Estado:
- arquitectura/modelo especificados;
- PBIP/Power BI Desktop para validar;
- fondos derivados de HTML posibles;
- publicación pública condicionada.

Error actual:
`El argumento 'dataType' no puede ser nulo. Nombre del parámetro: dataType`

Decisión:
- pausado durante cierre de TVs;
- retomar desde el 11/8/2026 o después de HTML;
- no tocar durante retoques estéticos.

Rol final:
- validación/modelo interno;
- no duplicar lógica del motor central.

---

## 23. HTML / dashboards / salidas

### TV1 — Visión general del negocio

**Pregunta:** lectura general de inventario, actividad y performance.

**KPIs finales visibles:**
- Core Comercial: 964.
- Unidades con campaña: 495 / 51,3%; junio 457; +38.
- YPF: 305 estaciones / 61,6% del total activo; junio 263; +42.
- Estático: 119 / 29,4%; junio 30,5%; -1,1 pp; YTD 28,6%.
- Digital calendario: 71 / 78,0%; junio 79,1%; -1,1 pp.
- Digital fill: 321 / 20,4%; junio 27,3%; -6,9 pp.

**Evolución Ene–Jul:**
- Digital: 69,71,72,72,71,72,71.
- Estático: 112,110,115,116,116,122,119.
- YPF: 0,0,0,0,233,263,305.

**Catálogo ampliado:** 1.065.
- Shoppings 42,2%.
- Pantallas 1,0%.
- YPF 42,3%.
- AA2000 4,9%.
- Cencomedia 8,3%.
- Otros 1,3%.

**Campañas únicas acumuladas a julio:** 425.

**Estado:** funcional/visual cerrado.  
**A validar Etapa 2:** porcentaje YPF y denominador Estático vs TV3.

---

### TV2 — Core Comercial Digital

**Scope:** Pantallas LED + Shoppings Digital + AA2000; sin YPF/APSA/London.

**KPIs:**
- 71 activos / 78,0%.
- 321 slots / 20,4% fill.
- Pantallas LED 11 / 100%.
- Shoppings Digital 60 / 88,2%.
- 1.254 slots disponibles / 79,6%.

**Ranking Pantallas:**
1. Cabildo 78,1% fill
2. Pilar 68,2%
3. Cerrito 47,4%
4. Remeros 47,4%
5. Avellaneda 45,3%

**Ranking Shoppings:**
1. Unicenter 30,0% fill
2. Palmas del Pilar 17,5%
3. Portal Escobar 15,0%

**Estado:** cerrado para entrega.  
**Tests:** 37/37.  
**A validar:** 68 vs ~61 elegibles y Remeros.

---

### TV3 — Core Comercial Estático

**Scope final actual:** CENCOSUD + REMEROS + AA2000 + PILAR_FRONTLIGHT.  
Cencomedia no se fuerza al Core.

**KPIs finales:**
- ocupación calendario: 119 / 28,2%;
- Shoppings Estático: 119 / 31,2%;
- AA2000 Estático: 0 / 0,0%, 0 de 40 elegibles;
- Shoppings con actividad: 7 / 41,2%;
- disponibles: 303 / 71,8%.

**Ranking:**
1. Palmas del Pilar — 23 ocup. / 7 disp. / 76,7%.
2. Remeros — 19 / 8 / 70,4%.
3. Factory Quilmes — 3 / 2 / 60,0%.
4. Unicenter — 31 / 31 / 50,0%.
5. Plaza Oeste — 19 / 30 / 38,8%.

**Soportes:**
- Otro 109.
- Frontlight 10.

**Evolución:** 112,110,115,116,116,122,119.

**Estado:** visual aceptado.  
**Tests:** A VALIDAR textual.  
**Pendientes Etapa 2:** `Otro`, denominadores/evolución, reconciliación con TV1.

---

### TV4 — Pipeline Comercial

**Corte:** 31/07/2026.

**KPIs:**
- 77 campañas únicas activas;
- 421 activaciones en curso;
- 4 reservas futuras / 15 activaciones;
- 4 campañas inician en 30 días;
- 9 finalizan en 30 días;
- 93,9% Pipeline en Core = 77/82.

**Estado pipeline:**
- Activas 421.
- Reservas 15.
- Finalizadas histórico 1.222.

**Activas por grupo:**
- Shoppings Digital 229.
- Shoppings Estático 124.
- Pantallas LED 68.
- AA2000/Pilar Frontlight 0.

**Próximas activaciones mostradas:**
- Netflix Moria
- Nintendo
- Oak Street
- PlayStation

**Estado:** cerrado.  
**Tests:** 35/35 PASS.

---

### TV5 — Pulso comercial YPF

**KPIs:**
- catálogo: 451 estaciones / 3.082 elementos;
- estaciones activas julio: 305 / 67,6%;
- junio: 263 / 58,3%; +42; +9,3 pp;
- campañas julio: 9;
- activaciones: 7.681;
- junio: 4 campañas / 4.286 activaciones;
- 853,4 activaciones/campaña;
- elementos activos: 2.651;
- junio: 2.371; +280;
- formato líder: Punteras, 9 campañas;
- junio: Torres (4).

**Histórico May–Jul:**
- Menu Board: 0 / 103 / 379.
- Torres: 1.560 / 2.878 / 4.760.
- Punteras: 708 / 1.305 / 2.542.

**Mapa final:**
- cartografía real Natural Earth;
- contorno y provincias;
- 79 puntos con geo de 305 estaciones activas;
- cobertura 25,9%;
- tamaño/color por activaciones Ene–Jul;
- assets locales;
- sin runtime online;
- versión con silueta aproximada queda descartada.

**Estado:** usuario aprobó la versión final después del mapa real.  
**Tests:** 34/34 antes de GEO; post-GEO A VALIDAR si se requiere formalidad.

---

### TV6 — Así se comporta la demanda

**Regla temporal final:**
- KPIs arriba = julio;
- cuerpo abajo = Ene–Jul.

**KPIs finales visibles:**
- Clientes directos activos · julio: **11**.
  - nota: excluye 32 vía agencia y 377 `A confirmar`.
- Marcas activas · julio: **86**.
- Agencias activas · julio: **2**.
  - nota: 377 activaciones con agencia `A Confirmar`.
- Cliente top identificado: **GCBA**.
  - 2.271 activaciones directas.
- Concentración Top 5: **77,6%**.
  - Top5 marcas del total de 8.203 activaciones.

**Top marcas Ene–Jul visible en captura final:**
1. GCBA — 2.356
2. Medife — 1.807
3. Turismo De Cordoba — 970
4. Bridgestone — 818
5. Dermaglos — 526
6. Milka — 374

**Demanda por circuito acumulada:**
Columnas:
- Pantallas LED
- Sh. Digital
- Sh. Estático
- AA2000 / Pilar Frontlight
- YPF

Valores visibles:
- GCBA: 39 / 49 / — / — / 2.268
- Medife: 1 / 28 / — / — / 1.778
- Turismo De Cordoba: — / — / — / — / 970
- Bridgestone: 3 / 14 / 10 / — / 791
- Dermaglos: 13 / — / — / — / 513
- Milka: — / — / — / — / 374

**Footer final:**
- Lectura: julio 86 marcas y 2 agencias; YTD lideran GCBA, MEDIFE, TURISMO DE CORDOBA; agencias CARAT, OMD, OSA; actividad concentrada principalmente en YPF (82,4%).
- Punto positivo: YTD identifica 5 agencias: CARAT, OMD, OSA, GROUPM, NEXT MEDIA, frente a 2 en julio.
- A atender: 112 activaciones programáticas Ene–Jul sin agencia imputada; `PROGRAMATICA` no permite asociarlas a una agencia identificada.

**Ranking inferior:** diseño híbrido MARCAS → AGENCIAS → PROGRAMÁTICA; la captura final mostrada corresponde a MARCAS.

**Cambio respecto de primeras previews:** valores como `31 clientes`, `62 marcas` o `McDonald’s` no son el estado final. La captura final anterior prevalece.

**Estado:** visual aceptado.  
**Tests finales:** A VALIDAR evidencia textual.

---

## 24. Automatización

### Existente
- builders manuales ejecutables con Python;
- Git/GitHub disponible;
- outputs regenerables.

### Planificada
`Excel actualizado → validación/Gates → builders → HTML → GitHub/hosting → CMS`

GitHub Actions se consideró para regeneración automática.

### SharePoint futuro
`SharePoint/Excel online → proceso → outputs → publicación`

### Descartada en Etapa 1
Automatización pesada con múltiples Office Scripts/Flows.

---

## 25. Referencias del chat de Claude e imágenes

### 25.1 Qué trabajo se hizo en Claude

- inspección de repo;
- Gates;
- auditorías de universos;
- builders;
- tests;
- previews;
- ajustes visuales;
- Git sanity;
- análisis YPF GEO;
- descarga/filtrado Natural Earth;
- integración del mapa;
- creación de las seis TVs.

### 25.2 Qué información pasó desde Claude

- métricas reales;
- resultados de tests;
- rutas;
- clasificación de scopes;
- previews 1920×1080;
- problemas de template;
- problema del header TV5;
- cobertura GEO;
- evidencia de cartografía.

### 25.3 Imágenes conocidas relevantes

#### Captura final TV1
**Qué muestra:** “Visión general del negocio” con 6 KPIs, evolución, composición y footer.  
**Conclusión:** referencia visual vigente TV1.  
**Vigencia:** final de trabajo.

#### Captura final TV2
**Qué muestra:** Core Comercial Digital, 5 KPIs, rankings, evolución y footer.  
**Vigencia:** final de trabajo.

#### Captura final TV3
**Qué muestra:** Core Comercial Estático, ranking, soportes, evolución.  
**Vigencia:** final de trabajo.

#### Captura final TV4
**Qué muestra:** Pipeline Comercial 77/421/4/4/9/93,9%, barras y próximas activaciones.  
**Vigencia:** final.

#### Captura final TV5
**Qué muestra:** Pulso comercial YPF con mapa real de provincias/región y puntos de actividad.  
**Vigencia:** final aprobada.

#### Captura final TV6
**Qué muestra:** Demanda con 11 clientes directos, 86 marcas, 2 agencias, GCBA y 77,6%.  
**Vigencia:** final.

### 25.4 Capturas de Claude

- captura de validación TV1: `Tests TV1: 39/39`, `Suite completa: 238/238`, `Commit: NO · Push: NO · Deploy: NO`.
- capturas TV4 con pytest/permisos y preview.
- captura TV5 donde Claude informó `All 34 tests still pass`.
- capturas de comandos GEO y Natural Earth.
- captura del mapa fallido con silueta aproximada.
- captura final con cartografía real.

### 25.5 Versiones visuales

TV5 mapa:
1. panel sin cartografía útil;
2. scatter + silueta manual;
3. **DESCARTADO** por no parecer Argentina/AMBA/CABA;
4. Natural Earth 1:10m real;
5. final aprobado.

TV6:
1. versión mensual;
2. detectada baja representación de agencias/programática;
3. híbrida mensual + YTD;
4. final con 11/86/2/GCBA/77,6%.

### 25.6 Información visual que no debe perderse

Brand:
- Poppins;
- azul `#1C60FF`;
- negro `#1D1D1B`;
- cream `#EEECE6`;
- coral `#FF4E46`;
- navy `#071124`;
- dark `#05070D`;
- positivo `#2FE084`;
- warning `#FFC857`.

Layout:
- 1920×1080;
- header superior;
- cards redondeadas;
- paneles oscuros;
- footer horizontal de 3 insights;
- sin scroll.

### 25.7 Imágenes que deben volver a cargarse

| Imagen a recuperar | Qué contiene | Por qué es necesaria | Decisiones asociadas | Prioridad |
|---|---|---|---|---|
| captura final TV1 10/8 | diseño/números finales | referencia estética | TV1 | IMPORTANTE |
| captura final TV2 10/8 | diseño final | referencia estética | TV2 | IMPORTANTE |
| captura final TV3 10/8 | diseño final | referencia estética | TV3 | IMPORTANTE |
| captura final TV4 10/8 | diseño final | referencia estética | TV4 | IMPORTANTE |
| captura final TV5 con mapa real | cartografía/puntos | evitar regresión | TV5 | CRÍTICA |
| captura final TV6 | números y composición final | evitar usar preview vieja | TV6 | CRÍTICA |
| manual Brand Plus | branding | mantener identidad | branding | IMPORTANTE |
| mapa TV5 fallido | silueta manual descartada | solo para recordar qué NO hacer | mapa | REFERENCIA |

---

## 26. Otros archivos externos que deberían recuperarse

1. `CONTEXTO_MAESTRO.md`.
2. `input/OCU26_BASE_DATOS.xlsx`.
3. `input_aux/YPF_GEO_COORDENADAS.xlsx.xlsx` o su fuente original `BASE TDT - ELEMENTOS - JUNIO26.xlsx`.
4. `config/business_semantics.json`.
5. scripts Gates.
6. builders TV1–TV6.
7. templates TV1–TV6.
8. tests TV1–TV6.
9. outputs JSON.
10. `tv1.html` … `tv6.html`.
11. `audit_sources/*`.
12. GeoJSON Argentina.
13. carpeta `powerbi/`.
14. PBIP/archivos Power BI existentes.
15. `BPBI_OCU26_V2_REPARADO_DESPLEGABLES_OK.xlsx`.
16. documentación Etapa 1/SharePoint.
17. manual de Brand Plus.
18. `GUIA_FINAL_TABLEROS_OCU26.pdf`.
19. capturas finales de las seis TVs.

---

## 27. Problemas encontrados

### Power BI dataType
**Impacto:** bloqueó avance inmediato.  
**Estado:** pendiente.

### Templates directos muestran S/D
**Causa:** sin payload inyectado.  
**Solución:** abrir HTML raíz generado.  
**Estado:** resuelto.

### TV5 header clipping
**Causa:** header 92px insuficiente para bloque derecho.  
**Solución:** 118px.  
**Estado:** resuelto.

### TV5 servidor local persistente
**Impacto:** Claude podía quedar “colgado”.  
**Solución:** matar servidor después de preview.  
**Estado:** resuelto/procedimiento aprendido.

### TV5 mapa sin referencia geográfica
**Causa:** solo scatter.  
**Estado:** resuelto.

### TV5 silueta aproximada
**Causa:** Claude inventó contorno manual.  
**Resultado:** rechazado.  
**Solución:** Natural Earth real.  
**Estado:** resuelto.

### GEO cobertura parcial
79/305.  
**Estado:** conocido; no inventar faltantes.

### TV6 mensual ocultaba profundidad
**Solución:** híbrido mensual/YTD.  
**Estado:** resuelto.

### TV2 universo Shoppings
~68 vs ~61.  
**Estado:** Etapa 2.

### TV3 `Otro`
109/119.  
**Estado:** Etapa 2.

---

## 28. Errores ya cometidos

1. Abrir template en vez de `tvN.html` y creer que faltan datos.
2. Usar números legacy `window.OCU_DATA`.
3. Diseñar un mapa YPF manualmente en lugar de usar geometría real.
4. Confundir campañas con activaciones.
5. Confundir estaciones con elementos.
6. Confundir fill rate con ocupación calendario.
7. Tratar unknown como 0.
8. Reabrir auditorías cerradas por una duda visual.
9. Consumir tokens con microauditorías repetidas.
10. Dejar un servidor local persistente.
11. Asumir que una preview de TV6 temprana era final.
12. Asumir que la numeración histórica de TVs seguía vigente.
13. Interpretar `pp` como porcentaje relativo.
14. Inventar capacidad histórica de catálogo YPF.
15. Dar por hecho que un resultado de tests pasó si no quedó evidencia.

---

## 29. Soluciones descartadas

### Silueta manual mapa
Descartada definitivamente.

### Fill total YPF
Descartado hasta contar con capacidad real.

### Power BI como único frontend
Pausado/no exclusivo.

### Lógica duplicada en HTML
Descartada.

### SharePoint Lists como primera etapa obligatoria
Postergada.

### Reauditar todas las TVs antes de cada ajuste
Descartado por costo/riesgo.

### Sesión Claude única para todos los retoques
No recomendada. Preferir sesión por TV.

---

## 30. Supuestos

### Confirmados posteriormente
- YPF tiene 451 estaciones actuales en catálogo.
- 305 estaciones activas julio.
- GEO auxiliar sí aporta lat/lon útiles.
- Natural Earth permitió resolver cartografía offline.

### Siguen vigentes
- la estructura del Excel se mantiene para regeneración;
- CMS/TV consumirá HTML/URL;
- no hay restricción de privacidad.

### Resultaron incorrectos
- TV4=YPF como numeración final;
- un mapa con silueta manual era suficiente;
- cuerpo TV6 totalmente mensual era suficiente.

### Pendientes
- universo Shoppings Digital canónico;
- clave APIE canónica;
- denominadores transversales;
- tests faltantes.

---

## 31. Dependencias

### Datos
Excel → Gates → MetricsEngine → builders.

### TV5
Excel principal + GEO auxiliar + GeoJSON local.

### Visual
Templates dependen de payload; HTML raíz depende del builder.

### Publicación
Git/repo → hosting → CMS → TV.

### Power BI
Depende de resolver `dataType`.

### SharePoint
Depende de base final/operación.

### Personas/proceso
Carga inicial 1 persona; futura 3–4.

---

## 32. Riesgos

- cambiar columnas Excel;
- cambiar IDs;
- modificar `ElementoID`/`IDCampaña`;
- mover archivos/rutas;
- alterar config semántica;
- tocar Gates por estética;
- borrar GeoJSON TV5;
- reemplazar cartografía por tiles externos sin necesidad;
- asumir 100% cobertura GEO;
- desplegar sin revisar Git;
- mezclar datos legacy;
- corregir TV1 YPF 61,6% sin decisión;
- “arreglar” TV3/TV2 sin reconciliación;
- depender de PC encendida;
- romper compatibilidad webOS;
- agotar créditos Claude con auditorías redundantes.

---

## 33. Datos que NO deben modificarse sin validación

- `input/OCU26_BASE_DATOS.xlsx`
- SHA esperado
- `ElementoID`
- `IDCampaña`
- `CargaID`
- `ClaveNegocio`
- Gates
- `config/business_semantics.json`
- `MetricStatus`
- APSA/London exclusiones
- regla YPF sin fill
- capacidades digitales
- scopes Core
- GeoJSON finales TV5
- aux GEO sin entender join
- branding
- stage 1920×1080
- archivos cerrados de otras TVs durante edición puntual
- `powerbi/`
- `audit_sources/*`

---

## 34. Pendientes completos

### P0 — Críticos / bloqueantes

No hay bloqueo crítico para presentar las seis TVs.

**Tarea:** preservar esta versión CM1.  
**Estado:** en ejecución con este archivo.  
**Resultado:** fuente de verdad actualizada.

### P1 — Alta prioridad

**Tarea:** ediciones estéticas pendientes TV por TV.  
**Estado:** postergadas deliberadamente.  
**Dependencia:** sesión Claude correspondiente.  
**Resultado:** pulido sin alterar datos.

**Tarea:** Etapa 2 reconciliaciones.  
Incluye:
- TV1 YPF 61,6% vs 67,6% de catálogo;
- TV1 Estático 29,4% vs TV3 28,2%;
- TV2 ~68 vs ~61;
- Remeros;
- TV3 `Otro`;
- TV3 evolución;
- cross-TV.

**Tarea:** evidencia tests faltantes TV3/TV5 final/TV6 final.

**Tarea:** revisar `git status`/`git diff`.

### P2 — Importantes

- commit/push aprobados;
- deploy;
- prueba física LG/webOS;
- integrar APIE real;
- completar GEO YPF;
- resolver agencia `A Confirmar`;
- resolver programática sin agencia;
- Power BI;
- SharePoint final.

### P3 — Mejoras futuras

- GitHub Actions/automatización;
- SharePoint Lists;
- Forms robusto;
- actualizaciones automáticas;
- forecast;
- nuevas métricas;
- mayor geografía YPF.

---

## 35. Roadmap actual

1. Guardar/copy este `CONTEXTO_MAESTRO.md` en `docs/CM1.md`.
2. No tocar las TVs mientras no haya un cambio concreto.
3. Cuando se retomen ajustes estéticos:
   - abrir sesión específica TV;
   - congelar lógica;
   - cambiar solo estética;
   - preview;
   - cerrar.
4. Ejecutar Etapa 2 de microauditorías/reconciliaciones.
5. Correr tests específicos solo donde hubo cambio material.
6. `git status` + `git diff`.
7. commit/push con aprobación.
8. deploy/hosting.
9. prueba física TV/CMS.
10. Power BI.
11. SharePoint/automatización.

---

## 36. Próximo paso exacto

Después de generar este archivo:

1. descargar `CONTEXTO_MAESTRO.md`;
2. reemplazar/actualizar `docs/CM1.md` en el repo cuando corresponda;
3. conservar las seis TVs sin reauditar;
4. si se inicia un ajuste visual, hacerlo en la sesión de Claude Code de esa TV con instrucción:
   `Solo ajustes estéticos; lógica y métricas congeladas salvo error evidente. No tocar Gates ni otras TVs.`

No hay que volver a implementar ninguna TV desde cero.

---

## 37. Punto exacto donde terminó este chat

Inmediatamente antes de pedir esta actualización CM1:

1. TV5 había quedado con mapa cartográfico real después de descargar una única vez Natural Earth 1:10m.
2. Se habían creado assets locales:
   - `argentina_provincias.geojson`
   - `argentina_pais.geojson`
3. Se había borrado la carpeta temporal de descarga.
4. `build_tv5_dashboard.py` fue ejecutado para regenerar.
5. Se mostró la captura final TV5 con provincias y puntos.
6. Se mostraron juntas las capturas finales TV1–TV6.
7. El usuario indicó que estaba perfecto/cerrado y pidió un PDF final simple con glosario y explicación.
8. Se generó `GUIA_FINAL_TABLEROS_OCU26.pdf`.
9. El usuario consultó cómo interpretar YTD y pp en TV1 Estático.
10. Se explicó:
    - 119 activos;
    - 29,4% ocupación;
    - junio 30,5%;
    - -1,1 pp;
    - YTD 28,6%;
    - `pp` no equivale a % relativo.
11. Se acordó que futuras ediciones estéticas se harán en cada sesión Claude correspondiente para ahorrar tokens.
12. El usuario pidió actualizar CM1 siguiendo el prompt exhaustivo adjunto.

Este archivo es esa actualización.

---

## 38. Preguntas abiertas

Solo dudas reales vigentes:

1. ¿Debe TV1 YPF cambiar en Etapa 2 de `61,6% del total activo` a `67,6% del catálogo YPF` para armonizar semántica?
2. ¿Cuál es el denominador canónico final de Shoppings Digital TV2?
3. ¿Cómo debe clasificarse Remeros sin solapamientos?
4. ¿Por qué TV1 Estático usa 29,4% y TV3 28,2% para 119 activos?
5. ¿Cómo desagregar `Otro` en TV3?
6. ¿Cuál será el APIE/ID estación canónico en capa central?
7. ¿Puede ampliarse cobertura GEO YPF >25,9% con la fuente disponible?
8. ¿Cuál fue el resultado textual final de tests TV3?
9. ¿Pasan los tests TV5 después del cambio GEO final?
10. ¿Cuál fue el resultado textual final de tests TV6 híbrida?
11. ¿Cuál es el estado Git exacto?
12. ¿Qué hosting/deploy final se usará para producción?
13. ¿Cómo se resuelve Power BI `dataType`?
14. ¿Cómo funcionará la prueba física webOS/CMS?

---

## 39. Glosario

**OCU26:** proyecto/base operativa de ocupación 2026.  
**Core Comercial:** universo comercial central definido por reglas semánticas.  
**Catálogo ampliado:** Core + componentes adicionales como Cencomedia/MAB según la vista.  
**Elemento:** soporte/unidad comercial identificada por `ElementoID`.  
**Campaña:** campaña única identificada por `IDCampaña`.  
**Activación:** presencia de una campaña en un elemento; TV5 usa `(IDCampaña, ElementoID)` distinto.  
**Estación activa YPF:** estación con al menos una campaña en período.  
**Ocupación por calendario:** % de elementos elegibles con campaña.  
**Fill rate:** % de capacidad de slots vendida/ocupada.  
**Slot:** espacio/cupo digital vendible.  
**Disponible:** inventario/capacidad no ocupada.  
**YTD:** Year To Date; referencia acumulada/del indicador en lo que va del año hasta el mes vigente.  
**pp:** puntos porcentuales.  
**APIE:** identificador real de estación YPF deseado para la capa canónica.  
**StationKey_TV1:** surrogate de estación usado en la lógica existente.  
**MetricStatus:** estado de disponibilidad/calidad de métrica.  
**Core Digital:** Pantallas LED + Shoppings Digital + AA2000 según TV2.  
**Core Estático:** scope actual TV3 CENCOSUD + REMEROS + AA2000 + PILAR_FRONTLIGHT.  
**Programática:** registros con `PROGRAMATICA == 'Si'`; no inferir agencia.  
**Gate:** etapa controlada del pipeline.  
**Builder:** script que prepara payload y HTML de una TV.  
**Payload:** JSON específico consumido por template.  
**Natural Earth:** fuente cartográfica usada para GeoJSON real local de TV5.  
**GeoJSON:** formato de geometría geográfica.  
**CMS:** sistema que consume/publica contenidos en las TVs.  
**webOS:** sistema operativo de las TVs LG objetivo.  
**CM1:** Contexto Maestro principal.

---

## 40. Personas, equipos, proveedores y actores

### Brand Plus
Dueño del sistema/tableros.

### Usuario / Gerencia de Producto
Define reglas de negocio, priorización, revisión visual y aprobación.

### Claude / Claude Code
Implementación técnica asistida.

### ChatGPT
Orquestación conceptual, migración, QA, prompts y documentación.

### Microsoft 365 / SharePoint
Infraestructura futura de carga/operación.

### GitHub / Netlify
Versionado/publicación potencial.

No incluir datos personales adicionales innecesarios.

---

## 41. Cronología relevante

### Antes de agosto 2026
- bases de ocupación;
- Power BI;
- incorporación YPF;
- diseño M365.

### 5/8
- Etapa 1 M365;
- V3/base piloto.

### 7/8
- Gate3A;
- semántica.

### 8/8
- Gate3B;
- 150 tests en una etapa;
- commits.

### 8–9/8
- Gate4A/4B;
- Parquet;
- 191→199 tests;
- commit/push.

### 9/8
- Power BI error `dataType`;
- Power BI pausado;
- HTML prioridad;
- arquitectura 6 TVs;
- TV1 y TV2 avanzan.

### 10/8 madrugada/mañana
- TV3 productiva;
- TV4 Pipeline productiva;
- decisión TV6 antes de TV5;
- TV6 híbrida mensual/YTD;
- TV5 YPF implementada.

### 10/8 antes de 16:14
- GEO auxiliar YPF;
- auditoría de cobertura;
- mapa scatter;
- silueta manual rechazada;
- Natural Earth Admin0/Admin1 1:10m descargado una vez;
- GeoJSON Argentina local;
- mapa real final;
- seis capturas finales;
- PDF guía final;
- aclaración YTD/pp;
- decisión de sesiones separadas para retoques estéticos;
- actualización CM1.

---

## 42. Reglas para el próximo ChatGPT

“Este documento representa el estado acumulado del proyecto al cierre del chat anterior.

El proyecto tuvo múltiples iteraciones y cambios de criterio.

No asumir que una decisión temprana sigue siendo válida cuando existe una decisión posterior.

Antes de modificar la solución:

1. leer este documento completo;
2. identificar las decisiones vigentes;
3. respetar las reglas de negocio;
4. no reconstruir componentes que ya funcionan;
5. no recomendar soluciones previamente descartadas sin una razón nueva;
6. mantener compatibilidad con los archivos y arquitectura actual;
7. verificar las dependencias antes de cambiar nombres, estructuras o IDs;
8. diferenciar claramente una corrección necesaria de un rediseño opcional;
9. utilizar las referencias visuales y archivos indicados cuando sean necesarios;
10. si una decisión depende de una imagen que no está disponible, solicitar específicamente esa imagen en lugar de asumir su contenido;
11. preguntar únicamente cuando la información no pueda resolverse utilizando este documento y los archivos disponibles.”

Reglas adicionales OCU26:

12. APSA/London fuera.
13. No fill total YPF.
14. No números dummy.
15. No lógica nueva en HTML.
16. No tocar Gates por estética.
17. No unknown→0.
18. No reabrir TVs cerradas salvo cambio explícito.
19. Para estética, usar sesión específica de esa TV.
20. No gastar tokens en auditorías repetidas.
21. No usar silueta geográfica inventada.
22. No afirmar tests sin evidencia.
23. No commit/push/deploy sin aprobación.

---

## 43. Instrucciones para reconstruir el contexto visual

Al recibir nuevamente las imágenes:

1. identificar TV y versión;
2. comparar con la captura final descrita aquí;
3. ignorar datos legacy de referencias;
4. verificar branding/1920×1080;
5. comprobar cards y footer;
6. TV5: confirmar que el mapa usa geometría real y muestra cobertura 79/305;
7. TV6: confirmar 11/86/2/GCBA/77,6%, no preview antigua;
8. si la imagen contradice HTML productivo actual, revisar archivo antes de asumir;
9. si falta una imagen crítica, pedirla específicamente.

Para retoques:
- no cambiar números por diferencia estética;
- no tocar otro dashboard;
- mantener jerarquía de información;
- confirmar preview final.

---

## 44. Checklist de migración de archivos

- [ ] `CONTEXTO_MAESTRO.md`
- [ ] `input/OCU26_BASE_DATOS.xlsx`
- [ ] `input_aux/YPF_GEO_COORDENADAS.xlsx.xlsx`
- [ ] `config/business_semantics.json`
- [ ] `scripts/validate_input.py`
- [ ] `scripts/transform_data.py`
- [ ] `scripts/semantic_model.py`
- [ ] `scripts/metrics_engine.py`
- [ ] `scripts/export_data.py`
- [ ] builders TV1–TV6
- [ ] templates TV1–TV6
- [ ] tests TV1–TV6
- [ ] `output/tv1_data.json` … `tv6_data.json`
- [ ] `tv1.html` … `tv6.html`
- [ ] `scripts/templates/assets/argentina_pais.geojson`
- [ ] `scripts/templates/assets/argentina_provincias.geojson`
- [ ] `audit_sources/*`
- [ ] `powerbi/`
- [ ] archivos PBIP/Power BI vigentes
- [ ] manual Brand Plus
- [ ] capturas finales TV1–TV6
- [ ] `GUIA_FINAL_TABLEROS_OCU26.pdf`
- [ ] documentación SharePoint/M365
- [ ] archivos históricos solo si se requiere trazabilidad adicional

---

## 45. Snapshot ejecutivo

1. Repo: `C:\brand plus\ocu26-dashboard`.
2. Branch: `main`.
3. Excel: `input/OCU26_BASE_DATOS.xlsx`.
4. SHA histórico esperado: `2f165e12...a2cd`.
5. Arquitectura centralizada.
6. Gates 1–4B terminados.
7. HTML no es motor de negocio.
8. APSA fuera.
9. London fuera.
10. YPF sin fill total.
11. TV1 = General.
12. TV2 = Digital.
13. TV3 = Estático.
14. TV4 = Pipeline.
15. TV5 = YPF.
16. TV6 = Demanda.
17. TV1: 964 Core; 495 activos.
18. TV1 YPF: 305; 61,6% del total activo; semántica a validar.
19. TV1 Estático: 119; 29,4%; YTD 28,6%.
20. TV2: 71/78,0%; fill 321/20,4%.
21. TV2: Shoppings 60/88,2%.
22. TV3: 119/422 = 28,2%.
23. TV3: 303 disponibles.
24. TV4: 77 campañas/421 activaciones.
25. TV4: 93,9% pipeline Core.
26. TV5: 451 estaciones/3.082 elementos.
27. TV5: 305 activas/67,6%.
28. TV5: 9 campañas/7.681 activaciones/2.651 elementos.
29. TV5: Punteras líder.
30. TV5 mapa: Natural Earth real, offline.
31. TV5 geo: 79/305 =25,9%.
32. TV6: 11 clientes directos, 86 marcas, 2 agencias.
33. TV6 top: GCBA.
34. TV6 Top5: 77,6% de 8.203.
35. TV6 cuerpo = Ene–Jul.
36. Programática: 112 activaciones sin agencia.
37. 377 activaciones con agencia A Confirmar.
38. Tests confirmados: TV1 39, TV2 37, TV4 35, TV5 34 pre-GEO.
39. TV3/TV5 post-GEO/TV6 tests finales = A VALIDAR.
40. Próximo: estética por sesión TV → Etapa 2 → Git/deploy → Power BI/SharePoint.

---

# ANEXO A — CONTEXTO MAESTRO ANTERIOR COMPLETO (~08:22 ART)

> Este anexo se conserva íntegramente para trazabilidad. Puede contener estados que fueron reemplazados durante el 10/8. **No prevalece** sobre las secciones 0–45 anteriores.


# CONTEXTO MAESTRO DEL PROYECTO

**Proyecto:** BRAND PLUS · OCU26 · Sistema de inteligencia comercial, ocupación y dashboards  
**Estado documentado:** 10 de agosto de 2026, aproximadamente 08:22 ART
**Repositorio principal:** `C:\brand plus\ocu26-dashboard`  
**Branch vigente:** `main`  
**Fuente local vigente:** `input/OCU26_BASE_DATOS.xlsx`  
**Prioridad inmediata:** cerrar TV6 · DEMANDA COMERCIAL con el enfoque híbrido confirmado: KPI cards superiores = foto mensual de JULIO 2026; cuerpo inferior = acumulado ENE–JUL 2026; ranking dinámico MARCAS → AGENCIAS → PROGRAMÁTICA; matriz Demanda por circuito acumulada; footer triple LECTURA / PUNTO POSITIVO / A ATENDER. TV5 · YPF queda después. No reabrir TV1–TV4 ni hacer microauditorías adicionales salvo contradicción real.
**Finalidad:** permitir continuar el proyecto en un chat nuevo de ChatGPT y/o una sesión nueva de Claude Code sin releer el chat original.

> **REGLA DE VIGENCIA:** este documento integra el contexto maestro anterior y todas las decisiones posteriores tomadas durante la implementación productiva de TV1, TV2 y TV3. Una decisión posterior reemplaza a una anterior cuando existe contradicción. La actualización más reciente al comienzo del archivo prevalece sobre actualizaciones anteriores y anexos históricos.


# ACTUALIZACIÓN VIGENTE CM1 — 10/8/2026 ~08:22 ART

> **ESTA ES LA CAPA DE VIGENCIA MÁS RECIENTE DEL DOCUMENTO.**
> Si contradice una actualización anterior, **prevalece esta sección**.
> Conservar secciones anteriores únicamente para trazabilidad.
> Regla transversal: no inventar datos, no completar vacíos, no asumir resultados de comandos cuya salida no haya sido vista. Marcar `A VALIDAR` cuando falte evidencia.

## C0. Punto exacto actual

La producción acelerada de OCU26 continúa con foco en cerrar **TV6 · DEMANDA COMERCIAL**.

Estado vigente:

- **TV1 — Core Comercial:** cerrada para entrega inmediata; quedan microajustes/Etapa 2.
- **TV2 — Core Comercial Digital:** cerrada para entrega inmediata.
- **TV3 — Core Comercial Estático:** cerrada para entrega inmediata.
- **TV4 — Pipeline Comercial:** aprobada para entrega inmediata, 35/35 tests PASS confirmados.
- **TV6 — Demanda Comercial:** IMPLEMENTADA en una primera versión productiva y luego **REDISEÑADA EN SU LÓGICA TEMPORAL** por decisión del usuario.
- **TV5 — YPF:** queda después de TV6.
- **Power BI:** pausado.
- **Git / deploy:** bloqueados hasta aprobación explícita.

### Decisión nueva y vigente de TV6

La TV6 ya NO debe leer todo con el mismo período.

Se adopta enfoque **HÍBRIDO MENSUAL + HISTÓRICO**:

**ARRIBA = JULIO 2026**
- 5 KPI cards superiores.

**ABAJO = ACUMULADO ENE–JUL 2026**
- ranking dinámico Marcas / Agencias / Programática;
- matriz Demanda por circuito;
- lectura ejecutiva histórica.

### Motivo del cambio

La primera implementación productiva de TV6, al limitar los rankings inferiores a julio, mostraba muy pocas agencias y poca o nula programática identificada. El usuario señaló que sabe que existen más agencias y actividad programática y propuso que, dado que la TV6 es una lectura comercial general, el cuerpo inferior represente el histórico acumulado hasta julio.

Se aceptó esta lógica porque:
- las cards superiores siguen mostrando la foto del mes;
- el cuerpo inferior muestra la estructura real de la demanda acumulada;
- evita que un mes puntual oculte agencias/programática que sí participan en otros meses;
- mantiene separadas dos preguntas distintas: “qué está pasando este mes” y “quiénes componen la demanda comercial del año”.

Esta decisión reemplaza el enfoque anterior de TV6 completamente mensual.

---

## C1. TV6 — referencia visual y producto

### Referencia visual

Archivo de referencia ya colocado en `audit_sources`.

Adjunto histórico: `TV6_REFERENCE.html.html`.

Advertencia:
- internamente la referencia legacy estaba rotulada como TV5;
- contenía `window.OCU_DATA`;
- sus números NO son fuente de verdad;
- se usa solamente para layout, jerarquía, interacción, paleta y composición.

### Producto TV6 creado/esperado

- `scripts/build_tv6_dashboard.py`
- `scripts/templates/tv6_template.html`
- `tests/test_build_tv6_dashboard.py`
- `output/tv6_data.json`
- `tv6.html`

### Arquitectura

Pipeline central validado → builder TV6 → payload TV6 → template → `tv6.html`.

Regla: la lógica de negocio queda en builder/capa de datos, no en HTML.

---

## C2. Universo vigente TV6

TV6 integra **CORE COMERCIAL COMPLETO + YPF**.

Circuitos usados en auditorías TV6:
- `CENCOSUD`
- `REMEROS`
- `PANTALLAS_LED`
- `PILAR_FRONTLIGHT`
- `AA2000`
- `YPF`

Filtro base: `OPERATIVO_GENERAL`.

### Exclusiones

Mantener siempre:
- APSA excluido;
- London Supply excluido.

### Deduplicación

No sumar Core + YPF aritméticamente para entidades únicas. Aplicar `distinct` sobre clave/campo canónico para Marca, Cliente, Agencia e `IDCampaña`.

Ejemplo: una marca presente en Core y YPF cuenta una sola vez como marca activa.

---

## C3. Auditorías ejecutadas para TV6 antes del enfoque híbrido

Se hicieron consultas de solo lectura en Claude Code usando `load_pipeline(vi.DEFAULT_INPUT_PATH)`, `semantic_result['maestro']`, `semantic_result['campanas']`, `filter_universe(..., 'OPERATIVO_GENERAL', config)` y, exploratoriamente, `engine._campanas_overlap(...)`.

### Nota sobre `_campanas_overlap`

Se utilizó como herramienta exploratoria/auditoría. No debe convertirse automáticamente en dependencia productiva si existe una vía pública/canónica equivalente.

### Consultas ejecutadas

1. Columnas y distribución de `PortfolioTier` y `CircuitoNegocio`.
2. Scope TV6 julio: elementos TV6, filas del scope, valores `PROGRAMATICA`, nulos/blanks de `Cliente` y `Agencia`, top values de Cliente y Agencia.
3. Desambiguación Cliente/Agencia: `Cliente` en `AGENCIA`, `A CONFIRMAR` o null vs `Agencia`; `Agencia == 'No'` vs Cliente; `GROUPM`/`CARAT`; `Marca` cuando `Cliente == 'AGENCIA'`.
4. Calidad adicional: `Agencia` nula con Cliente informado; `Agencia == 'A confirmar'`; confirmación del nombre de `IDCampaña`.
5. Marcas: total activaciones scope, `distinct IDCampaña`, `Marca.value_counts`, placeholders, `nunique Marca`.
6. Programática: registros `PROGRAMATICA == 'Si'`, Agencia/Cliente/Marca asociados y cruce global `Agencia × PROGRAMATICA`.
7. Matriz por familia: Pantallas LED, Shoppings Digital, Shoppings Estático, AA2000 / Pilar Frontlight, YPF.
8. Concentración Top 5: `activaciones Top 5 marcas / total activaciones del scope`.

### Regla importante sobre taxonomía

En auditorías se usó temporalmente una función manual `familia(row)` para reconciliar resultados. Esto es válido como auditoría, pero NO debe dejarse como taxonomía paralela hardcodeada si la capa semántica ya ofrece clasificación canónica.

---

## C4. Primera implementación productiva de TV6

Se ejecutó `build_tv6_data()` para inspeccionar el payload productivo y luego `./.venv/Scripts/python.exe scripts/build_tv6_dashboard.py`.

Después se ejecutaron tests específicos con `pytest tests/test_build_tv6_dashboard.py -q`.

### Evidencia disponible

Se vio una captura donde Claude había modificado el archivo de tests y proponía confirmar **39/39 PASS**, pero el resultado final de ese run NO quedó compartido explícitamente en este chat.

Por lo tanto:

**Resultado exacto de tests de la primera versión TV6: A VALIDAR**.

No afirmar 39/39 PASS sin evidencia textual/captura del resultado.

### Sanity / Git

Se ejecutaron comandos de solo lectura:
- `git status --porcelain=v1 -- [archivos protegidos]`
- `git status --porcelain=v1`

La salida concreta no fue compartida en el chat. Evidencia textual de `git status`: **A VALIDAR**.

---

## C5. Problema detectado en primera preview TV6

La primera preview productiva mostró una TV6 visualmente consistente, pero el usuario detectó un problema de lectura:
- sabe que existen más agencias;
- sabe que existe programática;
- en la vista de julio aparecían muy pocas agencias y la representación de programática parecía insuficiente.

Hipótesis aceptada: la pantalla estaba leyendo los rankings inferiores solo con JULIO 2026.

### Decisión del usuario

Mantener tarjetas superiores = mes correspondiente (Julio).

Cambiar cuerpo inferior = histórico acumulado hasta julio.

Además, el footer inferior debe adoptar el patrón:
- `LECTURA`
- `PUNTO POSITIVO`
- `A ATENDER`.

---

## C6. TV6 — lógica temporal vigente

### A. KPI CARDS SUPERIORES

Período: **JULIO 2026**.

Mantener las 5 cards:
1. `CLIENTES DIRECTOS ACTIVOS · JULIO`
2. `MARCAS ACTIVAS · JULIO`
3. `AGENCIAS ACTIVAS · JULIO`
4. `CLIENTE TOP IDENTIFICADO · JULIO`
5. `CONCENTRACIÓN TOP 5 · JULIO`

No pasar estas cards a histórico. No mezclar métricas de julio con acumulado sin rótulo explícito.

### B. CUERPO INFERIOR

Período: **01/01/2026 → 31/07/2026**.

Rotular `ACUMULADO ENE–JUL 2026`.

Aplica a ranking Marcas, ranking Agencias, ranking Programática y Demanda por circuito.

---

## C7. Ranking dinámico vigente

Ciclo exacto: **MARCAS → AGENCIAS → PROGRAMÁTICA**.

### Marcas
- Top 6 si entra limpio.
- Período Ene–Jul.
- Métrica acumulada canónica.
- Orden actividad desc → nombre asc.

### Agencias
- solo agencias identificadas reales;
- excluir placeholders `No`, `A confirmar`, vacíos, etc.;
- período Ene–Jul;
- Top 6.

### Programática

Programática = subset de agencias.

Usar únicamente registros con `PROGRAMATICA == 'Si'`.

Luego:
- asociar agencia solo si está realmente identificada;
- no inferir por nombre de cliente/marca;
- no hardcodear listas;
- no transformar vacíos en `No`.

Si existen activaciones programáticas sin agencia identificada:
- no eliminarlas;
- no asignarlas;
- preservarlas como pendiente;
- mostrar copy tipo `X activaciones programáticas pendientes de imputación de agencia`.

Una misma agencia puede aparecer en Agencias y en Programática si efectivamente tiene actividad programática.

---

## C8. Auditoría acumulada Ene–Jul ejecutada

Después de decidir el enfoque híbrido se ejecutaron consultas consolidadas sobre `2026-01-01 → 2026-07-31`.

### Auditoría 1 — agencias / marcas / programática

Se consultó:
- total activaciones Ene–Jul;
- `distinct IDCampaña`;
- `Agencia.value_counts`;
- agencias identificadas excluyendo `No` y `A confirmar`;
- ranking de agencias identificadas;
- Top marcas;
- `nunique Marca`;
- valores `PROGRAMATICA`;
- filas `PROGRAMATICA == 'Si'`;
- Agencia / Cliente / Marca de esas filas.

### Auditoría 2 — programática identificada vs pendiente

Se consultó:
- total `PROGRAMATICA == 'Si'`;
- distribución por Agencia;
- cantidad con agencia identificada;
- cantidad sin agencia identificada / pendiente.

### Auditoría 3 — matriz acumulada

Se consultó:
- Top 6 marcas Ene–Jul;
- cruce contra Pantallas LED, Shoppings Digital, Shoppings Estático, AA2000 / Pilar Frontlight y YPF;
- totales por familia;
- concentración Top 5 acumulada.

### Auditoría 4 — reconciliación de matriz

Se ejecutó otra consulta para imprimir la matriz completa y revisar `row sums`.

### Evidencia de valores

Las salidas numéricas concretas de estas consultas no fueron pegadas en el chat actual. Por lo tanto, metodología confirmada; números históricos exactos deben leerse del payload final o volver a pedirse a Claude si se requiere documentarlos. No inventarlos en CM1.

---

## C9. Concentración Top 5 — regla temporal

### Card superior

`CONCENTRACIÓN TOP 5 · JULIO` usa únicamente julio.

Definición usada en auditoría:
- Top 5 marcas;
- numerador = activaciones de esas Top 5;
- denominador = total activaciones del scope julio.

No presentar como campañas únicas.

### Histórico inferior

Si un insight menciona concentración Ene–Jul, recalcular expresamente sobre Ene–Jul. Nunca usar el porcentaje de julio como si fuera histórico.

---

## C10. Demanda por circuito — vigente

Panel derecho: `DEMANDA POR CIRCUITO`.

Período: **Ene–Jul 2026**.

Subtítulo recomendado: `Top marcas × grupo comercial · activaciones acumuladas Ene–Jul`.

Familias auditadas:
- Pantallas LED;
- Shoppings Digital;
- Shoppings Estático;
- AA2000 / Pilar Frontlight;
- YPF.

UI: `Estático`, nunca `Fijo`.

La matriz debe reconciliar con la misma unidad de actividad elegida para el ranking.

---

## C11. Footer TV6 — decisión nueva y vigente

La versión anterior mantenía un solo bloque `Lectura`. Eso queda reemplazado.

TV6 debe usar:

**LECTURA | PUNTO POSITIVO | A ATENDER**

### LECTURA
Combinar foto de julio arriba con estructura acumulada Ene–Jul abajo.

### PUNTO POSITIVO
Elegir un hecho factual real: diversidad de marcas, agencias identificadas, presencia programática, diversificación de circuitos o crecimiento real.

### A ATENDER
Prioridades:
1. programática sin agencia imputada;
2. pendientes / `A confirmar`;
3. concentración elevada;
4. dependencia fuerte de YPF;
5. otra señal material real.

No inventar causas.

---

## C12. Payload híbrido TV6

Después de implementar el enfoque híbrido se ejecutó `build_tv6_data()` y se imprimieron:
- `data['kpis']`
- `data['ranking']`
- `data['matriz']`
- `data['insights']`
- `data['universo']`

Luego se imprimieron específicamente:
- `insights['lectura']`
- `insights['punto_positivo']`
- `insights['a_atender']`

Esto confirma que el builder ya fue adaptado para producir una estructura híbrida con KPIs, ranking, matriz, insights triple y universo. Los valores concretos no quedaron pegados en el chat.

---

## C13. Regeneración híbrida TV6

Se ejecutó nuevamente `./.venv/Scripts/python.exe scripts/build_tv6_dashboard.py`.

Luego se ejecutó `./.venv/Scripts/python.exe -m pytest tests/test_build_tv6_dashboard.py -q`.

### Resultado de tests híbridos

La orden fue ejecutada/solicitada, pero el resultado final NO fue compartido.

**TV6 híbrida — test count/result: A VALIDAR**.

### Git status posterior

También se ejecutó `git status --porcelain=v1` después de la versión híbrida. La salida no fue compartida.

**TV6 híbrida — sanity Git: A VALIDAR**.

---

## C14. Estado visual TV6

### Primera preview

Se vio una preview 1920×1080 del primer producto TV6. El problema funcional detectado fue que la lectura mensual de abajo hacía parecer que había pocas agencias/programática.

### Preview híbrida

Después del cambio a Ene–Jul en el cuerpo, todavía NO se compartió una captura final en este chat.

**Preview final híbrida 1920×1080: PENDIENTE / A VALIDAR**.

No cerrar formalmente TV6 hasta verla.

---

## C15. Próximo paso exacto

No volver a auditar datos salvo contradicción real.

1. Obtener/confirmar resultado final de `tests/test_build_tv6_dashboard.py`.
2. Confirmar salida de sanity/Git si es necesario.
3. Abrir `tv6.html` de la raíz, no el template.
4. Revisar preview final 1920×1080.
5. Validar: cards arriba claramente Julio; cuerpo claramente Ene–Jul; ciclo Marcas → Agencias → Programática; programática identificada vs pendiente; matriz histórica; `Lectura / Punto positivo / A atender`; sin overflow; sin labels ambiguos.
6. Hacer solo correcciones visuales/copy críticas.
7. Cerrar TV6.
8. Pasar a TV5 YPF.
9. Luego revisión conjunta TV1–TV6.
10. Git/deploy solo con aprobación explícita.

---

## C16. A VALIDAR actuales

### TV6 — bloque inmediato

1. Resultado exacto de tests específicos de la versión híbrida.
2. Preview final híbrida 1920×1080.
3. Valores finales productivos de KPIs julio, Top marcas Ene–Jul, Top agencias Ene–Jul, Programática Ene–Jul, pendientes programáticos, matriz acumulada y concentración histórica si se usa.
4. Copy final de `Lectura / Punto positivo / A atender`.
5. Sanity final / `git status`.
6. Confirmación visual de que `tv6.html` final no tiene overflow.
7. Confirmación de que Programática muestra agencia identificada real y pendientes separados.

### Pendientes posteriores

8. TV5 YPF.
9. Revisión conjunta TV1–TV6.
10. TV2 ~68 vs ~61 Shoppings Digital elegibles.
11. TV3 clasificación `Otro`.
12. TV3 ajustes menores.
13. YPF surrogate → ID real/APIE si aparece.
14. Power BI `dataType`.
15. Integración/Netlify si sigue pendiente.

---

## C17. Errores / comportamientos que NO repetir

1. No usar un único mes para representar toda la estructura histórica de demanda si la pantalla busca una lectura general.
2. No asumir “2 agencias” como universo comercial anual cuando la métrica solo corresponde a julio.
3. No inferir agencia programática.
4. No eliminar programática sin agencia identificada; conservar como pendiente.
5. No mezclar concentración julio con histórico.
6. No mostrar `Fijo`; usar `Estático`.
7. No abrir `scripts/templates/tv6_template.html` para evaluar datos finales; abrir `tv6.html`.
8. No afirmar PASS de tests sin ver salida.
9. No convertir una taxonomía exploratoria `familia(row)` en una taxonomía productiva paralela si la semántica central ya resuelve el concepto.
10. No reabrir microauditorías ya cubiertas sin una contradicción real.

---

## C18. Snapshot ejecutivo actualizado

1. TV1 cerrada para entrega inmediata.
2. TV2 cerrada para entrega inmediata.
3. TV3 cerrada para entrega inmediata.
4. TV4 aprobada; 35/35 tests PASS.
5. TV6 es el dashboard activo.
6. TV6 integra Core Comercial completo + YPF.
7. APSA y London siguen excluidos.
8. TV6 ya tiene builder/template/tests/payload/HTML productivos.
9. La referencia TV6 es legacy y solo visual.
10. Primera TV6 fue mensual tanto arriba como abajo.
11. El usuario detectó que así aparecían pocas agencias/programática.
12. Nueva decisión: TV6 híbrida.
13. Cards superiores = Julio 2026.
14. Cuerpo inferior = acumulado Ene–Jul 2026.
15. Ranking = Marcas → Agencias → Programática.
16. Programática es subset de Agencias.
17. `PROGRAMATICA == 'Si'` es el criterio de auditoría.
18. No inferir agencia programática.
19. Programática sin agencia identificada queda pendiente.
20. Agencias excluyen placeholders `No` / `A confirmar` para rankings identificados.
21. Marcas abajo = Top acumuladas Ene–Jul.
22. Matriz abajo = acumulada Ene–Jul.
23. Familias: Pantallas LED / Shoppings Digital / Shoppings Estático / AA2000-Pilar / YPF.
24. Concentración card = Julio.
25. Concentración histórica, si se usa, se recalcula Ene–Jul.
26. Footer TV6 cambió a Lectura / Punto positivo / A atender.
27. Builder híbrido ya fue regenerado.
28. Payload híbrido ya fue inspeccionado.
29. Insights triple ya existen en payload.
30. Tests híbridos fueron lanzados pero resultado final no fue compartido.
31. Git status fue lanzado pero salida no fue compartida.
32. Preview híbrida final todavía no fue mostrada.
33. No más auditorías salvo contradicción real.
34. Próximo paso: tests/sanity final + preview 1920×1080.
35. Después cerrar TV6.
36. Luego TV5 YPF.
37. Después revisión conjunta TV1–TV6.
38. Git/deploy solo con aprobación.


# ACTUALIZACIÓN VIGENTE CM1 — 10/8/2026 ~06:20 ART

> **ESTA ES LA CAPA DE VIGENCIA MÁS RECIENTE DEL DOCUMENTO.**
> Si contradice una actualización anterior, **prevalece esta sección**.
> Conservar las secciones anteriores únicamente para trazabilidad.
> Regla: no inventar datos ni semánticas. Si algo no está respaldado por la base, el pipeline, un test, un payload o una decisión explícita del usuario, marcar `A VALIDAR`.

## B0. Punto exacto actual

La producción acelerada de las TVs OCU26 continúa para la presentación del lunes 10 de agosto de 2026.

Estado vigente:

- **TV1 — Core Comercial:** prácticamente cerrada para entrega.
- **TV2 — Core Comercial Digital:** cerrada para entrega inmediata; quedan microauditorías de Etapa 2.
- **TV3 — Core Comercial Estático:** productiva y aceptada para continuar; quedan mejoras menores de Etapa 2.
- **TV4 — Pipeline Comercial:** **APROBADA PARA LA ENTREGA INMEDIATA**.
  - La versión final visible fue revisada en 1920×1080.
  - Los tests específicos TV4 dieron **35/35 PASS**.
  - Se resolvió la ambigüedad `4 reservas` vs `15 reservas/activaciones`: la KPI superior expresa **4 campañas futuras** y el panel expresa **15 activaciones cargadas**.
  - No hacer más auditorías de TV4 ahora.
- **TV6 — Demanda Comercial:** pasa a ser el **SIGUIENTE DASHBOARD A IMPLEMENTAR**.
  - El usuario decidió trabajar TV6 antes de TV5 YPF.
  - Existe referencia visual cargada en `audit_sources` y también se adjuntó al chat como `TV6_REFERENCE.html.html`.
  - Esa referencia es legacy: internamente todavía se identifica como `TV5` y contiene `window.OCU_DATA`; debe usarse **solo como referencia visual**, nunca como fuente de verdad.
  - Cambio principal confirmado: el gráfico/ranking principal debe ciclar **Marcas → Agencias → Programática**.
  - `Programática` = **solamente agencias que hacen programática**, según dato canónico; no inferir manualmente.
  - El resto de la referencia debe replicarse, salvo reglas globales ya vigentes.
- **TV5 — YPF:** queda para después de TV6.
- **Power BI:** pausado.
- **Git / deploy:** siguen bloqueados hasta revisión conjunta y aprobación explícita.

### Orden vigente de trabajo

**TV6 Demanda Comercial → TV5 YPF → revisión conjunta TV1–TV6 → Git/deploy con aprobación.**

Este orden **reemplaza** el roadmap anterior `TV5 YPF → TV6 Demanda`.

---

## B1. TV4 — estado final aprobado

### Nombre

`TV4 · PIPELINE COMERCIAL`

### Resultado visual final

La versión final mostrada a las ~06:11 ART fue revisada y el usuario indicó que ya quedó.

Estructura final:

- 5 KPI cards superiores;
- `Estado del pipeline`;
- `Próximas activaciones · 30 días`;
- `LECTURA`;
- `PUNTO POSITIVO`;
- `A ATENDER`.

Sin comparativos mensuales en las KPI superiores.

### Valores finales visibles

#### Card 1 — ACTIVIDAD ACTUAL

- **77 campañas únicas activas**
- **421 activaciones en curso**
- Nota:
  `1 campaña puede estar en varios elementos`

Lectura:
`77` y `421` son granularidades diferentes y no deben sumarse ni tratarse como equivalentes.

#### Card 2 — RESERVAS FUTURAS

- **4 campañas**
- Nota:
  `inicio posterior al corte · 15 activaciones cargadas`

Esta corrección de copy resolvió la confusión visual anterior.

Regla:
- 4 = campañas únicas futuras;
- 15 = activaciones/reservas cargadas asociadas.

#### Card 3 — INICIAN · 30 DÍAS

- **4**

#### Card 4 — FINALIZAN · 30 DÍAS

- **9**

#### Card 5 — PIPELINE EN CORE

- **93,9%**
- Nota:
  **77 de 82 campañas del universo general**

Esta card reemplaza el 100% tautológico de la referencia legacy.

### Estado del pipeline

- Activas: **421**
- Reservas: **15 activaciones**
- Finalizadas: **1.222 histórico**

### Activaciones activas hoy · por grupo comercial

- Shoppings Digital: **229**
- Shoppings Estático: **124**
- Pantallas LED: **68**
- AA2000 / Pilar Frontlight: **0**

La suma de los tres grupos activos materiales:
229 + 124 + 68 = **421**.

### Próximas activaciones visibles

4 eventos/campañas, todos con fecha `01 Ago`:

1. **NETFLIX MORIA** — Shoppings Estático
2. **NINTENDO** — Shoppings Estático
3. **OAK STREET** — Shoppings Digital
4. **PLAYSTATION** — Shoppings Digital + Shoppings Estático

No rellenar una quinta fila artificial.

### Insights finales visibles

#### LECTURA
`Hoy corren 77 campañas únicas activas en 421 activaciones del Core Comercial. El pipeline registra 4 reservas futuras, con 4 campañas que inician y 9 que finalizan en los próximos 30 días.`

#### PUNTO POSITIVO
`Hay 4 reservas futuras cargadas, que suman 15 activaciones adicionales al pipeline.`

#### A ATENDER
`Finalizan 9 campañas en los próximos 30 días frente a 4 que inician.`

### Tests

Captura de Claude:
**All 35 tests pass.**

Por lo tanto:
- `tests/test_build_tv4_dashboard.py`: **35/35 PASS**.

Se rechazó correr nuevamente suites completas de TV1–TV3 por ser redundante para la entrega inmediata.

### Sanity

Se ejecutó un sanity mínimo que verificaba:
- estructura de payload TV4;
- ausencia de YPF/APSA/London;
- SHA del Excel;
- archivos protegidos vía `git status`.

La salida textual completa del sanity no quedó reproducida en este chat.
El usuario posteriormente indicó que TV4 ya quedó y mostró la preview final.

Para auditoría formal:
- resultado visual: confirmado;
- tests TV4: confirmado;
- resultado textual exacto del SHA: `A VALIDAR` si se necesita evidencia documental posterior.

### Decisión final

**TV4 queda cerrada para la entrega inmediata.**
No reabrir esta TV salvo error real o ajuste pedido por el usuario.

---

## B2. Cambio de prioridad: TV6 antes que TV5

### Decisión anterior

Después de TV4:
`TV5 YPF → TV6 Demanda`

### Decisión vigente

Después de cerrar TV4:
`TV6 Demanda → TV5 YPF`

### Motivo operativo

El usuario pidió avanzar ahora con la pantalla de demanda y dejar YPF para después.

### Consecuencia

El próximo Claude Code debe abrirse enfocado únicamente en TV6.

No invertir el orden por el roadmap histórico anterior.

---

## B3. TV6 — referencia visual actual

### Archivo adjunto en chat

`TV6_REFERENCE.html.html`

### Archivo de repo

El usuario indicó que la referencia ya fue agregada a `audit_sources`.

La ruta exacta dentro del repo debe usarse tal como existe.
Si el archivo fue normalizado, la ruta esperada es:

`audit_sources/TV6_REFERENCE.html`

Si el repo conserva doble extensión u otro nombre, **no inventar**: comprobar una sola vez.

### Advertencia crítica

La referencia es una versión histórica de Demanda Comercial y todavía contiene internamente:

- `<title>Brand Plus · TV5 · Demanda Comercial</title>`
- `<section ... id="tv5">`
- eyebrow `TV 5`
- data attributes `tv5`
- `window.OCU_DATA`
- valores legacy.

Por lo tanto:

> La referencia define estética/composición, NO numbering, datos ni lógica vigente.

La TV productiva nueva debe ser TV6 y usar builder/payload propios.

### Estructura visual legacy confirmada

Header:
- eyebrow legacy: `TV 5 · Demanda Comercial`
- título: `Así se comporta la demanda`
- subtítulo: `Clientes, marcas y agencias · dónde pautan`
- pregunta: `¿Quién pauta, en qué circuitos y con qué concentración?`

Fila KPI:
- 5 cards.

Cuerpo:
- panel izquierdo de ranking dinámico;
- panel derecho `Demanda por circuito`.

Panel izquierdo legacy:
- título dinámico `Top clientes`;
- ciclo:
  `clientes · marcas · agencias`;
- activaciones del mes.

Panel derecho:
- `Demanda por circuito`
- subtítulo:
  `Top clientes × grupo comercial · activaciones del mes`
- matriz por actor comercial y grupo.

Inferior:
- un único bloque vertical `Lectura`.

### Regla de referencia

Usar esta composición como base.
No serializar ni copiar `window.OCU_DATA`.

---

## B4. TV6 — universo comercial vigente

### Decisión acumulada

TV6 es una pantalla de **Demanda Comercial** y debe incorporar:

- **Core Comercial completo**
- **YPF**

La última instrucción del usuario refuerza explícitamente que ahora se integra **todo el Core Comercial**.

No hubo instrucción explícita que retire YPF de TV6, por lo que continúa vigente la decisión anterior:

**Universo TV6 = Core Comercial + YPF**

### Core Comercial

Usar las familias canónicas vigentes del Core y no una suma manual improvisada.

### YPF

YPF participa porque TV6 mide demanda/marcas/clientes, no capacidad física.

### Exclusiones globales

Mantener:
- APSA excluido;
- London Supply excluido.

### Regla de deduplicación

Core + YPF no se suman aritméticamente para entidades únicas.

Para:
- marca;
- cliente;
- agencia;
- campaña;

primero construir el universo combinado y luego aplicar `distinct` sobre la clave/campo canónico.

Ejemplo:
una marca presente simultáneamente en Core y YPF cuenta una sola vez como marca activa.

---

## B5. TV6 — cambio principal confirmado en el gráfico/ranking

### Legacy

La referencia cicla:

`Clientes → Marcas → Agencias`

### Nueva decisión del usuario

El gráfico/ranking principal debe ciclar:

**MARCAS → AGENCIAS → PROGRAMÁTICA**

Esto **reemplaza Clientes** dentro del ciclo del gráfico principal.

### Programática — definición obligatoria

`Programática` NO es un tercer tipo de cliente separado.

Es un **subconjunto de Agencias**:

> solo las agencias que efectivamente hacen programática.

### Reglas

- No inferir programática por nombre de agencia.
- No mantener una lista manual hardcodeada si existe un campo canónico.
- Auditar una sola vez cuál es el campo/flag real (`Programática`, equivalente o clasificación semántica vigente).
- Filtrar únicamente agencias con condición programática positiva.
- Deduplicar la agencia usando su clave/campo canónico.
- Si una agencia tiene actividad tradicional y programática, puede aparecer en `Agencias` y también en `Programática`; no es un error porque `Programática` es una vista/subconjunto.
- No convertir vacíos o `A confirmar` en `No`.
- Preservar `MetricStatus` / calidad de datos.
- Si el campo no permite saber con certeza qué agencias son programáticas, marcar `A VALIDAR` y no inventar clasificación.

### Métrica del ranking

La referencia legacy usa `activaciones del mes`.

La implementación productiva debe confirmar la unidad canónica antes de construir:
- activaciones;
- campañas únicas;
- otra métrica central validada.

No copiar valores legacy.

---

## B6. TV6 — “lo demás replicar”

Última instrucción del usuario:

> Fuera del cambio del gráfico a Marcas / Agencias / Programática, **replicar lo demás** de la referencia.

Por lo tanto, salvo una regla global ya confirmada que obligue a corregir un término o una inconsistencia:

### Mantener

- composición general de 5 KPI cards;
- header y pregunta ejecutiva;
- panel izquierdo dinámico;
- panel derecho `Demanda por circuito`;
- matriz de demanda por grupo comercial;
- jerarquías, tamaños y estética Brand Plus;
- 1920×1080;
- Poppins;
- colores institucionales;
- sin scroll/overflow.

### Regla global que sí prevalece

En UI:
`Shoppings Estático`

Nunca:
`Shoppings Fijo`

aunque la referencia legacy diga `Shoppings Fijo`.

### Insight inferior

La referencia legacy tiene un solo bloque `Lectura`.

La última instrucción `lo demás replicar` indica conservar esa composición salvo que el usuario posteriormente pida aplicar el patrón triple.

Esto **reemplaza como decisión visual de TV6** cualquier propuesta anterior del asistente que hubiese sugerido automáticamente:
`Lectura | Punto positivo | A atender`.

No cambiarlo sin nueva instrucción.

---

## B7. TV6 — KPI cards que la referencia sugiere conservar

La referencia legacy tiene cinco conceptos:

1. `CLIENTES DIRECTOS ACTIVOS · JULIO`
2. `MARCAS ACTIVAS · JULIO`
3. `AGENCIAS ACTIVAS · JULIO`
4. `CLIENTE TOP IDENTIFICADO`
5. `CONCENTRACIÓN TOP 5`

La última instrucción del usuario es replicar el resto, por lo que estos conceptos son la base visual/funcional actual.

### Importante

Los valores legacy de la referencia:
- 34;
- 57;
- 2;
- GCBA;
- 23%;

NO son valores vigentes.

Deben recalcularse sobre:
**Core Comercial completo + YPF**

con exclusiones y deduplicaciones correctas.

### Calidad de datos

La referencia histórica contiene conceptos como:
- `A confirmar`;
- vía agencia;
- pendientes a imputar.

No convertir registros pendientes en clientes/agencias identificados.

La nueva auditoría debe separar:
- identificado;
- pendiente;
- vía agencia;
- directo;
- programático.

---

## B8. TV6 — matriz “Demanda por circuito”

### Mantener concepto

Panel derecho:
`DEMANDA POR CIRCUITO`

La referencia cruza actores comerciales con grupos de inventario y activaciones del mes.

### Actualización requerida

La matriz productiva debe cubrir el **Core Comercial completo** y el universo definido para TV6.

No copiar columnas legacy de manera ciega.

Familias/grupos visibles deben provenir de taxonomía vigente.

### Nomenclatura

Si aparece estático:
`SH. ESTÁTICO`
o equivalente claro.

No `SH. FIJO`.

### YPF

Como TV6 incluye YPF en el universo de demanda, resolver en auditoría si YPF:
- necesita una columna propia;
- se representa por una familia/circuito existente;
- o se integra sin ampliar la matriz por restricciones de espacio.

No inventar. Elegir la opción que refleje mejor la semántica real y mantenga legibilidad.

---

## B9. TV6 — header vigente recomendado a partir de referencia

Mantener el concepto de la referencia, corrigiendo numbering:

Eyebrow:
`TV 6 · OCU26 · INTELIGENCIA COMERCIAL`

Título:
`ASÍ SE COMPORTA LA DEMANDA`

Subtítulo base:
`Clientes, marcas y agencias · dónde pautan`

Como el nuevo gráfico incorpora Programática, el copy puede requerir un ajuste mínimo para que no parezca omitida.

Opción compatible con la decisión actual:
`Marcas, agencias y programática · dónde pautan`

o:
`Clientes, marcas, agencias y programática · dónde pautan`

La redacción final queda `A VALIDAR` en implementación para no cambiar copy sin necesidad.

Pregunta de referencia:
`¿Quién pauta, en qué circuitos y con qué concentración?`

Mantener salvo nueva instrucción.

---

## B10. TV6 — implementación productiva esperada

Crear/usar:

- `scripts/build_tv6_dashboard.py`
- `scripts/templates/tv6_template.html`
- `tests/test_build_tv6_dashboard.py`
- `output/tv6_data.json`
- `tv6.html`

### Patrón

Pipeline validado
→ builder/capa de datos
→ payload TV6
→ template
→ HTML.

No lógica de negocio dentro del HTML.

### Referencia

`audit_sources/TV6_REFERENCE...`

read-only.

### Protecciones

No tocar:
- `scripts/validate_input.py`
- `scripts/transform_data.py`
- `scripts/semantic_model.py`
- `scripts/metrics_engine.py`
- `scripts/export_data.py` salvo bloqueo real;
- `config/business_semantics.json`;
- Excel;
- TV1;
- TV2;
- TV3;
- TV4;
- Power BI;
- `audit_sources/*`;
- Git/deploy.

---

## B11. TV6 — auditoría única requerida antes de implementar

Hacer **una sola auditoría consolidada**.

Resolver:

### Universo
- Core completo incluido;
- YPF incluido;
- APSA/London excluidos.

### Clientes directos
- definición canónica;
- activos julio;
- calidad/pending.

### Marcas
- distintas marcas activas julio;
- actividad asociada;
- ranking.

### Agencias
- agencias identificadas activas;
- pending / A confirmar;
- ranking.

### Programática
- campo/flag real;
- agencias con programática;
- ranking;
- actividad asociada.

### Cliente top
- excluir placeholders si corresponde;
- distinguir directo vs vía agencia.

### Concentración Top 5
Definir exactamente:
- Top 5 de qué;
- numerador;
- denominador;
- unidad de actividad.

No copiar el 23% legacy.

### Matriz por circuito
- grupos reales del Core completo;
- tratamiento YPF;
- activaciones/medida canónica.

Después:
**no microauditorías adicionales salvo error real.**

---

## B12. TV6 — tests mínimos esperados

1. payload solo TV6;
2. Core Comercial completo incluido;
3. YPF incluido;
4. APSA excluido;
5. London excluido;
6. marcas deduplicadas;
7. agencias deduplicadas;
8. programática es subset de agencias;
9. programática no incluye agencias sin flag/campo positivo;
10. no inferencia manual de programática;
11. clientes directos no mezclan placeholders indebidamente;
12. concentración Top 5 correcta;
13. ranking Marcas correcto;
14. ranking Agencias correcto;
15. ranking Programática correcto;
16. ciclo visual exactamente Marcas → Agencias → Programática;
17. matriz reconcilia con la medida de demanda elegida;
18. UI usa `Estático`;
19. no legacy `OCU_DATA`;
20. SHA Excel intacto;
21. TV1 intacta;
22. TV2 intacta;
23. TV3 intacta;
24. TV4 intacta;
25. referencia TV6 intacta.

No full suite.

---

## B13. TV6 — estado de referencia legacy

La referencia adjunta contiene todavía datos legacy de una antigua TV5 de demanda.

Entre otras cosas, el HTML histórico:
- cicla `Top clientes`, `Top marcas`, `Top agencias`;
- usa una matriz `Demanda por circuito`;
- contiene valores hardcodeados/legacy en `window.OCU_DATA`.

Estos datos históricos NO deben considerarse actuales.

Uso permitido:
- layout;
- estilos;
- orden visual;
- interacción de ciclo;
- estructura de matriz.

Uso prohibido:
- copiar números;
- asumir scopes;
- asumir identificadores;
- serializar todo `OCU_DATA`.

---

## B14. Etapa 2 — backlog actualizado

Además de pendientes anteriores, agregar:

### TV6 / Demanda
- revisar calidad de `Agencia`;
- validar semántica de Programática;
- resolver pendientes `A confirmar`;
- reconciliar demanda directa / vía agencia / programática;
- revisar concentración Top 5;
- validar deduplicación Core + YPF;
- documentar taxonomía de actores comerciales.

Estas revisiones profundas no deben frenar la entrega si el dashboard puede construirse correctamente con reglas canónicas actuales.

---

## B15. Próximo paso exacto

1. Abrir o continuar sesión Claude Code exclusiva para **TV6 Demanda Comercial**.
2. Leer focalizadamente esta actualización de CM1.
3. Confirmar una sola vez la ruta exacta de `TV6_REFERENCE` en `audit_sources`.
4. Auditoría consolidada:
   - Core completo + YPF;
   - marcas;
   - agencias;
   - programática;
   - clientes directos;
   - concentración;
   - matriz.
5. Implementar TV6.
6. Tests específicos TV6.
7. Sanity.
8. Preview 1920×1080.
9. Correcciones visuales críticas únicamente.
10. Cerrar TV6.
11. Pasar a TV5 YPF.
12. Revisión conjunta TV1–TV6.
13. Git/deploy solo con aprobación explícita.

---

## B16. A VALIDAR actuales

### Entrega inmediata
1. TV6: valores productivos de las 5 KPI cards.
2. TV6: campo canónico que identifica programática.
3. TV6: ranking exacto de Programática.
4. TV6: métrica/unidad final del ranking.
5. TV6: tratamiento de YPF dentro de la matriz por circuito.
6. TV6: copy final de subtítulo incorporando Programática.
7. TV6: concentración Top 5 recalculada.
8. TV6: tests y preview.
9. TV5 YPF: diseño final.
10. TV1: cambio menor de copy Core, si aún no fue aplicado.

### Etapa 2
11. TV2: ~68 vs ~61 elegibles Shoppings Digital.
12. TV3: clasificación `Otro`.
13. TV3: mejoras semánticas/visuales menores.
14. YPF: sustituir surrogate por APIE/Station ID real.
15. Power BI: error `dataType`.
16. Netlify: clasificación exacta de integración si sigue pendiente.

### Evidencia formal
17. TV4: SHA textual exacto del sanity, solo si se requiere documentación formal posterior.

---

## B17. Snapshot ejecutivo actualizado

1. TV1 Core Comercial: cerrada para entrega, con copy menor pendiente.
2. TV2 Core Digital: cerrada para entrega.
3. TV3 Core Estático: cerrada para entrega.
4. TV4 Pipeline: aprobada.
5. TV4 tiene 35/35 tests PASS.
6. TV4 final: 77 campañas únicas y 421 activaciones.
7. TV4 reservas: 4 campañas / 15 activaciones.
8. TV4 inician 30d: 4.
9. TV4 finalizan 30d: 9.
10. TV4 Pipeline en Core: 93,9%, 77 de 82 campañas.
11. TV4 grupos: 229 Shoppings Digital, 124 Estático, 68 Pantallas, 0 AA2000/Pilar.
12. TV4 no requiere más auditorías para esta entrega.
13. Siguiente dashboard: TV6 Demanda Comercial.
14. TV5 YPF queda después.
15. La referencia TV6 es un HTML legacy internamente rotulado TV5.
16. La referencia es solo visual, no fuente de datos.
17. TV6 conserva el concepto `Así se comporta la demanda`.
18. TV6 debe integrar el Core Comercial completo.
19. Por continuidad, TV6 mantiene YPF dentro del universo de demanda.
20. APSA y London siguen excluidos.
21. Core + YPF debe deduplicarse.
22. El gráfico principal ya NO cicla Clientes/Marcas/Agencias.
23. Nuevo ciclo: Marcas → Agencias → Programática.
24. Programática es subset de Agencias.
25. Programática debe salir de dato canónico, no inferencia manual.
26. El resto de la referencia debe replicarse.
27. Las 5 cards legacy son la base, pero se recalculan.
28. La matriz `Demanda por circuito` se conserva conceptualmente.
29. `Shoppings Fijo` debe mostrarse `Shoppings Estático`.
30. La referencia tiene un solo bloque `Lectura`; por `lo demás replicar`, se conserva salvo nueva orden.
31. No copiar `window.OCU_DATA`.
32. Builder TV6 debe producir payload propio.
33. Una auditoría consolidada antes de implementar.
34. Tests específicos TV6, no full suite.
35. Después de TV6: TV5 YPF.
36. Después: revisión conjunta, Git/deploy con aprobación.
37. Etapa 2 absorbe microauditorías profundas.


# ACTUALIZACIÓN VIGENTE CM1 — 10/8/2026 ~00:20 ART

> **ESTA ES LA CAPA DE VIGENCIA MÁS RECIENTE DEL DOCUMENTO.**
> Si contradice la actualización de 9/8/2026 ~23:27 ART o cualquier sección histórica posterior, **prevalece esta actualización**.
> El contenido anterior se conserva debajo por trazabilidad. Cuando algo todavía no fue comprobado por salida textual, test, payload o preview, marcarlo `A VALIDAR`.

## A0. Punto exacto actual

El proyecto continúa en modo de producción acelerada de las 6 TVs OCU26 para la presentación del lunes 10 de agosto de 2026.

Estado vigente:
- **TV1 — Core Comercial:** prácticamente cerrada. Pendiente menor para mañana: cambiar solo el copy `Universo comercial core` → `Shoppings · Pantallas · AA2000 · YPF`, sin recalcular 964.
- **TV2 — Core Comercial Digital:** productiva y visualmente aceptada. Pendiente Etapa 2: revisar ~68 elegibles de Shoppings Digital vs ~61 recordados por negocio.
- **TV3 — Core Comercial Estático:** productiva y comprendida funcionalmente. El usuario entendió la diferencia entre `Shoppings Estático` y `Shoppings con actividad`. Quedan ajustes/revisiones menores no bloqueantes.
- **TV4 — Pipeline Comercial:** **EN IMPLEMENTACIÓN PRODUCTIVA**. Se definió el diseño funcional, se hicieron auditorías de scope/temporalidad y se invocó `scripts/build_tv4_dashboard.py`. Falta confirmar resultado, tests específicos, sanity y preview.
- **TV5 — YPF:** será el dashboard específico de YPF.
- **TV6 — Demanda Comercial:** cerrará la secuencia y debe analizar demanda/marcas/clientes sobre **Core Comercial + YPF**, deduplicando entidades.
- **Power BI:** pausado.
- **Git/deploy:** bloqueados hasta revisión/aprobación explícita.

### Orden de trabajo vigente

**TV4 Pipeline → TV5 YPF → TV6 Demanda Comercial → revisión conjunta TV1–TV6 → Git/deploy con aprobación.**

Esta secuencia reemplaza la decisión anterior `TV5 → TV6 → TV4/YPF`.

---

## A1. Cambio material de arquitectura de TVs 4–6

### Decisión anterior reemplazada
- TV4 = YPF.
- TV5 = siguiente dashboard.
- TV6 después.
- TV4/YPF al final.

### Decisión vigente
1. **TV4 = Pipeline Comercial**
2. **TV5 = YPF**
3. **TV6 = Demanda Comercial**
   - universo conceptual: **Core Comercial + YPF**
   - foco: marcas, clientes, agencias, campañas, concentración y comportamiento de la demanda.

### Motivo
La narrativa queda:
TV1 general → TV2 Digital → TV3 Estático → TV4 movimiento operativo → TV5 YPF → TV6 demanda.

No volver a renumerar Pipeline salvo nueva instrucción explícita.

---

## A2. TV4 — concepto funcional final aprobado

### Nombre
`TV4 · PIPELINE COMERCIAL`

### Pregunta
`¿QUÉ ESTÁ CORRIENDO Y QUÉ VIENE?`

### Header
- Eyebrow: `TV 4 · OCU26 · INTELIGENCIA COMERCIAL`
- Título: `PIPELINE COMERCIAL`
- Subtítulo: `Campañas y reservas · movimiento comercial`
- Mantener logo, reloj, actualización y lenguaje visual Brand Plus de TV1–TV3.

### Referencia visual
Archivo subido: `TV4_pipeline_.html`.

La referencia usa:
- cinco KPI cards;
- panel `Estado del pipeline`;
- panel `Próximas activaciones · 30 días`;
- un único insight `Lectura`.

Decisión:
- conservar prácticamente toda la composición;
- actualizar con datos productivos;
- reemplazar el insight único por:
  - `LECTURA`
  - `PUNTO POSITIVO`
  - `A ATENDER`.

Ruta esperada si se copia al repo: `audit_sources/TV4_REFERENCE.html`.
**Estado exacto de esa copia: A VALIDAR.**

---

## A3. TV4 — no usar comparativos mensuales en KPI superiores

**Decisión confirmada:** no poner comparación Junio/Julio.

Motivo:
- `Actividad actual` = foto al corte;
- `Reservas futuras` = stock futuro al corte;
- `Inician · 30 días` y `Finalizan · 30 días` = ventanas móviles;
- `Pipeline en Core` = composición actual.

Forzar mes anterior sería menos claro y podría ser semánticamente incorrecto.

Las cards deben usar valor actual + aclaración breve, sin flechas/deltas mensuales.

---

## A4. TV4 — cinco KPI cards

### Card 1 — ACTIVIDAD ACTUAL
Formato partido:
- campañas únicas activas;
- activaciones en curso.

Regla:
`IDCampaña` puede repetirse en varias filas porque una campaña puede estar en varios elementos.
No confundir `distinct IDCampaña` con activaciones.

### Card 2 — RESERVAS FUTURAS
Cantidad real según semántica del pipeline.
Aclaración: `inicio posterior al corte` o equivalente real.

### Card 3 — INICIAN · 30 DÍAS
Ventana:
`cutoff < FechaInicio <= cutoff + 30 días`.

### Card 4 — FINALIZAN · 30 DÍAS
Ventana futura equivalente.
Aclaración permitida: `oportunidad de renovación`, sin afirmar causalidad.

### Card 5 — PIPELINE EN CORE
Solo si tiene denominador real y útil.
No mostrar un 100% tautológico por haber filtrado previamente al Core.
No usar valores legacy ni hardcodes.
Definición final: **A VALIDAR** hasta ver payload/tests.

---

## A5. TV4 — scope comercial vigente

TV4 representa Pipeline del **Core Comercial**.

Exclusiones:
- YPF;
- APSA;
- London Supply.

Circuitos explorados:
- `CENCOSUD`
- `REMEROS`
- `PANTALLAS_LED`
- `PILAR_FRONTLIGHT`
- `AA2000`

Familias visuales esperadas:
- Pantallas LED
- Shoppings Digital
- Shoppings Estático
- AA2000 / Pilar Frontlight, si corresponde según semántica real.

UI: mostrar `Shoppings Estático`, nunca `Shoppings Fijo`.
No renombrar valores técnicos internos arbitrariamente.

---

## A6. TV4 — lógica temporal auditada

### Corte usado
`2026-07-31`

### Ventana futura
30 días posteriores al corte.

### Actividad al corte
Conceptualmente:
`FechaInicio <= cutoff <= FechaFin efectiva`

Con manejo de:
- `FechaIndefinida == 'Si'`
- `FechaFin` nula
- fechas incompletas.

### Inicios
`cutoff < FechaInicio <= cutoff + 30 días`

### Finalizaciones
Para actividad en curso al corte:
`cutoff < FechaFin <= cutoff + 30 días`

### Histórico
`Finalizadas` debe mostrarse como histórico más tenue para no dominar la lectura actual.

---

## A7. TV4 — auditorías realizadas y errores evitados

Se inspeccionaron:
- `IDCampaña`;
- columnas requeridas;
- `business_semantics.json`;
- distribución de `Estado`;
- `FechaInicio`;
- `FechaFin`;
- `FechaIndefinida`;
- reservas;
- actividad a distintos cortes;
- scope `OPERATIVO_GENERAL`;
- `PERFORMANCE_CORE`;
- filas con fechas incompletas;
- campañas vacías;
- timeline de próximos inicios.

### Método privado
`engine._campanas_overlap(...)` se usó solo para auditoría.
No convertirlo automáticamente en lógica productiva si existe métrica pública/canónica.

### Hardcode exploratorio
Apareció un numerador literal `421` en una consulta de diagnóstico.
**No debe llegar al builder productivo.**

### Problema de merge
Una auditoría podía producir columnas duplicadas/sufijadas (`CircuitoNegocio_x/_y`, `Medio_x/_y`).

Solución de auditoría:
usar directamente:
`filter_universe(campanas, 'OPERATIVO_GENERAL', config)`
y luego filtrar `CircuitoNegocio`.

### Timeline
Se inspeccionaron:
- `IDCampaña`
- `Campaña`
- `FechaInicio`
- `CircuitoNegocio`
- `Medio`
- `SitioNegocio`
- `ElementoID`

Regla:
- evitar duplicados visuales por múltiples `ElementoID`;
- conservar eventos realmente distintos por fecha/circuito;
- máximo visual 5;
- ordenar por fecha ascendente y campaña.

---

## A8. TV4 — cuerpo visual

### Panel izquierdo
`ESTADO DEL PIPELINE`
Subtítulo: `activaciones cargadas`

Primer bloque:
- Activas
- Reservas
- Finalizadas

`Finalizadas` = histórico y debe ser más tenue.

Segundo bloque:
`ACTIVACIONES ACTIVAS HOY · POR GRUPO COMERCIAL`

Hasta 4 grupos reales, ordenados por actividad descendente.
No inventar `Otros` sin taxonomía canónica.

### Panel derecho
`PRÓXIMAS ACTIVACIONES · 30 DÍAS`
Subtítulo: `inician dentro de 30 días`

Máximo 5 filas:
- día;
- mes;
- campaña;
- circuito/familia;
- badge `Inicia`.

---

## A9. TV4 — insights inferiores

Reemplazar el bloque vertical único por:

### LECTURA
Resumen objetivo:
- campañas únicas activas;
- activaciones en curso;
- reservas/inicios próximos;
- contexto del pipeline.

### PUNTO POSITIVO
Prioridad:
1. nuevas activaciones próximas;
2. reservas futuras;
3. diversificación;
4. otro hecho positivo factual.

No inventar crecimiento mensual.

### A ATENDER
Prioridad:
1. muchas finalizaciones vs pocos inicios;
2. concentración en una familia;
3. falta de reservas;
4. `MetricStatus PARTIAL` material.

No inventar causas.

---

## A10. TV4 — archivos productivos y estado

Patrón esperado:
- `scripts/build_tv4_dashboard.py`
- `scripts/templates/tv4_template.html`
- `tests/test_build_tv4_dashboard.py`
- `output/tv4_data.json`
- `tv4.html`

Se invocó:
`./.venv/Scripts/python.exe scripts/build_tv4_dashboard.py`

Por precisión:
- builder invocado: **confirmado**;
- builder exitoso: **A VALIDAR**;
- payload final correcto: **A VALIDAR**;
- tests TV4: **pendientes**;
- preview TV4: **pendiente**.

---

## A11. Tests TV4 requeridos

Correr solo tests específicos TV4.

Validar mínimo:
1. payload solo TV4;
2. YPF excluido;
3. APSA excluido;
4. London excluido;
5. campañas activas distinct;
6. activaciones diferenciadas de campañas únicas;
7. reservas futuras;
8. ventana 30 días;
9. inicios;
10. finalizaciones;
11. timeline ordenada;
12. timeline máximo 5;
13. sin duplicados accidentales por `ElementoID`;
14. distribución por grupo reconcilia;
15. UI usa `Estático`;
16. `MetricStatus` preservado;
17. sin legacy `OCU_DATA`;
18. SHA Excel intacto;
19. TV1 intacta;
20. TV2 intacta;
21. TV3 intacta;
22. referencia TV4 intacta.

No correr full suite.

---

## A12. Próximo paso exacto

1. Confirmar si `scripts/build_tv4_dashboard.py` terminó sin error.
2. Si terminó bien, **no reauditar TV4**.
3. Ejecutar tests específicos TV4.
4. Sanity mínimo:
   - SHA Excel;
   - exclusiones YPF/APSA/London;
   - payload solo TV4;
   - protegidos intactos.
5. Preview `tv4.html` en 1920×1080.
6. Revisar 5 cards, nombres `Estático`, barras, timeline, insights y overflow.
7. Cerrar TV4.
8. Nueva sesión Claude Code para TV5 YPF.
9. Luego TV6 Demanda.
10. Revisión conjunta.
11. Solo después Git/deploy con aprobación.

---

## A13. TV5 — YPF

Número definitivo: `TV5`.

Datos/semántica YPF ya conocidos:
- YPF se excluye del Core estándar cuando corresponde.
- Surrogate vigente para estaciones en TV1:
  `prefijo numérico de ElementoID + localidad normalizada de Ubicacion`.
- Catálogo usado en TV1: 451 estaciones.
- Julio: 305 estaciones con campaña.
- Junio: 263.
- Serie Ene–Jul: `0,0,0,0,233,263,305`.
- Actividad observada actual: digital.
- No llamar APIE real al surrogate.
- Etapa 2: incorporar ID real de estación/APIE y retirar surrogate.

Diseño TV5: todavía no cerrado en este corte.

---

## A14. TV6 — Demanda Comercial

Número definitivo: `TV6`.

Universo conceptual:
**Core Comercial + YPF**

Foco:
- marcas;
- clientes;
- agencias;
- campañas;
- concentración;
- participación;
- comportamiento de demanda.

Regla crítica:
no sumar aritméticamente clientes/marcas/campañas Core + YPF si pueden repetirse.
Combinar universo y luego aplicar `distinct` sobre claves canónicas.

Exclusiones globales APSA/London se mantienen salvo cambio explícito.

KPIs definitivos: **A VALIDAR** al implementar TV6.

---

## A15. Etapa 2 — próxima semana

Objetivo:
- cálculos ampliados;
- reconciliaciones;
- microauditorías;
- revisión de universos/denominadores;
- consistencia entre dashboards;
- identificadores y surrogates.

Regla de entrega inmediata:
si una duda no bloquea la presentación, documentarla y seguir.

Backlog conocido:
- TV2: ~68 elegibles Shoppings Digital vs ~61 recordados.
- TV3: clasificación `Otro` demasiado amplia; revisar evolución/labels; ajuste menor `Disponibles`.
- YPF: reemplazar surrogate por APIE/ID real.
- TV4: inconsistencias temporales no bloqueantes si aparecen en tests/preview.
- Reconciliación transversal TV1–TV6.

---

## A16. Eficiencia Claude Code

Para TV4–TV6:
- sesión nueva por dashboard cuando convenga;
- prompt maestro;
- lectura focalizada CM1;
- una auditoría consolidada;
- implementación;
- tests específicos;
- sanity;
- preview;
- cierre.

Evitar:
- microauditorías repetitivas;
- releer todo el repo;
- full suite;
- reexplorar Python;
- scripts temporales innecesarios;
- commit/push/deploy antes de aprobación.

Entorno confirmado:
`./.venv/Scripts/python.exe`

---

## A17. Protecciones

No tocar sin bloqueo real:
- `scripts/validate_input.py`
- `scripts/transform_data.py`
- `scripts/semantic_model.py`
- `scripts/metrics_engine.py`
- `scripts/export_data.py`
- `config/business_semantics.json`
- `input/OCU26_BASE_DATOS.xlsx`
- `powerbi/`
- `audit_sources/*`
- archivos cerrados TV1/TV2/TV3

No `git add .`, commit, push ni deploy sin aprobación.

---

## A18. A VALIDAR actuales

1. TV4: resultado del builder.
2. TV4: valores finales de las 5 cards.
3. TV4: definición final `Pipeline en Core`.
4. TV4: tests específicos.
5. TV4: preview 1920×1080.
6. TV4: ruta exacta de referencia en repo.
7. TV2: 68 vs ~61 elegibles Shoppings Digital.
8. TV3: clasificación `Otro`.
9. TV3: ajuste visual/semántico menor.
10. TV1: cambio de copy Core.
11. TV5: diseño final YPF.
12. TV6: KPIs/visuales definitivos.
13. Netlify: integración A/B/C si sigue sin resolver.
14. Power BI: error `dataType` pausado.

---

## A19. Snapshot ejecutivo actualizado

1. Arquitectura centralizada: Excel → Gates → semántica/MetricsEngine → outputs → builders → HTML.
2. HTML sin lógica de negocio.
3. TV1 prácticamente cerrada.
4. TV2 cerrada para entrega.
5. TV3 productiva.
6. TV4 = Pipeline Comercial.
7. TV5 = YPF.
8. TV6 = Demanda Comercial.
9. La vieja decisión TV4=YPF queda reemplazada.
10. TV4 no compara cards contra junio.
11. TV4 = estado actual + 30 días.
12. TV4 excluye YPF/APSA/London.
13. Core TV4 explorado: CENCOSUD + REMEROS + PANTALLAS_LED + PILAR_FRONTLIGHT + AA2000.
14. Card 1 distingue campañas únicas vs activaciones.
15. Card 2 reservas futuras.
16. Card 3 inicios 30 días.
17. Card 4 finalizaciones 30 días.
18. Card 5 requiere denominador real.
19. Corte auditado: 2026-07-31.
20. Timeline máximo 5 y deduplicada.
21. UI usa `Estático`, no `Fijo`.
22. `_campanas_overlap` fue exploratorio.
23. Hardcode 421 no debe ser productivo.
24. Se evitó merge ambiguo usando `OPERATIVO_GENERAL` sobre campanas.
25. Builder TV4 invocado.
26. Tests/sanity/preview TV4 son el próximo paso.
27. TV5 conserva semántica YPF conocida hasta Etapa 2.
28. TV6 combina Core + YPF para demanda.
29. TV6 deduplica entidades únicas.
30. Etapa 2 próxima semana = microauditorías/cálculos ampliados.
31. TV2 68 vs ~61 queda Etapa 2.
32. TV3 `Otro` queda Etapa 2.
33. No commit/push/deploy.
34. Power BI pausado.
35. Próximo paso: tests TV4 → sanity → preview → TV5 YPF.


# ACTUALIZACIÓN VIGENTE CM1 — 9/8/2026 ~23:27 ART

> **ESTA ES LA CAPA DE VIGENCIA MÁS RECIENTE DEL DOCUMENTO.**
> Si contradice la actualización de ~22:46 ART o cualquier sección histórica posterior, **prevalece esta actualización**.
> El contenido anterior se conserva debajo por trazabilidad y para reconstruir la evolución del proyecto.

## A0. Punto exacto actual y prioridad de esta noche

El proyecto continúa en modo de producción acelerada de las 6 TVs para la presentación del lunes 10 de agosto de 2026.

Estado operativo vigente:

- **TV1:** prácticamente cerrada. No reabrir lógica. Sigue pendiente para mañana el cambio de copy en `CORE COMERCIAL`: `Universo comercial core` → `Shoppings · Pantallas · AA2000 · YPF`, sin recalcular el valor 964.
- **TV2:** productiva y visualmente cerrada para esta noche. No auditar de nuevo ahora. Queda pendiente para mañana/Etapa 2 revisar el denominador elegible de Shoppings Digital: dashboard actual ~68 vs ~61 recordado por negocio.
- **TV3:** `CORE COMERCIAL ESTÁTICO` ya fue construida y revisada visualmente a nivel funcional. La estructura y lectura general fueron aceptadas por el usuario. Quedan ajustes visuales/semánticos menores no bloqueantes documentados abajo.
- **TV4:** será **YPF**. Por decisión explícita del usuario, **se posterga para el final** porque conviene resolverla después de TV5/TV6.
- **TV5:** es el **próximo bloque inmediato**.
- **TV6:** debería seguir después de TV5, salvo cambio explícito.
- Orden operativo actual: **TV5 → TV6 → TV4/YPF**, con revisión conjunta final de las seis TVs.
- **Power BI:** continúa pausado.
- **Git/deploy:** siguen bloqueados hasta revisión/aprobación explícita. No commit, push ni deploy durante la producción de las TVs.

### Criterio de velocidad vigente

Para completar la entrega:

1. construir cada TV con un prompt maestro;
2. una sola auditoría consolidada cuando haga falta;
3. evitar microauditorías no bloqueantes;
4. builder → tests específicos → sanity mínimo → preview;
5. aceptar visualmente y congelar alcance;
6. registrar dudas menores para Etapa 2 en vez de detener la producción.

---

## A1. Decisión nueva — ETAPA 2 de auditoría profunda la próxima semana

**Decisión vigente y confirmada por el usuario:** la próxima semana se realizará una **Etapa 2** dedicada a cálculos más extensos, reconciliaciones y microauditorías específicas.

Objetivo de Etapa 2:

- recalcular universos y denominadores con mayor profundidad;
- revisar clasificaciones funcionales y posibles duplicaciones;
- reconciliar métricas entre TVs y capa semántica;
- ampliar cálculos temporales y de inventario;
- revisar casos especiales que hoy no bloquean la presentación;
- documentar/corregir diferencias entre expectativa de negocio y cálculo actual;
- dejar el sistema más robusto después de la entrega inmediata.

**Regla operativa actual:** durante la construcción de las TVs, una duda no crítica se documenta y se continúa. No gastar créditos/tiempo en microauditorías repetidas salvo que exista un error real que pueda falsear la presentación.

Backlog inicial para Etapa 2 ya identificado:

- TV2: Shoppings Digital elegibles ~68 actuales vs ~61 recordados por negocio.
- TV2: reconciliar clasificación Pantallas/Shoppings y Remeros si sigue siendo necesario después de revisar el denominador.
- TV3: revisar clasificación de soportes, especialmente la concentración de `Otro`.
- TV3: revisar semántica/copy del gráfico histórico si se decide expresar cantidad de elementos vs porcentaje de ocupación.
- cualquier nueva inconsistencia no bloqueante detectada durante TV5/TV6/TV4 debe agregarse a este backlog y no resolverse por microauditoría inmediata.

---

## A2. TV3 — CORE COMERCIAL ESTÁTICO — implementación productiva actual

### A2.1 Alcance real resuelto por el pipeline

Durante la auditoría de TV3 se verificó el universo estático Core con filtros del pipeline real.

Scope operativo usado para TV3:

- `CENCOSUD`
- `REMEROS`
- `AA2000`
- `PILAR_FRONTLIGHT`

con `Medio = Estático` y dentro del Core operativo correspondiente.

Exclusiones mantenidas:

- APSA
- London Supply
- YPF

Título visible actual:

`Core Comercial Estático`

Subtítulo visible actual:

`Rendimiento Shoppings Estático + AA2000 Estático + Pilar Frontlight`

Esta resolución reemplaza la etapa anterior en la que el scope TV3 figuraba como `A VALIDAR`.

### A2.2 Arquitectura y archivos TV3

TV3 sigue el patrón productivo ya establecido:

pipeline real → builder TV3 → `output/tv3_data.json` → template → `tv3.html`

Archivos esperados/creados:

- `scripts/build_tv3_dashboard.py`
- `scripts/templates/tv3_template.html`
- `tests/test_build_tv3_dashboard.py`
- `output/tv3_data.json`
- `tv3.html`
- referencia read-only: `audit_sources/TV3_REFERENCE.html`

El builder se ejecutó con:

`./.venv/Scripts/python.exe scripts/build_tv3_dashboard.py`

La suite específica fue lanzada con:

`./.venv/Scripts/python.exe -m pytest tests/test_build_tv3_dashboard.py -q`

**Resultado textual final de pytest no quedó visible en el chat de ChatGPT: `A VALIDAR`.**
No inventar cantidad de tests aprobados hasta leer el resultado real si fuese necesario.

También se ejecutó un `git status --porcelain` de solo lectura, pero su salida concreta no quedó documentada aquí: **A VALIDAR** antes del commit futuro.

### A2.3 KPI cards TV3 — valores actuales de julio 2026

#### KPI 1 — OCUPACIÓN POR CALENDARIO

- activos: **119**
- `% con campaña`: **28,2%**
- universo elegible implícito: **422 elementos**
- junio: **28,9%**
- delta: **-0,7 pp**
- semántica: rojo porque cae la ocupación.

Lectura aprobada:

`119 de 422 elementos del Core Comercial Estático tuvieron campaña en julio.`

#### KPI 2 — SHOPPINGS ESTÁTICO

- activos: **119**
- `% con campaña`: **31,2%**
- universo elegible: aproximadamente **381 elementos** de Shoppings Estático
- junio: **32,0%**
- delta: **-0,8 pp**

Lectura aprobada:

`Shoppings Estático` cuenta **elementos físicos/soportes**, no sedes.

El hecho de que esta tarjeta muestre también 119 implica que la actividad estática de julio está concentrada en Shoppings dentro del scope actual; AA2000 figura sin actividad y Pilar Frontlight no aporta actividad material visible en esta lectura.

#### KPI 3 — AA2000 ESTÁTICO

- activos: **0**
- `% con campaña`: **0,0%**
- elegibles: **40**
- junio: **0,0%**
- delta: **0,0 pp**

Interpretación confirmada:

AA2000 Estático sí tiene inventario elegible, pero no registra elementos con campaña en junio/julio según la métrica actual.

No convertir esta lectura en “AA2000 no existe”; significa `0 de 40 elegibles con campaña`.

#### KPI 4 — SHOPPINGS CON ACTIVIDAD

- shoppings/sitios activos: **7**
- `% de cobertura`: **41,2%**
- universo implícito: **17 shoppings/sitios**
- junio: **47,1%**
- delta: **-5,9 pp**

Distinción conceptual ya explicada al usuario y comprendida:

- **Shoppings Estático** = cuenta elementos/soportes físicos con campaña.
- **Shoppings con actividad** = cuenta sedes/shoppings que tienen al menos un elemento estático activo.

Ejemplo conceptual: si un shopping tiene 30 soportes y 23 están ocupados, en `Shoppings Estático` aporta 23 elementos activos; en `Shoppings con actividad` aporta 1 sitio activo.

El usuario indicó que esta diferencia quedó comprendida. No volver a explicarla salvo que sea necesario.

Posible mejora futura de copy: `SHOPPINGS ACTIVOS` podría ser más inmediato que `SHOPPINGS CON ACTIVIDAD`, pero **no está confirmada como cambio todavía**.

#### KPI 5 — DISPONIBLES

- disponibles: **303**
- `% del inventario`: **71,8%**
- cálculo: **422 - 119 = 303**
- junio: **71,1%**
- delta: **+0,7 pp**
- semántica: rojo porque más disponibilidad representa más inventario sin campaña.

Relación complementaria:

`28,2% ocupado + 71,8% disponible = 100%`.

### A2.4 Ranking de ocupación por shopping · julio

Top 5 visible actual:

1. **Palmas Del Pilar** — 23 ocupados · 7 disponibles — **76,7%**
2. **Remeros** — 19 ocupados · 8 disponibles — **70,4%**
3. **Factory Quilmes** — 3 ocupados · 2 disponibles — **60,0%**
4. **Unicenter** — 31 ocupados · 31 disponibles — **50,0%**
5. **Plaza Oeste** — 19 ocupados · 30 disponibles — **38,8%**

El diseño de barras/ranking fue aceptado visualmente.

### A2.5 Soportes más vendidos · julio

Panel actual:

`Soportes más vendidos · julio`

Subtítulo:

`Elementos con campaña por soporte`

Valores visibles actuales:

- `Otro` — **109**
- `Frontlight` — **10**

Aunque el diseño original pedía Top 3, en la preview solo aparecen estas dos categorías con dato. **No inventar una tercera categoría.**

Pendiente Etapa 2:

- revisar si `Otro = 109` es una clasificación demasiado amplia;
- determinar si existe información de soporte más granular que convenga normalizar o exponer.

Esto no bloquea la presentación actual.

### A2.6 Evolución mensual · Core Estático

Panel actual:

`Evolución mensual · Core Estático`

Subtítulo visible:

`Ocupación por calendario`

La serie visible está expresada como **cantidad de elementos con campaña**, no como porcentaje:

- Ene: **112**
- Feb: **110**
- Mar: **115**
- Abr: **116**
- May: **116**
- Jun: **122**
- Jul: **119**

Solo julio muestra label numérico destacado, siguiendo el criterio visual establecido.

**Pendiente no bloqueante / A VALIDAR:** el subtítulo `Ocupación por calendario` puede interpretarse como porcentaje, mientras el eje muestra cantidades. En una pasada posterior se puede:

- cambiar el subtítulo a `Elementos con campaña`, o
- convertir la serie a porcentaje si negocio lo prefiere.

No cambiar esta semántica durante producción rápida sin decisión explícita.

### A2.7 Insights inferiores TV3

TV3 ya replica el patrón de TV1/TV2:

- `LECTURA`
- `PUNTO POSITIVO`
- `A ATENDER`

Texto visible actual:

**LECTURA**

`Julio registra 119 de 422 elementos del Core Estático con campaña (28,2%). Shoppings Estático concentra toda la actividad del mes y quedan 303 elementos disponibles.`

**PUNTO POSITIVO**

`Palmas del Pilar lidera la ocupación de Shoppings Estático con 76,7% en julio, seguido por Remeros (70,4%).`

**A ATENDER**

`AA2000 Estático no registra elementos con campaña en julio (0 de 40 elegibles), igual que en junio.`

Estos textos son factuales; no incorporan causas inventadas.

### A2.8 Estado visual TV3 y pendiente menor señalado por el usuario

La preview 1920×1080 fue revisada por el usuario y la estructura general fue aceptada.

El usuario señaló una captura recortada de la tarjeta `DISPONIBLES` y dijo que **“luego cambiamos esto”** antes de pasar a TV5.

Lo observable en la captura es un **halo/ornamento claro superpuesto sobre el borde superior de la tarjeta `DISPONIBLES`**.

Estado:

- cambio visual pendiente;
- no bloquea continuar con TV5;
- **A VALIDAR** la corrección exacta antes de editar, pero la intención es limpiar ese artefacto/ornamento y mantener la tarjeta prolija.

No reabrir métricas de TV3 por este ajuste visual.

---

## A3. Decisión de secuencia — TV4 YPF se deja para el final

Decisión explícita del usuario al cierre de TV3:

- **TV4 será YPF**.
- No se implementará ahora.
- Se deja **para el final**.
- El próximo dashboard a trabajar es **TV5**.

Orden operativo recomendado bajo esta decisión:

1. TV5
2. TV6
3. TV4 / YPF
4. revisión conjunta TV1–TV6
5. ajustes críticos pendientes
6. Git/deploy solo con aprobación

Esta decisión reemplaza cualquier roadmap anterior que indicara TV3 → TV4 → TV5 → TV6 de forma estrictamente numérica.

---

## A4. Próximo paso exacto después de esta actualización

1. Guardar esta versión actualizada como `docs/CM1.md`.
2. Abrir una **nueva sesión de Claude Code para TV5** para ahorrar contexto/tokens.
3. Usar CM1 actualizado como fuente de verdad.
4. Cargar/usar la referencia visual de TV5 correspondiente como read-only.
5. Preparar un prompt maestro TV5 compacto pero completo.
6. Una auditoría consolidada única si TV5 necesita métricas nuevas.
7. Builder → tests específicos TV5 → sanity → preview.
8. No volver a TV2/TV3 salvo error crítico esta noche.
9. TV4/YPF queda expresamente para el final.

---

## A5. Snapshot operativo de esta actualización

1. TV1: cerrada funcionalmente; copy Core pendiente mañana.
2. TV2: productiva; 37 tests documentados como pass; denominador Shoppings Digital pendiente Etapa 2.
3. TV3: construida y preview funcional aceptada.
4. TV3 Core Estático usa CENCOSUD + REMEROS + AA2000 + PILAR_FRONTLIGHT.
5. TV3 Core Estático julio: 119/422 = 28,2% con campaña.
6. TV3 Shoppings Estático: 119 y 31,2%.
7. TV3 AA2000 Estático: 0/40 = 0,0%.
8. TV3 Shoppings con actividad: 7/17 = 41,2%.
9. TV3 disponibles: 303 = 71,8%.
10. Ranking TV3 líder: Palmas del Pilar 76,7%.
11. Remeros segundo: 70,4%.
12. Soportes visibles TV3: Otro 109; Frontlight 10.
13. Evolución de elementos con campaña TV3 Ene–Jul: 112,110,115,116,116,122,119.
14. Resultado textual de pytest TV3: A VALIDAR porque no quedó capturado en ChatGPT.
15. Artefacto/halo sobre tarjeta Disponibles TV3: pendiente visual menor.
16. Etapa 2 próxima semana: microauditorías y cálculos extendidos.
17. TV4 = YPF y queda para el final.
18. Próximo trabajo = TV5.
19. Después TV6 y luego TV4/YPF, salvo nueva decisión.
20. No commit/push/deploy todavía.

---

# ACTUALIZACIÓN VIGENTE CM1 — 9/8/2026 ~22:46 ART

> **ESTA ES LA CAPA DE VIGENCIA MÁS RECIENTE DEL DOCUMENTO.**
> Si contradice la actualización de ~21:08 ART, las secciones históricas o el snapshot anterior, **prevalece esta actualización**.
> El documento completo anterior se conserva debajo por trazabilidad.

## A. Punto exacto del proyecto al cierre de esta actualización

El proyecto está en producción acelerada de las 6 TVs para presentación del lunes 10 de agosto de 2026.

Estado operativo inmediato:

- **TV1:** prácticamente cerrada; no reabrir lógica ni auditorías. Queda para mañana un único cambio de copy en la tarjeta `CORE COMERCIAL`: `Universo comercial core` → `Shoppings · Pantallas · AA2000 · YPF`. El valor 964 no cambia.
- **TV2:** implementada, tests específicos aprobados y última preview visual aceptada para esta noche. No seguir auditando ni retocando salvo los pendientes expresamente documentados para mañana.
- **TV3:** siguiente trabajo. Debe iniciarse en **una sesión nueva de Claude Code** usando este CM1 actualizado y el HTML de referencia de TV3.
- **TV4–TV6:** todavía por implementar/revisar después de TV3.
- **Power BI:** pausado hasta después de la entrega de las TVs.
- **Git/deploy:** no hacer commit, push ni deploy durante la construcción de las TVs sin aprobación explícita. El resultado textual del último `git status` posterior a TV2 no quedó documentado en ChatGPT: **A VALIDAR** antes del commit futuro.

### Estrategia de tiempo vigente

Para llegar a la entrega:

1. congelar alcance de cada TV una vez visualmente aceptada;
2. prompt maestro único por TV;
3. una sola auditoría consolidada cuando haga falta dato nuevo;
4. builder → tests específicos → preview;
5. evitar full suite y micro-auditorías repetidas;
6. dejar refinamientos no críticos para la mañana previa a la presentación.

La sesión nueva de Claude Code es preferible para TV3 porque TV2 ya cerró su bloque y no conviene arrastrar su razonamiento, exploraciones y tokens.

---

## B. Entorno Python confirmado durante TV2

El entorno correcto del repositorio para ejecutar scripts/tests es:

`./.venv/Scripts/python.exe`

Se confirmó que ese entorno contiene `pandas` y pudo ejecutar el pipeline real, el builder y pytest.

Regla operativa:

- no volver a explorar múltiples intérpretes de Python salvo error real;
- usar directamente `./.venv/Scripts/python.exe` desde `C:\brand plus\ocu26-dashboard`.

---

## C. TV2 — CORE COMERCIAL DIGITAL — estado productivo actual

### C.1 Archivos productivos TV2

Archivos creados/usados durante la implementación:

- `scripts/build_tv2_dashboard.py`
- `scripts/templates/tv2_template.html`
- `tests/test_build_tv2_dashboard.py`
- `output/tv2_data.json`
- `tv2.html`
- referencia read-only esperada: `audit_sources/TV2_REFERENCE.html`

El builder se ejecutó con:

`./.venv/Scripts/python.exe scripts/build_tv2_dashboard.py`

Los tests específicos se ejecutaron con:

`./.venv/Scripts/python.exe -m pytest tests/test_build_tv2_dashboard.py -q`

**Resultado final documentado:** `37 tests pass`.

No volver a correr esos tests esta noche salvo que TV2 vuelva a modificarse por una razón real.

### C.2 Arquitectura TV2 confirmada

TV2 sigue el patrón productivo:

pipeline real → builder TV2 → `output/tv2_data.json` → template → `tv2.html`

El HTML no debe ser la fuente de verdad de negocio.
Los valores legacy de `window.OCU_DATA` presentes en la referencia vieja son solo referencia visual y **no deben reutilizarse como datos**.

### C.3 Scope vigente de TV2

Nombre visible:

`CORE COMERCIAL DIGITAL`

Subtítulo visible final:

`Rendimiento Pantallas LED + Shoppings Digital + AA2000`

Scope conceptual vigente:

- Pantallas LED
- Shoppings Digital
- AA2000 Digital

Exclusiones:

- YPF fuera de TV2
- APSA fuera
- London Supply fuera

AA2000:

- sí participa en los KPIs generales del Core Digital;
- sí participa en capacidad/slots/disponibilidad generales;
- no tiene ranking propio;
- no tiene tercera línea en el gráfico principal de evolución;
- aparece en el desglose compacto de disponibilidad cuando la métrica existe.

### C.4 Última preview TV2 aceptada para esta noche

La última preview generada muestra la siguiente estructura:

**Header**
- logo Brand Plus
- `TV 2 · OCU26 · INTELIGENCIA COMERCIAL`
- `Core Comercial Digital`
- subtítulo de universo
- reloj
- actualización
- `Periodo: Julio 2026`

**Fila KPI: 5 tarjetas**

#### KPI 1 — OCUPACIÓN POR CALENDARIO

- principal: **71**
- label principal: `ACTIVOS`
- secundario: **78,0%**
- label secundario: `% CON CAMPAÑA`
- junio: **79,1%**
- delta: **-1,1 pp**
- color del delta: rojo

Interpretación acordada con el usuario:

- 71 = cantidad de elementos digitales Core que tuvieron al menos una campaña en julio;
- 78,0% = 71 sobre el universo elegible de 91 elementos según la lógica actual del payload;
- por lo tanto, la lectura actual es `71 de 91` elementos Core Digital ocupados por calendario.

#### KPI 2 — FILL RATE DIGITAL

- principal: **321**
- label principal: `SLOTS OCUPADOS`
- secundario: **20,4%**
- label secundario: `% CAPACIDAD VENDIDA`
- junio: **27,3%**
- delta: **-6,9 pp**
- rojo

Interpretación:

- 321 = slots ocupados/vendidos en julio;
- 20,4% = esos slots sobre la capacidad total elegible del Core Digital.

#### KPI 3 — PANTALLAS LED

Esta tarjeta fue cambiada desde slots/fill a **ocupación calendario**, por decisión del usuario.

- principal: **11**
- label: `ACTIVOS`
- secundario: **100,0%**
- label: `% CON CAMPAÑA`
- junio: **100,0%**
- delta: **0,0 pp**
- neutro/gris

Interpretación:

- 11 Pantallas LED con campaña;
- 11 de 11 elegibles activas por calendario.

#### KPI 4 — SHOPPINGS DIGITAL

También fue cambiada desde slots/fill a **ocupación calendario**.

- principal: **60**
- label: `ACTIVOS`
- secundario: **88,2%**
- label: `% CON CAMPAÑA`
- junio: **89,7%**
- delta: **-1,5 pp**
- rojo

Interpretación actual del dashboard:

- 60 elementos de Shoppings Digital con campaña;
- 88,2% implica aproximadamente 68 elementos elegibles según la lógica actual.

**ESTE DENOMINADOR QUEDA PENDIENTE DE REVISIÓN MAÑANA.** Ver sección D.

#### KPI 5 — SLOTS DISPONIBLES

- principal: **1.254**
- label: `DISPONIBLES`
- secundario: **79,6%**
- label: `% DE CAPACIDAD`
- junio: **72,7%**
- delta: **+6,9 pp**
- se muestra rojo porque más capacidad disponible = peor utilización comercial

Desglose visible final:

- Pantallas: **61,2%** disponible
- Shoppings: **79,6%** disponible
- AA2000: **100,0%** disponible

Regla semántica importante:

- para ocupación/utilización, subir = verde;
- para disponibilidad, **bajar = verde** y subir = rojo.

La tarjeta Slots Disponibles fue ensanchada visualmente y las primeras cuatro se compactaron para mejorar uso del espacio.

### C.5 Rankings TV2

#### Pantallas LED · ranking de julio

La última preview muestra Top 5:

1. Cabildo — 78,1% fill · 100,0% ocupación calendario
2. Pilar — 68,2% fill · 100,0% ocupación
3. Cerrito — 47,4% fill · 100,0% ocupación
4. Remeros — 47,4% fill · 100,0% ocupación
5. Avellaneda — 45,3% fill · 100,0% ocupación

**A VALIDAR MAÑANA:** esta preview muestra `Remeros` dentro del ranking Pantallas LED, mientras una decisión anterior del diseño conceptual establecía que Remeros Digital debía considerarse dentro de Shoppings Digital y no duplicarse. No asumir todavía que la preview es incorrecta: revisar junto con el pendiente del denominador de Shoppings para resolver la clasificación real desde el pipeline.

#### Shoppings Digital · ranking de julio

Top 3 visible:

1. Unicenter — 30,0% fill · 97,1% ocupación calendario
2. Palmas del Pilar — 17,5% fill · 100,0% ocupación
3. Portal Escobar — 15,0% fill · 100,0% ocupación

No agregar AA2000 al ranking.

### C.6 Evolución mensual TV2

Panel visible:

`Evolución mensual · Core Digital`

Subtítulo:

`Fill rate · Pantallas LED vs Shoppings Digital`

Series:

- Pantallas LED (azul)
- Shoppings Digital (cream/blanco)

Meses:

- Ene–Jul únicamente
- sin Ago–Dic
- sin proyecciones

AA2000 queda incluido en KPIs generales pero no se agrega como tercera línea.

Ajuste visual final solicitado:

- mostrar valores numéricos **solo en julio**, el mes vigente;
- no mostrar labels sobre Ene–Jun.

La última preview muestra los labels del punto julio al extremo derecho.

### C.7 Insights inferiores TV2

Se reemplazó el insight vertical legacy por el patrón de TV1:

- `LECTURA`
- `PUNTO POSITIVO`
- `A ATENDER`

Texto/factualidad de la última preview:

**LECTURA**

`Julio registra 71 de 91 elementos del Core Digital con campaña (78,0%). El fill rate alcanza 20,4% y quedan 1.254 slots disponibles.`

**PUNTO POSITIVO**

`Shoppings Digital concentra la mayor parte de la actividad del Core (60 elementos activos), sosteniendo el negocio del mes.`

**A ATENDER**

`El fill rate de Pantallas LED cae 17,7 pp frente a junio.`

No inventar causas para estas variaciones.

### C.8 Ajustes visuales finales TV2

Durante el cierre de TV2 se hicieron estos cambios:

1. Pantallas LED y Shoppings Digital dejaron de mostrar slots/fill como KPI principal y pasaron a ocupación calendario.
2. Slots Disponibles mantuvo información de capacidad y el desglose por Pantallas/Shoppings/AA2000.
3. Las primeras cuatro tarjetas se hicieron algo más compactas.
4. Slots Disponibles se hizo más ancha para acomodar mejor su desglose.
5. Se aumentó moderadamente la jerarquía de los datos dentro de las KPI cards.
6. En la evolución mensual se dejaron labels numéricos solo para julio.
7. Se eliminó por completo un elemento/subtítulo residual del header que aparecía detrás de la fila de KPI cards, entre Shoppings y Slots.
8. Rankings e insights se mantuvieron sin rediseño posterior.

### C.9 Nota de troubleshooting del template TV2

Cuando se abrió directamente:

`scripts/templates/tv2_template.html`

la preview mostró `S/D` y `Sin datos de ranking`.

Eso **no era un bug de datos**: el template directo no contiene el payload productivo final. Después de regenerar `tv2.html` con el builder reaparecieron los datos reales.

Regla futura:

- revisar la salida final en `tv2.html`, no usar el template directo como evidencia de que faltan datos.

### C.10 UTF-8 TV2

Se verificó el archivo `output/tv2_data.json` en bytes buscando `campa\xc3\xb1a` para confirmar que `campaña` estaba realmente escrito en UTF-8.

Resultado esperado/confirmado en el flujo: no se trataba de un bug real de encoding.

---

## D. Pendiente explícito para mañana — Shoppings Digital TV2

El usuario recuerda que el universo elegible de Shoppings Digital debería ser aproximadamente **61**, pero la tarjeta final muestra:

- activos: 60
- ocupación: 88,2%

lo que implica un denominador de aproximadamente **68**.

Esto debe revisarse mañana, **no esta noche**.

### Qué comprobar mañana

Hacer una sola revisión focalizada para determinar:

1. cuál es el universo canónico de elementos elegibles de Shoppings Digital;
2. si 68 proviene de una regla de scope correcta o de una clasificación no deseada;
3. cómo se está clasificando Remeros Digital;
4. si Remeros aparece o no duplicado/trasladado entre Pantallas y Shoppings;
5. si existen elementos no comerciales/no elegibles incluidos en el denominador;
6. reconciliación `activos / elegibles = ocupación calendario`;
7. impacto sobre KPI Shoppings, KPI Core y ranking.

No modificar el dashboard antes de tener esa reconciliación.

**Estado:** `A VALIDAR`.

### Relación con tests

Los 37 tests específicos de TV2 pasaron.
Eso prueba consistencia con las expectativas codificadas, pero **no elimina la necesidad de validar el criterio de negocio del denominador** si la expectativa del test replica la misma clasificación actual.

---

## E. Semántica acordada para leer las cards TV2 y reutilizar patrón

El usuario pidió una explicación explícita de cómo leer `entero + porcentaje`.
Esta semántica debe reutilizarse en TV3 cuando corresponda:

- el **entero grande** representa cantidad absoluta;
- el **porcentaje secundario** explica qué proporción representa ese entero dentro del universo elegible/capacidad correspondiente;
- la línea inferior compara contra el mes anterior.

En TV2:

- tarjetas 1, 3 y 4 cuentan **elementos**;
- tarjetas 2 y 5 cuentan **slots**;
- no mezclar ambos granos en la misma interpretación.

Ejemplo calendario:

`71 | 78,0%` = 71 elementos con campaña, equivalentes al 78,0% del universo elegible.

Ejemplo fill:

`321 | 20,4%` = 321 slots ocupados, equivalentes al 20,4% de la capacidad.

Ejemplo disponibilidad:

`1.254 | 79,6%` = 1.254 slots no vendidos, equivalentes al 79,6% de la capacidad.

---

## F. TV3 — referencia recibida y estado

Archivo subido por el usuario:

`TV3_fijo_core_.html`

Para el repo debe utilizarse como referencia read-only, idealmente:

`audit_sources/TV3_REFERENCE.html`

La referencia legacy usa el título `Fijo Core` y estructura 1920×1080 con branding Brand Plus/Poppins.

Estructura visual legacy observada:

- header con `TV 3 · FIJO CORE · OOH`;
- título `Así rinde el fijo core`;
- subtítulo `Shoppings Fijo + soportes clave · ocupación calendario`;
- 5 KPI cards;
- panel izquierdo grande `Ocupación por shopping`;
- panel derecho `Soportes más vendidos`;
- un único insight inferior vertical `Lectura`.

El HTML legacy contiene `window.OCU_DATA` con valores viejos y hasta meses futuros. **No usar esos números como fuente de verdad.**

Valores legacy visibles en la referencia, solo para reconocerla, no para producción:

- elementos fijos core: 588
- ocupación fija julio: 9,6%
- ocupación fija 2026: 9,2%
- shopping top: P.Pilar
- disponibles 30 días: 533

Esos valores deben recalcularse desde el pipeline real.

---

## G. TV3 — decisión vigente del usuario

### G.1 Nomenclatura

Cambiar **toda nomenclatura visible** de `Fijo` a `Estático`.

Ejemplos:

- `Fijo Core` → `Core Comercial Estático`
- `ocupación fija` → `ocupación estática`
- `Shoppings Fijo` → `Shoppings Estático`
- `elementos fijos` → `elementos estáticos`

No renombrar valores internos del pipeline solo por copy si internamente `Fijo` sigue siendo categoría válida.

### G.2 Header

TV3 debe replicar la lógica de producto de TV2.

Eyebrow deseado:

`TV 3 · OCU26 · INTELIGENCIA COMERCIAL`

Título:

`CORE COMERCIAL ESTÁTICO`

Subtítulo:

debe enumerar el **universo real que compone el Core Estático**, siguiendo el patrón de TV2.

No asumir el universo antes del audit.
Debe resolverse desde las reglas/pipeline.

Ejemplo de forma, no dato confirmado:

`Rendimiento Shoppings Estático + AA2000 Estático + [...]`

No incluir Cencomedia/MAB/APSA/London automáticamente solo porque aparezcan en el catálogo ampliado o en estático general.

Pregunta guía:

`¿CÓMO RINDE EL CORE ESTÁTICO Y DÓNDE QUEDA DISPONIBILIDAD?`

### G.3 KPI cards TV3

El usuario quiere copiar el patrón visual/semántico de TV2:

**entero grande + porcentaje secundario + comparación contra junio**.

Reglas de color:

- aumento de ocupación = verde;
- disminución de ocupación = rojo;
- sin cambio = gris;
- para disponibilidad, menos disponibilidad puede significar mejor utilización y debe seguir semántica comercial consistente.

No usar cards que solo muestren un porcentaje grande sin el entero asociado si existe denominador interpretable.

### G.4 Gráficos TV3

El usuario aprueba la estética general de la referencia.

Mantener:

- ranking/ocupación por shopping a la izquierda;
- barras horizontales de soportes a la derecha.

Cambios:

- el gráfico de soportes debe ser **Top 3**, no 6;
- debajo del Top 3 de soportes agregar un **histórico lineal de ocupación de Estático**;
- evolución mensual Ene–Jul;
- sin meses futuros;
- mostrar label numérico solo en julio si sigue el patrón visual de TV2.

### G.5 Insights inferiores TV3

Reemplazar el insight único vertical de la referencia por el patrón horizontal de TV1/TV2:

- `LECTURA`
- `PUNTO POSITIVO`
- `A ATENDER`

Los textos deben salir de hechos reales del pipeline, sin causas inventadas.

### G.6 Scope estático — todavía no asumir

La composición exacta del `Core Comercial Estático` debe resolverse en una sola auditoría consolidada.

Especialmente confirmar:

- Shoppings Estático;
- AA2000 Estático, si corresponde;
- cualquier otro componente Core real;
- exclusiones de APSA/London;
- si Pilar Frontlight tiene tratamiento particular;
- Cencomedia/MAB solo si realmente pertenecen al Core y no simplemente al catálogo ampliado.

**Estado:** a resolver en la auditoría inicial de TV3.

---

## H. Prompt maestro TV3 vigente para nueva sesión de Claude Code

El siguiente es el contenido funcional que debe recibir Claude. Puede compactarse en forma sin perder ninguna regla.

```text
IMPLEMENTAR TV3 PRODUCTIVA — CORE COMERCIAL ESTÁTICO

Trabajar sobre la TV3 actual usando como referencia visual:
audit_sources/TV3_REFERENCE.html
(read-only)

La referencia original puede contener datos legacy/window.OCU_DATA: NO usar esos valores como fuente de verdad.

Mantener la arquitectura productiva ya establecida en TV1/TV2:
pipeline validado → builder → JSON → template → HTML.
No poner lógica de negocio nueva en el HTML.

Crear/usar:
- scripts/build_tv3_dashboard.py
- scripts/templates/tv3_template.html
- tests/test_build_tv3_dashboard.py
- output/tv3_data.json
- tv3.html

NO tocar:
- TV1
- TV2
- Gates 1–4
- Excel
- business_semantics.json salvo bloqueo real
- Power BI
- Git commit/push
- deploy

IMPORTANTE — PENDIENTE TV2
NO investigar ni corregir ahora la diferencia de Shoppings Digital elegibles (actualmente 68 vs ~61 esperado por negocio).
Eso queda pendiente para mañana.
No gastar tiempo/créditos en TV2.

1. CONCEPTO GENERAL TV3

Cambiar toda la nomenclatura VISIBLE de “Fijo” a “Estático”.
Ejemplos: Fijo Core → Core Comercial Estático; ocupación fija → ocupación estática; Shoppings Fijo → Shoppings Estático.
Cambiar copy/UI, no valores internos válidos del pipeline.

HEADER:
Eyebrow: TV 3 · OCU26 · INTELIGENCIA COMERCIAL
Título: CORE COMERCIAL ESTÁTICO
Subtítulo: indicar claramente qué universo incluye, siguiendo exactamente la lógica de TV2 Digital.
Primero resolver desde las reglas reales del pipeline cuáles son las familias que conforman el Core Comercial Estático y luego escribir el subtítulo con nombres ejecutivos.
NO inventar familias. NO incluir Cencomedia/MAB/APSA/London si no pertenecen al Core Comercial.
Mantener exclusiones APSA y London.
Pregunta guía: ¿CÓMO RINDE EL CORE ESTÁTICO Y DÓNDE QUEDA DISPONIBILIDAD?
Header/reloj/actualización/período igual a TV1/TV2. Período actual Julio 2026.

2. UNA SOLA AUDITORÍA CONSOLIDADA

Resolver para junio/julio y serie Ene–Jul:
CORE ESTÁTICO TOTAL: elegibles, con campaña, ocupación calendario %, disponibles, disponibilidad %, MetricStatus.
POR FAMILIA ESTÁTICA del Core: mismos campos.
POR SHOPPING: elegibles, con campaña, ocupación % para ranking julio.
POR TIPO DE SOPORTE: métrica canónica para soportes más vendidos/activados en julio, Top 3.
EVOLUCIÓN: ocupación calendario mensual del Core Estático Ene–Jul.
Reconciliar con TV1 cuando scope sea equivalente.
No unknown→0. Preservar MetricStatus.
No hacer micro-auditorías después salvo error real.

3. KPI CARDS — MISMO SISTEMA DE TV2

Mantener 5 tarjetas superiores con:
NÚMERO ENTERO GRANDE + PORCENTAJE secundario + comparación con JUNIO debajo.
Mejora verde, empeora rojo, igual gris.
Para ocupación, más = mejor. Para disponibilidad, menos puede ser mejor utilización.

CARD 1 — OCUPACIÓN POR CALENDARIO
principal = elementos estáticos Core con campaña en julio;
secundario = % sobre elegibles;
abajo = junio + delta pp.

CARD 2 — SHOPPINGS ESTÁTICO
principal = elementos Shoppings Estático con campaña;
secundario = % sobre elegibles;
abajo = junio + delta pp.

CARD 3 — AA2000 ESTÁTICO solo si pertenece realmente al Core y la métrica está disponible.
Si no, usar la familia real que corresponda. No inventar 0.

CARD 4 — ACTIVIDAD / INVENTARIO ESTÁTICO
Elegir del audit la métrica ejecutiva más útil que pueda representarse correctamente como entero + porcentaje + comparación mensual.
No mantener Shopping Top como KPI textual si rompe el sistema; el mejor shopping ya queda en ranking.

CARD 5 — DISPONIBLES
principal = cantidad absoluta de elementos estáticos disponibles según definición temporal válida;
secundario = % sobre inventario elegible;
abajo = junio + delta pp.
Si la definición canónica es próximos 30 días, mantener DISPONIBLES · 30 DÍAS.
No reutilizar el legacy 533: recalcular.

4. CUERPO

Mantener composición general referencia: panel grande izquierdo + columna derecha dividida.

IZQUIERDA:
OCUPACIÓN POR SHOPPING · JULIO
Ranking Shoppings Estático. Hasta Top 5 si entra limpio.
Cada fila: ranking, shopping, elementos con campaña/elegibles o disponibles, ocupación %, barra.
Orden: ocupación desc, cantidad activa desc, nombre asc.

DERECHA SUPERIOR:
SOPORTES MÁS VENDIDOS · JULIO
Reducir a TOP 3.
Mantener barras horizontales.
Nombre soporte + valor entero + barra.

DERECHA INFERIOR:
EVOLUCIÓN MENSUAL · CORE ESTÁTICO
Subtítulo: Ocupación por calendario
Serie: Core Comercial Estático
Ene–Jul únicamente, sin proyecciones.
Mostrar etiqueta numérica solo en julio.

5. INSIGHTS INFERIORES

Eliminar único bloque vertical Lectura.
Replicar TV1/TV2:
LECTURA | PUNTO POSITIVO | A ATENDER.
Solo hechos reales. No inventar causas.

6. ESTILO

Brand Plus, Poppins, azul #1C60FF, navy/dark, cream para Estático, verde positivo, rojo negativo, amarillo oportunidad, cards redondeadas, 1920×1080, sin scroll/overflow.
Usar “Estático” correctamente en UTF-8.

7. TESTS TV3

Validar mínimo:
payload solo TV3; APSA/London excluidos; scope Core Estático reconciliado; ocupación distinct ElementoID; no duplicación; Shoppings Estático correcto; AA2000 solo si corresponde; disponibles correctos; porcentajes; junio como previo; deltas pp; ranking ordenado; soportes max 3; evolución Ene–Jul; sin futuro; MetricStatus; sin OCU_DATA legacy; SHA Excel intacto; TV1/TV2/referencia intactas.
No correr full suite.

8. WORKFLOW EFICIENTE

1. Leer solo partes relevantes de CM1 y TV3_REFERENCE.
2. Leer mínimo TV1/TV2 productivas para patrón.
3. Una auditoría consolidada TV3.
4. Implementar builder + payload + template en una pasada.
5. Regenerar TV3.
6. Tests específicos TV3 una vez.
7. Sanity mínimo.
8. Preview 1920×1080.

No explicar arquitectura ya conocida. No scripts auxiliares innecesarios. No explorar repo repetidamente. No auditar TV2. No commit/push/deploy.

Al finalizar reportar de forma concisa:
- universo exacto Core Comercial Estático
- KPIs julio/junio
- resultado tests
- archivos creados/modificados
- preview lista para revisión visual
```

---

## I. Archivos/referencias que debe tener la nueva sesión TV3

Mínimo recomendable dentro del repo:

- `docs/CM1.md` actualizado con esta versión;
- `audit_sources/TV3_REFERENCE.html` copiado desde `TV3_fijo_core_.html` y tratado como read-only;
- archivos productivos TV1/TV2 ya existentes para patrón;
- pipeline/Gates existentes sin modificarlos.

No hace falta volver a cargar capturas de TV2 para implementar TV3 si los archivos productivos y este CM1 están presentes.

---

## J. Reglas de permisos Claude Code reforzadas

Para ahorrar créditos y reducir riesgo:

**Permitir una vez** normalmente para:

- lectura de archivos;
- `grep`/búsqueda puntual;
- auditoría consolidada read-only;
- builder de una TV cuando se espera que escriba sus outputs;
- pytest específico de esa TV;
- preview local;
- `git status` / `git diff` read-only.

**Denegar / pedir justificación** para:

- scripts placeholder/no-op sin valor claro;
- micro-auditorías repetidas;
- tocar Gates protegidos;
- tocar Excel fuente;
- modificar TV1/TV2 durante TV3;
- modificar `business_semantics.json` sin bloqueo real;
- `git add`, commit, push;
- deploy;
- dependencias nuevas.

---

## K. Pendientes vigentes reordenados al ~22:46 ART

### P0 — Entrega TVs

1. Implementar y cerrar TV3 Core Comercial Estático.
2. Implementar TV4.
3. Implementar TV5.
4. Implementar TV6.
5. Revisión visual conjunta 1920×1080 de las seis TVs.
6. Corregir solo errores críticos antes de presentación.

### P1 — mañana antes del cierre/publicación

1. TV1: cambiar copy Core Comercial a `Shoppings · Pantallas · AA2000 · YPF` sin recalcular 964.
2. TV2: validar denominador elegible de Shoppings Digital (~68 actual vs ~61 recordado por usuario).
3. TV2: en esa misma revisión, comprobar clasificación de Remeros y reconciliación Pantallas/Shoppings.
4. Revisar estado Git real y archivos modificados.

### P2 — después de TVs

- commit/push solo con aprobación;
- deploy/publicación después de revisión;
- retomar Power BI desde martes 11/8 o cuando cierre la entrega HTML.

---

## L. Próximo paso exacto

1. Guardar esta versión como `docs/CM1.md` en el repo.
2. Abrir **nueva sesión de Claude Code** en `C:\brand plus\ocu26-dashboard`.
3. Asegurar `audit_sources/TV3_REFERENCE.html` read-only.
4. Pegar el prompt maestro TV3 de la sección H.
5. Autorizar una única auditoría consolidada y luego seguir builder → tests TV3 → preview.
6. No reabrir TV2 esta noche.

---

## M. Snapshot ejecutivo actualizado — corte ~22:46 ART

1. TV1 prácticamente cerrada; solo copy pendiente mañana.
2. TV2 productiva implementada.
3. TV2: 37 tests específicos pasan.
4. TV2 Core Digital = Pantallas LED + Shoppings Digital + AA2000.
5. TV2 ocupación calendario general: 71 activos / 78,0% según universo actual de 91.
6. TV2 fill: 321 slots / 20,4%.
7. TV2 Pantallas calendario: 11 / 100,0%.
8. TV2 Shoppings calendario: 60 / 88,2%.
9. TV2 slots disponibles: 1.254 / 79,6%.
10. TV2 disponibilidad: Pantallas 61,2% · Shoppings 79,6% · AA2000 100,0%.
11. TV2 rankings: Pantallas Top 5, Shoppings Top 3.
12. TV2 evolución: Pantallas vs Shoppings, Ene–Jul, labels solo julio.
13. TV2 insights: Lectura / Punto positivo / A atender.
14. Elemento residual detrás de KPIs TV2 eliminado.
15. Abrir template directo con S/D no significa falta de datos; revisar `tv2.html` generado.
16. Pendiente mañana: Shoppings Digital elegibles ~68 actual vs ~61 recordado.
17. Pendiente relacionado: clasificación de Remeros A VALIDAR.
18. TV3 referencia recibida: `TV3_fijo_core_.html`.
19. TV3 debe llamarse visible `Core Comercial Estático`, no Fijo.
20. TV3 debe replicar entero + porcentaje + mes anterior en cards.
21. TV3 mantiene ranking por shopping.
22. TV3 soportes pasa a Top 3.
23. TV3 agrega evolución lineal de ocupación estática Ene–Jul.
24. TV3 agrega Lectura / Punto positivo / A atender.
25. Scope exacto Core Estático debe salir de una sola auditoría consolidada; no asumir.
26. Próximo trabajo: TV3 en nueva sesión Claude Code.
27. Python del repo confirmado: `./.venv/Scripts/python.exe`.
28. No full suite por defecto.
29. No commit/push/deploy durante construcción sin aprobación.
30. Power BI sigue pausado.

---


---

# ACTUALIZACIÓN VIGENTE CM1 — 9/8/2026 ~21:08 ART

> **ESTA SECCIÓN ES EL OVERRIDE MÁS RECIENTE DEL DOCUMENTO.**  
> Si cualquier sección posterior o anexo histórico contradice esta actualización, prevalece esta sección.  
> La razón es que TV1 continuó evolucionando después del corte anterior del contexto maestro.

## A. Punto exacto actual

TV1 quedó visual y conceptualmente aprobada en su última captura, con un único ajuste mínimo de copy postergado expresamente para mañana.

**No quedan auditorías de negocio abiertas para TV1.**
**No volver a investigar YPF, composición o el grano de estación.**
**No volver a rediseñar TV1.**

El único cambio pendiente es:

- tarjeta `CORE COMERCIAL`;
- valor se mantiene en `964`;
- subtítulo actual: `Universo comercial core`;
- subtítulo deseado para mañana: `Shoppings · Pantallas · AA2000 · YPF`;
- es solo copy;
- no recalcular;
- no reabrir la composición;
- luego hacer preview 1920×1080 y cerrar TV1.

Después:
1. revisar `git status`;
2. revisar `git diff`;
3. decidir commit/push;
4. no deploy salvo pedido explícito;
5. comenzar TV2 en una **sesión nueva de Claude Code** para reducir contexto/tokens.

## B. Estado visual y numérico TV1 — última versión

### Tarjeta 1 — CORE COMERCIAL
- valor: **964**
- subtítulo actual visible: `Universo comercial core`
- cambio pendiente mañana: `Shoppings · Pantallas · AA2000 · YPF`

Reconciliación matemática del Core:
- Shoppings: 449
- Pantallas: 11
- YPF: 451 estaciones
- AA2000: 52
- Pilar Frontlight: 1
- total: **964**

La leyenda ejecutiva que el usuario quiere mostrar mañana omite Pilar Frontlight por decisión de copy de alto nivel; eso **no cambia** el cálculo de 964.

### Tarjeta 2 — UNIDADES CON CAMPAÑA
- julio: **495**
- % del Core: **51,3%**
- junio: **457**
- delta: **+38**
- delta verde
- se eliminó definitivamente la línea redundante:
  `YPF 305 · Est 119 · Dig 71`
- no volver a agregar ese desglose dentro de la tarjeta.

### Tarjeta 3 — YPF
- julio: **305 estaciones con campaña**
- junio: **263**
- delta: **+42**
- % del total de unidades con campaña: **61,6%**
- `Presencia en activas: Est 0,0% · Dig 100,0%`

**Confirmación de negocio del usuario:** este 0% estático / 100% digital es correcto. Actualmente las campañas YPF vendidas son digitales; el estático se encuentra en negociación comercial. No mostrar warning ni volver a auditar.

### Tarjeta 4 — ESTÁTICO
- activos: **119**
- ocupación: **29,4%**
- jerarquía visual:
  - 119 blanco, grande;
  - 29,4% azul, secundario.
- junio: 30,5%
- delta: -1,1 pp en rojo
- YTD: 28,6%

### Tarjeta 5 — DIGITAL POR CALENDARIO
- activos: **71**
- % con actividad: **78,0%**
- 71 blanco/grande
- 78,0% azul/secundario
- junio: 79,1%
- delta: -1,1 pp rojo
- YPF excluido.

### Tarjeta 6 — DIGITAL POR FILL RATE
- slots ocupados: **321**
- % capacidad vendida: **20,4%**
- 321 blanco/grande
- 20,4% azul/secundario
- junio: 27,3%
- delta: -6,9 pp rojo
- YPF excluido.

## C. Evolución mensual TV1

Título:
`Evolución mensual 2026`

Subtítulo:
`Unidades con campaña`

Series vigentes:
- Digital: `[69, 71, 72, 72, 71, 72, 71]`
- Estático: `[112, 110, 115, 116, 116, 122, 119]`
- YPF estaciones: `[0, 0, 0, 0, 233, 263, 305]`

Interpretación YPF:
- es `COUNT DISTINCT StationKey_TV1`;
- cuenta estaciones con campaña, no elementos ni campañas;
- una campaña presente en muchas estaciones produce muchas estaciones con campaña pero sigue siendo 1 `IDCampaña`.

Footer:
`Campañas únicas acumuladas a julio: 425`

## D. Composición del catálogo comercial

Título:
`Composición del catálogo comercial`

Subtítulo:
`Participación sobre 1.065 unidades comerciales`

El panel utiliza **catálogo ampliado**, no actividad del mes.

Reconciliación:
- Core Comercial: **964**
- Cencomedia: **88**
- MAB: **13**
- catálogo ampliado: **1.065**

Familias visibles:
- Shoppings: **449 = 68 Digital + 381 Estático = 42,2%**
- Pantallas: **11 Digital = 1,0%**
- YPF: **451 estaciones = 42,3%**
- AA2000: **52 = 12 Digital + 40 Estático = 4,9%**
- Cencomedia: **88 Estático = 8,3%**
- Otros: **14 = MAB 13 + Pilar Frontlight 1 = 1,3%**

Notas visibles acordadas:
- `* Catálogo ampliado = Core Comercial 964 + Cencomedia 88 + MAB 13 = 1.065 unidades.`
- `* Otros = MAB 13 + Pilar Frontlight 1.`

Reglas:
- Remeros está dentro de **Shoppings**.
- Shoppings no significa solamente Cencosud.
- Pantallas `11` es cantidad y `1,0%` es participación; no confundir.
- YPF se muestra como una sola barra por estación.
- no abrir Digital/Estático YPF dentro de esta composición transversal.

## E. Insights inferiores

Los títulos son horizontales y arriba.

### LECTURA
Julio registra 495 unidades con campaña:
- 305 estaciones YPF;
- 119 unidades estáticas;
- 71 digitales;
- 425 campañas únicas acumuladas.

### PUNTO POSITIVO
- total +38 vs junio;
- impulsado por YPF +42;
- Digital -1;
- Estático -3.

### A ATENDER
- fill rate digital cae 6,9 pp;
- 27,3% → 20,4%.

## F. YPF — decisión estructural que NO debe perderse

La fuente actual **NO contiene una columna APIE**.

Auditoría realizada:
- YPF total observado: 3.082 ElementoID;
- 440 prefijos candidatos a estación;
- 36/440 prefijos tenían más de una `Ubicacion` (8,2%);
- prefijo solo fue descartado;
- se eligió temporalmente `prefijo + localidad normalizada`;
- clave temporal: `StationKey_TV1`;
- catálogo derivado usado en TV1: 451 estaciones;
- caso residual conocido: prefijo 824, Rosario, dos direcciones bajo mismo prefijo/localidad;
- riesgo máximo documentado: subconteo de 1 estación sobre 451.

**Etapa 2 obligatoria:** incorporar un identificador real de estación/APIE en la estructura y eliminar el surrogate.

## G. Validaciones acumuladas

Confirmadas históricamente durante TV1:
- TV1 después del cambio YPF: **39/39 PASS**
- suite completa previa: **238/238 PASS**
- SHA Excel verificado sin cambios
- APSA excluido
- London excluido
- payload TV1 sin tv2–tv6
- `TV1_REFERENCE` intacto
- UTF-8 / entidad `Estático` corregida
- preview local 1920×1080 sin scroll evidente

Última ejecución específica de test después de ajustes finales de copy:
- Claude la ejecutó;
- el resultado textual no quedó reproducido en esta conversación;
- si se necesita evidencia formal antes de commit: `A VALIDAR`;
- NO correr suite completa de nuevo por defecto.

## H. Git / estado de escritura

Durante TV1 se trabajó con:
- NO commit;
- NO push;
- NO deploy;
- NO `git add .`.

La UI de Claude mostró crecimiento de cambios locales durante la sesión. Eso no significa commit.

Antes de cualquier commit:
1. `git status`;
2. `git diff`;
3. confirmar archivos modificados;
4. confirmar que no se tocaron:
   - input Excel;
   - Gates;
   - config semántica protegida;
   - Power BI;
   - `TV1_REFERENCE.html`.

## I. Eficiencia Claude Code — regla nueva para TV2–TV6

TV1 consumió más contexto porque hubo:
- definición de arquitectura;
- problemas YPF;
- APIE inexistente;
- varias auditorías;
- iteraciones visuales.

Esto **NO debe repetirse**.

Desde TV2:
1. abrir sesión nueva de Claude Code;
2. usar `CM1.md` como contexto maestro;
3. ChatGPT define primero KPIs/scopes;
4. enviar un único prompt maestro;
5. una auditoría inicial agrupada;
6. implementación;
7. tests específicos;
8. preview;
9. cierre.

Evitar:
- micro-prompts;
- auditorías repetitivas;
- releer TV1;
- volver a explicar branding completo si el repo ya lo contiene;
- suites completas innecesarias.

## J. Power BI

Permanece pausado.

Error conocido:
`El argumento 'dataType' no puede ser nulo. Nombre del parámetro: dataType`

No troubleshooting hasta después de las TVs prioritarias.

## K. Próximo paso exacto mañana

1. abrir TV1;
2. cambiar SOLO:
   `Universo comercial core`
   por
   `Shoppings · Pantallas · AA2000 · YPF`;
3. regenerar;
4. preview rápida;
5. cerrar TV1;
6. revisar Git;
7. abrir sesión nueva Claude Code;
8. empezar TV2.

## L. Última referencia visual crítica

La última captura completa de TV1 mostrada en el chat el 9/8/2026 alrededor de las 21:08 es la **referencia visual vigente**.

Características visibles:
- fondo Brand Plus azul/navy;
- logo Brand Plus;
- 6 KPIs;
- Core 964;
- 495;
- YPF 305;
- Estático 119 / 29,4%;
- Digital 71 / 78,0%;
- Fill 321 / 20,4%;
- evolución mensual grande;
- composición 1.065;
- notas de catálogo;
- 3 insights inferiores horizontales.

### IMAGEN QUE DEBE VOLVER A CARGARSE

**Imagen / referencia:** última captura completa de TV1 del 9/8/2026 ~21:08.  
**Qué contiene:** versión visual prácticamente final de TV1.  
**Por qué es importante:** permite confirmar que mañana solo se modifica el subtítulo de Core y no se rediseña nada más.  
**Decisiones dependientes:** cierre TV1; baseline para continuidad visual.  
**Qué debe observar el próximo ChatGPT:** mantener todo igual salvo el subtítulo de Core.  
**Estado:** CRÍTICA.

---

## 0. Cómo utilizar este documento

Este archivo es una migración completa del conocimiento disponible sobre OCU26: conversaciones de ChatGPT, decisiones tomadas a partir de Claude/Claude Code, auditorías, Excel, scripts, tests, capturas, HTML y referencias visuales.

Debe usarse como **fuente de verdad inicial** antes de modificar la solución.

Reglas de lectura:

1. Leer primero las secciones 3, 4, 6, 7, 23, 34, 35, 36 y 37.
2. Las decisiones del 9/8/2026 por la tarde/noche sobre TV1 y YPF reemplazan las versiones previas.
3. No reconstruir Gates 1–4.
4. No reabrir preguntas ya auditadas salvo dato nuevo, error real o test fallido.
5. Cuando algo no esté confirmado, usar `A VALIDAR`.
6. Los HTML son salida; no deben convertirse en una segunda capa de lógica de negocio.
7. `audit_sources/TV1_REFERENCE.html` es baseline visual y no debe modificarse.
8. Power BI no forma parte del camino crítico de las TVs.
9. Claude Code debe recibir prompts maestros cerrados para ahorrar créditos.
10. Antes de commit/push/deploy se debe aprobar visualmente y revisar Git.

---

## 1. Objetivo general del proyecto

### Qué se está construyendo

OCU26 es el sistema de inteligencia comercial y ocupación de Brand Plus. Debe centralizar:

- inventario publicitario;
- campañas históricas, activas, reservadas y futuras;
- semántica de negocio;
- métricas;
- outputs reutilizables;
- seis dashboards HTML;
- futuro Power BI interno;
- futura operación M365/SharePoint.

### Por qué

La base histórica acumuló fórmulas, cascadas y lógica difícil de escalar. La incorporación de YPF mostró además que no todos los circuitos tienen el mismo grano comercial ni la misma disponibilidad de datos.

La solución debe evitar:

- lógica duplicada;
- métricas engañosas;
- fill rate inventado;
- unknown convertido a cero;
- inventarios legacy contaminando denominadores;
- dependencia de una PC encendida.

### Audiencia

- Dirección / gerencia Brand Plus.
- Comercial.
- Producto.
- Operaciones.
- Usuarios de 6 TVs.
- Futuro usuario de Power BI.

### Resultado final

```text
Excel / fuente operativa
        ↓
Gate 1 Validación
        ↓
Gate 2 Transformación
        ↓
Gate 3 Semántica + MetricsEngine
        ↓
Gate 4 Outputs / data mart
        ↓
Capa de dashboard por TV
        ↓
tv1.html ... tv6.html
        ↓
Hosting público
        ↓
CMS / TV
```

### Entregables

**Inmediatos**
- 6 HTML productivos.
- Datos reales.
- tests por TV.
- 1920×1080.
- diseño Brand Plus.
- publicación después de aprobación.

**Etapa 2**
- Power BI.
- SharePoint.
- automatización.
- APIE/ID estación YPF real.
- eliminación del surrogate de estación.

---

## 2. Resumen de la evolución del proyecto

### Etapa inicial

Arquitectura Microsoft 365 simple:

```text
Forms
↓
1 Power Automate
↓
Excel plano en SharePoint
↓
Power BI
↓
6 TVs
```

Principio inicial: el Excel guarda hechos, las capas posteriores calculan.

### Primeros cambios

Power BI dejó de ser requisito para TVs por:

- licencias;
- restricciones del tenant;
- no querer costos adicionales;
- no depender de una PC.

Se definió HTML público como salida práctica.

### Repositorio y Gates

Se creó:

`C:\brand plus\ocu26-dashboard`

Se implementaron:
- Gate 1;
- Gate 2;
- Gate 3A;
- Gate 3B;
- Gate 3B.1;
- Gate 3B.1.1;
- Gate 4A;
- Gate 4B.

### Una sola lógica central

La arquitectura abandonó la idea de calcular reglas por separado en cada HTML/Power BI.

Regla:
> La lógica de negocio vive en la capa central; las vistas consumen.

### Evolución de dashboards

Narrativa vigente:

1. TV1 — Visión general del negocio.
2. TV2 — Digital.
3. TV3 — Estático.
4. TV4 — YPF.
5. TV5 — Performance / demanda.
6. TV6 — Pipeline / proyección.

### APSA / London

Ambos quedaron fuera de dashboards estándar y denominadores.

### Power BI

Gate4B se especificó. Al aplicar en Power BI Desktop apareció:

`El argumento 'dataType' no puede ser nulo. Nombre del parámetro: dataType`

Power BI se pausó.

### Implementación productiva TV1

Se tomó `TV1_REFERENCE.html` como baseline. Claude Code construyó:

- builder TV1;
- template;
- payload JSON;
- tests;
- preview local.

Durante esa implementación se detectó que YPF no podía contarse por `ElementoID` en una visión transversal porque comercialmente se vende por estación.

### Cambio YPF

Se auditó si existía APIE:
- no existe en Excel;
- no existe en maestro resuelto.

Se auditó el prefijo de `ElementoID`:
- 440 prefijos candidatos;
- 36 con múltiples `Ubicacion`;
- 8,2% ambiguos.

Prefijo solo fue descartado.

Se eligió temporalmente:
`StationKey_TV1 = prefijo + localidad normalizada`.

La implementación previa produjo:
- 451 estaciones YPF de catálogo;
- 3.082 ElementoID YPF;
- 305 estaciones activas en julio;
- 0 StationKey no derivables;
- un caso residual conocido: prefijo 824 Rosario, dos direcciones bajo misma localidad; posible subconteo máximo 1 estación.

### Último cambio TV1 después de ver el dashboard

El usuario decidió:

- agregar una tarjeta YPF;
- quitar Campañas únicas de arriba si falta espacio;
- cambiar “Unidades con actividad” a “Unidades con campaña”;
- volver a validar específicamente las 305 estaciones YPF;
- cambiar el panel derecho a composición del catálogo;
- meter Remeros dentro de Shoppings;
- mantener split Digital/Estático dentro de Shoppings;
- títulos inferiores horizontales;
- corregir `Est&aacute;tico`;
- continuar con una sola iteración eficiente.

---

## 3. Estado actual exacto

### Terminado

- Auditoría V3.
- Gate 1.
- Gate 2.
- Gate 3A.
- Gate 3B.
- Gate 3B.1 / 3B.1.1.
- Gate 4A.
- Gate 4B como especificación.
- Commits y push de Gates.
- Outputs Parquet.
- Manifest.
- Sistema visual Brand Plus.
- Baseline TV1.
- Builder productivo TV1.
- Auditoría YPF para estación.
- Implementación surrogate temporal.
- Tests TV1 posteriores a YPF: **39/39 PASS**.
- Última suite completa ejecutada: **238/238 PASS**, antes de los ajustes conceptuales finales.
- SHA Excel sin cambios.
- APSA/London fuera del payload TV1.
- `TV1_REFERENCE` intacto.
- preview HTTP local ejecutada.

### Funcionando

- `scripts/validate_input.py`
- `scripts/transform_data.py`
- `scripts/semantic_model.py`
- `scripts/metrics_engine.py`
- `scripts/export_data.py`
- `scripts/build_tv1_dashboard.py`
- `scripts/templates/tv1_template.html`
- `tests/test_build_tv1_dashboard.py`
- `output/tv1_data.json`
- `tv1.html`

### Validado antes de la última revisión

- Core: **964**.
- YPF catálogo: **451 estaciones**.
- Campañas únicas YTD: **425**.
- Unidades con actividad/campaña: **495**.
- YPF julio: **305 estaciones**.
- Estático activos: **119**.
- Estático ocupación: **29,4%**.
- Digital calendario: **71/91 = 78,0%**.
- Digital fill: **321/1.575 = 20,4%**.
- Junio fill: **27,3%**, delta **-6,9 pp**.

### En desarrollo

Última iteración TV1:
- validar 305 YPF con campaña válida;
- agregar KPI YPF;
- cambiar KPI transversal a Unidades con campaña;
- remover Campañas únicas de fila superior si hace falta;
- composición del catálogo;
- Remeros dentro de Shoppings;
- aclarar/eliminar Otros;
- bottom titles horizontales;
- fix encoding `Estático`;
- tests TV1 afectados;
- preview final.

### Pendiente

- cerrar TV1;
- screenshot final;
- revisión visual;
- `git status` / `git diff`;
- commit/push solo después de aprobación;
- no deploy aún;
- TV2–TV6;
- hosting definitivo;
- CMS;
- Power BI;
- SharePoint;
- automatización;
- APIE real Etapa 2.

### Bloqueado

No hay bloqueo estructural.

### A validar

- 305 estaciones YPF con campaña en julio;
- mix estático/digital de estaciones activas YPF;
- composición exacta del catálogo por familia;
- contenido de “Otros” luego de reagrupar Remeros;
- tipo de integración Netlify A/B/C;
- compatibilidad física final webOS;
- causa `dataType`;
- `q`;
- `CantidadUnidades` Cencomedia;
- spots >10 s;
- exclusividad.

---

## 4. Decisiones vigentes

### Decisión 1 — lógica central

**Decisión vigente:** HTML y Power BI no recalculan negocio independientemente.  
**Estado:** confirmada.

### Decisión 2 — fuente

`input/OCU26_BASE_DATOS.xlsx`.  
Pipeline read-only.

### Decisión 3 — SHA

Valor esperado usado explícitamente en la sesión productiva TV1:

`2f165e12ad90c2f05963367a8e1717dcd1006e9ce669f976afa6e57470aca2cd`

Un contexto anterior contenía una variante tipográfica. Para continuidad, usar este valor reciente y recalcular si hubiera discrepancia.

### Decisión 4 — APSA

Fuera de todo dashboard estándar.

### Decisión 5 — London Supply

Fuera de todo dashboard estándar hasta nuevo aviso.

### Decisión 6 — Cencomedia

Entra al universo estático general sin inventar capacidad/ocupación.

### Decisión 7 — YPF transversal

No YPF:
`COUNT DISTINCT ElementoID`

YPF:
`COUNT DISTINCT estación`

### Decisión 8 — StationKey_TV1

Temporal:

`prefijo numérico ElementoID + localidad normalizada Ubicacion`

Sin fuzzy matching.  
Sin fallback a `ElementoID`.  
No llamarlo APIE real.

### Decisión 9 — Etapa 2 YPF

Agregar ID/APIE real estructural y eliminar surrogate.

### Decisión 10 — TV1 KPIs

Orden final previsto:

1. Core Comercial.
2. Unidades con Campaña.
3. YPF.
4. Estático.
5. Digital por Calendario.
6. Digital por Fill Rate.

Campañas únicas sale de arriba si hace falta espacio.

### Decisión 11 — Unidades con campaña

Una unidad cuenta solo si tiene campaña válida que intersecta el período.

### Decisión 12 — KPI YPF

Mostrar:
- estaciones con campaña;
- % sobre unidades con campaña;
- presencia estática/digital entre estaciones activas.

Los dos porcentajes de presencia no tienen por qué sumar 100%.

### Decisión 13 — Evolución

Título:
`EVOLUCIÓN MENSUAL 2026`

Subtítulo:
`UNIDADES CON CAMPAÑA`
o `UNIDADES COMERCIALES CON CAMPAÑA`.

Series:
- Digital BP: ElementoID distinct.
- Estático BP: ElementoID distinct.
- YPF: StationKey distinct.

### Decisión 14 — Composición

`COMPOSICIÓN DEL CATÁLOGO COMERCIAL`

Denominador:
Core comercial.

No participación sobre actividad.

### Decisión 15 — Shoppings

Shoppings no equivale a Cencosud.

Remeros debe incluirse dentro de Shoppings.

### Decisión 16 — Composición por familia

- Shoppings: Digital + Estático.
- Pantallas: Digital.
- YPF: una barra por estaciones.
- AA2000: Digital/Estático si aplica.
- Cencomedia: Estático.
- Otros: solo si queda algo real; explicar contenido.

### Decisión 17 — bottom insights

Títulos horizontales arriba.

### Decisión 18 — punto positivo

Si valores se mantienen:
`La actividad total crece en 38 unidades respecto a junio, impulsada por YPF (+42 estaciones); Digital −1 y Estático −3.`

### Decisión 19 — a atender

Fill rate 27,3% → 20,4%, -6,9 pp.

### Decisión 20 — campañas

`COUNT DISTINCT IDCampaña`.

### Decisión 21 — vs

Siempre mes anterior.

### Decisión 22 — YTD

Estático por elemento-día acumulado; no promedio de porcentajes.

### Decisión 23 — Power BI

Pausado.

### Decisión 24 — eficiencia Claude

ChatGPT define y cierra; Claude implementa por prompt maestro. Evitar micro-auditorías.

---

## 5. Historial de cambios de decisiones

### Motor
Excel/Power BI → Python/semántica central.

### TVs
Power BI Service → HTML público.

### YPF
ElementoID → estación.

### Identificador estación
prefijo → descartado por ambigüedad → prefijo + localidad normalizada temporal.

### KPI transversal
Elementos con actividad → Unidades con actividad → **Unidades con campaña**.

### Campañas únicas
KPI superior → puede salir de arriba y quedar en gráfico/lectura.

### Composición
actividad mensual → **catálogo comercial**.

### Remeros
circuito separado visualmente → dentro de **Shoppings** en TV1.

### Insights
títulos verticales → horizontales.

### Power BI
crítico → secundario → pausado.

---

## 6. Reglas de negocio

### `ElementoID`

Identifica elemento inventario.  
Vincula maestro/campañas.

### `IDCampaña`

Identifica campaña comercial.

Puede repetirse por elementos.  
Conteo siempre DISTINCT.

### `CargaID`

Fila/carga de campaña.

### `ClaveNegocio`

Clave compuesta de control.

### YPF

Formatos:
- MB = Menu Board, digital.
- TT = Torre, digital.
- PPUNTER = Puntera, digital.
- FB = Fotobox/Mupi, estático.

Ejemplos:
- `256 - MB - 1`
- `256 - TT - 1`
- `256 - PPUNTER - 1`
- `256 - FB - 1`

### No existe APIE físico

La fuente actual no contiene columna `APIE`.

### Surrogate

`StationKey_TV1 = prefijo + localidad normalizada`.

Datos de auditoría:
- 3.082 ElementoID YPF.
- 440 prefijos.
- 36 ambiguos por ubicación.
- 8,2%.
- 451 estaciones derivadas en implementación previa.
- 0 StationKey no derivables reportados.
- caso 824 Rosario residual, máximo 1 estación de subconteo.

### Actividad/campaña

Una unidad con campaña debe tener una campaña válida que intersecte el período.

YPF:
- 1 campaña en 1 estación = 1 estación.
- 1 campaña en 200 estaciones = 200 estaciones y 1 campaña única.
- varios formatos en misma estación = 1 estación.

### Core previo

`513 no YPF + 451 YPF = 964`.

### Digital calendario previo

`71/91 = 78,0%`.

YPF excluido.

### Fill previo

`321/1.575 = 20,4%`.

YPF excluido.

### Estático previo

119 activos.  
29,4% ocupación.

### Evolución previa

Digital:
`[69, 71, 72, 72, 71, 72, 71]`

Estático:
`[112, 110, 115, 116, 116, 122, 119]`

YPF:
`[0, 0, 0, 0, 233, 263, 305]`

La serie YPF y 305 quedan `A VALIDAR` una vez con definición estricta.

### Campañas únicas

YTD previo: 425.

### Digital segundos

72.000 s/día.  
1 salida = 1.800 s.  
2 = 3.600.  
3 = 5.400.  
4 = 7.200.

### Capacidades

- Tótem 20.
- Puente 13.
- Triedro 20 si reel.
- Pantalla 20.
- YPF formatos digitales 20 teóricos, sin fill real CMS.

### `SalidasVendidas` faltante

No asumir 0.

### MetricStatus

- `OK`
- `PARTIAL`
- `NO_APLICA`
- `REQUIERE_CONFIRMACION`

---

## 7. Arquitectura actual

```text
input/OCU26_BASE_DATOS.xlsx
        ↓
validate_input.py
        ↓
transform_data.py
        ↓
semantic_model.py + business_semantics.json
        ↓
metrics_engine.py
        ↓
export_data.py
        ↓
Parquet / manifest
        ↓
builder dashboard
        ↓
JSON específico
        ↓
template
        ↓
HTML
```

TV1:

```text
pipeline
↓
scripts/build_tv1_dashboard.py
↓
output/tv1_data.json
↓
scripts/templates/tv1_template.html
↓
tv1.html
```

No reabrir Gates por ajustes TV1.

---

## 8. Arquitecturas anteriores o descartadas

- Excel con fórmulas como motor.
- Power BI obligatorio.
- PC encendida.
- lógica propia en HTML.
- Office Scripts complejos inmediatos.
- prefijo solo YPF.
- fallback estación→ElementoID.
- fuzzy matching.
- composición TV1 por actividad como panel principal.
- títulos verticales.
- London/APSA en resumen.

---

## 9. Bases de datos

### `OCU26_BASE_DATOS.xlsx`

Ruta:
`input/OCU26_BASE_DATOS.xlsx`

Conteos históricos vigentes del pipeline:
- maestro 4.338;
- campañas 9.503;
- parámetros 23.

SHA reciente:
`2f165e12ad90c2f05963367a8e1717dcd1006e9ce669f976afa6e57470aca2cd`

### V3 auditada

`Base_ocupacion_26_FINAL_CON_YPF_AUDITADA_V3.xlsx`

Auditoría:
- 1.583 altas YPF;
- 7.758 activaciones YPF;
- 32 YA_CARGADOS excluidos;
- 0 duplicados nuevos;
- 0 campañas con ElementoID inexistente.

### Bases históricas

- `Final Base_ocupacion_26__4-8_CON_YPF.xlsx`
- `Final Base_ocupacion_26_ 4-8.xlsx`
- `YPF - Base campañas y elementos corregido estructural.xlsx`
- `OCU26_FORM_CARGA_LIVIANO(1).xlsx`
- `BPBI_OCU26_V2_REPARADO_DESPLEGABLES_OK.xlsx`
- `Base_ocupacion_26__4-8.xlsx`
- base YPF fuente.

---

## 10. Tablas, hojas y columnas

### Legacy
- `BASE_MAESTRA_ELEMENTOS`
- `BASE_CAMPAÑAS`
- `CONTROL_DISPONIBILIDAD`
- `AUX_CASCADA`

### Base plana
- `MAESTRO_ELEMENTOS`
- `CAMPANAS`
- `PARAMETROS`
- `tblElementos`
- `tblCampanas`

### Inventario
- `ElementoID`
- `Ciudad`
- `Medio`
- `CircuitoDashboard`
- `Subcircuito`
- `Ubicacion`
- `Nivel`
- `TipoInventario`
- `TipoInstalacion`
- `TipoCatalogo`
- `Material`
- `AplicaCantidad`
- `CapacidadSlotsReel`
- `SegundosDia`
- `DimensionOptico`
- `DimensionTotal`
- `Observaciones`
- `b`
- `h`
- `q`
- `m2`

### Campañas
- `CargaID`
- `ClaveNegocio`
- `FechaHoraCarga`
- `UsuarioCarga`
- `FuenteCarga`
- `ElementoID`
- `IDCampaña`
- `Campaña`
- `Cliente`
- `Marca`
- `Agencia`
- `Proveedor`
- `FechaInicio`
- `FechaFin`
- `FechaIndefinida`
- `Estado`
- `DuracionSpotSeg`
- `SalidasVendidas`
- `CantidadUnidades`
- `ModalidadPauta`
- `PROGRAMATICA`
- `CANJE`
- `TipoExclusividad`
- `HoraInicio`
- `HoraFin`
- `Observaciones`

### Semánticas
- `CircuitoNegocio`
- `SitioNegocio`
- `FormatoNegocio`
- `CoberturaCatalogo`
- `CompletitudMaestro`
- `CertezaDato`
- `ModoDisponibilidad`
- `PortfolioTier`
- `IncluyePerformanceCore`
- `IncluyeConteoGeneral`
- `VisiblePorDefecto`
- `TieneActividadComercial`
- `CantidadCampanasHistoricas`
- `FechaPrimeraCampana`
- `FechaUltimaCampana`
- `SlotsComerciales`
- `SegundosComerciales`

### Paralelas numéricas
- `BValor`
- `HValor`
- `M2Valor`

No `QValor`.

### `APIE`

No existe físicamente.

---

## 11. Identificadores

| ID | Qué identifica | Regla |
|---|---|---|
| `ElementoID` | elemento | estable |
| `IDCampaña` | campaña | puede repetirse por elementos |
| `CargaID` | fila/carga | fact/bridge |
| `ClaveNegocio` | combinación de negocio | control duplicados |
| `StationKey_TV1` | estación YPF temporal | prefijo + localidad |

`StationKey_TV1` NO es APIE real.

---

## 12. Archivos del proyecto

| Archivo | Función | Estado |
|---|---|---|
| `input/OCU26_BASE_DATOS.xlsx` | input | protegido |
| `scripts/validate_input.py` | Gate1 | protegido |
| `scripts/transform_data.py` | Gate2 | protegido |
| `scripts/semantic_model.py` | Gate3 | protegido |
| `scripts/metrics_engine.py` | métricas | protegido |
| `scripts/export_data.py` | Gate4A | protegido en tarea TV1 |
| `config/business_semantics.json` | semántica | protegido |
| `scripts/build_tv1_dashboard.py` | builder TV1 | activo |
| `scripts/templates/tv1_template.html` | template | activo |
| `tests/test_build_tv1_dashboard.py` | tests | activo |
| `output/tv1_data.json` | payload | generado |
| `tv1.html` | producto TV1 | en ajuste |
| `audit_sources/TV1_REFERENCE.html` | baseline | read-only |
| `output/*.parquet` | data mart | generado |
| `output/_export_manifest.json` | manifest | generado |
| `powerbi/*` | Power BI spec | pausado |
| `TV2_digital_core_(1).html` | baseline TV2 | referencia |
| `TV3_fijo_core_(1).html` | baseline TV3 | referencia |
| `TV4_pipeline_(1).html` | baseline pipeline | usar TV6 |
| `TV5_demanda_(1).html` | baseline demanda | usar TV5 |
| `OCU26_Etapa_1_Sistema_Simple.docx` | arquitectura M365 | referencia |
| `Auditoria_Independiente_OCU26_V3.pdf` | auditoría | referencia |

---

## 13. Estructura de carpetas

```text
ocu26-dashboard/
├── .claude/
├── .git/
├── .pytest_cache/
├── .venv/
├── audit_sources/
│   └── TV1_REFERENCE.html
├── config/
│   └── business_semantics.json
├── docs/
├── input/
│   └── OCU26_BASE_DATOS.xlsx
├── output/
│   ├── *.parquet
│   ├── _export_manifest.json
│   └── tv1_data.json
├── powerbi/
├── scripts/
│   ├── validate_input.py
│   ├── transform_data.py
│   ├── semantic_model.py
│   ├── metrics_engine.py
│   ├── export_data.py
│   ├── build_tv1_dashboard.py
│   └── templates/
│       └── tv1_template.html
├── tests/
│   └── test_build_tv1_dashboard.py
└── tv1.html
```

---

## 14. Código y scripts

### `validate_input.py`
Gate1, read-only.

### `transform_data.py`
Gate2, read-only en tarea dashboards.

### `semantic_model.py`
Resuelve semántica.

### `metrics_engine.py`
API genérica de métricas.

### `export_data.py`
Genera data mart y expone `load_pipeline()` usado en auditorías.

### `build_tv1_dashboard.py`
Builder productivo.

Responsabilidades:
- cargar pipeline;
- construir TV1;
- derivar estación YPF temporal;
- fallar si no puede derivar;
- generar JSON;
- render HTML.

### `tv1_template.html`
Template productivo.

### `test_build_tv1_dashboard.py`
Tests TV1.

---

## 15. Fórmulas y cálculos

### Core

```text
Core = ElementoID distinct no YPF + StationKey distinct YPF
```

Previo:
`513 + 451 = 964`.

### Unidades con campaña

```text
No YPF = ElementoID distinct con campaña válida
YPF = StationKey distinct con campaña válida
Total = suma de ambos universos
```

### Campañas

`COUNT DISTINCT IDCampaña`.

### Digital calendario

`activos / elegibles`.

### Fill

`slots ocupados / capacidad`.

### Estático

Elemento-día.

### YTD

Acumulado desde 1/1.

### KPI YPF mix

- % estaciones activas con al menos un formato estático.
- % estaciones activas con al menos un formato digital.

Pueden superponerse.

---

## 16. Validaciones

### Integridad
SHA Excel PASS.

### Estructura
Gates PASS.

### YPF
- APIE ausente confirmado.
- prefijo ambiguo confirmado.
- surrogate elegido.
- caso 824 documentado.

### TV1
- builder PASS.
- JSON.
- HTML.
- 39/39 tests.
- 238/238 suite previa.
- APSA/London excluidos.
- payload solo TV1.
- reference intacta.
- preview HTTP.

### Pendiente visual
Nueva iteración final.

---

## 17. Tests

Último estado:
- TV1: **39/39 PASS**.
- suite completa: **238/238 PASS**, 13m43s.

Después de la última revisión conceptual:
- ejecutar solo tests TV1 afectados;
- suite completa solo si cambio técnico lo justifica.

Coberturas importantes:
- IDCampaña DISTINCT;
- YPF por StationKey;
- varios elementos misma estación = 1;
- no fallback;
- no APSA/London;
- payload TV1;
- no future months;
- MetricStatus;
- SHA.

---

## 18. Herramientas utilizadas

- ChatGPT.
- Claude/Claude Code.
- Excel.
- Python.
- Pandas.
- PyArrow.
- HTML/CSS/JS.
- Chart.js.
- Git.
- GitHub.
- GitHub Desktop.
- Netlify.
- GitHub Pages.
- SharePoint.
- Microsoft Forms.
- Power Automate.
- Power BI Desktop.
- CMS / TV.

---

## 19. Restricciones de herramientas y licencias

- No pagar licencias extra.
- Power BI Free no resuelve compartir.
- Publicar en web depende de tenant.
- datos pueden ser públicos.
- no PC encendida.
- GitHub gratuito.
- Claude pago.
- M365 disponible.
- carga inicial manual posible.
- automatización después.
- minimizar consumo de Claude.

---

## 20. Git / GitHub

### Baseline

Commits:
- Gate3B `610d5b981fef011bd2e3ab991ba6ec828cf4faf2`
- Gate4A `7b54220186031f695f5edbf9b611fbf74d964a67`
- Gate4B `c8c1dbd0d40cf80db2be8dbc66fbc7546f30772f`

Pusheados a `main`.

### TV1 actual

Regla:
- no commit;
- no push;
- no deploy;
- no `git add .`.

Claude UI mostró durante la sesión hasta aproximadamente `+2,874 -0`, señal de cambios locales/session. No implica commit.

Antes de cualquier commit:
1. `git status`;
2. `git diff`;
3. revisar protegidos;
4. aprobación visual.

---

## 21. SharePoint / Microsoft 365

Sitio propio creado.

Estructura planificada:

```text
01_DATOS/
02_POWER_BI/
03_HISTORICOS/
04_DOCUMENTACION/
05_HTML/
```

Etapa 1:
- Excel plano;
- Forms;
- un Power Automate;
- concurrencia 1.

Etapa 2:
- robustecer;
- Lists si aporta;
- APIE real YPF;
- automatización.

---

## 22. Power BI

Tablas:
- DIM_ELEMENTOS.
- FACT_CAMPANAS.
- DIM_CALENDARIO.
- BRIDGE_CAMPANA_DIA.
- FACT_METRICAS_DIARIA.

Relaciones documentadas en Gate4B.

Error:
`dataType` nulo.

Sospecha histórica:
`FechaHoraCarga` completamente nula/tipo indefinido, NO confirmada.

Estado:
**PAUSADO.**

---

## 23. HTML / dashboards / salidas

### 23.1 Hardware

Destino:
**LG 43SM5KB-BD**

- horizontal;
- 1920×1080;
- webOS antiguo ~2.0;
- priorizar JS/CSS simple y estable.

### 23.2 Branding

Tipografía:
`Poppins`.

Colores:
- `#1C60FF`
- `#1D1D1B`
- `#EEECE6`
- `#FF4E46`
- navy `#071124`
- dark `#05070D`
- teal/cyan para Digital;
- gris para Estático;
- violeta para YPF;
- verde positivo;
- warning/coral solo cuando corresponda.

Logo oficial embebido en referencia.

### 23.3 Baseline TV1

`audit_sources/TV1_REFERENCE.html`

- no modificar;
- 1920×1080;
- logo izquierda;
- reloj derecha;
- fondo navy;
- cards;
- paneles;
- Chart.js;
- OCU_DATA legacy NO es fuente.

### 23.4 TV1 previa a ajuste final

Mostraba:
- Core 964.
- Campañas 425.
- Unidades 495.
- Estático 119 / 29,4%.
- Digital calendario 71 / 78%.
- Fill 321 / 20,4%.
- YPF 305.
- composición actividad.
- títulos inferiores verticales.

### 23.5 TV1 final deseada

**KPIs**
1. Core Comercial.
2. Unidades con Campaña.
3. YPF.
4. Estático.
5. Digital por Calendario.
6. Digital por Fill Rate.

**Evolución**
- título `EVOLUCIÓN MENSUAL 2026`;
- subtítulo `Unidades con campaña`;
- Digital/Estático/YPF.

**Composición**
- `COMPOSICIÓN DEL CATÁLOGO COMERCIAL`;
- participación sobre Core;
- Shoppings incluye Remeros;
- YPF una barra.

**Bottom**
- Lectura.
- Punto positivo.
- A atender.
- títulos horizontales arriba.

**Fix**
- `Estático`, no entidad HTML literal.

### 23.6 TV2

Digital sin YPF/APSA/London.

### 23.7 TV3

Estático sin YPF/APSA/London; con Cencomedia.

### 23.8 TV4

YPF con apertura por estación/formato/geografía/campaña.

### 23.9 TV5

Performance/demanda.

### 23.10 TV6

Pipeline/proyección.

---

## 24. Automatización

### Existente
Excel → Python → outputs.

### Planificada
Excel → CI/Python → JSON/HTML → hosting.

### Netlify

URL conocida:
`https://digitalcore-brandplus.netlify.app/`

Es hosting, no fuente de reglas.

A VALIDAR:
- A auto repo;
- B manual;
- C indeterminado.

No deploy durante ajuste TV1.

---

## 25. Referencias del chat de Claude e imágenes

### 25.1 Trabajo Claude

Claude Code:
- auditó repo;
- ejecutó Gates;
- creó export;
- Power BI spec;
- builder TV1;
- tests;
- YPF audit;
- preview;
- SHA;
- sanity.

### 25.2 YPF audit

Se realizaron auditorías focalizadas:
1. existencia APIE;
2. headers Excel;
3. patrón ElementoID;
4. ambigüedad de prefijo;
5. 36 casos;
6. clasificación por Medio/localidad;
7. prefijo 824;
8. nulos/malformed Ubicacion;
9. implementación surrogate.

### 25.3 Capturas de permisos

Regla:
- lectura/test = permitir una vez;
- comando mal escapado = denegar;
- writes TV1 esperados = evaluar;
- Gates/Excel/Git/deploy = detener salvo instrucción explícita.

### 25.4 Captura TV1

Mostró layout productivo 1920×1080 y permitió detectar:
- KPI YPF ausente;
- 495 confuso;
- YPF dominante;
- composición poco útil;
- títulos verticales;
- encoding Estático.

### 25.5 Imágenes que deben volver a cargarse

| Imagen / referencia | Qué contiene | Por qué | Prioridad |
|---|---|---|---|
| captura TV1 actual previa al último ajuste | versión productiva intermedia | comparar | CRÍTICA |
| captura TV1 final después del ajuste | versión a aprobar | cierre | CRÍTICA |
| render/HTML de `TV1_REFERENCE` | baseline visual | identidad | CRÍTICA |
| capturas YPF audit | evidencia surrogate | trazabilidad | IMPORTANTE |
| Power BI dataType | error | Etapa2 | REFERENCIA |
| GitHub Desktop | baseline Git | continuidad | REFERENCIA |
| baselines TV2–TV5 | próximas TVs | diseño | IMPORTANTE |

---

## 26. Otros archivos externos que deberían recuperarse

### TV1
- `CONTEXTO_MAESTRO.md`
- repo.
- Excel.
- `TV1_REFERENCE.html`
- `build_tv1_dashboard.py`
- template.
- test.
- JSON.
- HTML.
- captura.

### Próximas TVs
- baselines.
- config.
- scripts centrales.
- outputs.

### Etapa2
- `powerbi/`.
- documento M365.
- auditoría V3.
- manual Brand Plus.

---

## 27. Problemas encontrados

### YPF inflado por elementos
Resuelto temporalmente por estación.

### APIE ausente
Surrogate temporal.

### Prefijo ambiguo
Resuelto parcialmente con localidad.

### Caso 824
Residual conocido.

### 495 confuso
Nueva KPI YPF + desglose.

### 305 dudoso
A validar una vez.

### Composición redundante
Cambiar a catálogo.

### Otros opaco
Reagrupar Remeros; listar resto.

### Títulos verticales
Cambiar.

### Encoding
Corregir.

### Power BI
Pausado.

### Claude créditos
Prompts maestros.

---

## 28. Errores ya cometidos

1. 3 salidas = 7.200, incorrecto; son 5.400.
2. Pantallas LED lista de 10, incompleta.
3. digital calendario ≠ fill.
4. fill YPF inventado, no.
5. `TipoCatalogo` ≠ semántica completa.
6. London/APSA en denominadores, no.
7. OCU_DATA legacy como fuente, no.
8. YPF por ElementoID transversal, no.
9. llamar surrogate “APIE”, evitar.
10. prefijo solo estación, no.
11. títulos verticales, no.
12. Shoppings = Cencosud, no.
13. `git add .`, no.
14. micro-auditorías innecesarias, evitar.
15. suite completa repetida por cambio cosmético, evitar.

---

## 29. Soluciones descartadas

- Excel como motor.
- Power BI obligatorio.
- lógica duplicada.
- prefijo YPF solo.
- fallback a ElementoID.
- fuzzy matching.
- composición de actividad TV1.
- títulos verticales.
- London/APSA estándar.

---

## 30. Supuestos

### Confirmados
- datos públicos aceptables.
- HTML para TVs.
- YPF vende estación completa.
- Remeros es Shopping no Cencosud.
- APIE no existe.
- surrogate derivable para dataset actual.

### Incorrectos
- YPF por ElementoID.
- prefijo solo.
- composición actividad = mejor panel.
- Shoppings = Cencosud.

### A validar
- 305.
- mix YPF.
- composición catálogo.
- Otros.

---

## 31. Dependencias

Excel → Gates → outputs → builder → JSON → template → HTML.

YPF:
ElementoID + Ubicacion → StationKey.

TV1:
Core, unidades, evolución, composición dependen de StationKey.

TV4:
abre formatos internos.

---

## 32. Riesgos

- cambiar Excel;
- romper Gates;
- institucionalizar surrogate;
- usar 305 sin revalidación;
- duplicar estación por formatos;
- subcontar 824;
- forzar porcentajes YPF a 100%;
- dejar Remeros fuera de Shoppings;
- dejar Otros sin explicar;
- unknown→0;
- JS moderno incompatible webOS;
- CDN/red;
- commit prematuro;
- deploy prematuro;
- reabrir Power BI antes de TVs.

---

## 33. Datos que NO deben modificarse sin validación

- Excel.
- SHA.
- ElementoID.
- IDCampaña.
- CargaID.
- ClaveNegocio.
- Gates.
- config semántica.
- TV1_REFERENCE.
- capacidades.
- MetricStatus.
- exclusiones.
- regla YPF estación.
- branding.
- stage 1920×1080.

---

## 34. Pendientes completos

### P0 — críticos

**Cerrar TV1**
- validar 305;
- KPI YPF;
- Unidades con campaña;
- composición catálogo;
- Remeros en Shoppings;
- Otros;
- bottom horizontal;
- encoding;
- tests TV1;
- preview;
- screenshot;
- QA;
- Git diff.

**TV2–TV6**
Pendiente.

### P1
- hosting;
- CMS;
- commit/push aprobados;
- prueba física.

### P2
- Power BI;
- SharePoint;
- APIE real;
- automatización.

### P3
- Lists;
- Forms;
- exclusividad;
- >10s;
- forecast.

---

## 35. Roadmap actual

1. Cerrar TV1.
2. TV2.
3. TV3.
4. TV4.
5. TV5.
6. TV6.
7. publicar.
8. Power BI.
9. SharePoint/Etapa2.

Método TV2–TV6:
- definición ChatGPT;
- 1 prompt maestro;
- 1 auditoría agrupada;
- implementación;
- tests específicos;
- preview;
- cierre.

---

## 36. Próximo paso exacto

Ejecutar en Claude el prompt maestro de **ajuste final TV1** ya preparado.

Ese prompt debe:
1. validar solo YPF julio;
2. aplicar cambios;
3. tests TV1;
4. sanity mínimo;
5. preview;
6. screenshot.

Luego ChatGPT revisa visualmente.

No Power BI.  
No commit.  
No deploy.

---

## 37. Punto exacto donde terminó este chat

La última TV1 productiva fue mostrada en captura.

El usuario detectó:

- 495 requiere explicación;
- YPF necesita su propia tarjeta;
- campañas únicas puede salir de arriba;
- 305 YPF parece alto;
- evolución debe expresar unidades con campaña;
- composición debe ser del catálogo;
- Remeros va dentro de Shoppings;
- títulos de abajo horizontales;
- el resto de decisiones previas se mantiene.

Se entregó un prompt maestro para que Claude haga **una sola iteración** con todos esos cambios.

El siguiente chat debe continuar exactamente desde ahí.

---

## 38. Preguntas abiertas

1. ¿305 estaciones YPF tienen campaña válida julio?
2. % estaciones activas YPF con formato estático.
3. % estaciones activas YPF con formato digital.
4. composición exacta catálogo 964.
5. contenido final de Otros.
6. Netlify A/B/C.
7. prueba webOS física.
8. `dataType`.
9. `q`.
10. CantidadUnidades Cencomedia.
11. spots >10s.
12. exclusividad.

---

## 39. Glosario

**OCU26:** ocupación/inteligencia comercial 2026.  
**APIE:** identificador conceptual/real de estación YPF; no existe hoy como columna.  
**StationKey_TV1:** surrogate temporal.  
**MB:** Menu Board.  
**TT:** Torre.  
**PPUNTER:** Puntera.  
**FB:** Fotobox/Mupi.  
**Fill rate:** capacidad digital vendida / total cuando aplica.  
**MetricStatus:** estado métrica.  
**YTD:** acumulado anual.  
**Shoppings:** familia que incluye Cencosud y Remeros en TV1.  
**Core:** universo comercial principal.

---

## 40. Personas, equipos, proveedores y actores

- Brand Plus: dueño del proyecto.
- Responsable del proyecto: reglas/aceptación.
- ChatGPT: producto/prompts/QA.
- Claude Code: implementación.
- Microsoft: M365/Power BI.
- GitHub: versionado.
- Netlify: hosting.
- YPF: circuito sin CMS Brand Plus.
- AA2000: aeropuertos.
- Cencosud: shoppings.
- Remeros: shopping no Cencosud.
- Cencomedia: estático/flexible.
- APSA/London: fuera de estándar.

---

## 41. Cronología relevante

### Antes de agosto
Base legacy, fórmulas, incorporación YPF.

### 5/8
M365 Etapa1.

### 5–6/8
Auditoría V3.

### 7/8
Gate3A.

### 8/8
Gate3B.

### 8–9/8
Gate4A.

### 9/8 mañana
Gate4B, Power BI, dataType, pausa.

### 9/8 mediodía
6 TVs y scopes.

### 9/8 tarde
TV1 productiva.

### 9/8 tarde/noche
YPF estación:
- APIE ausente.
- 440/36.
- surrogate.
- 451 estaciones.
- 39/39.
- 238/238.

### 9/8 noche
revisión visual final y nuevo prompt de ajuste.

---

## 42. Reglas para el próximo ChatGPT

“Este documento representa el estado acumulado del proyecto al cierre del chat anterior.

El proyecto tuvo múltiples iteraciones y cambios de criterio.

No asumir que una decisión temprana sigue siendo válida cuando existe una decisión posterior.

Antes de modificar la solución:

1. leer este documento completo;
2. identificar las decisiones vigentes;
3. respetar las reglas de negocio;
4. no reconstruir componentes que ya funcionan;
5. no recomendar soluciones previamente descartadas sin una razón nueva;
6. mantener compatibilidad con los archivos y arquitectura actual;
7. verificar las dependencias antes de cambiar nombres, estructuras o IDs;
8. diferenciar claramente una corrección necesaria de un rediseño opcional;
9. utilizar las referencias visuales y archivos indicados cuando sean necesarios;
10. si una decisión depende de una imagen que no está disponible, solicitar específicamente esa imagen en lugar de asumir su contenido;
11. preguntar únicamente cuando la información no pueda resolverse utilizando este documento y los archivos disponibles.”

Reglas adicionales:
12. No APSA/London.
13. No fill real YPF.
14. No rediseñar branding.
15. `vs` mes anterior.
16. Power BI pausado.
17. No datos de ejemplo.
18. No reabrir Gates.
19. Mantener MetricStatus.
20. Priorizar TVs.
21. YPF transversal = estación.
22. surrogate temporal.
23. APIE real en Etapa2.
24. No llamar APIE al surrogate.
25. Remeros dentro de Shoppings TV1.
26. composición TV1 = catálogo.
27. prompt maestro por TV.
28. revisar Git antes de commit.
29. no deploy sin aprobación.

---

## 43. Instrucciones para reconstruir el contexto visual

Cuando se abra un nuevo chat:

1. subir este archivo;
2. subir captura TV1 actual/final;
3. subir `TV1_REFERENCE.html`;
4. para cada TV siguiente, subir su baseline;
5. comparar 1920×1080;
6. revisar branding, logo, Poppins, cards, charts, legibilidad;
7. no asumir contenido de imágenes ausentes;
8. si la validación de 305 ya ocurrió después de este documento, pedir ese resultado y actualizar contexto.

---

## 44. Checklist de migración de archivos

### Continuar TV1
- [ ] `CONTEXTO_MAESTRO.md`
- [ ] repo
- [ ] `tv1.html`
- [ ] `output/tv1_data.json`
- [ ] `scripts/build_tv1_dashboard.py`
- [ ] `scripts/templates/tv1_template.html`
- [ ] `tests/test_build_tv1_dashboard.py`
- [ ] `audit_sources/TV1_REFERENCE.html`
- [ ] captura TV1

### TV2–TV6
- [ ] baselines HTML
- [ ] Excel
- [ ] config
- [ ] scripts
- [ ] outputs

### Etapa2
- [ ] powerbi/
- [ ] captura dataType
- [ ] M365 doc
- [ ] auditoría V3
- [ ] manual Brand Plus
- [ ] APIE real cuando exista

---

## 45. Snapshot ejecutivo

1. OCU26 Brand Plus.
2. Repo `C:\brand plus\ocu26-dashboard`.
3. Branch `main`.
4. Excel vigente.
5. Gates cerrados.
6. Power BI pausado.
7. HTML prioridad.
8. TV1 builder existe.
9. TV1 JSON existe.
10. TV1 HTML existe.
11. referencia read-only.
12. TV1 39/39 previo.
13. suite 238/238 previa.
14. SHA intacto.
15. APSA fuera.
16. London fuera.
17. YPF transversal por estación.
18. APIE no existe.
19. StationKey = prefijo + localidad.
20. 36/440 ambiguos con prefijo solo.
21. 451 estaciones catálogo previo.
22. caso 824 residual.
23. Core 964 previo.
24. YPF julio 305 `A VALIDAR`.
25. Unidades 495 previo.
26. Estático 119 / 29,4%.
27. Digital calendario 71/91.
28. Fill 321/1.575.
29. Campañas YTD 425.
30. KPI YPF se agrega.
31. Campañas sale de arriba si hace falta.
32. KPI transversal = Unidades con campaña.
33. Evolución = unidades con campaña.
34. Composición = catálogo.
35. Remeros dentro de Shoppings.
36. YPF composición una barra.
37. KPI YPF puede mostrar presencia estático/digital.
38. Bottom horizontal.
39. Fix `Estático`.
40. No commit/push/deploy.
41. revisar Git.
42. Netlify existe.
43. LG 1920×1080 webOS antiguo.
44. TV2–TV6 con prompts maestros.
45. Etapa2 incorpora APIE real.

---

## ANEXO A — CONTEXTO MAESTRO ANTERIOR, CONSERVADO PARA TRAZABILIDAD

> El contenido siguiente corresponde a un corte anterior del mismo proyecto. Se conserva íntegro para no perder detalles técnicos, históricos, reglas, nombres, medidas, tests, archivos, Power BI, SharePoint y referencias de Claude. Si contradice las secciones 0–45 anteriores, prevalece la decisión más reciente de las secciones 0–45.

# CONTEXTO MAESTRO DEL PROYECTO

**Proyecto:** BRAND PLUS · OCU26 · Sistema de inteligencia comercial, ocupación y dashboards  
**Estado documentado:** 9 de agosto de 2026, aproximadamente 12:30 ART  
**Repositorio principal:** `C:\brand plus\ocu26-dashboard`  
**Branch vigente:** `main`  
**Uso de este documento:** migración exhaustiva para continuar indistintamente en ChatGPT, Claude o Claude Code sin depender del chat anterior.

---

## 0. Cómo utilizar este documento

Este archivo representa una migración del conocimiento acumulado del proyecto OCU26 desde conversaciones anteriores de ChatGPT, trabajo realizado en Claude/Claude Code, documentos de auditoría, bases Excel, scripts versionados, capturas de pantalla y los HTML de referencia visual.

Debe utilizarse como **fuente de verdad inicial** antes de modificar la solución.

Reglas de lectura:

1. La información más nueva reemplaza a la anterior cuando existe una contradicción explícita.
2. Las secciones **“Decisiones vigentes”**, **“Reglas de negocio”**, **“Estado actual exacto”**, **“Pendientes completos”**, **“Roadmap actual”** y **“Punto exacto donde terminó este chat”** tienen prioridad operativa.
3. Cuando una decisión haya cambiado, este documento intenta reconstruir la evolución y marca la regla vigente.
4. Cuando no existe evidencia suficiente se utiliza la marca `A VALIDAR`.
5. No se debe inferir que una clasificación histórica de Excel es equivalente a la clasificación comercial vigente.
6. No se debe reimplementar lógica de negocio dentro de los HTML, Power BI u otras vistas si esa lógica ya está resuelta en la capa semántica/motor central.
7. Los HTML subidos el 9/8/2026 son **fuente de verdad visual** para la presentación inmediata: se conserva su diseño Brand Plus; se cambia la información y la lógica de datos necesaria, no el sistema visual.
8. El objetivo inmediato al cierre de este documento es tener los **6 tableros HTML listos para presentar el lunes 10/8/2026**. Power BI queda para una etapa posterior a retomar desde el martes 11/8/2026.

---

# 1. Objetivo general del proyecto

## 1.1 Qué estamos construyendo

OCU26 es un sistema de inteligencia comercial y ocupación para Brand Plus que debe:

- centralizar el inventario de elementos publicitarios;
- centralizar campañas históricas, activas, reservadas y futuras;
- aplicar una única semántica de negocio;
- calcular métricas de inventario, actividad, ocupación, disponibilidad y performance;
- soportar circuitos con reglas distintas sin duplicar lógica;
- producir múltiples salidas de consumo;
- mostrar información ejecutiva en seis televisores;
- permitir posteriormente análisis interno en Power BI;
- evolucionar hacia una operación de carga centralizada en SharePoint/Microsoft 365;
- no depender de que una PC específica permanezca encendida para que las TVs funcionen.

## 1.2 Por qué

La base histórica de ocupación había acumulado fórmulas, cascadas, lógica repetida y una estructura difícil de escalar. La incorporación de YPF aumentó el volumen y la heterogeneidad del inventario. Además, Brand Plus necesita:

- entender rápidamente qué inventario tiene;
- saber qué parte se comercializa realmente;
- evitar porcentajes engañosos generados por inventario que casi nunca se vende;
- comparar meses;
- medir ocupación estática y fill rate digital con reglas correctas;
- visualizar YPF sin atribuirle métricas que Brand Plus no puede conocer;
- presentar tableros permanentes en TVs;
- disponer más adelante de un modelo analítico en Power BI;
- conservar trazabilidad y evitar que una vista invente su propia lógica.

## 1.3 Para quién

Principalmente:

- Dirección / gerencia de Brand Plus;
- área comercial;
- producto;
- operaciones;
- personas que consulten las seis TVs;
- posteriormente usuarios internos de Power BI.

## 1.4 Resultado final buscado

Arquitectura objetivo consolidada:

```text
FUENTE OPERATIVA / EXCEL
        ↓
GATE 1 — VALIDACIÓN
        ↓
GATE 2 — TRANSFORMACIÓN
        ↓
GATE 3 — SEMÁNTICA + MOTOR DE MÉTRICAS
        ↓
GATE 4 — DATA MART / OUTPUTS RESUELTOS
        ↓
 ┌────────────────────────────┬─────────────────────────────┐
 │                            │                             │
HTML / 6 TVs              Power BI interno            futuros consumos
(prioridad inmediata)     (segunda etapa)             / SharePoint
 │
 ↓
HOSTING PÚBLICO
 │
 ↓
CMS / Fire Stick / 6 TVs
```

## 1.5 Entregables

### Inmediatos
- 6 HTML finales, uno por TV.
- Datos reales y calculados desde la lógica central.
- Diseño visual idéntico al sistema Brand Plus ya construido.
- Comparaciones mensuales correctas.
- Presentación lista para el lunes 10/8/2026.

### Segunda etapa
- Power BI Desktop correctamente implementado.
- Modelo, relaciones, DAX, Power Query y validaciones.
- Continuación del flujo SharePoint/Microsoft 365.
- Hosting definitivo / automatización.
- Eventual actualización automática con GitHub Actions u otro flujo.

---

# 2. Resumen de la evolución del proyecto

## Etapa inicial

La primera solución propuesta para OCU26 fue una arquitectura simple Microsoft 365:

```text
Microsoft Forms
↓
1 flujo Power Automate
↓
OCU26_BASE_DATOS.xlsx en SharePoint, sin fórmulas
↓
Power BI
↓
6 páginas / 6 TVs
```

El objetivo era evitar que varias personas editaran una base pesada y evitar cientos de miles de fórmulas. Se diseñó una primera etapa con una persona de carga y un máximo futuro de 3–4 usuarios, con 1–2 simultáneos.

Principio inicial:
> El Excel guarda hechos; Power BI calcula resultados.

## Primeros cambios

Se decidió no depender exclusivamente de Power BI Service porque:

- Power BI Free no permite compartir de forma normal con terceros;
- la publicación pública depende de políticas del tenant;
- no se quería agregar costos de licencias;
- se quería evitar depender de una PC local;
- los datos no tienen una restricción de privacidad que impida URLs públicas.

Se definió entonces un **Plan B**, que pasó a convertirse en la opción práctica principal para las TVs:

```text
OCU26_BASE_DATOS.xlsx
↓
Python / Claude Code
↓
data / outputs
↓
tv1.html ... tv6.html
↓
GitHub Pages u hosting gratuito
↓
CMS / Fire Stick / TVs
```

Power BI quedó como herramienta de validación/modelado y luego como herramienta analítica interna, no como requisito para que las TVs funcionen.

## Normalización del repositorio

Se creó el repositorio:

`C:\brand plus\ocu26-dashboard`

con branch `main`.

Se implementaron Gates progresivos:

- Gate 1: validación del input.
- Gate 2: transformación.
- Gate 3A: definición semántica.
- Gate 3B: modelo semántico + motor de métricas.
- Gate 3B.1 / 3B.1.1: edge cases y políticas.
- Gate 4A: data mart / export.
- Gate 4B: especificación del modelo Power BI.

## Cambio material: una sola lógica central

Se abandonó la idea de que cada HTML o Power BI calculase su propia ocupación.

Regla vigente:
> Toda lógica de negocio debe vivir en la capa central. HTML y Power BI consumen resultados.

## Cambio material: semántica de circuitos

Se separaron conceptos que el Excel mezclaba:

- `TipoCatalogo`;
- cobertura real;
- completitud del maestro;
- certeza;
- modo de disponibilidad;
- portfolio;
- universo de análisis.

Se resolvieron reglas específicas para Cencosud, Remeros, Pantallas LED, Pilar Frontlight, AA2000, YPF, London Supply, MAB, Cencomedia y APSA.

## Cambio material: YPF

YPF pasó de incorporación de base a **parte central del negocio**.

Pero Brand Plus no administra el CMS de YPF, por lo que:
- no se conoce ocupación real de terceros;
- no se conoce fill rate real de slots;
- sí se conoce actividad de campañas Brand Plus sobre elementos registrados;
- sí se puede conocer capacidad teórica/comercial del formato;
- YPF debe tener un tablero propio.

## Cambio material: dashboards

Los primeros HTML diseñados separaban:
- TV1 Resumen Ejecutivo;
- TV2 Digital Core;
- TV3 Fijo Core;
- TV4 Pipeline;
- TV5 Demanda;
- existía lógica CSS para una TV6 de oportunidad, pero no se subió un HTML final de TV6 en este corte.

El 9/8/2026 se redefinió la narrativa a seis TVs:

1. General.
2. Digital sin YPF.
3. Estático sin YPF, incluyendo Cencomedia.
4. YPF.
5. Performance / demanda.
6. Pipeline / proyección.

## Cambio material: APSA y London Supply

### APSA
Ya estaba fuera del performance core y del conteo general en la semántica.

### London Supply
Antes se mantenía como complementario, conocido pero por consulta.

**Nueva decisión del 9/8/2026:**
> APSA y London Supply quedan fuera de TODO análisis y dashboard estándar hasta nuevo aviso.

Motivo:
- London Supply tiene muchos elementos que casi nunca se comercializan;
- solo se venden eventualmente algunos soportes, por ejemplo alguna columna de Ushuaia;
- su volumen ensucia inventarios, ocupaciones y porcentajes;
- APSA es legacy y sin acuerdo comercial vigente.

Se conservan en la fuente/histórico, pero no deben aparecer en la presentación ni contaminar denominadores.

## Cambio material: Power BI

Gate 4B dejó Power BI completamente especificado, pero durante la implementación manual en Power BI Desktop apareció:

`El argumento 'dataType' no puede ser nulo. Nombre del parámetro: dataType`

La cancelación de la carga quedó trabada.

El 9/8/2026 se decidió:
> Pausar Power BI. Retomarlo desde el martes 11/8/2026. No bloquear los tableros HTML por Power BI.

---

# 3. Estado actual exacto

## Terminado

- Auditoría estructural de la base V3.
- Gate 1.
- Gate 2.
- Gate 3A.
- Gate 3B.
- Gate 3B.1.
- Gate 3B.1.1.
- Gate 4A.
- Gate 4B.
- Repositorio Git.
- Commits Gate 3B, Gate 4A y Gate 4B.
- Push de los commits a `origin/main`.
- Export real de cinco tablas Parquet.
- Manifest de export.
- Tests automáticos.
- Definición del sistema visual Brand Plus de las TVs.
- Cinco HTML visuales de referencia subidos.

## Funcionando

- `validate_input.py`.
- `transform_data.py`.
- `semantic_model.py`.
- `metrics_engine.py`.
- `export_data.py`.
- generación de Parquet;
- lectura de los Parquet con PyArrow;
- modelo semántico;
- motor de métricas;
- universos existentes;
- MetricStatus;
- SHA read-only;
- Git/GitHub;
- branding y escalado 1920×1080 de los HTML.

## Validado

- suite completa: **199 passed, 0 failed** al cierre de Gate 4A/4B;
- Gate 4B: 8/8 validaciones específicas;
- SHA del Excel fuente sin cambios;
- archivos protegidos Gates 1–3 intactos;
- `.gitignore` ignora `output/`;
- Parquet round-trip;
- PK/FK e integridad referencial de outputs;
- no mezcla `pd.NA` con cero;
- columnas numéricas paralelas de Gate 4;
- 5 tablas / 5 relaciones Power BI;
- Power BI Desktop Free compatible en especificación.

## En desarrollo

- definición funcional detallada de TV1–TV6;
- actualización de los HTML para los nuevos scopes;
- integración de YPF como TV4;
- nueva lógica de comparación mes vs mes anterior;
- acumulado 2026/YTD especialmente para estático;
- regla de exclusión total APSA/London en dashboards;
- inclusión de Cencomedia en estático general.

## Pendiente

- generar los seis HTML finales;
- validar números reales de los seis tableros;
- presentar el lunes 10/8;
- definir hosting;
- conectar URLs al CMS/TV;
- automatizar regeneración;
- retomar Power BI el martes 11/8;
- retomar SharePoint/Forms/Power Automate cuando corresponda.

## Bloqueado

Ningún bloqueo estructural del pipeline.

Power BI Desktop está **pausado**, no bloquea el producto inmediato.

## A validar

- mecanismo exacto de hosting final;
- definición final de algunas tarjetas de TV1–TV6;
- si la nueva exclusión de London se incorpora como nuevo universo en `business_semantics.json` o solo como scope de salida;
- fórmula futura para spots >10 segundos;
- matemática futura de exclusividad;
- significado de la columna `q`;
- fuente futura de `CantidadUnidades` para Cencomedia;
- si se guardó accidentalmente algún PBIX durante la prueba manual: el chat indica “Sin título” y no hay evidencia de un PBIX guardado.

---

# 4. Decisiones vigentes

## 4.1 Lógica central

**Decisión vigente:** HTML y Power BI no deben tener lógica de negocio independiente.  
**Origen:** evolución Gate 3.  
**Por qué:** evitar inconsistencias.  
**Reemplaza:** cálculos propios en cada vista.  
**Consecuencia:** cambios de regla se resuelven centralmente.  
**Estado:** confirmada.

## 4.2 Fuente productiva del pipeline

**Decisión vigente:** `input/OCU26_BASE_DATOS.xlsx` es la fuente operativa local del pipeline actual.  
**Estado:** confirmada.

## 4.3 Protección del input

**Decisión vigente:** el pipeline no escribe el Excel.  
**Consecuencia:** SHA antes/después debe coincidir.  
**Estado:** confirmada.

## 4.4 APSA

**Decisión vigente:** fuera de todo análisis estándar, conteo, ocupación, performance y dashboards.  
**Solo:** histórico/consulta explícita.  
**Estado:** confirmada.

## 4.5 London Supply

**Decisión vigente al 9/8:** fuera de todo análisis y dashboards hasta nuevo aviso.  
**Motivo:** inventario voluminoso y de comercialización muy esporádica que distorsiona porcentajes.  
**Estado:** confirmada.

## 4.6 Cencomedia

**Decisión vigente al 9/8:** incluir en el tablero estático general, aunque hoy tenga poco/no tenga performance para mostrar.  
**Estado:** confirmada para dashboards.  
**Nota:** su `PortfolioTier` central continúa documentado como complementario hasta que se modifique explícitamente la semántica.

## 4.7 YPF

**Decisión vigente:** core del negocio; tablero propio.  
**Estado:** confirmada.

## 4.8 Digital TV2

**Decisión vigente:** digital sin YPF.  
Incluye digital gestionable/comparable de Brand Plus: Pantallas LED, Shoppings Digital y AA2000 Digital según disponibilidad de datos.  
**Estado:** confirmada como arquitectura de dashboard.

## 4.9 Estático TV3

**Decisión vigente:** estático sin YPF, APSA ni London; sumar Cencomedia.  
También considera Shoppings Fijo, Remeros, Pilar Frontlight y AA2000 Estático según la semántica y datos.  
**Estado:** confirmada como arquitectura de dashboard.

## 4.10 Comparaciones temporales

**Decisión vigente:** el “vs.” de una tarjeta es **siempre contra el mes anterior**.  
Ejemplo:
- Julio: 50.
- Junio: 48.
- Delta: +2.

Para porcentajes:
- Julio 32%.
- Junio 29%.
- Delta +3 pp.

**No usar YoY por defecto.**

## 4.11 Acumulado anual

**Decisión vigente:** puede mostrarse ocupación/acumulado 2026 (YTD), especialmente útil en estático.  
**Estado:** confirmada.

## 4.12 Diseño HTML

**Decisión vigente:** no tocar el sistema visual.  
Conservar:
- paleta;
- tipografía;
- fondos;
- cards;
- bordes;
- layout;
- logo;
- 1920×1080;
- escalado;
- estética Brand Plus.

Se pueden cambiar textos/títulos necesarios para reflejar el propósito de cada TV, pero sin rediseñar el sistema.

## 4.13 Power BI

**Decisión vigente:** postergado a etapa siguiente desde el martes 11/8.  
**No es requisito** para presentar las TVs el lunes.  
**Estado:** confirmada.

## 4.14 Prioridad

**Decisión vigente:** prioridad absoluta = 6 tableros listos para presentar el lunes 10/8/2026.  
**Estado:** confirmada.

## 4.15 Claude Code / tokens

**Decisión vigente:** ser más eficiente.  
- ChatGPT estructura y cierra requerimientos.
- Claude Code implementa con prompts maestros.
- Evitar iteraciones pequeñas si un prompt claro puede resolver un bloque completo.
- Mantener controles de seguridad en staging/commit/push.
**Estado:** confirmada.

---

# 5. Historial de cambios de decisiones

## Tema: motor de cálculo

**Inicial:** fórmulas en Excel / Power BI.  
**Intermedio:** Power BI como lugar principal de cálculos.  
**Final:** Python + semántica central + motor de métricas; Power BI y HTML consumen.  
**Regla actual:** no duplicar negocio en vistas.

## Tema: salida a TVs

**Inicial:** Power BI Service, seis páginas.  
**Intermedio:** evaluar “Publicar en web”.  
**Problema:** licencia/tenant/costo.  
**Final:** HTML público como salida primaria; Power BI interno posteriormente.

## Tema: APSA

**Inicial:** presente dentro de shoppings en base legacy.  
**Intermedio:** legacy, fuera performance.  
**Final:** fuera de todos los análisis estándar.

## Tema: London Supply

**Inicial:** inventario conocido, complementario, por consulta.  
**Intermedio:** podía participar en universo operativo general.  
**Final 9/8:** excluir completamente de análisis/dashboard hasta nuevo aviso.

## Tema: Cencomedia

**Inicial:** flexible/complementario, fuera performance core.  
**Final dashboard 9/8:** sumarlo al estático general; no darle peso artificial si no tiene actividad.

## Tema: YPF

**Inicial:** incorporación de nueva base.  
**Intermedio:** core pero con `TipoCatalogo=Abierto` histórico.  
**Final:** core + tablero específico; no medir fill rate de terceros.

## Tema: Power BI

**Inicial:** pieza crítica.  
**Intermedio:** modelo interno secundario.  
**Gate 4B:** especificación completa.  
**9/8:** intento manual falla por `dataType`; se pausa.  
**Regla actual:** retomar martes; no bloquear HTML.

## Tema: comparación

**Propuesta anterior:** mes anterior y eventualmente año/año anterior.  
**Final:** “vs” siempre mes anterior. YTD separado.

---

# 6. Reglas de negocio

## 6.1 Circuitos de negocio

### Cencosud / Shoppings
Universo cerrado/conocido.

Sedes mencionadas:
- Unicenter;
- Palmas del Pilar;
- Plaza Oeste;
- Portal Palermo;
- Factory San Martín;
- Factory Quilmes;
- Factory Brown;
- Portal Escobar;
- Portal Lomas;
- Portal Los Andes;
- Portal Rosario;
- Portal Tucumán;
- Portal Santiago;
- Portal Trelew;
- Portal Salta;
- Portal Patagonia.

Digital + estático según sede/elemento.

### Remeros
Shopping no Cencosud.
Digital + estático.
Core.

### Pantallas LED
11 ubicaciones confirmadas. Nombres conocidos en el contexto:
- 9 de Julio;
- Remeros;
- Pilar;
- Cabildo y Juramento;
- Cerrito;
- Avellaneda;
- Chacarita;
- Olazábal;
- Juan B. Justo y Panamericana;
- Córdoba.

**A VALIDAR:** recuperar del maestro/config el listado exacto completo de las 11 antes de escribirlo manualmente en una vista. La lista histórica de 10 estaba incompleta; Córdoba había sido omitida.

### Pilar
Mismo sitio con dos medios:
- Pantalla LED Pilar = digital.
- Pilar Frontlight = estático, contracara.
Ambos core.

### AA2000
Digital + estático.
Aeropuertos:
- Ezeiza;
- Aeroparque;
- Mendoza;
- Córdoba.

El universo comercial es conocido/cerrado, pero el maestro actual es incompleto en algunas plazas.
Regla central:
- `CoberturaCatalogo=COMPLETO`;
- `CompletitudMaestro=PARCIAL`;
- disponibilidad `MIXTO`.

No confundir incompletitud de carga con desconocimiento del negocio.

### YPF
Core.
El Excel histórico mantiene `TipoCatalogo=Abierto`.
No corregir ese valor legacy solo por semántica.

Identificación:
- estación/APIE;
- formato;
- secuencia.

Ejemplos:
- `256 - MB - 1`
- `256 - TT - 1`
- `256 - PPUNTER - 1`
- `256 - FB - 1`

Formatos:
- MB = Menu Board, digital.
- TT = Torre, digital.
- PPUNTER = Puntera, digital.
- FB = Fotobox/Mupi, estático.

Regla de certeza:
- si existe campaña, el elemento puede considerarse comercialmente confirmado para el propósito actual;
- no exigir auditoría física posterior como condición de existencia.

Brand Plus no administra el CMS YPF:
- no conoce ocupación total de terceros;
- no conoce fill rate real de slots;
- `slots_ocupados`, `slots_disponibles`, `fill_rate_slots`, `segundos_vendidos`, `segundos_disponibles`, `fill_rate_segundos` de actividad total de YPF = `NO_APLICA`;
- sí puede mostrar actividad de campañas registradas sobre elementos;
- sí puede mostrar capacidad teórica por formato si existe perfil.

### Cencomedia
Supermercados / flexible gráfico.
Formatos:
- cubre alarmas;
- stoppers;
- carros;
- floorgraphics.

Las cantidades se determinan por lo comprado en campaña.
Se agregan tiendas/elementos con el tiempo.
`CantidadUnidades` actualmente no ofrece base histórica suficiente.
Para dashboards 9/8:
- incluir en estático general;
- no forzar una métrica de unidades que no existe;
- no inventar ocupación.

### MAB
Móvil.
Abierto/flexible/complementario.
La ubicación actual no es permanente.

### APSA
Legacy.
Acuerdo comercial ya no vigente.
Nueva regla:
- excluir de absolutamente todo análisis estándar;
- conservar únicamente para histórico/consulta explícita.

### London Supply
Aeropuertos de Patagonia/otras plazas históricas.
Universo conocido pero comercialización muy esporádica.
Nueva regla 9/8:
- excluir de absolutamente todo análisis/dashboard estándar;
- conservar fuente/histórico;
- no permitir que sus numerosos elementos diluyan ocupación.

## 6.2 Universos semánticos implementados previamente

Antes del cambio del 9/8 existían:

### `OPERATIVO_GENERAL`
Excluía APSA, pero podía incluir complementarios como London.

### `PERFORMANCE_CORE`
Core:
- CENCOSUD;
- REMEROS;
- PANTALLAS_LED;
- PILAR_FRONTLIGHT;
- AA2000;
- YPF.

### `COMPLETO_HISTORICO`
Todo, incluyendo APSA.

### Nueva necesidad dashboard 9/8
Se necesita un scope que excluya:
- APSA;
- London Supply.

Y para TV3 incluya:
- Cencomedia.

**Importante:** esta necesidad nueva no debe implementarse con filtros ad hoc repetidos dentro de cada HTML. Debe resolverse en la capa de salida/configuración.

## 6.3 Temporalidad estática

- calendario inclusivo;
- mismo inicio y fin = 1 día ocupado;
- una campaña puede tener cualquier duración;
- `Reservada` bloquea capacidad futura;
- `Cancelado` no bloquea;
- `FechaIndefinida=Si` + `FechaFin` vacío puede extenderse hasta el final de la consulta;
- fechas faltantes sin `FechaIndefinida` son defecto de carga;
- no inventar fecha;
- resultado temporal afectado debe ser `PARTIAL` con warning.

## 6.4 Digital

Referencia comercial:
- 72.000 segundos/día.
- Spot base habitual: 10 s.

`SalidasVendidas`:
- 1 → 1.800 s/día.
- 2 → 3.600 s/día.
- 3 → 5.400 s/día.
- 4 → 7.200 s/día.

Corrección importante:
> 3 salidas NO son 7.200. Son 5.400.

Capacidades:
- Tótem: 20 slots.
- Puente LED: 13 slots comerciales.
- Triedro: 20 slots cuando funciona en reel.
- Pantalla LED: 20 slots.
- YPF TT/PPUNTER/MB: 20 slots teóricos.

Puente LED:
- legacy puede tener 10;
- capacidad comercial efectiva = 13;
- no sobrescribir legacy, aplicar override semántico.

Triedros:
- generalmente exclusividad/intervención;
- pueden operar reel;
- matemática de exclusividad no aprobada de forma universal.

Videos >10 s:
- sin fórmula general aprobada;
- `UnsupportedBusinessCaseError` si la métrica requiere actividad y el caso no está soportado.

Exclusividad:
- sin fórmula universal aprobada;
- no inventar.

## 6.5 `SalidasVendidas` faltante

No asumir 0.

- métricas basadas en segundos → `REQUIERE_CONFIRMACION`/NA cuando afectadas;
- slots pueden contar la campaña si la definición del slot no depende de `SalidasVendidas`.

## 6.6 AA2000

`ocupacion_calendario_pct`:
- `NO_APLICA` mientras `CompletitudMaestro != COMPLETO`.

Métricas digitales sobre elementos registrados:
- pueden calcularse;
- status `PARTIAL` si el maestro está incompleto.

## 6.7 MetricStatus

Estados que deben preservarse:
- `OK`;
- `PARTIAL`;
- `NO_APLICA`;
- `REQUIERE_CONFIRMACION`.

Nunca:
- convertir desconocido en 0;
- mostrar 0% por falta de dato;
- ocultar una advertencia relevante.

## 6.8 Comparación mensual de dashboards

Número principal = mes seleccionado.

Subdato obligatorio cuando exista mes anterior:
- valor mes anterior;
- delta.

Para porcentajes:
- delta en puntos porcentuales (`pp`) cuando corresponda.

YTD:
- puede mostrarse como tercer dato secundario;
- especialmente útil en ocupación estática.

No usar año contra año por defecto.

---

# 7. Arquitectura actual

```text
input/OCU26_BASE_DATOS.xlsx
        │
        ▼
scripts/validate_input.py
Gate 1
        │
        ▼
scripts/transform_data.py
Gate 2
        │
        ▼
scripts/semantic_model.py
+ config/business_semantics.json
Gate 3
        │
        ▼
scripts/metrics_engine.py
Gate 3
        │
        ▼
scripts/export_data.py
Gate 4A
        │
        ├── output/dim_elementos.parquet
        ├── output/fact_campanas.parquet
        ├── output/dim_calendario.parquet
        ├── output/bridge_campana_dia.parquet
        ├── output/fact_metricas_diaria.parquet
        └── output/_export_manifest.json
        │
        ├───────────────► Power BI spec (Gate 4B, pausado)
        │
        └───────────────► HTML / 6 TVs (prioridad actual)
```

## Principio

Cada capa consume la anterior.
Ninguna capa posterior vuelve a reinterpretar la base desde cero.

---

# 8. Arquitecturas anteriores o descartadas

## 8.1 Base Excel pesada con fórmulas

**Considerada porque:** era el sistema histórico.  
**Problema:** fórmulas, cascadas, crecimiento, mantenimiento, riesgo de lentitud.  
**Estado:** histórico; no volver como arquitectura central.

## 8.2 Forms + Power Automate + Excel + Power BI como único sistema

**Considerada:** primera etapa Microsoft 365.  
**Ventaja:** simple y corporativo.  
**Cambio:** Power BI dejó de ser obligatorio para TVs.

## 8.3 Power BI “Publicar en web” como salida primaria

**Considerada:** URL pública.  
**Problema:** depende de tenant/permiso y no se quiere pagar licencias.  
**Estado:** no usar como requisito.

## 8.4 Power BI local dependiente de PC

**Descartado para TVs:** se busca no depender de PC encendida.

## 8.5 Lógica en cada HTML

**Descartado:** genera divergencias.

## 8.6 Office Script en etapa inmediata

El usuario pidió evitar complejidad temprana.
Solo se reconsidera en etapa futura si aporta valor real.

---

# 9. Bases de datos

## 9.1 `OCU26_BASE_DATOS.xlsx`

**Función:** input productivo local actual.  
**Ruta:** `input/OCU26_BASE_DATOS.xlsx`.  
**Estado:** vigente.  
**Conteos Gate 2:**  
- maestro: 4.338;
- campañas: 9.503;
- parámetros: 23.

**SHA-256:**

`2f165e12ad90c2f05963367a8e1717dcd1006e9ce669f976af6e57470aca2cd`

Debe permanecer idéntico antes/después del pipeline.

## 9.2 V3 auditada histórica

`Base_ocupacion_26_FINAL_CON_YPF_AUDITADA_V3.xlsx`

Auditoría 5/8/2026:
- V3 válida estructuralmente;
- no fue necesaria V4;
- incorporación YPF preservada;
- pendientes externos: Excel Online, SharePoint, Power Automate, Power BI, concurrencia y rendimiento.

Incorporación YPF verificada en esa auditoría:
- 1.583 altas al maestro;
- 7.758 activaciones YPF;
- 32 registros `YA_CARGADOS` excluidos;
- 0 duplicados nuevos por YPF;
- 0 campañas con ElementoID inexistente.

## 9.3 Bases históricas mencionadas

- `Final Base_ocupacion_26__4-8_CON_YPF.xlsx`
- `Final Base_ocupacion_26_ 4-8.xlsx`
- `YPF - Base campañas y elementos corregido estructural.xlsx`
- `OCU26_FORM_CARGA_LIVIANO(1).xlsx`
- `BPBI_OCU26_V2_REPARADO_DESPLEGABLES_OK.xlsx`
- `Base_ocupacion_26__4-8.xlsx`
- Base YPF fuente original.

No todas alimentan runtime actual.

---

# 10. Tablas, hojas y columnas

## 10.1 Legacy V3

Hojas principales:
- `BASE_MAESTRA_ELEMENTOS`
- `BASE_CAMPAÑAS`
- `CONTROL_DISPONIBILIDAD`
- `AUX_CASCADA`

## 10.2 Base plana propuesta en Microsoft 365

- `MAESTRO_ELEMENTOS`
- `CAMPANAS`
- `PARAMETROS`

Tablas:
- `tblElementos`
- `tblCampanas`

## 10.3 Columnas de inventario relevantes

Nombres literales observados/confirmados:

- `ElementoID`
- `Ciudad`
- `Medio`
- `CircuitoDashboard`
- `Subcircuito`
- `Ubicacion`
- `Nivel`
- `TipoInventario`
- `TipoInstalacion`
- `TipoCatalogo`
- `Material`
- `AplicaCantidad`
- `CapacidadSlotsReel`
- `SegundosDia`
- `DimensionOptico`
- `DimensionTotal`
- `Observaciones`
- `b`
- `h`
- `q`
- `m2`

## 10.4 Columnas de campañas relevantes

- `CargaID`
- `ClaveNegocio`
- `FechaHoraCarga`
- `UsuarioCarga`
- `FuenteCarga`
- `ElementoID`
- `IDCampaña`
- `Campaña`
- `Cliente`
- `Marca`
- `Agencia`
- `Proveedor`
- `FechaInicio`
- `FechaFin`
- `FechaIndefinida`
- `Estado`
- `DuracionSpotSeg`
- `SalidasVendidas`
- `CantidadUnidades`
- `ModalidadPauta`
- `PROGRAMATICA`
- `CANJE`
- `TipoExclusividad`
- `HoraInicio`
- `HoraFin`
- `Observaciones`

**Nota:** consultar `export_data.py` para el listado completo vigente. No inventar columnas faltantes.

## 10.5 Dimensiones semánticas nuevas

Entre otras:
- `CircuitoNegocio`
- `SitioNegocio`
- `FormatoNegocio`
- `CoberturaCatalogo`
- `CompletitudMaestro`
- `CertezaDato`
- `ModoDisponibilidad`
- `PortfolioTier`
- `IncluyePerformanceCore`
- `IncluyeConteoGeneral`
- `VisiblePorDefecto`
- `TieneActividadComercial`
- `CantidadCampanasHistoricas`
- `FechaPrimeraCampana`
- `FechaUltimaCampana`
- `SlotsComerciales`
- `SegundosComerciales`

## 10.6 Gate 4 — columnas paralelas

Para preservar trazabilidad:
- originales `b`, `h`, `m2` siguen como texto;
- se agregaron numéricas nullable:
  - `BValor`
  - `HValor`
  - `M2Valor`

`QValor` **no se creó** porque el significado de `q` no está confirmado.

Los bloqueos de conversión observados eran:
- `No aplica`
- `No`
- espacio no separable `\xa0`

No se convierten a 0; pasan a NA en el campo numérico.

---

# 11. Identificadores

## `ElementoID`
Clave estable del inventario.

Debe:
- identificar un registro lógico;
- vincular maestro con campañas;
- no cambiar sin migración;
- soportar crecimiento.

Advertencia histórica:
- existían 22 `ElementoID` duplicados únicos en V3 legacy;
- fueron conservados históricamente y bloqueados para nuevas selecciones.

## `CargaID`
Identifica una fila/carga de campaña.

Histórico:
- formato tipo `HIST-00000001`.

Se usa en:
- `FACT_CAMPANAS`;
- `BRIDGE_CAMPANA_DIA`.

## `ClaveNegocio`
Clave estable/compuesta para prevenir duplicados exactos.

Diseño histórico:
- IDCampaña;
- ElementoID;
- fechas;
- horarios.

## `IDCampaña`
Identificador comercial de campaña.
Puede repetirse en varias filas si una campaña está asociada a varios elementos.

## YPF `ElementoID`
Derivado de:
- APIE;
- formato;
- secuencia.

---

# 12. Archivos del proyecto

| Archivo | Tipo | Función | Estado | Versión / vigencia | Dependencias | Observaciones |
|---|---|---|---|---|---|---|
| `input/OCU26_BASE_DATOS.xlsx` | XLSX | Input productivo | Activo | Vigente | Gate1 | SHA protegido |
| `scripts/validate_input.py` | Python | Gate 1 | Activo | Congelado | input | No modificar sin gate explícito |
| `scripts/transform_data.py` | Python | Gate 2 | Activo | Congelado | Gate1 | Read-only |
| `scripts/semantic_model.py` | Python | Gate 3B | Activo | Vigente | Gate2 + config | |
| `scripts/metrics_engine.py` | Python | Motor métricas | Activo | Vigente | semantic model | |
| `scripts/export_data.py` | Python | Gate 4A | Activo | Vigente | Gates1-3 | data mart |
| `config/business_semantics.json` | JSON | Reglas semánticas | Activo | Vigente | semantic model | |
| `tests/test_validate_input.py` | Python | Tests Gate1 | Activo | Vigente | | |
| `tests/test_transform_data.py` | Python | Tests Gate2 | Activo | Vigente | | |
| `tests/test_semantic_model.py` | Python | Tests Gate3 | Activo | Vigente | | |
| `tests/test_metrics_engine.py` | Python | Tests motor | Activo | Vigente | | |
| `tests/test_export_data.py` | Python | Tests Gate4A | Activo | Vigente | | |
| `docs/GATE3_SEMANTICA_NEGOCIO_PROPUESTA.md` | MD | Gate3A | Referencia | histórica | | |
| `.gitignore` | texto | Ignora output | Activo | Vigente | | `output/` |
| `requirements.txt` | texto | Dependencias | Activo | Vigente | | `pyarrow==25.0.0` |
| `powerbi/README.md` | MD | Gate4B | Activo | Vigente | | |
| `powerbi/model/relaciones.md` | MD | Relaciones | Activo | Vigente | | |
| `powerbi/validation/validate_gate4b.py` | Python | Validación PB | Activo | Vigente | outputs | 8/8 |
| `powerbi/validation/README.md` | MD | Resultados PB | Activo | Vigente | | |
| `powerbi/power_query/*.m` | M | Power Query | Activo | Vigente | Parquet | 6 archivos |
| `powerbi/dax/*` | DAX | Medidas | Activo | Vigente | modelo | 7 archivos |
| `output/dim_elementos.parquet` | Parquet | Dimensión | Generado | Regenerable | export | 4.338 × 53 |
| `output/fact_campanas.parquet` | Parquet | Fact campañas | Generado | Regenerable | export | 9.503 × 30 |
| `output/dim_calendario.parquet` | Parquet | Calendario | Generado | Regenerable | export | 1.186 × 10 |
| `output/bridge_campana_dia.parquet` | Parquet | Bridge | Generado | Regenerable | export | 881.210 × 2 |
| `output/fact_metricas_diaria.parquet` | Parquet | Fact diaria | Generado | Regenerable | export | 573.675 × 6 |
| `output/_export_manifest.json` | JSON | Manifest | Generado | Regenerable | export | SHA match |
| `TV1(1).html` | HTML | referencia visual TV1 | Referencia | 9/8 | | no rediseñar |
| `TV2_digital_core_(1).html` | HTML | referencia visual TV2 | Referencia | 9/8 | | |
| `TV3_fijo_core_(1).html` | HTML | referencia visual TV3 | Referencia | 9/8 | | |
| `TV4_pipeline_(1).html` | HTML | pipeline | Referencia | 9/8 | | reutilizar TV6 |
| `TV5_demanda_(1).html` | HTML | demanda | Referencia | 9/8 | | TV5 |
| `OCU26_Etapa_1_Sistema_Simple.docx` | DOCX | Arquitectura M365 inicial | Referencia | 5/8 | | |
| `Auditoria_Independiente_OCU26_V3.pdf` | PDF | Auditoría V3 | Referencia | 5/8 | | |

---

# 13. Estructura de carpetas

Reconstrucción conocida:

```text
ocu26-dashboard/
├── .git/
├── .gitignore
├── requirements.txt
├── input/
│   └── OCU26_BASE_DATOS.xlsx
├── audit_sources/
│   ├── [archivos históricos/auditoría]
│   └── INFORME_RECONCILIACION_MIGRACION.md
├── config/
│   └── business_semantics.json
├── docs/
│   └── GATE3_SEMANTICA_NEGOCIO_PROPUESTA.md
├── scripts/
│   ├── validate_input.py
│   ├── transform_data.py
│   ├── semantic_model.py
│   ├── metrics_engine.py
│   └── export_data.py
├── tests/
│   ├── test_validate_input.py
│   ├── test_transform_data.py
│   ├── test_semantic_model.py
│   ├── test_metrics_engine.py
│   └── test_export_data.py
├── output/
│   ├── dim_elementos.parquet
│   ├── fact_campanas.parquet
│   ├── dim_calendario.parquet
│   ├── bridge_campana_dia.parquet
│   ├── fact_metricas_diaria.parquet
│   └── _export_manifest.json
└── powerbi/
    ├── README.md
    ├── power_query/
    │   ├── 00_parametro_pRutaOutput.m
    │   └── [5 consultas de tablas]
    ├── dax/
    │   └── [7 archivos]
    ├── model/
    │   └── relaciones.md
    └── validation/
        ├── validate_gate4b.py
        └── README.md
```

**A VALIDAR:** nombres exactos de los 5 `.m` de tablas y de los 7 archivos DAX desde el repo, no inventarlos.

---

# 14. Código y scripts

## `scripts/validate_input.py`

**Ruta:** `scripts/validate_input.py`  
**Función:** Gate 1, validar estructura, dominios, fechas e integridad.  
**Input:** Excel vigente.  
**Output:** resultado de validación.  
**Escritura:** ninguna.  
**Estado:** congelado/validado.

## `scripts/transform_data.py`

**Ruta:** `scripts/transform_data.py`  
**Función:**
- llama a Gate1;
- transforma;
- reconstruye reglas heredadas necesarias;
- normaliza;
- verifica SHA.

**No escribe Excel.**

## `scripts/semantic_model.py`

**Gate:** 3B.

Responsabilidad:
1. cargar `business_semantics.json`;
2. resolver jerarquía:
   1. override ElementoID;
   2. CircuitoDashboard + Subcircuito;
   3. CircuitoDashboard;
   4. generic rules;
   5. defaults;
3. enriquecer maestro sin pisar originales;
4. enriquecer campañas.

## `scripts/metrics_engine.py`

API conceptual:

```python
MetricsEngine.query(
    metric,
    group_by,
    filters,
    universe,
    start_date,
    end_date
)
```

Reglas:
- genérico;
- sin funciones por circuito;
- MetricStatus;
- políticas por config;
- error explícito en casos no soportados.

## `scripts/export_data.py`

Gate 4A.

Genera:
- DIM_ELEMENTOS;
- FACT_CAMPANAS;
- DIM_CALENDARIO;
- BRIDGE_CAMPANA_DIA;
- FACT_METRICAS_DIARIA;
- manifest.

Funciones relevantes:
- carga pipeline;
- sanea columnas mixtas para Parquet;
- preserva originales;
- agrega campos numéricos paralelos;
- exporta determinísticamente.

Función agregada:
`_sanear_columnas_texto_mixto()`

Objetivo:
- detectar `object` con tipos Python mezclados;
- convertir a string nullable;
- preservar `pd.NA`;
- evitar fallas de Parquet;
- no tocar Gate1–3.

---

# 15. Fórmulas y cálculos

## Ocupación estática

Para elemento `e` y período `[Pini, Pfin]`:

- considerar campañas que bloquean;
- solapamiento inclusivo;
- contar días únicos ocupados por elemento;
- no sumar doble si dos campañas coinciden en el mismo elemento/día para la métrica binaria de ocupado.

Conceptualmente:

```text
elemento_dias_posibles = cantidad_elementos × cantidad_dias_periodo
elemento_dias_ocupados = suma de elemento-día con actividad bloqueante
ocupacion_pct = ocupados / posibles
```

Solo si el universo es apto.

## Actividad sobre registrados

Métrica transversal útil especialmente para YPF:

```text
actividad_sobre_registrados_pct =
elementos_registrados_con_actividad_periodo
/
elementos_registrados
```

No equivale a fill rate.

## Digital

```text
segundos_vendidos_dia =
SalidasVendidas × 1800
```

solo para spot base 10 s soportado.

## YTD estático

Debe calcularse sobre los días transcurridos del año/período anual definido, no como promedio ingenuo de porcentajes mensuales.

**A VALIDAR en implementación de dashboard:** fórmula exacta deseada de YTD visual. Recomendación coherente con el motor:
- ocupación acumulada por elemento-día desde 1/1 al cierre del mes seleccionado.

## Comparación mensual

Cantidades:
```text
delta_abs = mes_actual - mes_anterior
```

Porcentajes:
```text
delta_pp = pct_actual - pct_mes_anterior
```

---

# 16. Validaciones

### Integridad de archivos
- SHA Excel antes/después: PASS.

### Estructura
- Gate1: PASS con warnings documentados.

### Columnas
- export Gate4: PASS.

### Tipos de datos
- Parquet: nullable `Int64`, `Float64`, `boolean`;
- fechas: `datetime64[us]`;
- texto: string;
- sin columnas object mixtas en output final.

### IDs
- integridad maestro/campañas validada.
- legacy duplicados conocidos no deben entrar en nueva carga.

### Relaciones
Gate4B: 5/5.

### Reglas de negocio
Gate3B tests.

### Outputs
Conteos exactos verificados.

### Visualizaciones
HTML antiguos funcionan como referencia visual.
Los nuevos datos todavía no están terminados.

### Publicación
Pendiente.

### Automatizaciones
Pendiente.

---

# 17. Tests

## Estado final

**199 passed, 0 failed.**

Evolución:
- Gates1–3: 150;
- Gate4A inicial llevó la suite a 191;
- mejoras numéricas y tests adicionales: 199.

Gate4B:
- 8/8 validaciones específicas.

## Coberturas relevantes

Tests incluyen:
- fecha inicio=fin;
- `Reservada`;
- `Cancelado`;
- fechas indefinidas;
- Puente 13;
- Pantalla 20;
- YPF policy;
- pure capacity;
- SalidasVendidas vacía;
- unsupported spot/exclusividad;
- resolution order real;
- mixed groups;
- blocked YPF antes de actividad;
- SHA;
- PK;
- FK;
- bridge temporal;
- MetricStatus;
- Parquet round-trip;
- idempotencia;
- tipos nullable;
- BValor/HValor/M2Valor;
- `No aplica` / NBSP → NA y nunca 0;
- originales de texto preservados;
- `q` sin QValor.

---

# 18. Herramientas utilizadas

## ChatGPT
- diseño de arquitectura;
- reglas;
- prompts;
- revisión de Claude;
- decisiones de producto;
- explicación paso a paso;
- diseño de dashboards.

## Claude / Claude Code
- lectura/auditoría de repo;
- implementación de Gates;
- tests;
- validación;
- commits;
- scripts Power BI;
- export.

## Excel
Fuente histórica / base.

## Python
Pipeline central.

## Pandas
DataFrames.

## PyArrow
Parquet.

## Power BI Desktop
Segunda salida analítica; implementación pausada.

## JavaScript / HTML / CSS
Wallboards 6 TVs.

## Chart.js
Utilizado en HTML de referencia para gráficos.

## Git
Versionado.

## GitHub
Remote.

## GitHub Desktop
Push cuando CLI tuvo problemas de autenticación.

## GitHub Pages
Hosting gratuito planificado.

## SharePoint / Microsoft 365
Etapa operativa futura/inicial.

## Microsoft Forms
Carga futura/Etapa 1.

## Power Automate
Un único escritor de la base futura.

## CMS / Fire Stick
Consumo de URLs en TVs.

---

# 19. Restricciones de herramientas y licencias

- No pagar licencias adicionales para esta etapa.
- Power BI Free no resuelve compartición privada normal a terceros.
- “Publicar en web” depende del tenant.
- Datos de estos tableros pueden ser públicos.
- No depender de PC encendida.
- GitHub gratuito disponible.
- Claude pago disponible.
- M365 disponible.
- SharePoint disponible.
- Inicialmente la carga puede ser manual/cada 1–2 días.
- Automatización completa puede venir después.
- En Power Automate con Excel:
  - ideal único escritor;
  - concurrencia 1.

---

# 20. Git / GitHub

## Repositorio

`ocu26-dashboard`

Ruta local:
`C:\brand plus\ocu26-dashboard`

Branch:
`main`

## Commit Gate 3B

Hash:
`610d5b981fef011bd2e3ab991ba6ec828cf4faf2`

Mensaje:
`feat: add Gate 3B semantic model and metrics engine`

## Commit Gate 4A

Hash:
`7b54220186031f695f5edbf9b611fbf74d964a67`

Mensaje:
`feat: add Gate 4A Power BI data mart export layer`

Archivos:
- `.gitignore`
- `requirements.txt`
- `scripts/export_data.py`
- `tests/test_export_data.py`

## Commit Gate 4B

Hash:
`c8c1dbd0d40cf80db2be8dbc66fbc7546f30772f`

Mensaje:
`feat: add Gate 4B Power BI model specification`

17 archivos bajo `powerbi/`:
- power_query: 6;
- dax: 7;
- model: 1;
- validation: 2;
- README raíz: 1.

## Estado final Git

Tras GitHub Desktop:
- `Fetch origin`;
- branch `main`;
- `0 changed files`;
- `No local changes`.

Gate4B sincronizado a origin.

## Flujo de permisos Claude

- read-only: permitir una vez;
- tests con `PYTHONDONTWRITEBYTECODE=1` y `-p no:cacheprovider`;
- `git add` solo archivos explícitos;
- no `git add .`;
- commit solo luego de auditoría;
- push preferentemente GitHub Desktop si CLI dispara autenticación inestable.

---

# 21. SharePoint / Microsoft 365

## Arquitectura documentada Etapa 1

Sitio:
Brand Plus – OCU26 Datos.

Estructura planificada:

```text
01_DATOS/
    OCU26_BASE_DATOS.xlsx
02_POWER_BI/
    OCU26_TABLEROS.pbix
03_HISTORICOS/
    bases originales / YPF
04_DOCUMENTACION/
    diccionario / reglas / cambios
05_HTML/
    salidas TVs
```

## Estado

- sitio propio ya creado;
- base final online pendiente;
- listas SharePoint postergadas para segunda etapa;
- primera etapa buscaba Excel plano, un único flujo y Forms;
- la arquitectura Python/GitHub se desarrolló en paralelo y actualmente resuelve la urgencia de dashboards.

## Futuro

Migrar:
- `tblElementos` → lista;
- `tblCampanas` → lista;
sin cambiar lógica de negocio.

---

# 22. Power BI

## Rol

Herramienta analítica interna y de validación.
No es dependencia de las TVs.

## Gate4B

### Tablas
1. `DIM_ELEMENTOS`
2. `FACT_CAMPANAS`
3. `DIM_CALENDARIO`
4. `BRIDGE_CAMPANA_DIA`
5. `FACT_METRICAS_DIARIA`

### Relaciones

1. `DIM_ELEMENTOS[ElementoID]` → `FACT_CAMPANAS[ElementoID]`
2. `DIM_ELEMENTOS[ElementoID]` → `FACT_METRICAS_DIARIA[ElementoID]`
3. `DIM_CALENDARIO[Fecha]` → `FACT_METRICAS_DIARIA[Fecha]`
4. `DIM_CALENDARIO[Fecha]` → `BRIDGE_CAMPANA_DIA[Fecha]`
5. `FACT_CAMPANAS[CargaID]` → `BRIDGE_CAMPANA_DIA[CargaID]`

Todas:
- 1:N;
- activas;
- single direction;
- sin bidireccionales;
- sin many-to-many ambiguo;
- bridge como hoja/leaf.

### Power Query

- parámetro `pRutaOutput`;
- 5 consultas;
- `Parquet.Document(File.Contents(...))`.

Valor local usado en prueba:
`C:\brand plus\ocu26-dashboard\output`

### DAX

Aproximadamente 45 medidas en 7 archivos:
- universos;
- inventario;
- comercial;
- calendario;
- digital;
- temporal / time intelligence / RANKX;
- MetricStatus.

Regla:
- no `IF CircuitoNegocio = ...` hardcodeado;
- usar flags/campos resueltos.

## Implementación manual del 9/8

Se abrió Power BI Desktop “Sin título”.

Se creó:
`pRutaOutput = C:\brand plus\ocu26-dashboard\output`

Consultas:
- DIM_ELEMENTOS;
- FACT_CAMPANAS;
- DIM_CALENDARIO;
- BRIDGE_CAMPANA_DIA;
- FACT_METRICAS_DIARIA.

Verificación visual final:
- DIM_ELEMENTOS = 53 columnas;
- FACT_CAMPANAS = 30;
- DIM_CALENDARIO = 10;
- BRIDGE_CAMPANA_DIA = 2;
- FACT_METRICAS_DIARIA = 6.

Al `Cerrar y aplicar` apareció:
> `El argumento 'dataType' no puede ser nulo. Nombre del parámetro: dataType`

Se sospechó `FechaHoraCarga` completamente nula / tipo indefinido, pero **no se verificó definitivamente**.

La cancelación quedó trabada en:
- DIM_ELEMENTOS → Cancelando…
- FACT_CAMPANAS → Cancelando…

Decisión:
- forzar cierre si era necesario;
- pausar Power BI;
- retomar martes 11/8 con revisión de tipos de una sola vez.

No hacer troubleshooting de Power BI antes de terminar los HTML.

---

# 23. HTML / dashboards / salidas

## 23.1 Diseño visual obligatorio

Los HTML del 9/8 comparten:

### Paleta

```css
--bp-blue:#1C60FF;
--bp-blue-dark:#0B1D4D;
--bp-blue-mid:#123EAA;
--bp-blue-soft:#7EA6FF;
--bp-black:#1D1D1B;
--bp-black-deep:#05070D;
--bp-navy:#071124;
--bp-cream:#EEECE6;
--bp-white:#FFFFFF;
--bp-muted:#A7B0C2;
--bp-coral:#FF4E46;
--bp-warning:#FFC857;
--bp-success:#2FE084;
```

Panels:
```css
--panel:rgba(255,255,255,.055);
--panel-2:rgba(255,255,255,.09);
--panel-border:rgba(255,255,255,.12);
--panel-border-2:rgba(126,166,255,.28);
```

### Tipografía
`Poppins`

Pesos:
300, 400, 500, 600, 700, 800.

### Stage
1920×1080.

Escalado:
```javascript
k = min(innerWidth/1920, innerHeight/1080)
```

### Fondo
- azul/navy/negro;
- gradientes radiales;
- textura tipo LED muy sutil;
- glow inferior.

### Componentes
- header;
- logo;
- eyebrow con número de TV;
- título;
- subtítulo;
- reloj;
- fecha actualización;
- pregunta guía;
- 5 KPI cards;
- panels;
- rankings;
- chips;
- charts;
- bloque `Lectura`;
- watermark sutil.

### Regla
**NO REDISEÑAR.**

## 23.2 HTML subidos y rol futuro

### `TV1(1).html`
Actual:
- “Resumen Ejecutivo General”.
- pregunta: “¿Cómo está el negocio hoy, este mes y este año?”
- evolución mensual.
- mix.
- debe actualizarse a nueva lógica ejecutiva.

### `TV2_digital_core_(1).html`
Actual:
- “Así rinde el digital core”.
- Pantallas LED + Shoppings Digital.
- rankings;
- fill rate + ocupación calendario.
- se mantiene como base de TV2.
- excluir YPF.
- incorporar AA2000 digital cuando aplique, sin falsear completitud.

### `TV3_fijo_core_(1).html`
Actual:
- “Así rinde el fijo core”.
- shopping fijo;
- ocupación calendario;
- rankings.
- se mantiene como base TV3.
- excluir YPF/APSA/London.
- incluir Cencomedia.

### `TV4_pipeline_(1).html`
Actual:
- pipeline;
- campañas/reservas;
- próximas activaciones.
- debe pasar conceptualmente a **TV6**.

### `TV5_demanda_(1).html`
Actual:
- clientes;
- marcas;
- agencias;
- demanda por circuito.
- debe evolucionar a TV5 Performance / demanda.

### TV4 nueva — YPF
No existe todavía un HTML específico final.
Debe construirse reutilizando exactamente el mismo sistema visual.

### TV6
No se subió un HTML TV6 final en este corte.
Se reutiliza la estructura del actual TV4 Pipeline.

## 23.3 Nueva arquitectura TV1–TV6

### TV1 — Resumen ejecutivo del negocio

Pregunta:
> ¿Qué tenemos, qué parte estamos trabajando y cómo rindió el mes?

Debe ordenar:
1. catálogo visible;
2. core / universo comercial relevante;
3. actividad del mes;
4. performance digital;
5. performance estática.

Nunca mezclar fill rate YPF con fill rate gestionable.

Métrica transversal posible:
- elementos con actividad / elementos registrados del scope.

Comparación:
- siempre mes anterior.

YTD:
- donde aporte valor.

### TV2 — Digital

Scope:
- sin YPF;
- sin APSA;
- sin London;
- Pantallas LED;
- Shoppings Digital;
- AA2000 digital, con status si parcial.

Mostrar:
- fill rate;
- capacidad;
- ranking;
- slots;
- evolución;
- mes vs anterior.

### TV3 — Estático

Scope:
- Shoppings Fijo/Cencosud;
- Remeros;
- Pilar Frontlight;
- AA2000 Estático;
- Cencomedia;
- sin YPF;
- sin APSA;
- sin London.

Mostrar:
- elementos;
- ocupación mensual;
- mes anterior;
- YTD 2026;
- disponibilidad;
- ranking;
- soportes/ubicaciones.

### TV4 — YPF

Todo YPF junto:
- MB;
- TT;
- PPUNTER;
- FB.

Mostrar:
- estaciones/APIE;
- elementos registrados;
- elementos con actividad;
- campañas del mes;
- marcas/clientes;
- formato;
- geografía;
- evolución mensual.

No mostrar fill rate total de slots como si Brand Plus conociera el CMS.

### TV5 — Performance / demanda

Pregunta:
> ¿Qué se vende, qué elementos funcionan y quién compra?

Mostrar:
- clientes;
- marcas;
- campañas;
- activaciones;
- top elementos;
- top ubicaciones;
- demanda por circuito;
- agencia si el dato aporta.

### TV6 — Pipeline / proyección

Pregunta:
> ¿Qué está corriendo y qué viene?

Mostrar:
- campañas activas;
- reservas;
- próximos inicios;
- próximos finales;
- pipeline;
- próximos 30 días;
- proyección basada solo en campañas/reservas cargadas.

No hacer forecast inventado.

---

# 24. Automatización

## Automatización existente

Pipeline local:
```text
Excel → Python → Parquet
```

Git versionado.

## Automatización planificada

```text
Excel actualizado
↓
Python / CI
↓
data/output
↓
HTML
↓
GitHub Pages
```

La idea futura:
- reemplazar Excel por otro con mismo nombre y estructura;
- regenerar outputs;
- regenerar HTML.

## Automatización descartada por ahora

- múltiples Power Automate;
- Office Scripts complejos;
- dependencia de Power BI Service para TVs.

---

# 25. Referencias del chat de Claude e imágenes

## 25.1 Qué trabajo se hizo en Claude

Claude/Claude Code fue utilizado para:

- inspeccionar el repo;
- diseñar Gate4A;
- implementar `export_data.py`;
- escribir tests;
- auditar tipos mixtos;
- evaluar conversión de `b`, `h`, `q`, `m2`;
- corregir `.gitignore`;
- validar outputs;
- construir Gate4B;
- crear Power Query;
- crear DAX;
- crear documentación de relaciones;
- crear validaciones;
- ejecutar tests;
- preparar commits.

## 25.2 Qué información pasó desde Claude a este proyecto

Resultados críticos:
- 199 tests;
- outputs Parquet;
- SHA;
- 17 archivos Gate4B;
- 5 relaciones;
- tipos;
- compatibilidad Power BI Desktop Free;
- estado Git limpio.

## 25.3 Qué imágenes/capturas se cargaron

### Capturas de permisos Claude Code
Mostraban comandos:
- pytest;
- scripts scratchpad;
- git diff;
- git add;
- git commit.

Se utilizaron para decidir:
- permitir una vez;
- denegar cuando un comando podía escribir algo no previsto;
- evitar pytest cache;
- evitar bytecode.

### Captura Gate4A resultado
Mostraba:
- 191 tests inicialmente;
- output counts;
- Git diff;
- SHA.

Luego se amplió a 199 tests.

### Capturas auditoría b/h/q/m2
Mostraron:
- b 96,11% convertible;
- h 96,26%;
- q 86,77%;
- m2 96,34%.

Valores bloqueantes:
- `No aplica`;
- `No`;
- `\xa0`.

Decisión:
- crear BValor/HValor/M2Valor;
- no QValor.

### Capturas GitHub Desktop
Confirmaron:
- main;
- Fetch origin;
- 0 changes;
- pushes completados.

### Capturas Power BI
Mostraron:
- creación de parámetro;
- consultas;
- conteos de columnas;
- error `dataType`;
- cancelación trabada.

## 25.4 Versiones visuales

HTML de agosto 2026:
- ya presentan diseño Brand Plus consolidado.
- No deben reemplazarse por un rediseño.

La evolución futura es **de contenido**, no de identidad.

## 25.5 Información visual que no debe perderse

- fondo azul/navy/negro;
- `Poppins`;
- cards con glass effect;
- acentos verticales;
- azul primario;
- green/warning/coral solo para estados;
- logo a la izquierda;
- reloj/actualización a la derecha;
- 5 KPI cards;
- body 2 columnas;
- bloque `Lectura`;
- 1920×1080.

## 25.6 Imágenes que deben volver a cargarse

| Imagen a recuperar | Qué contiene | Por qué es necesaria | Decisiones asociadas | Prioridad |
|---|---|---|---|---|
| Render/captura TV1, si existe | Apariencia | Control visual | diseño no tocar | IMPORTANTE |
| Render/captura TV2 | Apariencia | Ranking digital | diseño | IMPORTANTE |
| Render/captura TV3 | Apariencia | Fijo | diseño | IMPORTANTE |
| Render/captura TV4 pipeline | Apariencia | Reutilizar TV6 | diseño | IMPORTANTE |
| Render/captura TV5 demanda | Apariencia | Reutilizar TV5 | diseño | IMPORTANTE |
| Error Power BI `dataType` | Mensaje | Retomar martes | troubleshooting | REFERENCIA |
| GitHub Desktop sincronizado | Repo limpio | Continuidad | commits | REFERENCIA |

Los HTML son más importantes que las capturas; con los HTML disponibles, las capturas no son críticas.

---

# 26. Otros archivos externos que deberían recuperarse

Prioridad alta:

1. `CONTEXTO_MAESTRO.md`
2. `input/OCU26_BASE_DATOS.xlsx`
3. `config/business_semantics.json`
4. `scripts/semantic_model.py`
5. `scripts/metrics_engine.py`
6. `scripts/export_data.py`
7. tests asociados
8. los cinco HTML del 9/8
9. `powerbi/` si se retoma Power BI

Documentales:
- `GATE3_SEMANTICA_NEGOCIO_PROPUESTA.md`
- `Auditoria_Independiente_OCU26_V3.pdf`
- `OCU26_Etapa_1_Sistema_Simple.docx`
- manual de marca Brand Plus si se va a cambiar diseño.

---

# 27. Problemas encontrados

## Problema: Excel legacy no escalaba bien
**Contexto:** crecimiento YPF.  
**Impacto:** fórmulas/cascadas pesadas.  
**Solución:** pipeline plano + Python.  
**Estado:** resuelto arquitectónicamente.

## Problema: límites $4000
**Contexto:** V3.  
**Causa:** rangos legacy.  
**Solución:** rangos ampliados.  
**Estado:** resuelto.

## Problema: YPF y métricas CMS
**Causa:** Brand Plus no gestiona CMS.  
**Solución:** `NO_APLICA` a fill rate real; actividad sobre registrados.  
**Estado:** resuelto.

## Problema: AA2000 maestro parcial
**Solución:** separar cobertura del negocio de completitud del maestro.  
**Estado:** resuelto semánticamente.

## Problema: Parquet con tipos mixtos
**Causa:** columnas object con str/int/float.  
**Solución:** saneo genérico a string nullable en Gate4.  
**Estado:** resuelto.

## Problema: numéricos guardados como texto
**Causa:** legacy Excel.  
**Solución:** columnas paralelas numéricas.  
**Estado:** resuelto para b/h/m2.

## Problema: Power BI `dataType`
**Estado:** pendiente, pausado.

## Problema: cancelación Power BI trabada
**Estado:** cierre forzado / pausa.

## Problema: Git CLI auth
**Solución:** GitHub Desktop.

## Problema: iteraciones/tokens Claude
**Solución:** prompts maestros más cerrados.

---

# 28. Errores ya cometidos

1. Asumir 3 salidas = 7.200 s.  
   Correcto: 5.400.

2. Olvidar Pantalla LED Córdoba en una lista de 10.  
   Correcto: 11 ubicaciones.

3. Tratar ocupación digital como ocupación binaria de calendario.  
   Incorrecto: puede quedar capacidad de slots.

4. Aplicar fill rate de YPF como si se conociera CMS.  
   No hacer.

5. Confundir `TipoCatalogo` con semántica completa.  
   No hacer.

6. Usar London en denominadores ejecutivos.  
   Nueva regla: no.

7. Duplicar consultas Power Query y cambiar solo nombre, no archivo.  
   Ocurrió el 9/8; se corrigió.

8. Dejar placeholder `NOMBRE_ARCHIVO.parquet`.  
   Ocurrió en FACT_CAMPANAS; corregido.

9. Intentar aplicar Power BI sin revisar todos los tipos finales.  
   produjo `dataType`.

10. Ejecutar pytest sin bloquear caché en fase estrictamente read-only.  
    Usar flags.

---

# 29. Soluciones descartadas

## Lógica por circuito con funciones separadas
**Motivo:** rigidez.  
**Reconsiderar:** no salvo caso excepcional demostrado.

## DAX que reinterpreta semántica
**Motivo:** duplicación.  
**Reconsiderar:** no.

## Mega tabla plana única
**Motivo:** modelo difícil de mantener.  
**Reconsiderar:** no para BI principal.

## Duplicar físicamente universos
**Motivo:** redundancia.  
**Reconsiderar:** solo si hay necesidad de performance demostrada.

## Power BI obligatorio para TVs
**Motivo:** costo/licencia/tenant/dependencia.  
**Reconsiderar:** solo si cambia contexto de licencia.

## London en resumen general
**Motivo:** distorsiona.  
**Reconsiderar:** solo por pedido explícito.

## APSA en dashboard estándar
**Motivo:** legacy.  
**Reconsiderar:** solo consulta histórica.

---

# 30. Supuestos

## Supuestos confirmados posteriormente

- Puente LED = 13.
- 72.000 referencia comercial.
- Reservada bloquea.
- Córdoba es la 11ª Pantalla LED.
- Pilar tiene digital + frontlight.
- MAB flexible.

## Supuestos que siguen vigentes

- HTML puede ser público.
- 6 TVs consumen URL.
- no hace falta privacidad.
- actualización inicial puede ser manual.

## Supuestos incorrectos

- 3 salidas → 7.200.
- lista Pantallas LED de 10.
- London útil en denominador general.
- Power BI necesario para poder presentar TVs.

## Supuestos pendientes

- spot >10s;
- exclusividad;
- `q`;
- CantidadUnidades Cencomedia;
- Power BI dataType.

---

# 31. Dependencias

## Archivos
Excel → Gate1 → Gate2 → Gate3 → Gate4.

## Semántica
`business_semantics.json` → semantic model → metrics engine.

## Dashboards
outputs/metric engine → dataset/HTML.

## Publicación
HTML → hosting → URL → CMS → TV.

## Power BI
Parquet → Power Query → modelo → DAX.

## Operación futura
Forms → Power Automate → SharePoint/Excel/Listas.

---

# 32. Riesgos

- cambiar columnas del Excel;
- renombrar `ElementoID`;
- introducir duplicados;
- mover input sin actualizar path;
- meter lógica en HTML;
- usar London/APSA por error;
- calcular YPF fill rate;
- convertir unknown a 0;
- ignorar MetricStatus;
- promediar porcentajes mensuales incorrectamente para YTD;
- hardcodear filas o ubicaciones;
- cambiar branding;
- depender de Google Fonts sin fallback si la red falla;
- hosting no disponible;
- GitHub Pages no configurado;
- Power BI no resuelto;
- Cencomedia sin `CantidadUnidades`;
- master AA2000 parcial.

---

# 33. Datos que NO deben modificarse sin validación

- `input/OCU26_BASE_DATOS.xlsx`
- SHA esperado
- `ElementoID`
- `CargaID`
- `ClaveNegocio`
- `scripts/validate_input.py`
- `scripts/transform_data.py`
- `scripts/semantic_model.py`
- `scripts/metrics_engine.py`
- `config/business_semantics.json`
- capacidad Puente = 13;
- 72.000 s;
- SalidasVendidas mapping;
- políticas YPF;
- status semantics;
- relaciones Gate4B;
- paleta HTML;
- layout 1920×1080;
- exclusión APSA/London actual.

---

# 34. Pendientes completos

## P0 — Críticos / bloqueantes

### Tarea: cerrar especificación TV1–TV6
**Estado:** en curso.  
**Dependencias:** decisiones de negocio.  
**Resultado esperado:** ficha exacta por TV.

### Tarea: implementar 6 HTML
**Estado:** pendiente.  
**Dependencias:** especificación.  
**Resultado esperado:** archivos finales.

### Tarea: validar datos reales
**Estado:** pendiente.  
**Dependencias:** outputs/motor.  
**Resultado esperado:** números no inventados.

### Tarea: probar 1920×1080
**Estado:** pendiente.  
**Resultado esperado:** legible.

### Tarea: tener archivos presentables lunes
**Estado:** pendiente.

## P1 — Alta prioridad

### Hosting
**Estado:** pendiente.

### URLs en CMS
**Estado:** pendiente.

### Exclusión APSA/London
**Estado:** regla definida; implementación dashboard pendiente.

### Cencomedia
**Estado:** incluir TV3 sin métricas falsas.

### TV4 YPF
**Estado:** crear.

## P2 — Importantes

### Power BI
**Estado:** retomar martes.

### SharePoint
**Estado:** continuar.

### Automatización
**Estado:** futura.

### Calidad AA2000
**Estado:** completar maestro cuando haya datos.

## P3 — Mejoras futuras

- Lists;
- Forms productivo;
- Power Automate;
- exclusividad;
- spots no 10s;
- automatización total;
- performance avanzada;
- forecast real si se define metodología.

---

# 35. Roadmap actual

## Domingo 9/8
1. Congelar Power BI.
2. Definir 6 TVs.
3. Cerrar universo y KPIs.
4. Preparar prompt maestro Claude.
5. Implementar HTML.
6. Validar.

## Lunes 10/8
1. Presentar tableros.
2. Recoger feedback.
3. Ajustes mínimos si son necesarios.

## Martes 11/8 en adelante
1. Power BI.
2. resolver dataType.
3. implementar modelo Gate4B.
4. SharePoint.
5. hosting/automatización definitiva.
6. evolución de carga.

---

# 36. Próximo paso exacto

El próximo ChatGPT debe:

1. NO volver a Power BI ahora.
2. Tomar los 5 HTML subidos como baseline visual.
3. Cerrar el contenido de TV1.
4. Continuar TV2–TV6.
5. Verificar que:
   - APSA no aparece;
   - London no aparece;
   - YPF está separado;
   - Cencomedia entra en estático;
   - `vs` = mes anterior;
   - YTD se muestra donde corresponde.
6. Crear un prompt maestro para Claude Code.
7. Pedir a Claude implementar todos los HTML con mínima iteración.
8. Auditar outputs.
9. No tocar Gates1–4 salvo necesidad real.

---

# 37. Punto exacto donde terminó este chat

El 9/8/2026:

1. Gate4B ya estaba committeado y pusheado.
2. Se intentó implementar Power BI manualmente.
3. Power Query llegó a leer correctamente:
   - 53 columnas;
   - 30;
   - 10;
   - 2;
   - 6.
4. `Cerrar y aplicar` lanzó error `dataType`.
5. Cancelar quedó trabado.
6. Se decidió posponer Power BI.
7. Se subieron 5 HTML antiguos.
8. Se confirmó que el diseño se mantiene **tal cual**.
9. Se redefinieron las 6 TVs.
10. Se decidió excluir completamente APSA y London.
11. Se decidió incluir Cencomedia en estático general.
12. Se confirmó que el “vs” es siempre el mes anterior.
13. Se confirmó interés en YTD 2026, especialmente estático.
14. Se pidió generar este `CONTEXTO_MAESTRO.md`.

El siguiente trabajo es definir e implementar los seis HTML.

---

# 38. Preguntas abiertas

1. ¿Cuál será el hosting definitivo inmediato?
2. ¿Cómo se implementará centralmente el nuevo scope “sin APSA ni London” sin duplicar lógica?
3. ¿Cencomedia debe cambiar formalmente de `PortfolioTier` o solo incluirse en TV3?
4. ¿Qué significa exactamente `q`?
5. ¿De dónde se cargará `CantidadUnidades` para Cencomedia?
6. ¿Cómo se tratarán spots >10s?
7. ¿Cuál será la matemática de exclusividad?
8. ¿Cuál es la causa exacta del `dataType` de Power BI?
9. ¿Existe/queda algún PBIX guardado de la prueba? A VALIDAR.
10. ¿Qué campos exactos conformarán cada una de las 5 tarjetas de TV1–TV6? En definición.

---

# 39. Glosario

**OCU26:** proyecto de ocupación 2026.  
**OOH:** Out of Home.  
**DOOH:** Digital Out of Home.  
**Core:** portfolio principal.  
**Complementario:** se comercializa pero no forma parte del core.  
**Legacy:** histórico/no operativo estándar.  
**APIE:** identificador de estación YPF usado en estructuración.  
**MB:** Menu Board.  
**TT:** Torre.  
**PPUNTER:** Puntera.  
**FB:** Fotobox/Mupi.  
**Fill rate:** capacidad digital vendida / total cuando aplica.  
**Elemento-día:** unidad temporal estática.  
**YTD:** acumulado año a la fecha.  
**MetricStatus:** estado de calidad/aplicabilidad de una métrica.  
**Gate:** etapa versionada de desarrollo/validación.  
**Data mart:** capa de outputs para BI/HTML.  
**Bridge:** tabla puente campaña-día.

---

# 40. Personas, equipos, proveedores y actores

## Brand Plus
Dueño del proyecto y reglas de negocio.

## Usuario / responsable del proyecto
Define:
- criterios comerciales;
- qué se muestra;
- prioridades;
- aceptación.

## ChatGPT
Estrategia, documentación, prompts y control funcional.

## Claude Code
Implementación técnica y tests.

## Microsoft
SharePoint, Forms, Power Automate, Power BI.

## GitHub
Versionado / hosting futuro.

## YPF
Circuito externo cuyo CMS no administra Brand Plus.

## AA2000
Circuito aeroportuario core.

## Cencosud
Shopping core.

## London Supply
Inventario fuera de análisis actual.

---

# 41. Cronología relevante

## Antes de agosto 2026
- base histórica con fórmulas;
- OCU26 Power BI/base ocupación;
- incorporación YPF.

## 5/8/2026
- documento Etapa1 M365;
- Forms + Power Automate + Excel plano + Power BI.

## 5–6/8
- auditoría V3;
- V3 aprobada para piloto técnico;
- no V4.

## 7/8
- Gate3A semántica;
- decisiones sobre YPF, London, APSA, Pilar, MAB, 72.000, Puente, etc.

## 8/8
- Gate3B;
- edge cases;
- 150 tests;
- commit y push.

## 8–9/8
- Gate4A;
- Parquet;
- 191 → 199 tests;
- commit/push.

## 9/8
- Gate4B;
- 17 archivos;
- commit/push.
- Power BI intento manual.
- error `dataType`.
- Power BI pausado.
- HTML pasa a prioridad.
- nueva arquitectura de 6 TVs.
- APSA + London totalmente fuera.
- Cencomedia entra en estático.
- comparación mes anterior.
- YTD 2026.

---

# 42. Reglas para el próximo ChatGPT

“Este documento representa el estado acumulado del proyecto al cierre del chat anterior.

El proyecto tuvo múltiples iteraciones y cambios de criterio.

No asumir que una decisión temprana sigue siendo válida cuando existe una decisión posterior.

Antes de modificar la solución:

1. leer este documento completo;
2. identificar las decisiones vigentes;
3. respetar las reglas de negocio;
4. no reconstruir componentes que ya funcionan;
5. no recomendar soluciones previamente descartadas sin una razón nueva;
6. mantener compatibilidad con los archivos y arquitectura actual;
7. verificar las dependencias antes de cambiar nombres, estructuras o IDs;
8. diferenciar claramente una corrección necesaria de un rediseño opcional;
9. utilizar las referencias visuales y archivos indicados cuando sean necesarios;
10. si una decisión depende de una imagen que no está disponible, solicitar específicamente esa imagen en lugar de asumir su contenido;
11. preguntar únicamente cuando la información no pueda resolverse utilizando este documento y los archivos disponibles.”

Instrucciones adicionales específicas OCU26:

12. No introducir APSA ni London en dashboards estándar.
13. No calcular fill rate total YPF.
14. No alterar el diseño visual de los HTML.
15. `vs` = mes anterior.
16. Power BI está pausado hasta después de la presentación del lunes.
17. No usar números de ejemplo como números reales.
18. No reabrir Gate3/4 por preferencias visuales.
19. Mantener MetricStatus.
20. Priorizar entregar las TVs.

---

# 43. Instrucciones para reconstruir el contexto visual

Cuando el usuario vuelva a subir los HTML:

1. comprobar que correspondan a:
   - TV1;
   - TV2 digital;
   - TV3 fijo;
   - TV4 pipeline;
   - TV5 demanda.
2. leer CSS variables;
3. no cambiar paleta;
4. no cambiar Poppins;
5. no cambiar stage;
6. no cambiar proporciones de cards;
7. identificar IDs/data attributes que el JS usa;
8. reemplazar únicamente data bindings y contenido requerido;
9. si TV4 YPF necesita una nueva disposición, reutilizar componentes existentes;
10. si TV4 pipeline se transforma en TV6, conservar su composición visual;
11. comparar render 1920×1080 antes y después;
12. verificar legibilidad a distancia.

Si se vuelve a subir el manual de marca:
- usarlo como validación de la identidad;
- no reinterpretar el diseño sin pedido.

---

# 44. Checklist de migración de archivos

## Mínimo para continuar dashboards

- [ ] `CONTEXTO_MAESTRO.md`
- [ ] `TV1(1).html`
- [ ] `TV2_digital_core_(1).html`
- [ ] `TV3_fijo_core_(1).html`
- [ ] `TV4_pipeline_(1).html`
- [ ] `TV5_demanda_(1).html`
- [ ] acceso al repo `ocu26-dashboard` o archivos de `scripts/`, `config/`, `output/`

## Para lógica

- [ ] `input/OCU26_BASE_DATOS.xlsx`
- [ ] `config/business_semantics.json`
- [ ] `scripts/semantic_model.py`
- [ ] `scripts/metrics_engine.py`
- [ ] `scripts/export_data.py`
- [ ] tests

## Para Power BI más adelante

- [ ] carpeta `powerbi/`
- [ ] captura/error `dataType` si se necesita
- [ ] este contexto

## Referencia

- [ ] `GATE3_SEMANTICA_NEGOCIO_PROPUESTA.md`
- [ ] `Auditoria_Independiente_OCU26_V3.pdf`
- [ ] `OCU26_Etapa_1_Sistema_Simple.docx`
- [ ] manual Brand Plus si habrá cambios visuales

---

# 45. Snapshot ejecutivo

1. OCU26 es el sistema de inventario, ocupación y performance de Brand Plus.
2. Repo: `C:\brand plus\ocu26-dashboard`.
3. Branch `main`.
4. Fuente vigente: `input/OCU26_BASE_DATOS.xlsx`.
5. Maestro: 4.338.
6. Campañas: 9.503.
7. Parámetros: 23.
8. SHA: `2f165e12ad90c2f05963367a8e1717dcd1006e9ce669f976af6e57470aca2cd`.
9. Gate1 terminado.
10. Gate2 terminado.
11. Gate3A terminado.
12. Gate3B terminado.
13. Gate4A terminado.
14. Gate4B terminado.
15. Suite: 199 passed.
16. Gate4B validación: 8/8.
17. Parquet: 5 tablas + manifest.
18. `DIM_ELEMENTOS`: 4.338 × 53.
19. `FACT_CAMPANAS`: 9.503 × 30.
20. `DIM_CALENDARIO`: 1.186 × 10.
21. `BRIDGE_CAMPANA_DIA`: 881.210 × 2.
22. `FACT_METRICAS_DIARIA`: 573.675 × 6.
23. Gate3B commit: `610d5b981fef011bd2e3ab991ba6ec828cf4faf2`.
24. Gate4A commit: `7b54220186031f695f5edbf9b611fbf74d964a67`.
25. Gate4B commit: `c8c1dbd0d40cf80db2be8dbc66fbc7546f30772f`.
26. Todo pusheado a main.
27. Power BI está especificado pero pausado.
28. Power BI dio error `dataType` al aplicar.
29. Se retoma desde martes 11/8.
30. Prioridad: HTML para lunes 10/8.
31. Diseño HTML Brand Plus NO se toca.
32. Paleta base usa `#1C60FF`, navy/negro, Poppins.
33. APSA queda fuera de todo análisis.
34. London Supply queda fuera de todo análisis hasta nuevo aviso.
35. Cencomedia entra en estático general.
36. YPF tiene TV propia.
37. YPF no tiene fill rate real de CMS.
38. TV1 General; TV2 Digital; TV3 Estático; TV4 YPF; TV5 Performance; TV6 Pipeline.
39. Comparación = siempre mes anterior.
40. YTD 2026 se usa especialmente para estático.
41. No convertir desconocidos en cero.
42. No duplicar negocio en DAX/HTML.
43. Claude Code debe recibir prompts cerrados para ahorrar tokens.
44. Hosting público queda pendiente.
45. Próximo paso: cerrar contenido TV1 y luego TV2–TV6.

---

`[CIERRE HISTÓRICO DEL CM1 ANTERIOR — reemplazado por el cierre final de esta consolidación]`


---



---

# CIERRE ACTUAL DE CM1

Al momento de generar `CM1.md`:

- TV1 está prácticamente cerrada.
- No hay que volver a discutir YPF.
- No hay que volver a modificar composición.
- No hay que volver a poner el desglose YPF/Est/Dig en Unidades con Campaña.
- Mañana solo se cambia el copy de Core.
- Luego se pasa a TV2 en una sesión nueva de Claude Code.
- Para TV2–TV6 prima eficiencia: prompt maestro, auditoría agrupada, tests específicos, preview y cierre.


`[CIERRE HISTÓRICO DEL CM1 ANTERIOR — reemplazado por el cierre final de esta consolidación]`


---

# CIERRE MÁS RECIENTE — 9/8/2026 ~22:46 ART

La vigencia operativa se encuentra en la sección `ACTUALIZACIÓN VIGENTE CM1 — 9/8/2026 ~22:46 ART` al comienzo de este archivo. TV2 se considera cerrada para esta noche; el siguiente paso es TV3 en una nueva sesión de Claude Code. Los pendientes TV1 copy y TV2 denominador Shoppings/Remeros quedan explícitamente para mañana.

`[CIERRE HISTÓRICO DEL CM1 ANTERIOR — reemplazado por el cierre final de esta consolidación]`

---

# CIERRE MÁS RECIENTE — 9/8/2026 ~23:27 ART

La vigencia operativa se encuentra en la sección `ACTUALIZACIÓN VIGENTE CM1 — 9/8/2026 ~23:27 ART` al comienzo de este archivo.

Punto actual: TV3 Core Comercial Estático ya está construida y comprendida funcionalmente; queda un ajuste visual menor en la tarjeta `DISPONIBLES`. TV4 será YPF y se deja para el final. El próximo bloque inmediato es TV5. La próxima semana habrá una Etapa 2 de microauditorías y cálculos ampliados para profundizar validaciones no bloqueantes detectadas durante la producción acelerada.

`[CIERRE HISTÓRICO DEL CM1 ANTERIOR — reemplazado por el cierre final de esta consolidación]`

---

# CIERRE MÁS RECIENTE — 10/8/2026 ~00:20 ART

La vigencia operativa se encuentra en `ACTUALIZACIÓN VIGENTE CM1 — 10/8/2026 ~00:20 ART`.

Punto exacto: TV4 Pipeline Comercial ya pasó por definición funcional y auditorías de scope/temporalidad. El builder productivo fue invocado; falta confirmar salida, correr tests específicos TV4, sanity y preview 1920×1080. Después se implementará TV5 YPF y finalmente TV6 Demanda Comercial sobre Core Comercial + YPF. Las microauditorías profundas quedan para Etapa 2 la próxima semana.

`[CIERRE HISTÓRICO DEL CM1 ANTERIOR — reemplazado por el cierre final de esta consolidación]`

---

# CIERRE MÁS RECIENTE — 10/8/2026 ~06:20 ART

La vigencia operativa se encuentra en `ACTUALIZACIÓN VIGENTE CM1 — 10/8/2026 ~06:20 ART`.

Punto exacto: TV4 Pipeline Comercial quedó aprobada con 35/35 tests PASS y preview final. El siguiente bloque es TV6 Demanda Comercial, usando la referencia legacy de demanda únicamente como guía visual. TV6 debe integrar Core Comercial completo + YPF, con gráfico dinámico Marcas → Agencias → Programática; Programática contiene únicamente las agencias que efectivamente hacen programática según la data canónica. El resto de la referencia se replica salvo reglas globales vigentes. Después se implementará TV5 YPF.

`[CIERRE HISTÓRICO DEL CM1 ANTERIOR — reemplazado por el cierre final de esta consolidación]`

---

# CIERRE MÁS RECIENTE — 10/8/2026 ~08:22 ART

La vigencia operativa se encuentra en `ACTUALIZACIÓN VIGENTE CM1 — 10/8/2026 ~08:22 ART`.

Punto exacto: TV6 Demanda Comercial fue reformulada con enfoque híbrido. Las 5 KPI cards superiores siguen siendo JULIO 2026; rankings Marcas/Agencias/Programática y matriz Demanda por circuito pasan a acumulado ENE–JUL 2026; Programática usa únicamente registros `PROGRAMATICA == 'Si'` con agencia identificada real, preservando como pendientes los casos no imputados; el footer pasa a `LECTURA / PUNTO POSITIVO / A ATENDER`. Builder/payload híbridos ya fueron regenerados e inspeccionados. Falta confirmar el resultado final de tests específicos, sanity/Git y revisar la preview híbrida 1920×1080 antes de cerrar TV6. Luego continúa TV5 YPF.

`[CIERRE HISTÓRICO DEL CM1 ANTERIOR — reemplazado por el cierre final de esta consolidación]`

`FIN DEL CONTEXTO MAESTRO — ESTADO DEL PROYECTO AL MOMENTO DE LA MIGRACIÓN`
