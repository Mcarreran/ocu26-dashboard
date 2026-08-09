# Relaciones del modelo (referencia rápida para armarlo en Power BI Desktop)

Crear en Modelado > Administrar relaciones > Nueva. Las 5 son
**single-direction** (filtro `Único`, de "Desde" hacia "Hacia"), **activas**,
cardinalidad `Uno a varios`.

| # | Desde (lado 1) | Columna | Hacia (lado N) | Columna | Cross-filter |
|---|---|---|---|---|---|
| R1 | DIM_ELEMENTOS | ElementoID | FACT_CAMPANAS | ElementoID | Único (DIM→FACT) |
| R2 | DIM_ELEMENTOS | ElementoID | FACT_METRICAS_DIARIA | ElementoID | Único (DIM→FACT) |
| R3 | DIM_CALENDARIO | Fecha | FACT_METRICAS_DIARIA | Fecha | Único (DIM→FACT) |
| R4 | DIM_CALENDARIO | Fecha | BRIDGE_CAMPANA_DIA | Fecha | Único (DIM→bridge) |
| R5 | FACT_CAMPANAS | CargaID | BRIDGE_CAMPANA_DIA | CargaID | Único (FACT→bridge) |

Post-setup:
1. `DIM_CALENDARIO` → clic derecho > **Marcar como tabla de fechas** → columna `Fecha`.
2. Verificar en la vista de modelo que **no** aparece ninguna relación punteada
   (inactiva) ni ícono de "Ambas direcciones" — si aparece alguna, se creó una
   relación de más (ver `powerbi/README.md` Sección 2, "Relaciones que NO se
   crean").
3. `BRIDGE_CAMPANA_DIA` debe quedar con exactamente 2 relaciones entrantes
   (R4, R5) y ninguna saliente.

Ver `powerbi/README.md` para la justificación de cardinalidad, dirección, y
por qué esta forma evita loops/ambigüedad sin necesidad de relaciones
bidireccionales.
