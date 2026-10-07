# OCU26 — Roadmap de automatización (decisión 2026-10-07)

La rama `etapa2-ocu26` es el **checkpoint del desarrollo hasta la Etapa 2C**
(reglas centrales 2A, migración histórica 2B/2B.1/2B.2 y depuración 2C). Está
publicada como rama paralela; `main` no se modificó. No es el proceso operativo.

Antecedente: `CM3_MARCO_RECTOR_ESPACIOS_PUBLICITARIOS_2026-08-19.md` (§8.6 y §9)
dejaba Office Script y GitHub Actions para una etapa futura. Esta es esa etapa:
el usuario fija la arquitectura objetivo de abajo.

## Arquitectura objetivo

```
Power Apps / Form
  → OCU26_ENTRADAS (SharePoint)
  → Power Automate
  → OCU_ELEMENTOS / OCU_CAMPANAS
  → Office Script
  → Excel OCU26
  → dataset técnico de publicación
  → GitHub Actions
  → generador TV1–TV6
  → HTML
  → Netlify
```

## Se construye una vez

- Reglas comerciales centralizadas (`config/business_semantics.json` + `scripts/semantic_model.py`).
- Generador TV1–TV6, KPIs, filtros y plantillas.
- Estructura del dataset técnico.
- GitHub Actions y vínculo GitHub ↔ Netlify.
- Puente Power Automate / Office Script / dataset.

## Corre en cada actualización (sin intervención manual)

Form → SharePoint → Power Automate → Office Script → dataset → GitHub Action →
generador → HTML → Netlify.

Una actualización normal (p. ej. lunes/miércoles) no requiere copiar archivos,
correr scripts a mano ni abrir Claude.

## Rol de Claude

Capa de **desarrollo y evolución**. Se usa cuando cambia una regla o un KPI,
cambia el diseño, aparece una nueva TV o hay que depurar. **No** es necesario
para cargar campañas normales ni para regenerar TV1–TV6 con datos nuevos.

## Migración histórica 2026 (excepcional)

`scripts/staging_import_historico.py`, `scripts/cierre_historico_2b2.py` y
`scripts/cierre_historico_2c.py` se conservan como herramienta de migración,
auditoría, recuperación y trazabilidad. **No** son una etapa del refresh: exigen
el SHA exacto de su base de partida y no vuelven a correr sobre la base vigente.
Reglas confirmadas y resultados: `ETAPA2B_HISTORICO_OCUPACION.md`,
`ETAPA2C_REVISION_FINAL_HISTORICO.md`. Quedan 331 asignaciones para revisión
humana (fuera del repo).

Reglas históricas que deben preservarse: OT = IDCampaña numérico; `B <n>` = OT
`<n>` bonificada (la B solo como observación); Remeros por posición (REM-TS n →
n-ésimo REM-DB del maestro, sin crear soportes); PALS/Alsina dado de baja; una
extensión solo con evidencia (misma OT + elemento + inicio, FIN posterior); no
cambiar FechaInicio sin evidencia inequívoca; si CAMPANAS ya cubre todos los
meses con presencia real, prevalece CAMPANAS; periodos disjuntos demostrados
pueden ser filas nuevas; lo ambiguo queda para revisión humana.

## Requisito de datos (repositorio público)

- La base Excel completa **no** se publica; tampoco datos comerciales sensibles.
- El dataset técnico futuro contiene solo los campos estrictamente necesarios
  para los tableros, o se usa un mecanismo de transferencia que no exponga
  información privada.
- Diseño e implementación: etapa de automatización (todavía no decidido).

## Estado al checkpoint

Base local validada (CAMPANAS 16.382 filas, 633 OT). Todavía sin SharePoint,
Power Apps, Power Automate, Office Script, GitHub Actions ni conexión de esta
rama a Netlify. TV1–TV6 no regeneradas.
