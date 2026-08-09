// Query Power BI: BRIDGE_CAMPANA_DIA
// Origen: output/bridge_campana_dia.parquet (Gate 4A), grain (CargaID, Fecha)
// vigente. Tabla puente minimalista (solo 2 columnas): NO agregar columnas
// descriptivas aca (Cliente/Marca/etc. se traen desde FACT_CAMPANAS via
// relacion CargaID, ver powerbi/README.md Sec.2). ~881k filas: es la tabla
// mas grande del modelo mecanicamente, deliberado (evita reimplementar
// overlap de fechas en DAX).
let
    Origen = Parquet.Document(File.Contents(pRutaOutput & "\bridge_campana_dia.parquet")),
    #"Tipos clave" = Table.TransformColumnTypes(Origen, {{"CargaID", type text}, {"Fecha", type date}})
in
    #"Tipos clave"
