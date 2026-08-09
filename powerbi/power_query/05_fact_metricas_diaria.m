// Query Power BI: FACT_METRICAS_DIARIA
// Origen: output/fact_metricas_diaria.parquet (Gate 4A), grain (ElementoID,
// Fecha) CON actividad (WIDE, sparse). Un dia sin fila = actividad 0, NO
// "dia inexistente": cualquier medida "sobre todos los dias del periodo"
// debe iterar DIM_CALENDARIO (relacion R3), nunca asumir densidad de esta
// tabla (ver powerbi/README.md Sec.1 y dax/06_metric_status.dax).
let
    Origen = Parquet.Document(File.Contents(pRutaOutput & "\fact_metricas_diaria.parquet")),
    #"Tipos clave" = Table.TransformColumnTypes(Origen, {{"ElementoID", type text}, {"Fecha", type date}})
in
    #"Tipos clave"
