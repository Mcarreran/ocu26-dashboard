// Query Power BI: FACT_CAMPANAS
// Origen: output/fact_campanas.parquet (Gate 4A), 1 fila / CargaID.
// IDCampaña es columna de agrupación comercial, NO clave - no se crea
// ninguna relación de modelo sobre IDCampaña (ver powerbi/README.md Sec.2).
let
    Origen = Parquet.Document(File.Contents(pRutaOutput & "\fact_campanas.parquet")),
    #"Tipos clave" = Table.TransformColumnTypes(Origen, {
        {"CargaID", type text},
        {"ElementoID", type text},
        {"IDCampaña", type text},
        {"FechaInicio", type date},
        {"FechaFin", type date}
    })
in
    #"Tipos clave"
