// Query Power BI: DIM_CALENDARIO
// Origen: output/dim_calendario.parquet (Gate 4A), 1 fila / día, rango
// dinámico [min FechaInicio .. max FechaFin] de FACT_CAMPANAS, sin huecos
// (generado con pd.date_range en Python: no se regenera en M).
// Tras cargar: Herramientas de tabla > Marcar como tabla de fechas > columna
// "Fecha" (habilita time intelligence nativo: SAMEPERIODLASTYEAR, DATEADD, etc.
// - ver powerbi/dax/05_temporal.dax).
let
    Origen = Parquet.Document(File.Contents(pRutaOutput & "\dim_calendario.parquet")),
    #"Tipo Fecha" = Table.TransformColumnTypes(Origen, {{"Fecha", type date}, {"InicioMes", type date}, {"FinMes", type date}})
in
    #"Tipo Fecha"
