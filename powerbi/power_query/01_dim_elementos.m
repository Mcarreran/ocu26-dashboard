// Query Power BI: DIM_ELEMENTOS
// Origen: output/dim_elementos.parquet (Gate 4A), 1 fila / ElementoID.
// Requiere el parámetro pRutaOutput (ver 00_parametro_pRutaOutput.m).
// No se renombran ni reinterpretan columnas: Parquet.Document ya preserva
// los dtypes escritos por pandas/pyarrow (Int64/Float64 nullable, boolean,
// datetime, string) - no hace falta un paso Table.TransformColumnTypes.
let
    Origen = Parquet.Document(File.Contents(pRutaOutput & "\dim_elementos.parquet")),
    #"Tipo ElementoID texto" = Table.TransformColumnTypes(Origen, {{"ElementoID", type text}})
in
    #"Tipo ElementoID texto"
