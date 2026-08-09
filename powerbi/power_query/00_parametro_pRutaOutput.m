// Parametro de Power Query: pRutaOutput
// ---------------------------------------------------------------------------
// Crear en Power BI Desktop: Inicio > Administrar parámetros > Nuevo parámetro.
//   Nombre:        pRutaOutput
//   Tipo:          Texto
//   Valor actual:  C:\brand plus\ocu26-dashboard\output
// (sin barra final). Las 5 queries de tabla (01_.. a 05_..) referencian este
// parámetro en vez de repetir la ruta absoluta 5 veces: al mover el repo o
// regenerar el Excel/Parquet en otra máquina, se edita un solo valor.
//
// Definición equivalente en M (por si se crea a mano en el editor avanzado
// en vez de por UI):
let
    pRutaOutput = "C:\brand plus\ocu26-dashboard\output" meta [IsParameterQuery = true, Type = "Text", IsParameterQueryRequired = true]
in
    pRutaOutput
