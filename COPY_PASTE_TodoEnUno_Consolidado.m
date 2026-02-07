// ═══════════════════════════════════════════════════════════════════════════
// SCRIPT FINAL CORREGIDO (MANEJO DE NULOS)
// ═══════════════════════════════════════════════════════════════════════════
// CORRECCIÓN: Se agrega un paso para convertir "null" en "0" después de la unión.
// Esto evita que la fórmula "if [CreditCount] > 7" falle en chats vacíos.
// ═══════════════════════════════════════════════════════════════════════════

let
    RutaCarpeta = "C:\Users\A365\Documents\QuinaQuery\2025\04. Abril",

    // 1. RDC (RESUMEN)
    Files = Folder.Files(RutaCarpeta),
    FileRDC = Table.SelectRows(Files, each Text.Contains([Name], "RDC_") and Text.EndsWith([Name], ".xlsx")){0},
    RDC_Raw = Excel.Workbook(File.Contents(FileRDC[Folder Path] & FileRDC[Name]), null, true){0}[Data],
    RDC_Head = Table.PromoteHeaders(RDC_Raw, [PromoteAllScalars=true]),
    RDC_Select = Table.SelectColumns(RDC_Head, {"ID Chat", "ID", "F.Inicio Chat", "Tipificación Chat"}),
    RDC_Type = Table.TransformColumnTypes(RDC_Select, {{"F.Inicio Chat", type datetime}, {"ID", type text}}),
    
    // Regla 24h
    RDC_Sorted = Table.Sort(Table.SelectRows(RDC_Type, each [ID] <> null), {{"ID", Order.Ascending}, {"F.Inicio Chat", Order.Ascending}}),
    RDC_Buffer = Table.Buffer(RDC_Sorted), 
    RDC_Index = Table.AddIndexColumn(RDC_Buffer, "Idx", 0, 1, Int64.Type),
    RDC_WithFlag = Table.AddColumn(RDC_Index, "Es_Cobrable", each 
        if [Idx]=0 then 1 
        else if [ID] <> RDC_Buffer[ID]{[Idx]-1} then 1 
        else if Duration.TotalHours([#"F.Inicio Chat"] - RDC_Buffer[#"F.Inicio Chat"]{[Idx]-1}) >= 24.0 then 1
        else 0, Int64.Type),
    RDC_Final = Table.SelectColumns(RDC_WithFlag, {"ID Chat", "Es_Cobrable"}),

    // 2. DDC (DETALLE)
    FilesDDC = Table.SelectRows(Files, each Text.Contains([Name], "DDC_") and Text.EndsWith([Name], ".xlsx")),
    DDC_List = List.Transform(Table.ToRecords(FilesDDC), each Excel.Workbook(File.Contents([Folder Path] & [Name]), null, true){0}[Data]),
    DDC_Union = Table.Combine(List.Transform(DDC_List, each Table.PromoteHeaders(_, [PromoteAllScalars=true]))),
    DDC_Select = Table.SelectColumns(DDC_Union, {"ID Chat", "Mensaje", "Fecha Hora", "Tipo"}),
    DDC_Type = Table.TransformColumnTypes(DDC_Select, {{"Fecha Hora", type datetime}, {"ID Chat", type text}, {"Tipo", type text}, {"Mensaje", type text}}),
    DDC_Clean = Table.SelectRows(DDC_Type, each [#"Fecha Hora"] <> null and [#"ID Chat"] <> null),

    // Agrupamiento
    DDC_Stats = Table.Group(DDC_Clean, {"ID Chat"}, {
        {"Stats", (t) => 
            let
                Times = Table.Column(t, "Fecha Hora"),
                RowsAgent = Table.SelectRows(t, each [Tipo] = "NOTIFICATION"),
                TimeAgent = if Table.IsEmpty(RowsAgent) then null else List.Min(Table.Column(RowsAgent, "Fecha Hora")),
                RowsCredit = Table.SelectRows(t, each [Mensaje] <> null and Text.Contains([Mensaje], "evaluar si tienes un crédito", Comparer.OrdinalIgnoreCase)),
                TimeCredit = if Table.IsEmpty(RowsCredit) then null else List.Min(Table.Column(RowsCredit, "Fecha Hora")),
                
                Facturables = List.Count(List.Select(Times, each (TimeAgent = null or _ < TimeAgent) and (TimeCredit = null or _ < TimeCredit))),
                CreditCount = if TimeCredit = null then 0 else List.Count(List.Select(Times, each _ >= TimeCredit))
            in
                [Facturable = Facturables, CreditCount = CreditCount], type record
        }
    }),
    DDC_Expanded = Table.ExpandRecordColumn(DDC_Stats, "Stats", {"Facturable", "CreditCount"}, {"FACTURABLE_BANCO", "CreditCount"}),

    // 3. UNIÓN Y LIMPIEZA DE NULOS (SOLUCIÓN ERROR)
    Join = Table.NestedJoin(RDC_Final, {"ID Chat"}, DDC_Expanded, {"ID Chat"}, "DDC", JoinKind.LeftOuter),
    Expand = Table.ExpandTableColumn(Join, "DDC", {"FACTURABLE_BANCO", "CreditCount"}, {"FACTURABLE_BANCO", "CreditCount"}),
    
    // IMPORTANTE: Convertir 'null' a 0 para que las fórmulas funcionen
    FillNulls = Table.ReplaceValue(Expand, null, 0, Replacer.ReplaceValue, {"FACTURABLE_BANCO", "CreditCount"}),
    
    // Clasificación final
    Result = Table.AddColumn(FillNulls, "Clasificacion_Mesa", each 
        if [CreditCount] > 7 then "Cobrar en ambas mesas"
        else if [CreditCount] > 0 then "Cobrar Mesa Comercial"
        else null, type text
    )
in
    Result
