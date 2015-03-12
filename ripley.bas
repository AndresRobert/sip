Sub Procesar()
    'Limpiar
    Columns("B:O").ClearContents
    Columns("B:O").ClearFormats
    'Separar
    Columns("A:A").Select
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("A1"), _
        DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, _
        Tab:=False, _
        Semicolon:=False, _
        Comma:=True, _
        Space:=False, _
        Other:=False, _
        TrailingMinusNumbers:=True
    
    'Borrar columnas sin uso
    Columns("K:K").Delete
    Columns("G:G").Delete
    Columns("A:A").Delete
    
    'Ordenar por SKU
    Columns("A:H").Select
    ActiveWorkbook.Worksheets("PRECIO").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("PRECIO").Sort.SortFields.Add _
        Key:=Range("B2"), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("PRECIO").Sort
        .SetRange Range("A1", Range("H1").End(xlDown))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Eliminar duplicados
    Range("B2").Select
    While ActiveCell.Value <> ""
        If ActiveCell.Value = ActiveCell.Offset(1, 0).Value Then
            ActiveCell.Offset(0, 5).FormulaR1C1 = ActiveCell.Offset(0, 5).Value + ActiveCell.Offset(1, 5).Value
            ActiveCell.Offset(1, 0).EntireRow.Delete
        Else
            ActiveCell.Offset(1, 0).Select
        End If
    Wend
    
    
    
    'Etiquetar
    Range("A1").FormulaR1C1 = "ATS"
    Range("B1").FormulaR1C1 = "SKU"
    Range("C1").FormulaR1C1 = "UPC"
    Range("D1").FormulaR1C1 = "DEP"
    Range("E1").FormulaR1C1 = "LIN"
    Range("F1").FormulaR1C1 = "DES"
    Range("G1").FormulaR1C1 = "CAN"
    Range("H1").FormulaR1C1 = "VAL"
    Range("I1").FormulaR1C1 = "DEPDES"
    Range("J1").FormulaR1C1 = "LINDES"
    Range("K1").FormulaR1C1 = "DEP"
    Range("L1").FormulaR1C1 = "DES"
    Range("M1").FormulaR1C1 = "UPC"
    
    'Llenar información faltante
    Range("A2").FormulaR1C1 = "=VLOOKUP(TEXT(RC[1],0),MAE!C:C[1],2,0)"
    Range("I2").FormulaR1C1 = "=VLOOKUP(LEFT(RC[-5],4),MAE!C[-5]:C[-4],2,0)"
    Range("J2").FormulaR1C1 = "=VLOOKUP(TEXT(RC[-5],""000000""),MAE!C[-3]:C[-2],2,0)"
    Range("K2").FormulaR1C1 = "=TRIM(RC[-7])"
    Range("L2").FormulaR1C1 = "=TRIM(RC[-6])"
    Range("M2").FormulaR1C1 = "=LEFT(RC[-10],12)"
    On Error Resume Next
    Range("A2").AutoFill Destination:=Range("A2", Range("A1").End(xlDown))
    Range("I2:M2").AutoFill Destination:=Range("I2", Range("H1").End(xlDown).Offset(0, 5))
    Columns("A:M").Copy
    Range("A1").PasteSpecial Paste:=xlValues
    Columns("F:F").Delete
    Columns("C:D").Delete
    Columns("B:B").Insert
    Columns("E:E").Copy
    Columns("B:B").PasteSpecial Paste:=xlValues
    Columns("E:E").Delete
    Columns("A:J").AutoFit
    Application.CutCopyMode = False
    Imprimir = Range("A1", Range("B1").End(xlDown)).AddressLocal
    ActiveSheet.PageSetup.PrintArea = Imprimir
    Range("A1", Range("B1").End(xlDown)).Select
    Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
    Selection.Borders(xlInsideVertical).Weight = xlHairline
    Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    Selection.Borders(xlInsideHorizontal).Weight = xlHairline
    Range("A1").Select
End Sub
