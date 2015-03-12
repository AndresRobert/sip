Sub Preparar()
    
    'Eliminar columnas sin uso
    Columns("L:L").Delete
    Columns("H:I").Delete
    Columns("E:E").Delete
    Columns("A:A").Delete
    
    'Etiquetar
    Range("A1").FormulaR1C1 = "DEP"
    Range("B1").FormulaR1C1 = "UPC"
    Range("C1").FormulaR1C1 = "SKU"
    Range("D1").FormulaR1C1 = "DES"
    Range("E1").FormulaR1C1 = "COL"
    Range("F1").FormulaR1C1 = "CAN"
    Range("G1").FormulaR1C1 = "VAL"
    Range("H1").FormulaR1C1 = "UPC"
    Range("I1").FormulaR1C1 = "ATS"
    Range("J1").FormulaR1C1 = "TAL"
    
    'Borrar duplicados
    Range("C2").Select
    While ActiveCell.Value <> ""
        If ActiveCell.Value = ActiveCell.Offset(1, 0).Value Then
            ActiveCell.Offset(0, 3).FormulaR1C1 = ActiveCell.Offset(0, 3).Value + ActiveCell.Offset(1, 3).Value
            ActiveCell.Offset(1, 0).EntireRow.Delete
        Else
            ActiveCell.Offset(1, 0).Select
        End If
    Wend
    
    'Llenar datos no disponibles
    Range("H2").FormulaR1C1 = "=LEFT(RC[-6],12)"
    Range("I2").FormulaR1C1 = "=VLOOKUP(VALUE(RC[-6]),MAE!C[-8]:C[-7],2,0)"
    Range("J2").FormulaR1C1 = "=VLOOKUP(MID(RC[-1],8,3),MAE!C[-6]:C[-5],2,0)"
    If Range("A3").Value <> "" Then
        Range("H2:J2").AutoFill Destination:=Range("H2", Range("G1").End(xlDown).Offset(0, 3))
    End If
    Columns("A:J").Copy
    Range("A1").PasteSpecial Paste:=xlValues
    Columns("B:B").Delete
    Columns("H:H").Insert
    Columns("E:E").Copy
    Columns("H:H").PasteSpecial Paste:=xlValues
    Columns("E:E").Delete
    Imprimir = Range("G1", Range("H1").End(xlDown)).AddressLocal
    ActiveSheet.PageSetup.PrintArea = Imprimir
    Range("G1", Range("H1").End(xlDown)).Select
    Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
    Selection.Borders(xlInsideVertical).Weight = xlHairline
    Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    Selection.Borders(xlInsideHorizontal).Weight = xlHairline
    Columns("A:J").AutoFit
    Range(Range("A1").End(xlDown).Offset(1, 0), Range("A1").End(xlDown).End(xlDown)).EntireRow.Delete
    Range("A1").Select
End Sub
