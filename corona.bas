Sub Preparar()
    Columns("J:J").Delete
    Columns("C:C").Delete
    Columns("A:A").Delete
    Range("L2").FormulaR1C1 = "=CONCATENATE(LEFT(RC[-6],3),""."",RC[-11],""."",TEXT(VALUE(RC[-10]),""000""))"
    Range("M2").FormulaR1C1 = "=CONCATENATE(LEFT(RC[-7],3),""-"",LEFT(RC[-8],3))"
    Range("N2").FormulaR1C1 = "=LEFT(RC[-3],12)"
    Range("L1").FormulaR1C1 = "SKU"
    Range("M1").FormulaR1C1 = "SBR"
    Range("N1").FormulaR1C1 = "UPS"
    Range("L2:N2").Select
    If Range("A3").Value <> "" Then
        Selection.AutoFill Destination:=Range("L2", Range("K1").End(xlDown).Offset(0, 3))
    End If
    Columns("A:N").Copy
    Range("A1").PasteSpecial Paste:=xlValues
    Columns("A:B").Delete
    Columns("C:D").Delete
    Columns("G:G").Delete
    Range("A1").FormulaR1C1 = "TAL"
    Range("B1").FormulaR1C1 = "DES"
    Range("C1").FormulaR1C1 = "ATS"
    Range("D1").FormulaR1C1 = "VAL"
    Range("E1").FormulaR1C1 = "TEM"
    Range("F1").FormulaR1C1 = "CAN"
    Columns("D:D").Insert
    Columns("G:G").Copy
    Range("D1").PasteSpecial Paste:=xlValues
    Columns("G:G").Delete
    Columns("A:I").AutoFit
    
    Imprimir = Range("C1", Range("D1").End(xlDown)).AddressLocal
    ActiveSheet.PageSetup.PrintArea = Imprimir
    Range("C1", Range("D1").End(xlDown)).Select
    Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
    Selection.Borders(xlInsideVertical).Weight = xlHairline
    Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    Selection.Borders(xlInsideHorizontal).Weight = xlHairline
    
End Sub
