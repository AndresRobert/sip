Sub Preparar()
    Columns("K:L").ClearFormats
    Columns("O:O").Delete
    Columns("M:M").Delete
    Columns("J:J").Delete
    Columns("C:C").Delete
    Columns("A:A").Delete
    Range("A1").FormulaR1C1 = "UPS"
    Range("B1").FormulaR1C1 = "TAL"
    Range("C1").FormulaR1C1 = "YEA"
    Range("D1").FormulaR1C1 = "TEM"
    Range("E1").FormulaR1C1 = "VAL"
    Range("F1").FormulaR1C1 = "DEP"
    Range("G1").FormulaR1C1 = "DES"
    Range("H1").FormulaR1C1 = "COL"
    Range("I1").FormulaR1C1 = "CAN"
    Range("J1").FormulaR1C1 = "NAC"
    Range("K1").FormulaR1C1 = "SKU"
    Range("K2").FormulaR1C1 = "=RIGHT(RC[-10],9)"
    Range("L1").FormulaR1C1 = "ATS"
    Range("L2").FormulaR1C1 = "=VLOOKUP(TEXT(RC[-11],0),MAE!C[-11]:C[-10],2,0)"
    Range("K2:L2").AutoFill Destination:=Range("K2", Range("J1").End(xlDown).Offset(0, 2))
    Columns("A:L").Copy
    Range("A1").PasteSpecial Paste:=xlValues
    Columns("I:I").Copy
    Range("M1").PasteSpecial Paste:=xlValues
    Columns("I:I").Delete
    
    Imprimir = Range("K1", Range("L1").End(xlDown)).AddressLocal
    ActiveSheet.PageSetup.PrintArea = Imprimir
    Range("K1", Range("L1").End(xlDown)).Select
    Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
    Selection.Borders(xlInsideVertical).Weight = xlHairline
    Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    Selection.Borders(xlInsideHorizontal).Weight = xlHairline
    
    Columns("A:L").AutoFit
    Range("A1").Select
End Sub
