Sub Preparar()
    Columns("B:AM").ClearContents
    Columns("B:AM").ClearFormats
    Columns("A:A").Select
    Application.CutCopyMode = False
    Selection.TextToColumns _
        Destination:=Range("A1"), _
        DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, _
        Tab:=True, _
        Semicolon:=False, _
        Comma:=False, _
        Space:=False, _
        Other:=True, _
        OtherChar:="|", _
        TrailingMinusNumbers:=True
    Columns("AG:AM").Delete
    Columns("Y:AD").Delete
    Columns("W:W").Delete
    Columns("A:U").Delete
    Range("B1").FormulaR1C1 = "DES"
    Range("C1").FormulaR1C1 = "VAL"
    Range("D1").FormulaR1C1 = "CAN"
    Range("E1").FormulaR1C1 = "ATS"
    Range("E2").FormulaR1C1 = "=VLOOKUP(VALUE(MID(RC[-4],4,7)),MAE!C[-4]:C[-3],2,0)"
    Range("E2").AutoFill Destination:=Range("E2", Range("D2").End(xlDown).Offset(0, 1))
    Range("E2", Range("E2").End(xlDown)).Copy
    Range("E2").PasteSpecial Paste:=xlValues
    Range("F2").FormulaR1C1 = "=LEFT(RC[-5],12)"
    Range("F2").AutoFill Destination:=Range("F2", Range("E2").End(xlDown).Offset(0, 1))
    Range("F2", Range("F2").End(xlDown)).Copy
    Range("A2").PasteSpecial Paste:=xlValues
    Columns("F:F").ClearContents
    
    Imprimir = Range("D1", Range("E1").End(xlDown)).AddressLocal
    ActiveSheet.PageSetup.PrintArea = Imprimir
    Range("D1", Range("E1").End(xlDown)).Select
    Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
    Selection.Borders(xlInsideVertical).Weight = xlHairline
    Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    Selection.Borders(xlInsideHorizontal).Weight = xlHairline
    
    Columns("A:E").AutoFit
    Range("A1").Select
End Sub
