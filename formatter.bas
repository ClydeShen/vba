Sub Anz(sheetName As String)

    Sheets(sheetName).Select
    Columns("G:G").Select
    Selection.Cut
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Columns("G:G").Select
    Selection.Cut
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    
    SplitAmount
    addTypeColumn
    
End Sub

Sub Westpac(sheetName As String)

    Sheets(sheetName).Select
    Columns("C:C").Select
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    
    SplitAmount
    addTypeColumn

End Sub

Sub Asb(sheetName As String)

    Sheets(sheetName).Select
    Rows("1:6").Select
    Selection.Delete Shift:=xlUp
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    Columns("F:F").Select
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Columns("G:G").Select
    Selection.Cut
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    
    SplitAmount
    addTypeColumn
    
End Sub

Sub SplitAmount()
    
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "In+"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Out-"
    

    Dim LastRow As Long
    LastRow = Cells(Rows.Count, "C").End(xlUp).Row
    Range("C2:C" & LastRow).Select
    Selection.TextToColumns Destination:=Range("D2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="-", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True

End Sub

Sub addTypeColumn()

    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Type"
    
End Sub


Sub Formatter()
    ' format csv file
    Anz "C-ANZ-go"
    Anz "C-ANZ-saving"
    Anz "S-ANZ-loan"
    Westpac "S-Westpac"
    Asb "Y-ASB"
End Sub
