Private Const None As Long = -1
Private Const LightBlue As Long = &HADD8E6 ' RBG(173, 216, 230)
Private Const DarkBlue As Long = &H286EAA ' RBG(40, 110, 170)
Private Const LightYellow As Long = &HFFFF99 ' RBG(255, 255, 153)
Private Const LightRed As Long = &HFFA0A0 ' RBG(255, 160, 160)

Private Sub Anz(sheetName As String, tabColor As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    ws.Select
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
    setAutoFilter
    SetTabColor ws, tabColor
    
End Sub

Private Sub Bnz(sheetName As String, tabColor As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    ws.Select
    Columns("G:G").Select
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    
    SplitAmount
    addTypeColumn
    setAutoFilter
    SetTabColor ws, tabColor
    
End Sub

Private Sub Westpac(sheetName As String, tabColor As Long)

     Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    ws.Select

    Columns("C:C").Select
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    
    SplitAmount
    addTypeColumn
    setAutoFilter
    SetTabColor ws, tabColor

End Sub

Private Sub Asb(sheetName As String, tabColor As Long)
    
     Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    ws.Select

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
    setAutoFilter
    SetTabColor ws, tabColor
    
End Sub

Private Sub SplitAmount()
    
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

Private Sub addTypeColumn()

    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Type"
    
End Sub

Private Sub SetTabColor(ws As Worksheet, ByVal tabColor As Long)
   If tabColor <> vbNullColor Then
      ws.Tab.Color = tabColor
   End If
End Sub

Private Sub setAutoFilter()
    Columns("A:L").Select
    Selection.AutoFilter
End Sub

Sub Formatter()
    ' set color reference
    ' Anz "C-ANZ-go", None
    ' Bnz "C-BNZ-go", LightBlue
    ' Bnz "S-BNZ-loan", DarkBlue
    Westpac "S-Westpac", LightRed
    ' Asb "Y-ASB", LightYellow
End Sub