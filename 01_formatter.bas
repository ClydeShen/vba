
Private Const None As Long = -1
' RGB(173, 216, 230)
Private Const LightBlue As Long = 15773696
' RGB(40, 110, 170)
Private Const DarkBlue As Long = 12611584
' RGB(255, 255, 153)
Private Const LightYellow As Long = 49407
' RGB(255, 160, 160)
Private Const LightRed As Long = 255

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
    AddTypeColumn
    SetAutoFilter
    ConverDateFormat ws
    SetTabColor ws, tabColor
    
End Sub

Private Sub Bnz(sheetName As String, tabColor As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    ws.Select
    ' update the date format first
    ConverBNZDateFormat ws

    Columns("G:G").Select
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    
    SplitAmount
    AddTypeColumn
    SetAutoFilter
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
    AddTypeColumn
    SetAutoFilter
    ConverDateFormat ws
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
    AddTypeColumn
    SetAutoFilter
    ConverDateFormat ws
    SetTabColor ws, tabColor
    
End Sub

Private Sub ConverBNZDateFormat(ws As Worksheet)
    Dim LastRow As Long
    Dim dateParts As Variant
    Dim newDate As String
    Dim dayPart As String
    Dim monthPart As String
    Dim yearPart As String
    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Loop through all the rows and assuming row 1 is header
    For i = 2 To LastRow
    '    check column A and N for dates
        For Each dateCell In ws.Range("A" & i & ", N" & i)
            ' check if the cell contains a date-like string
            If InStr(1, dateCell.Text, "/", vbTextCompare) > 0 Then
                ' split the date string
                dateParts = Split(dateCell.Text, "/")
                ' Ensure we have 3 parts
                If UBound(dateParts) = 2 Then
                    ' remove the first 2 characters from the first part
                    dayPart = Right(dateParts(0), Len(dateParts(0)) - 2)

                    ' add "0" to the month part if it is a single digit
                    If Len(dateParts(1)) = 1 Then
                        monthPart = "0" & dateParts(1)
                    Else
                        monthPart = dateParts(1)
                    End If

                    ' add "20" to the start of the year part
                    yearPart = "20" & dateParts(2)

                    ' combine the parts into a new date format "dd/mm/yyyy"
                    newDate = dayPart & "/" & monthPart & "/" & yearPart
                    ' set the cell text to the new date
                    dateCell.Value = CDate(newDate)
                    dateCell.NumberFormat = "dd/mm/yyyy"
                End If
            End If
        Next dateCell
    Next i
End Sub

Private Sub ConverDateFormat(ws As Worksheet)
    Dim LastRow As Long

    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Loop through all the rows and assuming row 1 is header
    For i = 2 To LastRow
    '    check column A and N for dates
        For Each dateCell In ws.Range("A" & i)
            ' check if the cell is a date type
            If IsDate(dateCell.Text) Then
                ' set the cell text to the new date
                dateCell.Value = CDate(dateCell.Text)
                dateCell.NumberFormat = "dd/mm/yyyy"
            End If
        Next dateCell
    Next i
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

Private Sub AddTypeColumn()

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

Private Sub SetAutoFilter()
    Columns("A:L").Select
    Selection.AutoFilter
End Sub

Sub Formatter()
    ' set color reference
    Anz "C-ANZ-go", None
    Bnz "C-BNZ-go", LightBlue
    Bnz "S-BNZ-loan", DarkBlue
    Westpac "S-Westpac", LightRed
    Asb "Y-ASB", LightYellow
End Sub