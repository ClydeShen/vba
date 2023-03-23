Sub FindMatchingCell(ByVal searchStr As String, ByVal startRow As Long, ByRef resultCell As Range)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Summary")
    
    Dim lastRow As Long
    lastRow = ws.Cells(Rows.Count, 6).End(xlUp).Row  'get last row of column A
    
    Dim foundCell As Range
    Set foundCell = ws.Range("F" & startRow & ":F" & lastRow).find(What:=searchStr, LookIn:=xlValues, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    
    If Not foundCell Is Nothing Then
        Set resultCell = foundCell 'set the result cell to the matching cell
    Else
        Set resultCell = Nothing 'set the result cell to nothing if no matching cell is found
    End If

End Sub

' set the type cell
Sub SetTypeCell(ByVal emptyCell As Range, ByVal typeCell As Range)
  ' if typeCell is Nothing then print the address of emptyCell
    If typeCell Is Nothing Then
        Debug.Print emptyCell.Address
    End


   If Not typeCell Is Nothing Then ' check if typeCell is not Nothing
        emptyCell.Value = typeCell.Value
        emptyCell.Interior.Color = typeCell.Interior.Color
    End If
End Sub


Sub HightlightTypes()
    Dim summary As Worksheet
    Dim dict As Object
    Dim dataRow As Long
    Dim midRow As Long
    
    Set summary = ThisWorkbook.Sheets("Summary")
    
    dataRow = summary.Cells(summary.Rows.Count, 1).End(xlUp).Row
    midRow = dataRow + 1

    Dim i As Long
    Dim j As Long
    Dim typeCell As Range
    Dim otherPartyCell As Range
    

    For j = 1 To midRow
        Set otherPartyCell = summary.Cells(j, 2)
        ' category: groceries
        If InStr(1, otherPartyCell, "Countdown", vbTextCompare) > 0 _
            Or (InStr(1, otherPartyCell, "Pak N Save", vbTextCompare) > 0 And _
              InStr(1, otherPartyCell, "Pak N Save Fuel", vbTextCompare) = 0) _
            Or InStr(1, otherPartyCell, "New World", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Taiping", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Golden Apple", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Seasons Markets", vbTextCompare) > 0 Then

            FindMatchingCell "Groceries", midRow, typeCell
            SetTypeCell summary.Cells(j, 6), typeCell
        End If
        ' category: Home & contents
        If InStr(1, otherPartyCell, "AA Insurance Pre", vbTextCompare) > 0 Then
            FindMatchingCell "Home & contents", midRow, typeCell
            SetTypeCell summary.Cells(j, 6), typeCell
        End If
        
        ' category: Mortgage repayments/rent
        If InStr(1, otherPartyCell, "Loan Payment", vbTextCompare) > 0 Then
            FindMatchingCell "Mortgage repayments/rent", midRow, typeCell
            SetTypeCell summary.Cells(j, 6), typeCell
        End If

        ' category: Electricity & Gas & Internet
        If InStr(1, otherPartyCell, "Contact Energy", vbTextCompare) > 0 Then
            FindMatchingCell "Electricity & Gas & Internet", midRow, typeCell
            SetTypeCell summary.Cells(j, 6), typeCell
        End If

        ' category: Travel
        If InStr(1, otherPartyCell, "AT HOP", vbTextCompare) > 0  _ 
            Or InStr(1, otherPartyCell, "Gull", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "BP", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Pak N Save Fuel", vbTextCompare) > 0 Then
            FindMatchingCell "Travel", midRow, typeCell
            SetTypeCell summary.Cells(j, 6), typeCell
        End If

        ' category: Telephone
        If InStr(1, otherPartyCell, "Vodafone", vbTextCompare) > 0 _ 
          Or InStr(1, otherPartyCell, "Skinny", vbTextCompare) > 0 Then
            FindMatchingCell "Telephone", midRow, typeCell
            SetTypeCell summary.Cells(j, 6), typeCell
        End If

        ' category: Council Rate
        If InStr(1, otherPartyCell, "Auckland Council", vbTextCompare) > 0 Then
            FindMatchingCell "Council Rate", midRow, typeCell
            SetTypeCell summary.Cells(j, 6), typeCell
        End If

    
            
    Next j
            
End Sub
