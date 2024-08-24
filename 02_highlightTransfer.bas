
Private Const None As Long = -1
' RGB(173, 216, 230)
Private Const LightBlue As Long = 15773696
' RGB(40, 110, 170)
Private Const DarkBlue As Long = 12611584
' RGB(255, 255, 153)
Private Const LightYellow As Long = 49407
' RGB(255, 160, 160)
Private Const LightRed As Long = 255

Private Function MarkHighlightColor(sheet As Worksheet, row As Long, color As Long)
    sheet.Range("A" & row & ":K" & row).Interior.color = color
End Function

Private Sub ANZHighlightMatchingRows(sheetName1 As String, sheetName2 As String, sheet1Color As Long, sheet2Color As Long)
    Dim sheet1 As Worksheet
    Dim sheet2 As Worksheet
    Dim lastRow1 As Long
    Dim lastRow2 As Long
    Dim row1 As Long
    Dim row2 As Long
    Dim particulars As String
    Dim amount As Double
    
    ' set sheet references
    Set sheet1 = ThisWorkbook.Sheets(sheetName1)
    Set sheet2 = ThisWorkbook.Sheets(sheetName2)
    
    ' find the last row of data in each sheet
    lastRow1 = sheet1.Cells(sheet1.Rows.Count, "A").End(xlUp).Row
    lastRow2 = sheet2.Cells(sheet2.Rows.Count, "A").End(xlUp).Row
    
    ' loop through each row in sheet1
    For row1 = 2 To lastRow1
        particulars = sheet1.Cells(row1, 8).Value ' Particulars column
        amount = sheet1.Cells(row1, 3).Value ' Amount column
        
        ' loop through each row in sheet2
        For row2 = 2 To lastRow2
            ' check if the Particulars and Amount match
            If sheet2.Cells(row2, 8).Value = particulars And Abs(sheet2.Cells(row2, 3).Value) = Abs(amount) Then
                ' highlight the matching rows in both sheets
                MarkHighlightColor sheet1, row1, sheet1Color
                MarkHighlightColor sheet2, row2, sheet2Color
            End If
        Next row2
    Next row1
End Sub

Private Sub BNZHighlightMatchingRows(sheetName1 As String, sheetName2 As String, sheet1Color As Long, sheet2Color As Long)
    Dim sheet1 As Worksheet
    Dim sheet2 As Worksheet
    Dim lastRow1 As Long
    Dim lastRow2 As Long
    Dim row1 As Long
    Dim row2 As Long
    Dim particulars As String
    Dim amount As Double
    
    ' set sheet references
    Set sheet1 = ThisWorkbook.Sheets(sheetName1)
    Set sheet2 = ThisWorkbook.Sheets(sheetName2)
    
    ' find the last row of data in each sheet
    lastRow1 = sheet1.Cells(sheet1.Rows.Count, "A").End(xlUp).Row
    lastRow2 = sheet2.Cells(sheet2.Rows.Count, "A").End(xlUp).Row
    
    ' loop through each row in sheet1
    For row1 = 2 To lastRow1
        ' Particulars column
        particulars = sheet1.Cells(row1, 8).Value 
        ' Amount column
        amount = sheet1.Cells(row1, 3).Value 
        
        ' loop through each row in sheet2
        For row2 = 2 To lastRow2
            ' check if the Particulars and Amount match
            If sheet2.Cells(row2, 8).Value = particulars And Abs(sheet2.Cells(row2, 3).Value) = Abs(amount) Then
                ' highlight the matching rows in both sheets
                MarkHighlightColor sheet1, row1, sheet1Color
                MarkHighlightColor sheet2, row2, sheet2Color
            End If
        Next row2
    Next row1
End Sub

Private Sub ANZToWestpacHighlightMatchingRows(sheetName1 As String, sheetName2 As String, sheet1Color As Long, sheet2Color As Long)

    Dim sheet1 As Worksheet
    Dim sheet2 As Worksheet
    Dim lastRow1 As Long
    Dim lastRow2 As Long
    Dim row1 As Long
    Dim row2 As Long
    
    Set sheet1 = Worksheets(sheetName1)
    Set sheet2 = Worksheets(sheetName2)
    
    lastRow1 = sheet1.Cells(Rows.Count, 1).End(xlUp).Row
    lastRow2 = sheet2.Cells(Rows.Count, 1).End(xlUp).Row
    
    For row1 = 2 To lastRow1
        For row2 = 2 To lastRow2
            If Abs(sheet1.Cells(row1, 3).Value) = Abs(sheet2.Cells(row2, 3).Value) And _
                sheet1.Cells(row1, 10).Value = sheet2.Cells(row2, 8).Value And _
                sheet2.Cells(row2, 2).Value = "Mr L Shen" Then
                MarkHighlightColor sheet1, row1, sheet1Color
                MarkHighlightColor sheet2, row2, sheet2Color
            End If
        Next row2
    Next row1

End Sub

Private Sub BNZToWestpacHighlightMatchingRows(sheetName1 As String, sheetName2 As String, sheet1Color As Long, sheet2Color As Long)

    Dim sheet1 As Worksheet
    Dim sheet2 As Worksheet
    Dim lastRow1 As Long
    Dim lastRow2 As Long
    Dim row1 As Long
    Dim row2 As Long
    
    Set sheet1 = Worksheets(sheetName1)
    Set sheet2 = Worksheets(sheetName2)
    
    lastRow1 = sheet1.Cells(Rows.Count, 1).End(xlUp).Row
    lastRow2 = sheet2.Cells(Rows.Count, 1).End(xlUp).Row
    
    For row1 = 2 To lastRow1
        For row2 = 2 To lastRow2
            ' check amount, reference, and particulars
            If Abs(sheet1.Cells(row1, 3).Value) = Abs(sheet2.Cells(row2, 3).Value) And _
                sheet1.Cells(row1, 8).Value = sheet2.Cells(row2, 9).Value And _ 
                sheet1.Cells(row1, 10).Value = sheet2.Cells(row2, 8).Value Then
                MarkHighlightColor sheet1, row1, sheet1Color
                MarkHighlightColor sheet2, row2, sheet2Color
            End If
        Next row2
    Next row1

End Sub

Private Sub ASBToWestpacHighlightMatchingRows(sheetName1 As String, sheetName2 As String, sheet1Color As Long, sheet2Color As Long)

    Dim sheet1 As Worksheet
    Dim sheet2 As Worksheet
    Dim lastRow1 As Long
    Dim lastRow2 As Long
    Dim row1 As Long
    Dim row2 As Long
    
    Set sheet1 = Worksheets(sheetName1)
    Set sheet2 = Worksheets(sheetName2)
    
    lastRow1 = sheet1.Cells(Rows.Count, 1).End(xlUp).Row
    lastRow2 = sheet2.Cells(Rows.Count, 1).End(xlUp).Row
    
    For row1 = 2 To lastRow1
        For row2 = 2 To lastRow2
            If Abs(sheet1.Cells(row1, 3).Value) = Abs(sheet2.Cells(row2, 3).Value) And _
                InStr(1, sheet2.Cells(row2, 2).Value, "Y Zhang") > 0 And _
                sheet2.Cells(row2, 9).Value = "Yannic" And _
                (InStr(1, sheet1.Cells(row1, 2).Value, "Cost") > 0 Or InStr(1, sheet1.Cells(row1, 2).Value, "Living") > 0) Then
                MarkHighlightColor sheet1, row1, sheet1Color
                MarkHighlightColor sheet2, row2, sheet2Color
            End If
        Next row2
    Next row1

End Sub

' Private Sub ANZLoanHighlightMatchingRows(sheetName1 As String, sheetName2 As String, sheet1Color As Long, sheet2Color As Long)
'     Dim sheet1 As Worksheet
'     Dim sheet2 As Worksheet
'     Dim lastRow1 As Long
'     Dim lastRow2 As Long
'     Dim row1 As Long
'     Dim row2 As Long
'     Dim amount As Double
    
'     ' set sheet references
'     Set sheet1 = ThisWorkbook.Sheets(sheetName1)
'     Set sheet2 = ThisWorkbook.Sheets(sheetName2)
    
'     ' find the last row of data in each sheet
'     lastRow1 = sheet1.Cells(sheet1.Rows.Count, "A").End(xlUp).Row
'     lastRow2 = sheet2.Cells(sheet2.Rows.Count, "A").End(xlUp).Row
    
'     ' loop through each row in sheet1
'     For row1 = 2 To lastRow1
'         amount = sheet1.Cells(row1, 3).Value ' Amount column
'         ' loop through each row in sheet2
'         For row2 = 2 To lastRow2
'             ' check if the Particulars and Amount match
'             If Abs(sheet2.Cells(row2, 3).Value) = Abs(amount) And _
'                 sheet2.Cells(row2, 7).Value = "Mr L Shen" Then
'                 ' highlight the matching rows in both sheets
'                 MarkHighlightColor sheet1, row1, sheet1Color
'                 MarkHighlightColor sheet2, row2, sheet2Color
'             End If
'         Next row2
'     Next row1
' End Sub

' Private Sub ASBLoanHighlightMatchingRows(sheetName1 As String, sheetName2 As String, sheet1Color As Long, sheet2Color As Long)
'     Dim sheet1 As Worksheet
'     Dim sheet2 As Worksheet
'     Dim lastRow1 As Long
'     Dim lastRow2 As Long
'     Dim row1 As Long
'     Dim row2 As Long
'     Dim amount As Double
    
'     ' set sheet references
'     Set sheet1 = ThisWorkbook.Sheets(sheetName1)
'     Set sheet2 = ThisWorkbook.Sheets(sheetName2)
    
'     ' find the last row of data in each sheet
'     lastRow1 = sheet1.Cells(sheet1.Rows.Count, "A").End(xlUp).Row
'     lastRow2 = sheet2.Cells(sheet2.Rows.Count, "A").End(xlUp).Row
    
'     ' loop through each row in sheet1
'     For row1 = 2 To lastRow1
'         amount = sheet1.Cells(row1, 3).Value ' Amount column
'         ' loop through each row in sheet2
'         For row2 = 2 To lastRow2
'             ' check if the Particulars and Amount match
'             If Abs(sheet2.Cells(row2, 3).Value) = Abs(amount) And _
'                 InStr(1, sheet1.Cells(row1, 2).Value, "Home Loan 19B Tonkin") > 0 And _
'                 sheet2.Cells(row2, 7).Value = "Miss Y Zhang" Then
'                 ' highlight the matching rows in both sheets
'                 MarkHighlightColor sheet1, row1, sheet1Color
'                 MarkHighlightColor sheet2, row2, sheet2Color
'             End If
'         Next row2
'     Next row1
' End Sub

Private Sub ASBLoanHighlightMatchingRows(sheetName1 As String, sheetName2 As String, sheet1Color As Long, sheet2Color As Long)
    Dim sheet1 As Worksheet
    Dim sheet2 As Worksheet
    Dim lastRow1 As Long
    Dim lastRow2 As Long
    Dim row1 As Long
    Dim row2 As Long
    Dim amount As Double
    
    ' set sheet references sheet1 = ASB, sheet2 = BNZ
    Set sheet1 = ThisWorkbook.Sheets(sheetName1)
    Set sheet2 = ThisWorkbook.Sheets(sheetName2)
    
    ' find the last row of data in each sheet
    lastRow1 = sheet1.Cells(sheet1.Rows.Count, "A").End(xlUp).Row
    lastRow2 = sheet2.Cells(sheet2.Rows.Count, "A").End(xlUp).Row
    
    ' loop through each row in ASB
    For row1 = 2 To lastRow1
        amount = sheet1.Cells(row1, 3).Value ' Amount column
        ' loop through each row in BNZ
        For row2 = 2 To lastRow2
            ' check if the Particulars and Amount match
            If Abs(sheet2.Cells(row2, 3).Value) = Abs(amount) And _
                InStr(1, sheet1.Cells(row1, 2).Value, "A/P Yannic BNZ") > 0 And _
                sheet2.Cells(row2, 10).Value = "19B Tonkin" And _
                sheet2.Cells(row2, 8).Value = "Yannic" Then
                ' highlight the matching rows in both sheets
                MarkHighlightColor sheet1, row1, sheet1Color
                MarkHighlightColor sheet2, row2, sheet2Color
            End If
        Next row2
    Next row1
End Sub

Sub HightlightTransfer()    
    ' find and hightlight matching rows
    ' ANZHighlightMatchingRows "C-ANZ-go", "C-ANZ-saving", LightBlue  'lightBlue
    ' ANZHighlightMatchingRows "C-ANZ-go", "S-ANZ-loan", LightBlue 'lightBlue
    BNZHighlightMatchingRows "C-BNZ-go", "S-BNZ-loan", DarkBlue, DarkBlue
    ANZToWestpacHighlightMatchingRows "C-ANZ-go", "S-Westpac", LightRed, LightBlue
    ' BNZToWestpacHighlightMatchingRows "C-BNZ-go", "S-Westpac", LightRed 'lightRed
    ASBToWestpacHighlightMatchingRows "Y-ASB", "S-Westpac", LightRed , LightYellow
    ' ANZLoanHighlightMatchingRows "C-ANZ-go", "S-ANZ-loan", DarkBlue 'darkBlue
    ' ASBLoanHighlightMatchingRows "Y-ASB", "S-ANZ-loan", LightBlue 'lightBlue
    ASBLoanHighlightMatchingRows "Y-ASB", "S-BNZ-loan", DarkBlue, LightYellow
End Sub
