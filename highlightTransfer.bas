
Sub ANZHighlightMatchingRows(sheetName1 As String, sheetName2 As String, highlightColor As Long)
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
                sheet1.Rows(row1).Interior.Color = highlightColor
                sheet2.Rows(row2).Interior.Color = highlightColor
            End If
        Next row2
    Next row1
End Sub



Sub ANZToWestpacHighlightMatchingRows(sheetName1 As String, sheetName2 As String, highlightColor As Long)

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
                sheet1.Range("A" & row1 & ":K" & row1).Interior.Color = highlightColor
                sheet2.Range("A" & row2 & ":K" & row2).Interior.Color = highlightColor
            End If
        Next row2
    Next row1

End Sub

Sub ASBToWestpacHighlightMatchingRows(sheetName1 As String, sheetName2 As String, highlightColor As Long)

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
                sheet1.Range("A" & row1 & ":K" & row1).Interior.Color = highlightColor
                sheet2.Range("A" & row2 & ":K" & row2).Interior.Color = highlightColor
            End If
        Next row2
    Next row1

End Sub

Sub ANZLoanHighlightMatchingRows(sheetName1 As String, sheetName2 As String, highlightColor As Long)
    Dim sheet1 As Worksheet
    Dim sheet2 As Worksheet
    Dim lastRow1 As Long
    Dim lastRow2 As Long
    Dim row1 As Long
    Dim row2 As Long
    Dim amount As Double
    
    ' set sheet references
    Set sheet1 = ThisWorkbook.Sheets(sheetName1)
    Set sheet2 = ThisWorkbook.Sheets(sheetName2)
    
    ' find the last row of data in each sheet
    lastRow1 = sheet1.Cells(sheet1.Rows.Count, "A").End(xlUp).Row
    lastRow2 = sheet2.Cells(sheet2.Rows.Count, "A").End(xlUp).Row
    
    ' loop through each row in sheet1
    For row1 = 2 To lastRow1
        amount = sheet1.Cells(row1, 3).Value ' Amount column
        ' loop through each row in sheet2
        For row2 = 2 To lastRow2
            ' check if the Particulars and Amount match
            If Abs(sheet2.Cells(row2, 3).Value) = Abs(amount) And _
                sheet2.Cells(row2, 7).Value = "Mr L Shen" Then
                ' highlight the matching rows in both sheets
                sheet1.Rows(row1).Interior.Color = highlightColor
                sheet2.Rows(row2).Interior.Color = highlightColor
            End If
        Next row2
    Next row1
End Sub

Sub ASBLoanHighlightMatchingRows(sheetName1 As String, sheetName2 As String, highlightColor As Long)
    Dim sheet1 As Worksheet
    Dim sheet2 As Worksheet
    Dim lastRow1 As Long
    Dim lastRow2 As Long
    Dim row1 As Long
    Dim row2 As Long
    Dim amount As Double
    
    ' set sheet references
    Set sheet1 = ThisWorkbook.Sheets(sheetName1)
    Set sheet2 = ThisWorkbook.Sheets(sheetName2)
    
    ' find the last row of data in each sheet
    lastRow1 = sheet1.Cells(sheet1.Rows.Count, "A").End(xlUp).Row
    lastRow2 = sheet2.Cells(sheet2.Rows.Count, "A").End(xlUp).Row
    
    ' loop through each row in sheet1
    For row1 = 2 To lastRow1
        amount = sheet1.Cells(row1, 3).Value ' Amount column
        ' loop through each row in sheet2
        For row2 = 2 To lastRow2
            ' check if the Particulars and Amount match
            If Abs(sheet2.Cells(row2, 3).Value) = Abs(amount) And _
                InStr(1, sheet1.Cells(row1, 2).Value, "Home Loan 19B Tonkin") > 0 And _
                sheet2.Cells(row2, 7).Value = "Miss Y Zhang" Then
                ' highlight the matching rows in both sheets
                sheet1.Rows(row1).Interior.Color = highlightColor
                sheet2.Rows(row2).Interior.Color = highlightColor
            End If
        Next row2
    Next row1
End Sub


Sub HightlightTransfer()
    
    Dim lightBlue As Long
    Dim darkBlue As Long
    Dim lightYellow As Long
    Dim lightRed As Long
    
    ' set color reference
    lightBlue = RGB(173, 216, 230)
    darkBlue = RGB(40, 110, 170)
    lightYellow = RGB(255, 255, 153)
    lightRed = RGB(255, 160, 160)
    
    ' find and hightlight matching rows
    ANZHighlightMatchingRows "C-ANZ-go", "C-ANZ-saving", lightBlue
    ANZHighlightMatchingRows "C-ANZ-go", "S-ANZ-loan", lightBlue
    ANZToWestpacHighlightMatchingRows "C-ANZ-go", "S-Westpac", lightRed
    ASBToWestpacHighlightMatchingRows "Y-ASB", "S-Westpac", lightYellow
    ANZLoanHighlightMatchingRows "C-ANZ-go", "S-ANZ-loan", darkBlue
    ASBLoanHighlightMatchingRows "Y-ASB", "S-ANZ-loan", lightBlue
    
End Sub
