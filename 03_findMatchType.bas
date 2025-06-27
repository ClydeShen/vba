Private Sub FindMatchingCell(ByVal searchStr As String, ByVal startRow As Long, ByRef resultCell As Range)

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
 Private Sub SetTypeCell(ByVal emptyCell As Range, ByVal typeCell As Range)
    If Not typeCell Is Nothing Then
        Range(typeCell.Address).Select
        Selection.Copy
        Range(emptyCell.Address).Select
        ActiveSheet.Paste
    Else
        Debug.Print emptyCell.Address
    End If
End Sub


Sub HightlightTypes()
    Dim summary As Worksheet
    Dim dict As Object
    Dim dataRow As Long
    Dim midRow As Long
    
    Set summary = ThisWorkbook.Sheets("Summary")
    summary.Select
    dataRow = summary.Cells(summary.Rows.Count, 1).End(xlUp).Row
    midRow = dataRow + 1

    Dim i As Long
    Dim j As Long

    Dim categoryCell As Range

    Dim otherPartyCell As Range
    Dim typeCell As Range
    Dim descriptionCell As Range
    Dim referenceCell As Range
    Dim particularsCell As Range
    Dim analysisCodeCell As Range




    For j = 1 To midRow
        
        Set otherPartyCell = summary.Cells(j, 2)
        Set typeCell = summary.Cells(j, 6)
        Set descriptionCell = summary.Cells(j, 7)
        Set referenceCell = summary.Cells(j, 8)
        Set particularsCell = summary.Cells(j, 9)
        Set analysisCodeCell = summary.Cells(j, 10)

        ' category: groceries
        If InStr(1, otherPartyCell, "Woolworths", vbTextCompare) > 0 _
            Or (InStr(1, otherPartyCell, "Pak N Save", vbTextCompare) > 0 And _
              InStr(1, otherPartyCell, "Pak N Save Fuel", vbTextCompare) = 0) _
            Or InStr(1, otherPartyCell, "New World", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Taiping", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Tai Ping", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Golden Apple", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Wang Foodmarket", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Freshchoice", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "DH Supermarket", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Wang Food", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Young", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Rui Feng Kitchen", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "FoodieAsianSuperm", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Seasons Markets", vbTextCompare) > 0 Then

            FindMatchingCell "Groceries", midRow, categoryCell
            SetTypeCell typeCell, categoryCell
        End If

        ' category: Eating out
         If InStr(1, otherPartyCell, "4140Edison", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "9180Edison", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Chen,Dong",vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Hu,Nan", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Hungrypanda", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Golden City Cuisine", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Gui Rice Noodle", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Hello Mister Wyny", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Jinweide Noodle", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Kingsmade Noodle", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "1981 Noodle House", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "NO1 BEEF RAMAN", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "No1 Beef Raman", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Sunnynook Fast Food", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Mr LI Takeaway", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Daily Bread Galway", vbTextCompare) > 0 _ 
            Or InStr(1, otherPartyCell, "The Coffee Club", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Chongqing Noodles", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Doordash", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Double Happy", vbTextCompare) > 0 Then
            FindMatchingCell "Eating out", midRow, categoryCell
            SetTypeCell typeCell, categoryCell
        End If

        ' category: Home & contents Insurance
        If InStr(1, otherPartyCell, "AA Insurance Pre", vbTextCompare) > 0 Then
            FindMatchingCell "Home & contents", midRow, categoryCell
            SetTypeCell typeCell, categoryCell
        End If
        
        ' category: Health Insurance
        If InStr(1, otherPartyCell, "Southern Cross", vbTextCompare) > 0 _ 
            Or InStr(1, analysisCodeCell, "Southern Cross", vbTextCompare) > 0 Then
            FindMatchingCell "Health", midRow, categoryCell
            SetTypeCell typeCell, categoryCell
        End If

        ' category: Mortgage repayments
        If InStr(1, referenceCell, "LOAN PAYMT", vbTextCompare) > 0  Then
            FindMatchingCell "Mortgage repayments", midRow, categoryCell
            SetTypeCell typeCell, categoryCell
        End If

        ' category: Electricity & Gas & Internet
        If InStr(1, otherPartyCell, "Skinny Fixed", vbTextCompare) > 0  _
            Or InStr(1, otherPartyCell, "Flick Energy Limited", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Rockgas Limited", vbTextCompare) > 0 Then
            FindMatchingCell "Electricity & Gas & Internet", midRow, categoryCell
            SetTypeCell typeCell, categoryCell
        End If

        ' category: Travel
        If InStr(1, otherPartyCell, "AT HOP", vbTextCompare) > 0  _ 
            Or InStr(1, otherPartyCell, "Gull", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "BP", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "KIWI FUELS", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Caltex", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "AT Public", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Z Constell", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Mobil Wairau", vbTextCompare) > 0 _
            Or InStr(1, analysisCodeCell, "AT PUBLIC TRANSPORT", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Pak N Save Fuel", vbTextCompare) > 0 Then
            FindMatchingCell "Travel", midRow, categoryCell
            SetTypeCell typeCell, categoryCell
        End If

        ' category: Telephone
        If InStr(1, otherPartyCell, "One New Zealand", vbTextCompare) > 0 _ 
          Or InStr(1, otherPartyCell, "Rocket Mobile", vbTextCompare) > 0 Then
            FindMatchingCell "Telephone", midRow, categoryCell
            SetTypeCell typeCell, categoryCell
        End If

        ' category: Council Rate
        If InStr(1, otherPartyCell, "Auckland Council", vbTextCompare) > 0 Then
            FindMatchingCell "Council Rate", midRow, typeCell
            SetTypeCell typeCell, typeCell
        End If

        ' category: Water
        If InStr(1, otherPartyCell, "Watercare", vbTextCompare) > 0 Then
            FindMatchingCell "Water", midRow, categoryCell
            SetTypeCell typeCell, categoryCell
        End If

        ' category: Entertainment subscription
        If InStr(1, otherPartyCell, "Google YouTube", vbTextCompare) > 0 _ 
          Or InStr(1, otherPartyCell, "Google Lumosity", vbTextCompare) > 0 Then
            FindMatchingCell "Entertainment subscriptions", midRow, categoryCell
            SetTypeCell typeCell, categoryCell
        End If

   
        ' category: Home maintenance/repairs
        If InStr(1, otherPartyCell, "Bunnings", vbTextCompare) > 0  _ 
            Or InStr(1, otherPartyCell, "Kmart", vbTextCompare) > 0 Then
            FindMatchingCell "Home maintenance/repairs", midRow, categoryCell
            SetTypeCell typeCell, categoryCell
        End If

        ' category: Vehicle maintenance/repairs
        If InStr(1, otherPartyCell, "SUPERCHEAP", vbTextCompare) > 0 Then
            FindMatchingCell "Vehicle maintenance/repairs", midRow, categoryCell
            SetTypeCell typeCell, categoryCell
        End If
        
        ' category: Salary
        If InStr(1, analysisCodeCell, "FROM HAWKINS LIMITED", vbTextCompare) > 0 _ 
          Or InStr(1, referenceCell, "DELOITTE SAL", vbTextCompare) > 0 Then
            FindMatchingCell "Salary", midRow, categoryCell
            SetTypeCell typeCell, categoryCell
        End If

        ' category: Rent
        If InStr(1, otherPartyCell, "Tong, Z", vbTextCompare) > 0 _ 
          Or InStr(1, otherPartyCell, "Wang,", vbTextCompare) > 0 _ 
          Or InStr(1, referenceCell, "19b", vbTextCompare) > 0 _ 
          Or InStr(1, particularsCell, "rent", vbTextCompare) > 0 Then
            FindMatchingCell "Rent", midRow, categoryCell
            SetTypeCell typeCell, categoryCell
        End If

        ' category: Family Visit & Event
        If InStr(1, otherPartyCell, "balancing budget", vbTextCompare) > 0 Then
            FindMatchingCell "Family Visit & Event", midRow, categoryCell
            SetTypeCell typeCell, categoryCell
        End If

        ' category: Investment
        If InStr(1, otherPartyCell, "mylotto.co.nz", vbTextCompare) > 0 _ 
            Or InStr(1, otherPartyCell, "EF207562 Wealth Mgmt", vbTextCompare) > 0 _
            Or InStr(1, otherPartyCell, "Westpac Bonus Saver", vbTextCompare) > 0 _
            Or InStr(1, analysisCodeCell, "SUPERLIFE", vbTextCompare) > 0 _
            Or InStr(1, analysisCodeCell, "SMART INVEST", vbTextCompare) > 0 Then
            FindMatchingCell "Investment", midRow, categoryCell
            SetTypeCell typeCell, categoryCell
        End If
            

        ' category: Personal care
        If InStr(1, otherPartyCell, "CW ", vbTextCompare) > 0  _
            Or InStr(1, otherPartyCell, "Chemist Warehouse", vbTextCompare) > 0 Then
            FindMatchingCell "Personal care", midRow, categoryCell
            SetTypeCell typeCell, categoryCell
        End If

        ' category: Car/Motor
        If InStr(1, otherPartyCell, "CARD 0780 AMI INSURANC", vbTextCompare) > 0 _
             Then
            FindMatchingCell "Car/Motor", midRow, categoryCell
            SetTypeCell typeCell, categoryCell
        End If

    Next j
            
End Sub
