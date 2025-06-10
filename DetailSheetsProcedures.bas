Sub CreateDetailOneFromTB1(targetWorkbook As Workbook, TB1Sheet As Worksheet)
    Dim detailWorksheet As Worksheet
    Dim lastRowTB1 As Long
    Dim row As Long
    
    Set detailWorksheet = InitializeDetailWorksheet(targetWorkbook)
    lastRowTB1 = TB1Sheet.Cells(TB1Sheet.Rows.Count, 2).End(xlUp).row
    
    ' Add header details
    detailWorksheet.Range("A4").Value = "รายละเอียดประกอบที่ 1"
    detailWorksheet.Range("A4").Font.Bold = True
    detailWorksheet.Range("I4").Value = "หน่วย : บาท"
    
    row = 5 ' Start row for details
    
    ' Process inventory if applicable
    If HasInventoryFromTB1(TB1Sheet) Then
        row = ProcessInventoryFromTB1(detailWorksheet, TB1Sheet, row)
    End If
    
    ' Process service costs
    row = ProcessServiceCostsFromTB1(detailWorksheet, TB1Sheet, row)
    
    ' Format worksheet
    FormatDetailWorksheet detailWorksheet
End Sub

Function InitializeDetailWorksheet(targetWorkbook As Workbook) As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = targetWorkbook.Sheets("DT1")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = targetWorkbook.Sheets.Add(After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count))
        ws.Name = "DT1"
    End If
    
    CreateHeader ws, "Details"
    
    Set InitializeDetailWorksheet = ws
End Function

Function HasInventoryFromTB1(TB1Sheet As Worksheet) As Boolean
    Dim i As Long
    Dim lastRow As Long
    
    lastRow = TB1Sheet.Cells(TB1Sheet.Rows.Count, 2).End(xlUp).row
    
    ' Check for any 1510 account
    For i = 2 To lastRow
        If TB1Sheet.Cells(i, 2).Value = "1510" Then
            HasInventoryFromTB1 = True
            Exit Function
        End If
    Next i
    
    ' If no 1510 found, check for purchase accounts (5010)
    For i = 2 To lastRow
        Dim accountCode As String
        accountCode = TB1Sheet.Cells(i, 2).Value
        
        ' Check for main purchase account and its subdivisions
        If Left(accountCode, 4) = "5010" Then
            If TB1Sheet.Cells(i, 5).Value <> 0 Or TB1Sheet.Cells(i, 4).Value <> 0 Then
                HasInventoryFromTB1 = True
                Exit Function
            End If
        End If
    Next i
    
    HasInventoryFromTB1 = False
End Function

Function ProcessInventoryFromTB1(ws As Worksheet, TB1Sheet As Worksheet, startRow As Long) As Long
    Dim row As Long
    Dim inventoryForSale As Double
    Dim endingInventory As Double
    Dim costOfGoodsSold As Double
    
    row = startRow
    
    ws.Cells(row, 2).Value = "ต้นทุนสินค้าที่ขาย"
    ws.Cells(row, 2).Font.Bold = True
    row = row + 1
    
    ' Add beginning inventory
    inventoryForSale = AddBeginningInventoryFromTB1(ws, TB1Sheet, row)
    row = row + 1
    
    ' Process purchases
    Dim totalPurchases As Double
    totalPurchases = ProcessPurchasesFromTB1(ws, TB1Sheet, row)
    inventoryForSale = inventoryForSale + totalPurchases
    
    ' Add "สินค้าไว้เพื่อขาย"
    ws.Cells(row, 2).Value = "สินค้าไว้เพื่อขาย"
    ws.Cells(row, 9).Value = inventoryForSale
    row = row + 1
    
    ' Add ending inventory
    endingInventory = AddEndingInventoryFromTB1(ws, TB1Sheet, row)
    row = row + 1
    
    ' Calculate and add "ต้นทุนสินค้าที่ขาย"
    costOfGoodsSold = inventoryForSale - endingInventory
    ws.Cells(row, 2).Value = "ต้นทุนสินค้าที่ขาย"
    ws.Cells(row, 9).Value = costOfGoodsSold
    ws.Cells(row, 2).Font.Bold = True
    row = row + 1
    
    ProcessInventoryFromTB1 = row
End Function

Function AddBeginningInventoryFromTB1(ws As Worksheet, TB1Sheet As Worksheet, row As Long) As Double
    Dim i As Long
    Dim lastRow As Long
    Dim amount As Double
    
    lastRow = TB1Sheet.Cells(TB1Sheet.Rows.Count, 2).End(xlUp).row
    
    ws.Cells(row, 2).Value = "สินค้าคงเหลือต้นงวด"
    For i = 2 To lastRow
        If TB1Sheet.Cells(i, 2).Value = "1510" Then
            ' Use Column C (previous period) for beginning inventory
            amount = TB1Sheet.Cells(i, 3).Value
            ws.Cells(row, 9).Value = amount
            AddBeginningInventoryFromTB1 = amount
            Exit Function
        End If
    Next i
End Function

Function ProcessPurchasesFromTB1(ws As Worksheet, TB1Sheet As Worksheet, ByRef row As Long) As Double
    Dim i As Long
    Dim lastRow As Long
    Dim accountCode As String
    Dim amount As Double
    Dim totalPurchases As Double
    
    lastRow = TB1Sheet.Cells(TB1Sheet.Rows.Count, 2).End(xlUp).row
    
    ' Handle 5010 first
    For i = 2 To lastRow
        If TB1Sheet.Cells(i, 2).Value = "5010" Then
            amount = TB1Sheet.Cells(i, 5).Value  ' Column E for current period
            ws.Cells(row, 2).Value = "บวก"
            ws.Cells(row, 3).Value = "ซื้อสินค้า"
            ws.Cells(row, 9).Value = amount
            totalPurchases = totalPurchases + amount
            row = row + 1
            Exit For
        End If
    Next i
    
    ' Handle other cases
    For i = 2 To lastRow
        accountCode = TB1Sheet.Cells(i, 2).Value
        Select Case accountCode
            Case "5010.1", "5010.2"
                amount = -TB1Sheet.Cells(i, 6).Value  ' UPDATED: Use Column F with negative sign
                ws.Cells(row, 2).Value = "หัก"
                ws.Cells(row, 3).Value = IIf(accountCode = "5010.1", "ส่วนลดสินค้า", "ส่วนลดรับ")
                ws.Cells(row, 9).Value = amount
                totalPurchases = totalPurchases - amount
                row = row + 1
            Case "5010.3"
                amount = TB1Sheet.Cells(i, 6).Value  ' UPDATED: Column F for freight in
                ws.Cells(row, 2).Value = "บวก"
                ws.Cells(row, 3).Value = "ค่าขนส่งเข้า"
                ws.Cells(row, 9).Value = amount
                totalPurchases = totalPurchases + amount
                row = row + 1
            Case "5010.4"
                amount = TB1Sheet.Cells(i, 5).Value  ' Column E for current period
                ws.Cells(row, 3).Value = "ค่าแรงงานทางตรง"
                ws.Cells(row, 9).Value = amount
                totalPurchases = totalPurchases + amount
                row = row + 1
        End Select
    Next i
    
    ' Add total purchases
    ws.Cells(row, 3).Value = "รวมซื้อสุทธิ"
    ws.Cells(row, 9).Value = totalPurchases
    row = row + 1
    
    ProcessPurchasesFromTB1 = totalPurchases
End Function

Function AddEndingInventoryFromTB1(ws As Worksheet, TB1Sheet As Worksheet, row As Long) As Double
    Dim i As Long
    Dim lastRow As Long
    Dim firstOccurrence As Boolean
    Dim endingInventory As Double
    
    lastRow = TB1Sheet.Cells(TB1Sheet.Rows.Count, 2).End(xlUp).row
    firstOccurrence = True
    
    ws.Cells(row, 2).Value = "หัก สินค้าคงเหลือปลายงวด"
    For i = 2 To lastRow
        If TB1Sheet.Cells(i, 2).Value = "1510" Then
            If firstOccurrence Then
                firstOccurrence = False
            Else
                endingInventory = TB1Sheet.Cells(i, 5).Value  ' Column E for current period
                ws.Cells(row, 9).Value = endingInventory
                AddEndingInventoryFromTB1 = endingInventory
                Exit Function
            End If
        End If
    Next i
End Function

Function ProcessServiceCostsFromTB1(ws As Worksheet, TB1Sheet As Worksheet, startRow As Long) As Long
    Dim serviceCosts(10) As Double
    Dim serviceNames(10) As String
    Dim hasAmount(10) As Boolean
    Dim i As Long, row As Long
    Dim lastRow As Long
    Dim accountCode As String, accountName As String, mainCode As String
    Dim amount As Double
    Dim index As Integer
    Dim totalServiceCost As Double
    
    lastRow = TB1Sheet.Cells(TB1Sheet.Rows.Count, 2).End(xlUp).row
    row = startRow
    
    ' Initialize arrays
    For i = 1 To 10
        serviceCosts(i) = 0
        serviceNames(i) = ""
        hasAmount(i) = False
    Next i
    
    ' Process all accounts
    For i = 2 To lastRow
        accountCode = TB1Sheet.Cells(i, 2).Value
        accountName = TB1Sheet.Cells(i, 1).Value
        amount = TB1Sheet.Cells(i, 5).Value  ' Column E for current period
        
        ' Check for service cost accounts (5011-5020 and their decimals)
        If Left(accountCode, 3) = "501" Then
            mainCode = Left(accountCode, 4)  ' Get the main account code (e.g., 5011)
            If Val(Right(mainCode, 2)) >= 11 And Val(Right(mainCode, 2)) <= 20 Then
                index = Val(Right(mainCode, 2)) - 10  ' Index for arrays (1 to 10)
                
                ' Always store the main account name
                If InStr(accountCode, ".") = 0 Then  ' If it's a main account
                    serviceNames(index) = accountName
                End If
                
                ' Store the amount, whether it's from the main account or a decimal
                If amount <> 0 Then
                    serviceCosts(index) = serviceCosts(index) + amount  ' Add to existing amount
                    hasAmount(index) = True
                End If
            End If
        End If
    Next i
    
    ' Display service costs
    If Application.WorksheetFunction.Sum(serviceCosts) > 0 Then
        ws.Cells(row, 2).Value = "ต้นทุนบริการ"
        ws.Cells(row, 2).Font.Bold = True
        row = row + 1
        
        For i = 1 To 10
            If hasAmount(i) Then
                ws.Cells(row, 3).Value = serviceNames(i)
                ws.Cells(row, 9).Value = serviceCosts(i)
                totalServiceCost = totalServiceCost + serviceCosts(i)
                row = row + 1
            End If
        Next i
        
        ' Display total service cost
        ws.Cells(row, 3).Value = "รวมต้นทุนบริการ"
        ws.Cells(row, 9).Value = totalServiceCost
        ws.Cells(row, 3).Font.Bold = True
        
        ' Add borders to the total cell
        With ws.Cells(row, 9)
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlDouble
        End With
        
        row = row + 1
    End If
    
    ProcessServiceCostsFromTB1 = row
End Function

Sub FormatDetailWorksheet(ws As Worksheet)
    ' Apply Thai Sarabun font and font size 14 to the detail worksheet
    ws.Cells.Font.Name = "TH Sarabun New"
    ws.Cells.Font.Size = 14
    
    ' Set number format to use comma style for column I in the detail worksheet
    ws.Columns("I").NumberFormat = "#,##0.00"
End Sub

Sub CreateDetailTwoFromTB1(targetWorkbook As Workbook, TB1Sheet As Worksheet)
    Dim detailWorksheet As Worksheet
    Dim lastRowTB1 As Long
    Dim i As Long
    Dim accountCode As String
    Dim accountName As String
    Dim amount As Double
    Dim row As Long
    Dim dataStartRow As Long
    Dim financialCosts As Double
    
    ' Initialize and set up worksheet
    Set detailWorksheet = InitializeWorksheetTwo(targetWorkbook)
    lastRowTB1 = TB1Sheet.Cells(TB1Sheet.Rows.Count, 2).End(xlUp).row
    
    ' Create headers
    CreateDetailHeaders detailWorksheet
    
    ' Initialize the starting row for account details
    row = 6
    dataStartRow = row
    
    ' Process TB1 data
    For i = 2 To lastRowTB1
        accountCode = TB1Sheet.Cells(i, 2).Value
        accountName = TB1Sheet.Cells(i, 1).Value
        amount = TB1Sheet.Cells(i, 5).Value  ' Column E for current period
        
        ' Check if the account code is within our range and not 5910
        If accountCode >= "5300" And accountCode <= "5999" And accountCode <> "5910" Then
            If accountCode >= "5360" And accountCode <= "5364" Then
                ' Sum financial costs separately
                financialCosts = financialCosts + amount
            Else
                ' Process other accounts
                ProcessAccountFromTB1 detailWorksheet, accountCode, accountName, amount, row
            End If
        End If
    Next i
    
    ' Add "รวม" row and calculate totals
    row = AddTotalRow(detailWorksheet, dataStartRow, row)
    
    ' Add financial costs row
    row = AddFinancialCostsRow(detailWorksheet, financialCosts, row)
    
    ' Apply formatting
    FormatDetailWorksheetTwo detailWorksheet, dataStartRow, row
End Sub

Function InitializeWorksheetTwo(targetWorkbook As Workbook) As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = targetWorkbook.Sheets("DT2")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = targetWorkbook.Sheets.Add(After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count))
        ws.Name = "DT2"
    End If
    
    CreateHeader ws, "Details"
    
    Set InitializeWorksheetTwo = ws
End Function

Sub CreateDetailHeaders(ws As Worksheet)
    With ws.Range("A4")
        .Value = "รายละเอียดประกอบที่ 2"
        .Font.Bold = True
    End With
    
    With ws
        .Range("A5:F5").Merge
        .Range("A5").Value = "ค่าใช้จ่ายในการขายและบริหาร"
        .Range("A5").HorizontalAlignment = xlLeft
    
        .Range("G5").Value = "ค่าใช้จ่ายในการขาย"
        .Range("H5").Value = "ค่าใช้จ่ายในการบริหาร"
        .Range("I5").Value = "ค่าใช้จ่ายอื่น"
    
        With .Range("A5:I5")
            .Font.Bold = True
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
        
        .Range("G5:I5").HorizontalAlignment = xlCenter
    End With
End Sub

Sub ProcessAccountFromTB1(ws As Worksheet, accountCode As String, accountName As String, amount As Double, ByRef row As Long)
    With ws.Range(ws.Cells(row, 1), ws.Cells(row, 6))
        .Merge
        .Value = accountName
        .HorizontalAlignment = xlLeft
    End With
    
    ' Initialize all columns with zero
    ws.Cells(row, 7).Value = 0
    ws.Cells(row, 8).Value = 0
    ws.Cells(row, 9).Value = 0
    
    ' Check if the account code is not in the financial costs range (5360-5364)
    If Not (accountCode >= "5360" And accountCode <= "5364") Then
        Select Case Left(accountCode, 4)
            Case "5300" To "5311"
                ' Add the amount to "ค่าใช้จ่ายในการขาย" (column G)
                ws.Cells(row, 7).Value = amount
            Case "5312" To "5350"
                ' Add the amount to "ค่าใช้จ่ายในการบริหาร" (column H)
                ws.Cells(row, 8).Value = amount
            Case "5351" To "5359", "5365" To "5999"
                ' Add the amount to "ค่าใช้จ่ายอื่น" (column I)
                ws.Cells(row, 9).Value = amount
        End Select
        
        row = row + 1
    End If
End Sub

Function AddTotalRow(ws As Worksheet, dataStartRow As Long, row As Long) As Long
    Dim i As Long ' Declare the loop variable
    
    With ws.Range(ws.Cells(row, 1), ws.Cells(row, 6))
        .Merge
        .Value = "รวม"
        .HorizontalAlignment = xlLeft
        .Font.Bold = True
    End With
    
    For i = 7 To 9
        ws.Cells(row, i).Formula = "=SUM(" & ws.Cells(dataStartRow, i).Address & ":" & ws.Cells(row - 1, i).Address & ")"
        ws.Cells(row, i).Font.Bold = True
    Next i
    
    AddTotalRow = row + 1
End Function

Function AddFinancialCostsRow(ws As Worksheet, financialCosts As Double, row As Long) As Long
    
    With ws.Range(ws.Cells(row, 1), ws.Cells(row, 6))
        .Merge
        .Value = "ค่าใช้จ่ายต้นทุนทางการเงิน"
        .HorizontalAlignment = xlLeft
        .Font.Bold = False
    End With
    
    ' Initialize all columns with zero
    ws.Cells(row, 7).Value = 0
    ws.Cells(row, 8).Value = 0
    ws.Cells(row, 9).Value = financialCosts
    ws.Cells(row, 9).Font.Bold = False
    
    AddFinancialCostsRow = row + 1
End Function

Sub FormatDetailWorksheetTwo(ws As Worksheet, dataStartRow As Long, lastRow As Long)
    ' Apply Thai Sarabun font and font size 14 to the detail worksheet
    ws.Cells.Font.Name = "TH Sarabun New"
    ws.Cells.Font.Size = 14
    
    ' Set number format to use comma style for columns G, H, and I in the detail worksheet
    ws.Columns("G:I").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    
    ' Set column widths
    ws.Columns("A").ColumnWidth = 13
    ws.Columns("B").ColumnWidth = 8
    ws.Columns("C:F").ColumnWidth = 3.5
    ws.Columns("G").ColumnWidth = 18
    ws.Columns("H").ColumnWidth = 20
    ws.Columns("I").ColumnWidth = 18
    
    ' Add borders to the table
    With ws.Range(ws.Cells(5, 1), ws.Cells(lastRow - 1, 9))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End With
    
    ' Add double bottom border for the total row and financial costs row
    ws.Range(ws.Cells(lastRow - 2, 1), ws.Cells(lastRow - 2, 9)).Borders(xlEdgeBottom).LineStyle = xlDouble
    ws.Range(ws.Cells(lastRow - 1, 1), ws.Cells(lastRow - 1, 9)).Borders(xlEdgeBottom).LineStyle = xlDouble
End Sub
