Sub GenerateProfitLossStatement()
    Dim ws As Worksheet
    Dim trialPLSheets As Collection
    
    ' Get the TrialPL sheets
    Set trialPLSheets = GetWorksheetsWithPrefix("Trial PL")
    
    ' Check the number of TrialPL sheets
    If trialPLSheets.Count = 1 Then
        CreateSingleYearProfitLossStatement trialPLSheets(1)
    ElseIf trialPLSheets.Count = 2 Then
        CreateMultiYearProfitLossStatement trialPLSheets
    Else
        MsgBox "Invalid number of TrialPL sheets. Please ensure there are either one or two sheets.", vbExclamation
    End If
End Sub

Sub CreateSingleYearProfitLossStatement(trialPLSheet As Worksheet)
    Dim wsPL As Worksheet
    Dim row As Long
    Dim totalRevenue As Double
    Dim totalExpenses As Double
    Dim targetWorkbook As Workbook
    
    Set targetWorkbook = trialPLSheet.Parent
    
    ' Create new sheet for the Profit and Loss statement
    Set wsPL = targetWorkbook.Sheets.Add(After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count))
    wsPL.Name = "PL"
    
    ' Create the header
    CreateHeader wsPL, "Profit and Loss Statement"
    
    ' Add details
    row = 5 ' Start below the header
    
    ' Add "หมายเหตุ" and "หน่วย:บาท"
    AddProfitLossHeaderDetails wsPL, row
    
    row = row + 1
    
    ' Add revenue section
    totalRevenue = AddRevenueAccounts(wsPL, trialPLSheet, row)
    
    ' Add a blank row for better readability
    row = row + 1
    
    ' Add expense section
    totalExpenses = AddExpenseAccounts(wsPL, trialPLSheet, row)
    
    ' Calculate and add net profit/loss before financial costs and income tax
    With wsPL.Cells(row, 2)
        .Value = "กำไรก่อนต้นทุนทางการเงินและภาษีเงินได้"
        .Font.Bold = True
    End With
    
    Dim profitBeforeFinCostTaxRow As Long
    profitBeforeFinCostTaxRow = row
    With wsPL.Cells(row, 9)
        .Formula = "=" & wsPL.Names("RevenueTotalRow").RefersToRange.Cells(1).Address & "-" & _
                  wsPL.Names("ExpenseTotalRow").RefersToRange.Cells(1).Address
    End With
    row = row + 1
    
    ' Add financial costs
    Dim financialCostsRow As Long
    financialCostsRow = row
    wsPL.Cells(row, 2).Value = "ต้นทุนทางการเงิน"
    wsPL.Cells(row, 9).Value = GetFinancialCosts(trialPLSheet)
    With wsPL.Cells(row, 9)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    row = row + 1
    
    ' Calculate profit/loss before income tax
    With wsPL.Cells(row, 2)
        .Value = "กำไร(ขาดทุน)ก่อนภาษีเงินได้"
        .Font.Bold = True
    End With
    Dim profitBeforeTaxRow As Long
    profitBeforeTaxRow = row
    With wsPL.Cells(row, 9)
        .Formula = "=" & wsPL.Cells(profitBeforeFinCostTaxRow, 9).Address & "-" & _
                  wsPL.Cells(financialCostsRow, 9).Address
    End With
    row = row + 1
    
    ' Add income tax
    Dim incomeTaxRow As Long
    incomeTaxRow = row
    wsPL.Cells(row, 2).Value = "ภาษีเงินได้"
    wsPL.Cells(row, 9).Value = GetIncomeTax(trialPLSheet)
    row = row + 1
    
    ' Calculate and add net profit/loss
    With wsPL.Cells(row, 2)
        .Value = "กำไร(ขาดทุน)สุทธิ"
        .Font.Bold = True
    End With
    With wsPL.Cells(row, 9)
        .Formula = "=" & wsPL.Cells(profitBeforeTaxRow, 9).Address & "-" & _
                  wsPL.Cells(incomeTaxRow, 9).Address
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With
    
    ' Format the worksheet
    FormatPLWorksheet wsPL
End Sub

Sub CreateMultiYearProfitLossStatement(trialPLSheets As Collection)
    Dim wsPL As Worksheet
    Dim row As Long
    Dim totalRevenue As Double
    Dim totalRevenuePrevious As Double
    Dim totalExpenses As Double
    Dim totalExpensesPrevious As Double
    Dim infoSheet As Worksheet
    Dim year As String
    Dim PreviousYear As String
    Dim targetWorkbook As Workbook
    Dim financialCosts As Double
    Dim financialCostsPrevious As Double
    Dim incomeTax As Double
    Dim incomeTaxPrevious As Double
    
    Set targetWorkbook = trialPLSheets(1).Parent
    
    ' Get the year from the "Info" sheet
    Set infoSheet = targetWorkbook.Sheets("Info")
    year = infoSheet.Range("B3").Value
    PreviousYear = CStr(CLng(year) - 1)
    
    ' Create new sheet for the Profit and Loss statement
    Set wsPL = targetWorkbook.Sheets.Add(After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count))
    wsPL.Name = "PLM"
    
    ' Create the header
    CreateHeader wsPL, "Profit and Loss Statement"
    
    ' Add details
    row = 5 ' Start below the header
    
    ' Add "หมายเหตุ" and "หน่วย:บาท"
    AddProfitLossHeaderDetails wsPL, row
    
    row = row + 1
    
    ' Add years
    With wsPL.Cells(row, 7)
        .Value = year
        .Font.Underline = xlUnderlineStyleSingle
    End With
    
    With wsPL.Cells(row, 9)
        .Value = PreviousYear
        .Font.Underline = xlUnderlineStyleSingle
    End With
    
    row = row + 1
    
    ' Add revenue section
    totalRevenue = AddMultiYearRevenueAccounts(wsPL, trialPLSheets, row, totalRevenuePrevious)
    
    ' Add a blank row for better readability
    row = row + 1
    
    ' Add expense section
    totalExpenses = AddMultiYearExpenseAccounts(wsPL, trialPLSheets, row, totalExpensesPrevious)
    
    ' Calculate and add net profit/loss before financial costs and income tax using formula
    With wsPL.Cells(row, 2)
        .Value = "กำไรก่อนต้นทุนทางการเงินและภาษีเงินได้"
        .Font.Bold = True
    End With
    
    Dim profitBeforeFinCostTaxRow As Long
    profitBeforeFinCostTaxRow = row
    With wsPL.Cells(row, 7)
        .Formula = "=" & wsPL.Names("RevenueTotalRow").RefersToRange.Cells(1).Address & "-" & _
                  wsPL.Names("ExpenseTotalRow").RefersToRange.Cells(1).Address
    End With
    With wsPL.Cells(row, 9)
        .Formula = "=" & wsPL.Names("RevenueTotalRow").RefersToRange.Cells(3).Address & "-" & _
                  wsPL.Names("ExpenseTotalRow").RefersToRange.Cells(3).Address
    End With
    row = row + 1
    
    ' Add financial costs (store row for reference)
    Dim financialCostsRow As Long
    financialCostsRow = row
    wsPL.Cells(row, 2).Value = "ต้นทุนทางการเงิน"
    wsPL.Cells(row, 7).Value = GetFinancialCosts(trialPLSheets(1))
    wsPL.Cells(row, 9).Value = GetFinancialCosts(trialPLSheets(2))
    With wsPL.Range(wsPL.Cells(row, 7), wsPL.Cells(row, 9))
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    row = row + 1
    
    ' Calculate profit/loss before income tax using formula
    With wsPL.Cells(row, 2)
        .Value = "กำไร(ขาดทุน)ก่อนภาษีเงินได้"
        .Font.Bold = True
    End With
    Dim profitBeforeTaxRow As Long
    profitBeforeTaxRow = row
    With wsPL.Cells(row, 7)
        .Formula = "=" & wsPL.Cells(profitBeforeFinCostTaxRow, 7).Address & "-" & _
                  wsPL.Cells(financialCostsRow, 7).Address
    End With
    With wsPL.Cells(row, 9)
        .Formula = "=" & wsPL.Cells(profitBeforeFinCostTaxRow, 9).Address & "-" & _
                  wsPL.Cells(financialCostsRow, 9).Address
    End With
    row = row + 1
    
    ' Add income tax (store row for reference)
    Dim incomeTaxRow As Long
    incomeTaxRow = row
    wsPL.Cells(row, 2).Value = "ภาษีเงินได้"
    wsPL.Cells(row, 7).Value = GetIncomeTax(trialPLSheets(1))
    wsPL.Cells(row, 9).Value = GetIncomeTax(trialPLSheets(2))
    row = row + 1
    
    ' Calculate and add net profit/loss using formula
    With wsPL.Cells(row, 2)
        .Value = "กำไร(ขาดทุน)สุทธิ"
        .Font.Bold = True
    End With
    With wsPL.Cells(row, 7)
        .Formula = "=" & wsPL.Cells(profitBeforeTaxRow, 7).Address & "-" & _
                  wsPL.Cells(incomeTaxRow, 7).Address
    End With
    With wsPL.Cells(row, 9)
        .Formula = "=" & wsPL.Cells(profitBeforeTaxRow, 9).Address & "-" & _
                  wsPL.Cells(incomeTaxRow, 9).Address
    End With
    With wsPL.Range(wsPL.Cells(row, 7), wsPL.Cells(row, 9))
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With
    
    ' Format the worksheet
    FormatPLWorksheet wsPL
End Sub

Function AddRevenueAccounts(ws As Worksheet, trialPLSheet As Worksheet, ByRef row As Long) As Double
    Dim lastRow As Long
    Dim i As Long
    Dim totalRevenue As Double
    Dim salesRevenue As Double
    Dim otherIncome As Double
    Dim accountCode As String
    Dim accountName As String
    Dim amount As Double
    
    ' Store starting row for revenue details
    Dim revenueStartRow As Long
    revenueStartRow = row + 1  ' Skip the "รายได้" header row
    
    lastRow = trialPLSheet.Cells(trialPLSheet.Rows.Count, 2).End(xlUp).row
    totalRevenue = 0
    salesRevenue = 0
    otherIncome = 0
    
    ' Add "รายได้" header
    With ws.Cells(row, 2)
        .Value = "รายได้"
        .Font.Bold = True
    End With
    row = row + 1
    
    ' Add "รายได้จากการขายหรือการให้บริการ"
    ws.Cells(row, 3).Value = "รายได้จากการขายหรือการให้บริการ"
    
    ' Loop through trial balance and sum amounts for revenue accounts
    For i = 2 To lastRow
        accountCode = trialPLSheet.Cells(i, 2).Value
        If accountCode >= "4010" And accountCode <= "4019" Then
            accountName = trialPLSheet.Cells(i, 1).Value
            amount = trialPLSheet.Cells(i, 7).Value
            salesRevenue = salesRevenue + amount
        ElseIf accountCode >= "4020" And accountCode <= "4210" Then
            accountName = trialPLSheet.Cells(i, 1).Value
            amount = trialPLSheet.Cells(i, 6).Value
            otherIncome = otherIncome + amount
        End If
    Next i
    
    ' Add total amount for sales revenue
    ws.Cells(row, 9).Value = salesRevenue
    row = row + 1
    
    ' Add "รายได้อื่น"
    ws.Cells(row, 3).Value = "รายได้อื่น"
    ws.Cells(row, 9).Value = otherIncome
    row = row + 1
    
    ' Add total revenue with formula
    With ws.Cells(row, 2)
        .Value = "รวมรายได้"
        .Font.Bold = True
    End With
    With ws.Cells(row, 9)
        .Formula = "=SUM(" & ws.Range(ws.Cells(revenueStartRow, 9), ws.Cells(row - 1, 9)).Address & ")"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    
    ' Store the row reference in a name for later use
    ws.Names.Add "RevenueTotalRow", ws.Cells(row, 9)
    row = row + 1
    
    AddRevenueAccounts = totalRevenue
End Function

Function AddMultiYearRevenueAccounts(ws As Worksheet, trialPLSheets As Collection, ByRef row As Long, ByRef previousYearTotal As Double) As Double
    Dim lastRow As Long
    Dim i As Long
    Dim totalRevenue As Double
    Dim totalRevenuePrevious As Double
    Dim salesRevenue As Double
    Dim salesRevenuePrevious As Double
    Dim otherIncome As Double
    Dim otherIncomePrevious As Double
    Dim accountCode As String
    Dim accountName As String
    Dim amount As Double
    Dim amountPrevious As Double
    
    lastRow = trialPLSheets(1).Cells(trialPLSheets(1).Rows.Count, 2).End(xlUp).row
    totalRevenue = 0
    totalRevenuePrevious = 0
    salesRevenue = 0
    salesRevenuePrevious = 0
    otherIncome = 0
    otherIncomePrevious = 0
    
    ' Store starting row for revenue details
    Dim revenueStartRow As Long
    revenueStartRow = row + 1  ' Skip the "รายได้" header row
    
    ' Add "รายได้" header
    With ws.Cells(row, 2)
        .Value = "รายได้"
        .Font.Bold = True
    End With
    row = row + 1
    
    ' Add "รายได้จากการขายหรือการให้บริการ"
    ws.Cells(row, 3).Value = "รายได้จากการขายหรือการให้บริการ"
    
    ' Loop through trial balance and sum amounts for revenue accounts
    For i = 2 To lastRow
        accountCode = trialPLSheets(1).Cells(i, 2).Value
        If accountCode >= "4010" And accountCode <= "4019" Then
            accountName = trialPLSheets(1).Cells(i, 1).Value
            amount = trialPLSheets(1).Cells(i, 7).Value ' Note: Using column G for sales revenue accounts
            amountPrevious = GetAmountFromPreviousPeriodPL(trialPLSheets(2), accountCode, True)
            salesRevenue = salesRevenue + amount
            salesRevenuePrevious = salesRevenuePrevious + amountPrevious
        ElseIf accountCode >= "4020" And accountCode <= "4210" Then
            accountName = trialPLSheets(1).Cells(i, 1).Value
            amount = trialPLSheets(1).Cells(i, 6).Value ' Using column F for other income accounts
            amountPrevious = GetAmountFromPreviousPeriodPL(trialPLSheets(2), accountCode, False)
            otherIncome = otherIncome + amount
            otherIncomePrevious = otherIncomePrevious + amountPrevious
        End If
    Next i
    
    ' Add total amount for sales revenue
    ws.Cells(row, 7).Value = salesRevenue
    ws.Cells(row, 9).Value = salesRevenuePrevious
    row = row + 1
    
    ' Add "รายได้อื่น"
    ws.Cells(row, 3).Value = "รายได้อื่น"
    ws.Cells(row, 7).Value = otherIncome
    ws.Cells(row, 9).Value = otherIncomePrevious
    row = row + 1
    
    totalRevenue = salesRevenue + otherIncome
    totalRevenuePrevious = salesRevenuePrevious + otherIncomePrevious
    
    ' Add total revenue
    With ws.Cells(row, 2)
        .Value = "รวมรายได้"
        .Font.Bold = True
    End With
    With ws.Cells(row, 7)
        .Formula = "=SUM(" & ws.Range(ws.Cells(revenueStartRow, 7), ws.Cells(row - 1, 7)).Address & ")"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    With ws.Cells(row, 9)
        .Formula = "=SUM(" & ws.Range(ws.Cells(revenueStartRow, 9), ws.Cells(row - 1, 9)).Address & ")"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    ' Store the row number for future reference
    Dim revenueTotal As Range
    Set revenueTotal = ws.Cells(row, 7).Resize(1, 3)
    
    row = row + 1
    
    previousYearTotal = totalRevenuePrevious
    AddMultiYearRevenueAccounts = totalRevenue
    
    ' Store the row reference in a name for later use
    ws.Names.Add "RevenueTotalRow", revenueTotal
    
End Function

Function GetCostOfGoodsSold(trialPLSheet As Worksheet) As Double
    Dim lastRow As Long
    Dim i As Long
    Dim cogsCount As Integer
    
    lastRow = trialPLSheet.Cells(trialPLSheet.Rows.Count, 2).End(xlUp).row
    cogsCount = 0
    
    For i = 2 To lastRow
        If trialPLSheet.Cells(i, 2).Value = "1510" Then
            cogsCount = cogsCount + 1
            If cogsCount = 2 Then
                GetCostOfGoodsSold = trialPLSheet.Cells(i, 7).Value ' Using column G for cost of goods sold
                Exit Function
            End If
        End If
    Next i
    
    GetCostOfGoodsSold = 0 ' Return 0 if second occurrence of 1510 is not found
End Function

Function GetMultiYearCostOfGoodsSold(trialPLSheets As Collection, ByRef costOfGoodsSoldPrevious As Double) As Double
    Dim currentYearCOGS As Double
    Dim previousYearCOGS As Double
    
    currentYearCOGS = GetCostOfGoodsSold(trialPLSheets(1))
    previousYearCOGS = GetCostOfGoodsSold(trialPLSheets(2))
    
    costOfGoodsSoldPrevious = previousYearCOGS
    GetMultiYearCostOfGoodsSold = currentYearCOGS
End Function

Function AddExpenseAccounts(ws As Worksheet, trialPLSheet As Worksheet, ByRef row As Long) As Double
    Dim lastRow As Long
    Dim i As Long
    Dim totalExpenses As Double
    Dim costOfGoodsSold As Double
    Dim sellingExpenses As Double
    Dim adminExpenses As Double
    Dim otherExpenses As Double
    Dim accountCode As String
    Dim accountName As String
    Dim amount As Double
    
    ' Store starting row for expense details
    Dim expenseStartRow As Long
    expenseStartRow = row + 1  ' Skip the "ค่าใช้จ่าย" header row
    
    lastRow = trialPLSheet.Cells(trialPLSheet.Rows.Count, 2).End(xlUp).row
    totalExpenses = 0
    sellingExpenses = 0
    adminExpenses = 0
    otherExpenses = 0
    
    ' Add "ค่าใช้จ่าย" header
    With ws.Cells(row, 2)
        .Value = "ค่าใช้จ่าย"
        .Font.Bold = True
    End With
    row = row + 1
    
    ' Add "ต้นทุนขายหรือต้นทุนการให้บริการ"
    ws.Cells(row, 3).Value = "ต้นทุนขายหรือต้นทุนการให้บริการ"
    
    ' Get cost of goods sold
    costOfGoodsSold = GetCostOfGoodsSold(trialPLSheet)
    ws.Cells(row, 9).Value = costOfGoodsSold
    row = row + 1
    
    ' Calculate other expenses
    For i = 2 To lastRow
        accountCode = trialPLSheet.Cells(i, 2).Value
        accountName = trialPLSheet.Cells(i, 1).Value
        amount = trialPLSheet.Cells(i, 6).Value ' Using column F for other expenses
        
        Select Case Left(accountCode, 4)
            Case "5310" To "5319"
                sellingExpenses = sellingExpenses + amount
            Case "5309", "5320" To "5350"
                adminExpenses = adminExpenses + amount
            Case "5351" To "5399"
                otherExpenses = otherExpenses + amount
        End Select
    Next i
    
    ' Add expense categories to the worksheet
    ws.Cells(row, 3).Value = "ค่าใช้จ่ายในการขาย"
    ws.Cells(row, 9).Value = sellingExpenses
    row = row + 1
    
    ws.Cells(row, 3).Value = "ค่าใช้จ่ายในการบริหาร"
    ws.Cells(row, 9).Value = adminExpenses
    row = row + 1
    
    ws.Cells(row, 3).Value = "ค่าใช้จ่ายอื่น"
    ws.Cells(row, 9).Value = otherExpenses
    row = row + 1
    
    ' Add total expenses with formula
    With ws.Cells(row, 2)
        .Value = "รวมค่าใช้จ่าย"
        .Font.Bold = True
    End With
    With ws.Cells(row, 9)
        .Formula = "=SUM(" & ws.Range(ws.Cells(expenseStartRow, 9), ws.Cells(row - 1, 9)).Address & ")"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    
    ' Store the row reference in a name for later use
    ws.Names.Add "ExpenseTotalRow", ws.Cells(row, 9)
    row = row + 1
    
    AddExpenseAccounts = totalExpenses
End Function

Function AddMultiYearExpenseAccounts(ws As Worksheet, trialPLSheets As Collection, ByRef row As Long, ByRef previousYearTotal As Double) As Double
    Dim lastRow As Long
    Dim i As Long
    Dim totalExpenses As Double
    Dim totalExpensesPrevious As Double
    Dim costOfGoodsSold As Double
    Dim costOfGoodsSoldPrevious As Double
    Dim sellingExpenses As Double
    Dim sellingExpensesPrevious As Double
    Dim adminExpenses As Double
    Dim adminExpensesPrevious As Double
    Dim otherExpenses As Double
    Dim otherExpensesPrevious As Double
    Dim accountCode As String
    Dim accountName As String
    Dim amount As Double
    Dim amountPrevious As Double
    
    lastRow = trialPLSheets(1).Cells(trialPLSheets(1).Rows.Count, 2).End(xlUp).row
    totalExpenses = 0
    totalExpensesPrevious = 0
    sellingExpenses = 0
    sellingExpensesPrevious = 0
    adminExpenses = 0
    adminExpensesPrevious = 0
    otherExpenses = 0
    otherExpensesPrevious = 0
    
    ' Store starting row for expense details
    Dim expenseStartRow As Long
    expenseStartRow = row + 1  ' Skip the "ค่าใช้จ่าย" header row
    
    ' Add "ค่าใช้จ่าย" header
    With ws.Cells(row, 2)
        .Value = "ค่าใช้จ่าย"
        .Font.Bold = True
    End With
    row = row + 1
    
    ' Add "ต้นทุนขายหรือต้นทุนการให้บริการ"
    ws.Cells(row, 3).Value = "ต้นทุนขายหรือต้นทุนการให้บริการ"
    
    ' Get cost of goods sold for both years
    costOfGoodsSold = GetMultiYearCostOfGoodsSold(trialPLSheets, costOfGoodsSoldPrevious)
    ws.Cells(row, 7).Value = costOfGoodsSold
    ws.Cells(row, 9).Value = costOfGoodsSoldPrevious
    row = row + 1
    
    ' Calculate other expenses for both years
    For i = 2 To lastRow
        accountCode = trialPLSheets(1).Cells(i, 2).Value
        accountName = trialPLSheets(1).Cells(i, 1).Value
        amount = trialPLSheets(1).Cells(i, 6).Value ' Using column F for other expenses
        amountPrevious = GetAmountFromPreviousPeriodPL(trialPLSheets(2), accountCode, False)
        
        Select Case Left(accountCode, 4)
            Case "5310" To "5319"
                sellingExpenses = sellingExpenses + amount
                sellingExpensesPrevious = sellingExpensesPrevious + amountPrevious
            Case "5309", "5320" To "5350"
                adminExpenses = adminExpenses + amount
                adminExpensesPrevious = adminExpensesPrevious + amountPrevious
            Case "5351" To "5399"
                otherExpenses = otherExpenses + amount
                otherExpensesPrevious = otherExpensesPrevious + amountPrevious
        End Select
    Next i
    
    ' Add expense categories to the worksheet
    ws.Cells(row, 3).Value = "ค่าใช้จ่ายในการขาย"
    ws.Cells(row, 7).Value = sellingExpenses
    ws.Cells(row, 9).Value = sellingExpensesPrevious
    row = row + 1
    
    ws.Cells(row, 3).Value = "ค่าใช้จ่ายในการบริหาร"
    ws.Cells(row, 7).Value = adminExpenses
    ws.Cells(row, 9).Value = adminExpensesPrevious
    row = row + 1
    
    ws.Cells(row, 3).Value = "ค่าใช้จ่ายอื่น"
    ws.Cells(row, 7).Value = otherExpenses
    ws.Cells(row, 9).Value = otherExpensesPrevious
    row = row + 1
    
    ' Calculate and add total expenses
    totalExpenses = costOfGoodsSold + sellingExpenses + adminExpenses + otherExpenses
    totalExpensesPrevious = costOfGoodsSoldPrevious + sellingExpensesPrevious + adminExpensesPrevious + otherExpensesPrevious
    With ws.Cells(row, 2)
        .Value = "รวมค่าใช้จ่าย"
        .Font.Bold = True
    End With
    With ws.Cells(row, 7)
        .Formula = "=SUM(" & ws.Range(ws.Cells(expenseStartRow, 7), ws.Cells(row - 1, 7)).Address & ")"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    With ws.Cells(row, 9)
        .Formula = "=SUM(" & ws.Range(ws.Cells(expenseStartRow, 9), ws.Cells(row - 1, 9)).Address & ")"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    ' Store the row number for future reference
    Dim expenseTotal As Range
    Set expenseTotal = ws.Cells(row, 7).Resize(1, 3)
    row = row + 1
    previousYearTotal = totalExpensesPrevious
    AddMultiYearExpenseAccounts = totalExpenses
    
    ' Store the row reference in a name for later use
    ws.Names.Add "ExpenseTotalRow", expenseTotal
    
End Function

Function GetFinancialCosts(trialPLSheet As Worksheet) As Double
    Dim lastRow As Long
    Dim i As Long
    Dim totalFinancialCosts As Double
    Dim accountCode As String
    
    lastRow = trialPLSheet.Cells(trialPLSheet.Rows.Count, 2).End(xlUp).row
    totalFinancialCosts = 0
    
    For i = 2 To lastRow
        accountCode = trialPLSheet.Cells(i, 2).Value
        If accountCode >= "5360" And accountCode <= "5364" Then
            totalFinancialCosts = totalFinancialCosts + trialPLSheet.Cells(i, 6).Value
        End If
    Next i
    
    GetFinancialCosts = totalFinancialCosts
End Function

Function GetIncomeTax(trialPLSheet As Worksheet) As Double
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = trialPLSheet.Cells(trialPLSheet.Rows.Count, 2).End(xlUp).row
    
    For i = 2 To lastRow
        If trialPLSheet.Cells(i, 2).Value = "5910" Then
            GetIncomeTax = trialPLSheet.Cells(i, 7).Value  ' Changed from column 6 (F) to column 7 (G)
            Exit Function
        End If
    Next i
    
    GetIncomeTax = 0 ' Return 0 if account code 5910 is not found
End Function




Function GetAmountFromPreviousPeriodPL(trialPLSheet As Worksheet, accountCode As String, isSalesRevenue As Boolean) As Double
    Dim i As Long
    Dim lastRow As Long
    Dim columnIndex As Integer
    Dim cogsCount As Integer
    
    lastRow = trialPLSheet.Cells(trialPLSheet.Rows.Count, 1).End(xlUp).row
    columnIndex = IIf(isSalesRevenue, 7, 6) ' Use column G for sales revenue, F for other income
    cogsCount = 0
    
    For i = 2 To lastRow
        If trialPLSheet.Cells(i, 2).Value = accountCode Then
            If accountCode = "1510" Then
                cogsCount = cogsCount + 1
                If cogsCount = 2 Then
                    GetAmountFromPreviousPeriodPL = trialPLSheet.Cells(i, 7).Value ' Always use column G for COGS
                    Exit Function
                End If
            Else
                GetAmountFromPreviousPeriodPL = trialPLSheet.Cells(i, columnIndex).Value
                Exit Function
            End If
        End If
    Next i
    
    GetAmountFromPreviousPeriodPL = 0
End Function

Sub FormatPLWorksheet(ws As Worksheet)
    ' Apply Thai Sarabun font and font size 14 to the worksheet
    ws.Cells.Font.Name = "TH Sarabun New"
    ws.Cells.Font.Size = 14
    
    ' Set number format to use comma style for columns G and I (for multi-year) or column I (for single-year)
    ws.Columns("G:I").NumberFormat = "#,##0.00"
    
    ' Adjust column widths
    ws.Columns("A").ColumnWidth = 5
    ws.Columns("B").ColumnWidth = 7
    ws.Columns("C").ColumnWidth = 8
    ws.Columns("D").ColumnWidth = 7
    ws.Columns("E").ColumnWidth = 28
    ws.Columns("F").ColumnWidth = 7
    ws.Columns("G").ColumnWidth = 14
    ws.Columns("H").ColumnWidth = 2
    ws.Columns("I").ColumnWidth = 14
    
    ' Center align headers
    ws.Range("A1:I4").HorizontalAlignment = xlCenter
    
    ' Right align amount columns
    ws.Columns("G:I").HorizontalAlignment = xlRight
    
    ' Format G6 and I6
    With ws.Range("G6,I6")
        .NumberFormat = "General"
        .Font.Underline = xlUnderlineStyleSingle
        .HorizontalAlignment = xlCenter
    End With
End Sub

Sub AddProfitLossHeaderDetails(ws As Worksheet, row As Long)
    With ws.Cells(row, 6)
        .Value = "หมายเหตุ"
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
    End With
    
    With ws.Cells(row, 9)
        .Value = "หน่วย:บาท"
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
        .HorizontalAlignment = xlRight
    End With
End Sub




