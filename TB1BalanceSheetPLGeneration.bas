' TB1-based Balance Sheet and Profit & Loss Generation (Single TB1 Sheet)
' This module creates Balance Sheet and P&L statements directly from TB1 format
' Following the same data extraction logic as NoteGenerationCore but outputting in BalanceSheetGeneration and ProfitLossGeneration format
' Modified to work with only one TB1 sheet using column C (Previous Period) and column D (Current Period)

Sub GenerateBalanceSheetFromTB1()
    Dim TB1Sheets As Collection
    Dim targetWorkbook As Workbook
    
    ' Get TB1 sheets using the pattern from NoteGenerationCore
    Set TB1Sheets = GetWorksheetsWithPrefix("TB1")
    
    If TB1Sheets.Count = 0 Then
        MsgBox "No TB1 sheets found. Please ensure there is at least one sheet with 'TB1' prefix.", vbExclamation
        Exit Sub
    End If
    
    Set targetWorkbook = TB1Sheets(1).Parent
    
    ' Use only the first TB1 sheet for generation
    CreateMultiPeriodBalanceSheetFromTB1 TB1Sheets(1)
End Sub

Sub GenerateProfitLossFromTB1()
    Dim TB1Sheets As Collection
    Dim targetWorkbook As Workbook
    
    ' Get TB1 sheets
    Set TB1Sheets = GetWorksheetsWithPrefix("TB1")
    
    If TB1Sheets.Count = 0 Then
        MsgBox "No TB1 sheets found. Please ensure there is at least one sheet with 'TB1' prefix.", vbExclamation
        Exit Sub
    End If
    
    Set targetWorkbook = TB1Sheets(1).Parent
    
    ' Use only the first TB1 sheet for generation
    CreateMultiPeriodProfitLossFromTB1 TB1Sheets(1)
End Sub

Sub CreateMultiPeriodBalanceSheetFromTB1(TB1Sheet As Worksheet)
    Dim wsAsset As Worksheet
    Dim wsLiability As Worksheet
    Dim targetWorkbook As Workbook
    
    Set targetWorkbook = TB1Sheet.Parent
    
    ' Create new sheets for the balance sheet
    Set wsAsset = targetWorkbook.Sheets.Add(After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count))
    wsAsset.Name = "MPA_TB1"
    Set wsLiability = targetWorkbook.Sheets.Add(After:=wsAsset)
    wsLiability.Name = "MPL_TB1"
    
    ' Asset Side
    CreateMultiPeriodAssetBalanceSheetFromTB1 wsAsset, TB1Sheet
    
    ' Liability Side
    CreateMultiPeriodLiabilityBalanceSheetFromTB1 wsLiability, TB1Sheet
    
    ' Format both worksheets
    FormatWorksheet wsAsset
    FormatWorksheet wsLiability
End Sub

Sub CreateMultiPeriodAssetBalanceSheetFromTB1(ws As Worksheet, TB1Sheet As Worksheet)
    Dim row As Long
    Dim currentAssetsStartRow As Long
    Dim nonCurrentAssetsStartRow As Long
    Dim years As Variant
    Dim targetWorkbook As Workbook
    Dim tempResults As Variant

    ' Get the target workbook
    Set targetWorkbook = ws.Parent

    ' Get financial years
    years = GetFinancialYears(ws, True)
    If IsArray(years) Then
        If Left(years(1), 5) = "Error" Then
            MsgBox "Failed to get financial years: " & years(1), vbExclamation
            Exit Sub
        End If
    Else
        MsgBox "Failed to get financial years", vbExclamation
        Exit Sub
    End If
    
    ' Create header and details
    CreateHeader ws, "Balance Sheet"
    row = 5
    AddBalanceSheetHeaderDetails ws, row
    row = row + 1
    
    ' Add "สินทรัพย์"
    With ws.Cells(row, 2)
        .Value = "สินทรัพย์"
        .Font.Bold = True
    End With
    row = row + 1

    ' Add headers and years
    With ws.Cells(row, 2)
        .Value = "สินทรัพย์หมุนเวียน"
        .Font.Bold = True
    End With
    With ws.Cells(row, 7)
        .Value = years(1)
        .Font.Underline = xlUnderlineStyleSingle
    End With
    With ws.Cells(row, 9)
        .Value = years(2)
        .Font.Underline = xlUnderlineStyleSingle
    End With
    row = row + 1

    ' Store current assets start row
    currentAssetsStartRow = row

    ' Add current assets
    tempResults = AddMultiPeriodAccountGroupFromTB1(ws, TB1Sheet, row, "เงินสดและรายการเทียบเท่าเงินสด", "1010", "1099", "", True, False, True, False)
    tempResults = AddMultiPeriodAccountGroupFromTB1(ws, TB1Sheet, row, "ลูกหนี้การค้าและลูกหนี้หมุนเวียนอื่น", "1140", "1299", "1141", True, False, True, False)
    tempResults = AddMultiPeriodAccountGroupFromTB1(ws, TB1Sheet, row, "เงินให้กู้ยืมระยะสั้น", "1141", "1141", "", True, True, True, True)
    tempResults = AddMultiPeriodAccountGroupFromTB1(ws, TB1Sheet, row, "สินค้าคงเหลือ", "1510", "1530", "", True, False, True, False)

    ' Add total current assets with formulas
    With ws.Cells(row, 2)
        .Value = "รวมสินทรัพย์หมุนเวียน"
        .Font.Bold = True
    End With
    Dim currentAssetsRow As Long
    currentAssetsRow = row
    With ws.Cells(row, 7)
        .Formula = "=SUM(" & ws.Range(ws.Cells(currentAssetsStartRow, 7), ws.Cells(row - 1, 7)).Address & ")"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    With ws.Cells(row, 9)
        .Formula = "=SUM(" & ws.Range(ws.Cells(currentAssetsStartRow, 9), ws.Cells(row - 1, 9)).Address & ")"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    row = row + 1

    ' Add "สินทรัพย์ไม่หมุนเวียน"
    With ws.Cells(row, 2)
        .Value = "สินทรัพย์ไม่หมุนเวียน"
        .Font.Bold = True
    End With
    row = row + 1

    ' Store non-current assets start row
    nonCurrentAssetsStartRow = row

    ' Add non-current assets
    tempResults = AddMultiPeriodAccountGroupFromTB1(ws, TB1Sheet, row, "ที่ดิน อาคารและอุปกรณ์", "1600", "1659", "", True, False, True, False)
    tempResults = AddMultiPeriodAccountGroupFromTB1(ws, TB1Sheet, row, "สินทรัพย์ไม่หมุนเวียนอื่น", "1660", "1700", "", True, False, True, False)

    ' Add total non-current assets with formulas
    With ws.Cells(row, 2)
        .Value = "รวมสินทรัพย์ไม่หมุนเวียน"
        .Font.Bold = True
    End With
    Dim nonCurrentAssetsRow As Long
    nonCurrentAssetsRow = row
    With ws.Cells(row, 7)
        .Formula = "=SUM(" & ws.Range(ws.Cells(nonCurrentAssetsStartRow, 7), ws.Cells(row - 1, 7)).Address & ")"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    With ws.Cells(row, 9)
        .Formula = "=SUM(" & ws.Range(ws.Cells(nonCurrentAssetsStartRow, 9), ws.Cells(row - 1, 9)).Address & ")"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    row = row + 1

    ' Add total assets with formulas
    With ws.Cells(row, 2)
        .Value = "รวมสินทรัพย์"
        .Font.Bold = True
    End With
    With ws.Cells(row, 7)
        .Formula = "=" & ws.Cells(currentAssetsRow, 7).Address & "+" & ws.Cells(nonCurrentAssetsRow, 7).Address
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With
    With ws.Cells(row, 9)
        .Formula = "=" & ws.Cells(currentAssetsRow, 9).Address & "+" & ws.Cells(nonCurrentAssetsRow, 9).Address
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With
End Sub

' Add the missing AddBalanceSheetHeaderDetails subroutine
Sub AddBalanceSheetHeaderDetails(ws As Worksheet, row As Long)
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

Sub CreateMultiPeriodLiabilityBalanceSheetFromTB1(ws As Worksheet, TB1Sheet As Worksheet)
    Dim row As Long
    Dim currentLiabilitiesStartRow As Long
    Dim nonCurrentLiabilitiesStartRow As Long
    Dim equityStartRow As Long
    Dim partnersEquity As Collection
    Dim equityItem As Variant
    Dim tempResults As Variant
    Dim isLimitedPartnership As Boolean
    Dim liabilityAndEquityTerm As String
    Dim equityTerm As String
    Dim targetWorkbook As Workbook
    Dim years As Variant
    
    Set targetWorkbook = ws.Parent
    
    ' Get financial years
    years = GetFinancialYears(ws, True)
    If IsArray(years) Then
        If Left(years(1), 5) = "Error" Then
            MsgBox "Failed to get financial years: " & years(1), vbExclamation
            Exit Sub
        End If
    Else
        MsgBox "Failed to get financial years", vbExclamation
        Exit Sub
    End If
    
    ' Check if it's a limited partnership
    isLimitedPartnership = (targetWorkbook.Sheets("Info").Range("B2").Value = "ห้างหุ้นส่วนจำกัด")
    
    ' Set terms based on company type
    If isLimitedPartnership Then
        liabilityAndEquityTerm = "หนี้สินและส่วนของผู้เป็นหุ้นส่วน"
        equityTerm = "ส่วนของผู้เป็นหุ้นส่วน"
    Else
        liabilityAndEquityTerm = "หนี้สินและส่วนของผู้ถือหุ้น"
        equityTerm = "ส่วนของผู้ถือหุ้น"
    End If
    
    ' Initial setup
    CreateHeader ws, "Balance Sheet"
    row = 5
    AddBalanceSheetHeaderDetails ws, row
    row = row + 1

    ' Add main headers with years
    With ws.Cells(row, 2)
        .Value = liabilityAndEquityTerm
        .Font.Bold = True
    End With
    row = row + 1

    With ws.Cells(row, 2)
        .Value = "หนี้สินหมุนเวียน"
        .Font.Bold = True
    End With
    With ws.Cells(row, 7)
        .Value = years(1)
        .Font.Underline = xlUnderlineStyleSingle
    End With
    With ws.Cells(row, 9)
        .Value = years(2)
        .Font.Underline = xlUnderlineStyleSingle
    End With
    row = row + 1

    ' Store current liabilities start row
    currentLiabilitiesStartRow = row

    ' Add current liabilities
    tempResults = AddMultiPeriodAccountGroupFromTB1(ws, TB1Sheet, row, "เงินเบิกเงินบัญชีและเงินกู้ยืมระยะสั้นจากสถาบันการเงิน", "2001", "2009", "", True, False, True, False)
    tempResults = AddMultiPeriodAccountGroupFromTB1(ws, TB1Sheet, row, "เจ้าหนี้การค้าและเจ้าหนี้หมุนเวียนอื่น", "2010", "2999", "2030,2045,2050,2051,2052,2100,2120,2121,2122,2123", True, False, True, False)
    tempResults = AddMultiPeriodAccountGroupFromTB1(ws, TB1Sheet, row, "ส่วนของหนี้สินระยะยาวที่ถึงกำหนดชำระภายในหนึ่งปี", "0", "0", "", True, True, True, True)
    tempResults = AddMultiPeriodAccountGroupFromTB1(ws, TB1Sheet, row, "เงินกู้ยืมระยะสั้น", "2030", "2030", "", True, True, True, True)
    tempResults = AddMultiPeriodAccountGroupFromTB1(ws, TB1Sheet, row, "ภาษีเงินได้นิติบุคคลค้างจ่าย", "2045", "2045", "", True, True, True, True)

    ' Add total current liabilities with formulas
    With ws.Cells(row, 2)
        .Value = "รวมหนี้สินหมุนเวียน"
        .Font.Bold = True
    End With
    Dim currentLiabilitiesRow As Long
    currentLiabilitiesRow = row
    With ws.Cells(row, 7)
        .Formula = "=SUM(" & ws.Range(ws.Cells(currentLiabilitiesStartRow, 7), ws.Cells(row - 1, 7)).Address & ")"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    With ws.Cells(row, 9)
        .Formula = "=SUM(" & ws.Range(ws.Cells(currentLiabilitiesStartRow, 9), ws.Cells(row - 1, 9)).Address & ")"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    row = row + 1

    ' Add non-current liabilities header
    With ws.Cells(row, 2)
        .Value = "หนี้สินไม่หมุนเวียน"
        .Font.Bold = True
    End With
    row = row + 1

    ' Store non-current liabilities start row
    nonCurrentLiabilitiesStartRow = row

    ' Add non-current liabilities
    tempResults = AddMultiPeriodAccountGroupFromTB1(ws, TB1Sheet, row, "เงินกู้ยืมระยะยาวจากสถาบันการเงิน", "2120", "2123", "2121", True, False, True, False)
    
    tempResults = AddMultiPeriodAccountGroupFromTB1(ws, TB1Sheet, row, "เงินกู้ยืมระยะยาว", "2050", "2052", "", True, False, True, False)

    ' Add total non-current liabilities with formulas
    With ws.Cells(row, 2)
        .Value = "รวมหนี้สินไม่หมุนเวียน"
        .Font.Bold = True
    End With
    Dim nonCurrentLiabilitiesRow As Long
    nonCurrentLiabilitiesRow = row
    With ws.Cells(row, 7)
        .Formula = "=SUM(" & ws.Range(ws.Cells(nonCurrentLiabilitiesStartRow, 7), ws.Cells(row - 1, 7)).Address & ")"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    With ws.Cells(row, 9)
        .Formula = "=SUM(" & ws.Range(ws.Cells(nonCurrentLiabilitiesStartRow, 9), ws.Cells(row - 1, 9)).Address & ")"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    row = row + 1

    ' Add total liabilities with formulas
    With ws.Cells(row, 2)
        .Value = "รวมหนี้สิน"
        .Font.Bold = True
    End With
    Dim totalLiabilitiesRow As Long
    totalLiabilitiesRow = row
    With ws.Cells(row, 7)
        .Formula = "=" & ws.Cells(currentLiabilitiesRow, 7).Address & "+" & ws.Cells(nonCurrentLiabilitiesRow, 7).Address
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With
    With ws.Cells(row, 9)
        .Formula = "=" & ws.Cells(currentLiabilitiesRow, 9).Address & "+" & ws.Cells(nonCurrentLiabilitiesRow, 9).Address
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With
    row = row + 1

    ' Add equity section header
    With ws.Cells(row, 2)
        .Value = equityTerm
        .Font.Bold = True
    End With
    row = row + 1

    ' Store equity start row
    equityStartRow = row
    
    ' Declare paidUpCapitalStartRow here to use in both cases
    Dim paidUpCapitalStartRow As Long

    ' Add equity accounts
    If isLimitedPartnership Then
        ' For limited partnership, get partners equity using the same logic as main branch
        paidUpCapitalStartRow = equityStartRow
        
        ' Get partners equity using the same logic as main branch
        Set partnersEquity = GetPartnersEquity(targetWorkbook)
        For Each equityItem In partnersEquity
            ws.Cells(row, 3).Value = equityItem(0) ' Account name
            ws.Cells(row, 7).Value = equityItem(1) ' Current year
            ws.Cells(row, 9).Value = equityItem(1) ' Previous year (same as current from CSV)
            row = row + 1
        Next equityItem
    Else
        ' Add registered capital section
        ws.Cells(row, 3).Value = "ทุนจดทะเบียน"
        row = row + 1
        
        Dim shares As Long
        Dim shareValue As Double
        shares = CLng(targetWorkbook.Sheets("Info").Range("B6").Value)
        shareValue = CDbl(targetWorkbook.Sheets("Info").Range("B7").Value)
        
        ws.Cells(row, 4).Value = "หุ้นสามัญ " & Format(shares, "#,##0") & " หุ้น มูลค่าหุ้นละ " & Format(shareValue, "#,##0.00") & " บาท"
        ws.Cells(row, 7).Value = shares * shareValue
        ws.Cells(row, 9).Value = shares * shareValue
        With ws.Cells(row, 7)
            .Borders(xlEdgeBottom).LineStyle = xlDouble
        End With
        With ws.Cells(row, 9)
            .Borders(xlEdgeBottom).LineStyle = xlDouble
        End With
        row = row + 1
        
        ' Add paid-up capital section
        ws.Cells(row, 3).Value = "ทุนที่ออกและชำระแล้ว"
        row = row + 1
        ws.Cells(row, 4).Value = "หุ้นสามัญ " & Format(shares, "#,##0") & " หุ้น มูลค่าหุ้นละ " & Format(shareValue, "#,##0.00") & " บาท"
        paidUpCapitalStartRow = row
        ' Set paid-up capital to be the same as registered capital
        ws.Cells(row, 7).Value = shares * shareValue
        ws.Cells(row, 9).Value = shares * shareValue
        row = row + 1
        
        ' Add retained earnings for regular companies using the same calculation logic
        Dim retainedEarningsAmounts As Variant
        retainedEarningsAmounts = CalculateRetainedEarningsFromTB1(TB1Sheet)
        
        ws.Cells(row, 3).Value = "กำไร ( ขาดทุน ) สะสมยังไม่ได้จัดสรร"
        ws.Cells(row, 7).Value = retainedEarningsAmounts(1) ' Current period
        ws.Cells(row, 9).Value = retainedEarningsAmounts(2) ' Previous period
        row = row + 1
    End If

    ' Add total equity with formulas
    With ws.Cells(row, 2)
        .Value = "รวม" & equityTerm
        .Font.Bold = True
    End With
    Dim equityTotalRow As Long
    equityTotalRow = row
    
    ' Use paidUpCapitalStartRow which is set appropriately for both cases
    With ws.Cells(row, 7)
        .Formula = "=SUM(" & ws.Range(ws.Cells(paidUpCapitalStartRow, 7), ws.Cells(row - 1, 7)).Address & ")"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    With ws.Cells(row, 9)
        .Formula = "=SUM(" & ws.Range(ws.Cells(paidUpCapitalStartRow, 9), ws.Cells(row - 1, 9)).Address & ")"
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    row = row + 1

    ' Add total liabilities and equity with formulas
    With ws.Cells(row, 2)
        .Value = "รวม" & liabilityAndEquityTerm
        .Font.Bold = True
    End With
    With ws.Cells(row, 7)
        .Formula = "=" & ws.Cells(totalLiabilitiesRow, 7).Address & "+" & ws.Cells(equityTotalRow, 7).Address
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With
    With ws.Cells(row, 9)
        .Formula = "=" & ws.Cells(totalLiabilitiesRow, 9).Address & "+" & ws.Cells(equityTotalRow, 9).Address
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With
End Sub

' Function to add multi-period account group from TB1 data
Function AddMultiPeriodAccountGroupFromTB1(ws As Worksheet, TB1Sheet As Worksheet, ByRef row As Long, groupName As String, startCode As String, endCode As String, Optional excludeCode As String = "", Optional includeDecimal As Boolean = False, Optional listIndividualAccounts As Boolean = False, Optional isMultiPeriod As Boolean = True, Optional isSingleAccount As Boolean = False) As Variant
    Dim lastRow As Long
    Dim i As Long
    Dim totalAmount As Double
    Dim totalAmountPrevious As Double
    Dim accountCode As String
    Dim accountName As String
    Dim amount As Double
    Dim amountPrevious As Double
    Dim result(1 To 2) As Double
    Dim foundMatch As Boolean
    
    lastRow = TB1Sheet.Cells(TB1Sheet.Rows.Count, 2).End(xlUp).row
    totalAmount = 0
    totalAmountPrevious = 0
    foundMatch = False
    
    If groupName <> "" Then
        ws.Cells(row, 3).Value = groupName
    End If
    
    For i = 2 To lastRow
        accountCode = TB1Sheet.Cells(i, 2).Value
        If (isSingleAccount And accountCode = startCode) Or _
           (Not isSingleAccount And accountCode >= startCode And accountCode <= endCode) Then
            If InStr(1, excludeCode, accountCode) = 0 And (includeDecimal Or InStr(1, accountCode, ".") = 0) Then
                foundMatch = True
                accountName = TB1Sheet.Cells(i, 1).Value
                amount = TB1Sheet.Cells(i, 4).Value ' Column D = Current Period
                totalAmount = totalAmount + amount
                If isMultiPeriod Then
                    amountPrevious = TB1Sheet.Cells(i, 3).Value ' Column C = Previous Period
                    totalAmountPrevious = totalAmountPrevious + amountPrevious
                End If
                If listIndividualAccounts Or isSingleAccount Then
                    ws.Cells(row, 3).Value = groupName
                    ws.Cells(row, 7).Value = amount
                    ws.Cells(row, 9).Value = amountPrevious
                    row = row + 1
                End If
                If isSingleAccount Then Exit For
            End If
        End If
    Next i
    
    ' Handle case where no matching account codes were found
    If Not foundMatch Then
        If listIndividualAccounts Or isSingleAccount Then
            ws.Cells(row, 3).Value = groupName
            ws.Cells(row, 7).Value = 0
            ws.Cells(row, 9).Value = 0
            row = row + 1
        End If
    End If
    
    If Not listIndividualAccounts And Not isSingleAccount Then
        ws.Cells(row, 7).Value = totalAmount
        ws.Cells(row, 9).Value = totalAmountPrevious
        row = row + 1
    End If
    
    ' Set the result values
    result(1) = IIf(foundMatch, totalAmount, 0) ' Current year amount
    result(2) = IIf(isMultiPeriod, totalAmountPrevious, 0) ' Previous year amount
    
    AddMultiPeriodAccountGroupFromTB1 = result
End Function

Sub CreateMultiPeriodProfitLossFromTB1(TB1Sheet As Worksheet)
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
    
    Set targetWorkbook = TB1Sheet.Parent
    
    ' Get the year from the "Info" sheet
    Set infoSheet = targetWorkbook.Sheets("Info")
    year = infoSheet.Range("B3").Value
    PreviousYear = CStr(CLng(year) - 1)
    
    ' Create new sheet for the Profit and Loss statement
    Set wsPL = targetWorkbook.Sheets.Add(After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count))
    wsPL.Name = "PLM_TB1"
    
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
    totalRevenue = AddMultiYearRevenueAccountsFromTB1(wsPL, TB1Sheet, row, totalRevenuePrevious)
    
    ' Add a blank row for better readability
    row = row + 1
    
    ' Add expense section
    totalExpenses = AddMultiYearExpenseAccountsFromTB1(wsPL, TB1Sheet, row, totalExpensesPrevious)
    
    ' Calculate and add net profit/loss before financial costs and income tax using formula
    With wsPL.Cells(row, 2)
        .Value = "กำไรก่อนต้นทุนทางการเงินและภาษีเงินได้"
        .Font.Bold = True
    End With
    
    Dim profitBeforeFinCostTaxRow As Long
    profitBeforeFinCostTaxRow = row
    With wsPL.Cells(row, 7)
        .Formula = "=" & wsPL.Names("RevenueToTotalRow").RefersToRange.Cells(1).Address & "-" & _
                        wsPL.Names("ExpenseToTotalRow").RefersToRange.Cells(1).Address
    End With
    With wsPL.Cells(row, 9)
        .Formula = "=" & wsPL.Names("RevenueToTotalRow").RefersToRange.Cells(3).Address & "-" & _
                        wsPL.Names("ExpenseToTotalRow").RefersToRange.Cells(3).Address
    End With
    row = row + 1
    
    ' Add financial costs (store row for reference)
    Dim financialCostsRow As Long
    financialCostsRow = row
    wsPL.Cells(row, 2).Value = "ต้นทุนทางการเงิน"
    wsPL.Cells(row, 7).Value = GetFinancialCostsFromTB1(TB1Sheet)
    wsPL.Cells(row, 9).Value = GetFinancialCostsFromTB1Previous(TB1Sheet)
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
    wsPL.Cells(row, 7).Value = GetIncomeTaxFromTB1(TB1Sheet)
    wsPL.Cells(row, 9).Value = GetIncomeTaxFromTB1Previous(TB1Sheet)
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

Function AddMultiYearRevenueAccountsFromTB1(ws As Worksheet, TB1Sheet As Worksheet, ByRef row As Long, ByRef previousYearTotal As Double) As Double
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
    
    lastRow = TB1Sheet.Cells(TB1Sheet.Rows.Count, 2).End(xlUp).row
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
        accountCode = TB1Sheet.Cells(i, 2).Value
        If accountCode >= "4010" And accountCode <= "4019" Then
            accountName = TB1Sheet.Cells(i, 1).Value
            amount = TB1Sheet.Cells(i, 4).Value ' Column D for current year
            amountPrevious = TB1Sheet.Cells(i, 3).Value ' Column C for previous year
            salesRevenue = salesRevenue + amount
            salesRevenuePrevious = salesRevenuePrevious + amountPrevious
        ElseIf accountCode >= "4020" And accountCode <= "4210" Then
            accountName = TB1Sheet.Cells(i, 1).Value
            amount = TB1Sheet.Cells(i, 4).Value ' Column D for current year
            amountPrevious = TB1Sheet.Cells(i, 3).Value ' Column C for previous year
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
    AddMultiYearRevenueAccountsFromTB1 = totalRevenue
    
    ' Store the row reference in a name for later use
    ws.Names.Add "RevenueToTotalRow", revenueTotal
    
End Function

Function AddMultiYearExpenseAccountsFromTB1(ws As Worksheet, TB1Sheet As Worksheet, ByRef row As Long, ByRef previousYearTotal As Double) As Double
    Dim totalExpenses As Double
    Dim totalExpensesPrevious As Double
    
    totalExpenses = 0
    totalExpensesPrevious = 0
    
    ' Store starting row for expense details
    Dim expenseStartRow As Long
    expenseStartRow = row + 1  ' Skip the "ค่าใช้จ่าย" header row
    
    ' Add "ค่าใช้จ่าย" header
    With ws.Cells(row, 2)
        .Value = "ค่าใช้จ่าย"
        .Font.Bold = True
    End With
    row = row + 1
    
    ' Add expense categories with blank values for now
    ws.Cells(row, 3).Value = "ต้นทุนขายหรือต้นทุนการให้บริการ"
    ws.Cells(row, 7).Value = 0
    ws.Cells(row, 9).Value = 0
    row = row + 1
    
    ws.Cells(row, 3).Value = "ค่าใช้จ่ายในการขาย"
    ws.Cells(row, 7).Value = 0
    ws.Cells(row, 9).Value = 0
    row = row + 1
    
    ws.Cells(row, 3).Value = "ค่าใช้จ่ายในการบริหาร"
    ws.Cells(row, 7).Value = 0
    ws.Cells(row, 9).Value = 0
    row = row + 1
    
    ws.Cells(row, 3).Value = "ค่าใช้จ่ายอื่น"
    ws.Cells(row, 7).Value = 0
    ws.Cells(row, 9).Value = 0
    row = row + 1
    
    ' Add total expenses with formula (will sum the zeros for now)
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
    AddMultiYearExpenseAccountsFromTB1 = totalExpenses
    
    ' Store the row reference in a name for later use
    ws.Names.Add "ExpenseToTotalRow", expenseTotal
    
End Function

Function GetFinancialCostsFromTB1(TB1Sheet As Worksheet) As Double
    Dim lastRow As Long
    Dim i As Long
    Dim totalFinancialCosts As Double
    Dim accountCode As String
    
    lastRow = TB1Sheet.Cells(TB1Sheet.Rows.Count, 2).End(xlUp).row
    totalFinancialCosts = 0
    
    For i = 2 To lastRow
        accountCode = TB1Sheet.Cells(i, 2).Value
        If accountCode >= "5360" And accountCode <= "5364" Then
            totalFinancialCosts = totalFinancialCosts + TB1Sheet.Cells(i, 4).Value ' Column D
        End If
    Next i
    
    GetFinancialCostsFromTB1 = totalFinancialCosts
End Function

Function GetFinancialCostsFromTB1Previous(TB1Sheet As Worksheet) As Double
    Dim lastRow As Long
    Dim i As Long
    Dim totalFinancialCosts As Double
    Dim accountCode As String
    
    lastRow = TB1Sheet.Cells(TB1Sheet.Rows.Count, 2).End(xlUp).row
    totalFinancialCosts = 0
    
    For i = 2 To lastRow
        accountCode = TB1Sheet.Cells(i, 2).Value
        If accountCode >= "5360" And accountCode <= "5364" Then
            totalFinancialCosts = totalFinancialCosts + TB1Sheet.Cells(i, 3).Value ' Column C
        End If
    Next i
    
    GetFinancialCostsFromTB1Previous = totalFinancialCosts
End Function

Function GetIncomeTaxFromTB1(TB1Sheet As Worksheet) As Double
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = TB1Sheet.Cells(TB1Sheet.Rows.Count, 2).End(xlUp).row
    
    For i = 2 To lastRow
        If TB1Sheet.Cells(i, 2).Value = "5910" Then
            GetIncomeTaxFromTB1 = TB1Sheet.Cells(i, 4).Value ' Column D
            Exit Function
        End If
    Next i
    
    GetIncomeTaxFromTB1 = 0 ' Return 0 if account code 5910 is not found
End Function

Function GetIncomeTaxFromTB1Previous(TB1Sheet As Worksheet) As Double
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = TB1Sheet.Cells(TB1Sheet.Rows.Count, 2).End(xlUp).row
    
    For i = 2 To lastRow
        If TB1Sheet.Cells(i, 2).Value = "5910" Then
            GetIncomeTaxFromTB1Previous = TB1Sheet.Cells(i, 3).Value ' Column C
            Exit Function
        End If
    Next i
    
    GetIncomeTaxFromTB1Previous = 0 ' Return 0 if account code 5910 is not found
End Function

Sub FormatPLWorksheet(ws As Worksheet)
    ' Apply Thai Sarabun font and font size 14 to the worksheet
    ws.Cells.Font.Name = "TH Sarabun New"
    ws.Cells.Font.Size = 14
    
    ' Set number format to use comma style for columns G and I (for multi-year)
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

' Simple function to get retained earnings text
Function GetRetainedEarningsTextSimple(targetWorkbook As Workbook) As String
    ' Check if it's a limited partnership
    Dim isLimitedPartnership As Boolean
    isLimitedPartnership = (targetWorkbook.Sheets("Info").Range("B2").Value = "ห้างหุ้นส่วนจำกัด")
    
    If isLimitedPartnership Then
        GetRetainedEarningsTextSimple = "กำไรสะสม"
    Else
        GetRetainedEarningsTextSimple = "กำไรสะสม"
    End If
End Function

' Function to calculate retained earnings using special TB1 logic
Function CalculateRetainedEarningsFromTB1(TB1Sheet As Worksheet) As Variant
    Dim lastRow As Long
    Dim i As Long
    Dim accountCode As String
    Dim previousPeriodRetainedEarnings As Double
    Dim currentPeriodRetainedEarnings As Double
    Dim sumCredit As Double
    Dim sumDebit As Double
    Dim result(1 To 2) As Double
    
    lastRow = TB1Sheet.Cells(TB1Sheet.Rows.Count, 2).End(xlUp).row
    sumCredit = 0
    sumDebit = 0
    previousPeriodRetainedEarnings = 0
    
    ' Step 1: Get previous period retained earnings from column C and multiply by -1
    For i = 2 To lastRow
        accountCode = TB1Sheet.Cells(i, 2).Value
        If accountCode = "3020" Then
            previousPeriodRetainedEarnings = TB1Sheet.Cells(i, 3).Value * (-1) ' Column C * -1
            Exit For
        End If
    Next i
    
    ' Step 2 & 3: Sum credit side (column F) and debit side (column E) for accounts 4000-5999
    For i = 2 To lastRow
        accountCode = TB1Sheet.Cells(i, 2).Value
        If accountCode >= "4000" And accountCode <= "5999" Then
            sumCredit = sumCredit + TB1Sheet.Cells(i, 6).Value ' Column F (Credit)
            sumDebit = sumDebit + TB1Sheet.Cells(i, 5).Value   ' Column E (Debit)
        End If
    Next i
    
    ' Step 4: Calculate current period retained earnings
    ' Sum of Credit - Sum of Debit + Previous Period Amount
    currentPeriodRetainedEarnings = sumCredit - sumDebit + previousPeriodRetainedEarnings
    
    ' Return results: (1) = Current Period, (2) = Previous Period
    result(1) = currentPeriodRetainedEarnings
    result(2) = previousPeriodRetainedEarnings
    
    CalculateRetainedEarningsFromTB1 = result
End Function

' Function to get partners equity (from main branch)
Function GetPartnersEquity(targetWorkbook As Workbook) As Collection
    Dim entityNumber As String
    Dim entityFilePath As String
    Dim fso As Object
    Dim ts As Object
    Dim line As String
    Dim parts() As String
    Dim equityData As New Collection
    Dim accountName As String
    Dim amount As Double
    
    ' Get the entity number from the correct location in the target workbook
    entityNumber = targetWorkbook.Sheets("Info").Range("B4").Value
    
    ' Construct the file path using the target workbook's path
    entityFilePath = targetWorkbook.Path & "\ExtractWebDBD\" & entityNumber & ".csv"
    
    ' Check if the file exists
    If Dir(entityFilePath) = "" Then
        MsgBox "Entity file not found: " & entityFilePath, vbExclamation
        Exit Function
    End If
    
    ' Read the CSV file
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(entityFilePath, 1)
    
    ' Skip the header
    ts.SkipLine
    
    ' Read data
    Do Until ts.AtEndOfStream
        line = ts.ReadLine
        parts = Split(line, ",")
        
        ' Check if we have enough columns for the account name
        If UBound(parts) >= 8 Then
            accountName = Trim(parts(8))
            If accountName <> "" And Left(accountName, 5) <> "ลงทุน" Then
                ' Read the next line for the amount
                If Not ts.AtEndOfStream Then
                    line = ts.ReadLine
                    parts = Split(line, ",")
                    If UBound(parts) >= 9 Then
                        If IsNumeric(Trim(parts(9))) Then
                            amount = CDbl(Trim(parts(9)))
                            equityData.Add Array(accountName, amount)
                        End If
                    End If
                End If
            End If
        End If
    Loop
    
    ts.Close
    Set GetPartnersEquity = equityData
End Function

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

Sub FormatWorksheet(ws As Worksheet)
    ' Apply Thai Sarabun font and font size 14 to the worksheet
    ws.Cells.Font.Name = "TH Sarabun New"
    ws.Cells.Font.Size = 14
    
    ' Set number format to use comma style for columns G and I
    ws.Columns("G").NumberFormatLocal = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    ws.Columns("I").NumberFormatLocal = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    
    ' Format row 7 (years) separately
    With ws.Range("G7:I7")
        .NumberFormat = "General"
        .HorizontalAlignment = xlCenter
    End With
    
    ' Adjust column widths
    ws.Columns("A").ColumnWidth = 5
    ws.Columns("B").ColumnWidth = 7
    ws.Columns("C").ColumnWidth = 8
    ws.Columns("D:F").ColumnWidth = 7
    ws.Columns("E").ColumnWidth = 28
    ws.Columns("G").ColumnWidth = 14
    ws.Columns("H").ColumnWidth = 2
    ws.Columns("I").ColumnWidth = 14
    
    ' Center align headers
    ws.Range("A1:I4").HorizontalAlignment = xlCenter
End Sub

