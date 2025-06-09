Function CreateLongTermLoansNoteFromTB1(ws As Worksheet, TB1Sheet As Worksheet) As Boolean
    Dim lastRow As Long
    Dim i As Long
    Dim accountCode As String
    Dim accountName As String
    Dim amountCurrent As Double, amountPrevious As Double
    Dim noteCreated As Boolean
    Dim noteRow As Long
    Dim totalAmountCurrent As Double, totalAmountPrevious As Double
    Dim noteStartRow As Long
    Dim uniqueAccountCodes As New Collection
    Dim years As Variant
    Dim targetWorkbook As Workbook
    Dim currentPortionCurrent As Double, currentPortionPrevious As Double
    
    ' Get the target workbook and years
    Set targetWorkbook = ws.Parent
    years = GetFinancialYears(ws, True)
    If IsArray(years) Then
        If Left(years(1), 5) = "Error" Then
            MsgBox "Failed to get financial years: " & years(1), vbExclamation
            Exit Function
        End If
    Else
        MsgBox "Failed to get financial years", vbExclamation
        Exit Function
    End If
    
    ' Initialize variables
    noteCreated = False
    totalAmountCurrent = 0: totalAmountPrevious = 0
    
    ' Set up note header and position
    lastRow = TB1Sheet.Cells(TB1Sheet.Rows.Count, 2).End(xlUp).row
    noteRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    noteStartRow = noteRow
    gNoteOrder = gNoteOrder + 1
    
    ' Create note header
    ws.Cells(noteRow, 1).Value = gNoteOrder + 2
    ws.Cells(noteRow, 1).HorizontalAlignment = xlCenter
    ws.Cells(noteRow, 2).Value = "เงินกู้ยืมระยะยาวจากสถาบันการเงิน"
    ws.Cells(noteRow, 9).Value = "หน่วย : บาท"
    ws.Cells(noteRow + 1, 7).Value = years(1)
    ws.Cells(noteRow + 1, 9).Value = years(2)
    noteRow = noteRow + 2
    
    ' Process current period accounts
    For i = 2 To lastRow
        accountCode = TB1Sheet.Cells(i, 2).Value
        accountName = TB1Sheet.Cells(i, 1).Value
        amountCurrent = TB1Sheet.Cells(i, 4).Value  ' Column D = Current Period
        
        If accountCode >= "2120" And accountCode <= "2123" And accountCode <> "2121" Then
            If Not ContainsAccountCode(uniqueAccountCodes, accountCode) Then
                uniqueAccountCodes.Add accountCode, CStr(accountCode)
                amountPrevious = GetAmountFromTB1PreviousPeriod(TB1Sheet, accountCode)
                
                If (amountCurrent <> 0 And Not IsEmpty(amountCurrent)) Or (amountPrevious <> 0 And Not IsEmpty(amountPrevious)) Then
                    ws.Cells(noteRow, 3).Value = accountName
                    ws.Cells(noteRow, 7).Value = amountCurrent
                    ws.Cells(noteRow, 9).Value = amountPrevious
                    noteRow = noteRow + 1
                    noteCreated = True
                End If
                
                totalAmountCurrent = totalAmountCurrent + amountCurrent
                totalAmountPrevious = totalAmountPrevious + amountPrevious
            End If
        End If
    Next i
    
    ' Add total row
    If noteCreated Then
        ws.Cells(noteRow, 3).Value = "รวม"
        With ws.Cells(noteRow, 7)
            .Value = totalAmountCurrent
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
        End With
        With ws.Cells(noteRow, 9)
            .Value = totalAmountPrevious
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
        End With
        noteRow = noteRow + 1
        
        ' Get current portions from user and store in global variable
        gLoanCurrentPortion.CurrentYear = CDbl(InputBox("กรุณากรอกส่วนของหนี้สินระยะยาวที่ถึงกำหนดชำระภายในหนึ่งปี สำหรับปี " & years(1), "Current Portion - Current Year"))
        gLoanCurrentPortion.PreviousYear = CDbl(InputBox("กรุณากรอกส่วนของหนี้สินระยะยาวที่ถึงกำหนดชำระภายในหนึ่งปี สำหรับปี " & years(2), "Current Portion - Previous Year"))
        
        ' Add current portion row
        ws.Cells(noteRow, 3).Value = "หัก ส่วนของหนี้สินระยะยาวที่ถึงกำหนดชำระภายในหนึ่งปี"
        ws.Cells(noteRow, 7).Value = gLoanCurrentPortion.CurrentYear
        ws.Cells(noteRow, 9).Value = gLoanCurrentPortion.PreviousYear
        noteRow = noteRow + 1
        
        ' Add net amount row
        With ws.Cells(noteRow, 3)
            .Value = "เงินกู้ยืมระยะยาวสุทธิจากส่วนที่ถึงกำหนดคืนภายในหนึ่งปี"
        End With
        With ws.Cells(noteRow, 7)
            .Value = totalAmountCurrent - gLoanCurrentPortion.CurrentYear
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlDouble
        End With
        With ws.Cells(noteRow, 9)
            .Value = totalAmountPrevious - gLoanCurrentPortion.PreviousYear
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlDouble
        End With
        noteRow = noteRow + 1
        
        ' Add EndOfNote mark
        ws.Cells(noteRow, 1).Value = "EndOfNote"
        ws.Cells(noteRow, 1).Font.Color = vbWhite
        
        ' Check if note exceeds page limit and create new sheet if needed
        If noteRow > 34 Then
            Set ws = HandleNoteExceedingRow34(ws, "เงินกู้ยืมระยะยาวจากสถาบันการเงิน", noteStartRow, noteRow, TB1Sheet)
            ' Rename the new sheet to N-series format
            ws.Name = "N" & gNoteOrder
        End If
    Else
        ws.Range(ws.Cells(noteStartRow, 1), ws.Cells(noteRow, 11)).ClearContents
        gNoteOrder = gNoteOrder - 1
    End If
    
    FormatNote ws, noteStartRow, noteRow
    CreateLongTermLoansNoteFromTB1 = noteCreated
End Function

Function GetAmountFromTB1PreviousPeriod(TB1Sheet As Worksheet, accountCode As String) As Double
    Dim i As Long
    Dim lastRow As Long
    
    lastRow = TB1Sheet.Cells(TB1Sheet.Rows.Count, 2).End(xlUp).row
    
    For i = 2 To lastRow
        If TB1Sheet.Cells(i, 2).Value = accountCode Then
            GetAmountFromTB1PreviousPeriod = TB1Sheet.Cells(i, 3).Value  ' Column C = Previous Period
            Exit Function
        End If
    Next i
    
    GetAmountFromTB1PreviousPeriod = 0
End Function

' Global variable type for loan current portion
Type LoanCurrentPortion
    CurrentYear As Double
    PreviousYear As Double
End Type

' Global variable declaration (add this to a module)
Public gLoanCurrentPortion As LoanCurrentPortion
