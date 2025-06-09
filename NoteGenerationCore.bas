' Global note counter for consistent numbering across all modules
Public gNoteOrder As Integer

Sub CreateMultiPeriodNotesFromTB1(ws As Worksheet, TB1Sheet As Worksheet)
    ' Create notes using the new TB1 format
    ' TB1 format: Column A = Account Name, Column B = Account Code, Column C = Previous Period, Column D = Current Period
    
    ' Initialize global note counter
    gNoteOrder = 0
    
    Call CreateCashNoteFromTB1(ws, TB1Sheet)
    Call CreateMultiPeriodNoteFromTB1(ws, TB1Sheet, "ลูกหนี้การค้าและลูกหนี้หมุนเวียนอื่น", "1140", "1215", "1141")
    Call CreateMultiPeriodNoteFromTB1(ws, TB1Sheet, "เงินให้ยืมระยะสั้น", "1141", "1141")
    Call CreateNoteForLandBuildingEquipmentFromTB1(ws, TB1Sheet)  ' PPE note 1600-1659
    Call CreateMultiPeriodNoteFromTB1(ws, TB1Sheet, "สินทรัพย์อื่น", "1660", "1700")
    Call CreateMultiPeriodNoteFromTB1(ws, TB1Sheet, "เงินเบิกเกินบัญชีและเงินกู้ยืมระยะสั้นจากสถาบันการเงิน", "2001", "2009")
    Call CreateMultiPeriodNoteFromTB1(ws, TB1Sheet, "เจ้าหนี้การค้าและเจ้าหนี้หมุนเวียนอื่น", "2010", "2999", "2030,2045,2050,2051,2052,2100,2120,2121,2122,2123")
    Call CreateMultiPeriodNoteFromTB1(ws, TB1Sheet, "เงินกู้ยืมระยะสั้นจากบุคคลหรือกิจการที่เกี่ยวข้องกัน", "2030", "2030")
    Call CreateLongTermLoansNoteFromTB1(ws, TB1Sheet)
    Call CreateMultiPeriodNoteFromTB1(ws, TB1Sheet, "เงินกู้ยืมระยะยาว", "2050", "2052")
    Call CreateMultiPeriodNoteFromTB1(ws, TB1Sheet, "เงินกู้ยืมระยะยาวจากบุคคลหรือกิจการที่เกี่ยวข้องกัน", "2100", "2100")
    Call CreateMultiPeriodNoteFromTB1(ws, TB1Sheet, "รายได้อื่น", "4020", "4999")
    Call CreateExpensesByNatureNote(ws)
    Call CreateFinancialApprovalNote(ws)
End Sub

Function CreateMultiPeriodNoteFromTB1(ws As Worksheet, TB1Sheet As Worksheet, noteName As String, accountCodeStart As String, accountCodeEnd As String, Optional excludeAccountCodes As String = "") As Boolean
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
    
    ' Get the target workbook
    Set targetWorkbook = ws.Parent
    
    ' Get the financial years using the new function
    years = GetFinancialYears(ws, True)
    If IsArray(years) Then
        If Left(years(1), 5) = "Error" Then
            MsgBox "Failed to get financial years: " & years(1), vbExclamation
        End If
    Else
        MsgBox "Failed to get financial years", vbExclamation
    End If
    
    ' Initialize noteCreated to False
    noteCreated = False
    
    ' Initialize totalAmountCurrent and totalAmountPrevious to zero
    totalAmountCurrent = 0
    totalAmountPrevious = 0
    
    ' Get the last row of data in the TB1 sheet
    lastRow = TB1Sheet.Cells(TB1Sheet.Rows.Count, 2).End(xlUp).row
    
    ' Find the first empty row after the "EndOfNote" mark
    noteRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    noteStartRow = noteRow
    
    ' Increment the global note order
    gNoteOrder = gNoteOrder + 1
    
    ' Create the note header
    ws.Cells(noteRow, 1).Value = gNoteOrder + 2  ' Start from 3
    ws.Cells(noteRow, 1).HorizontalAlignment = xlCenter
    ws.Cells(noteRow, 2).Value = noteName
    ws.Cells(noteRow, 9).Value = "หน่วย : บาท"
    ws.Cells(noteRow + 1, 7).Value = years(1)
    ws.Cells(noteRow + 1, 9).Value = years(2)
    noteRow = noteRow + 2
    
    ' Loop through the TB1 data for current period
    For i = 2 To lastRow
        accountCode = TB1Sheet.Cells(i, 2).Value  ' Column B = Account Code
        accountName = TB1Sheet.Cells(i, 1).Value  ' Column A = Account Name
        amountCurrent = TB1Sheet.Cells(i, 4).Value  ' Column D = Current Period
        amountPrevious = TB1Sheet.Cells(i, 3).Value  ' Column C = Previous Period
        
        If accountCode >= accountCodeStart And accountCode <= accountCodeEnd Then
            If InStr(1, excludeAccountCodes, accountCode) = 0 Then
                ' Account code is within the range and not in the exclude list
                If Not ContainsAccountCode(uniqueAccountCodes, accountCode) Then
                    ' Account code is unique, add it to the collection
                    uniqueAccountCodes.Add accountCode, CStr(accountCode)
                    
                    ' Only add the account detail if either amount is not zero or blank
                    If (amountCurrent <> 0 And Not IsEmpty(amountCurrent)) Or (amountPrevious <> 0 And Not IsEmpty(amountPrevious)) Then
                        ' Add the account detail to the note
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
        End If
    Next i
    
    ' Check if any account details were added to the note
    If noteCreated Then
        ' Add the total amounts to the note
        ws.Cells(noteRow, 3).Value = "รวม"
        ws.Cells(noteRow, 7).Value = totalAmountCurrent
        ws.Cells(noteRow, 9).Value = totalAmountPrevious
        
        ' Add top and double bottom borders only to the amount cells (columns 7 and 9)
        With ws.Cells(noteRow, 7)
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlDouble
        End With
        With ws.Cells(noteRow, 9)
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlDouble
        End With
        
        noteRow = noteRow + 1
        
        ' Add the "EndOfNote" mark to the final row of the note and color it white
        ws.Cells(noteRow, 1).Value = "EndOfNote"
        ws.Cells(noteRow, 1).Font.Color = vbWhite
        
        ' Check if note exceeds page limit and create new sheet if needed
        If noteRow > 34 Then
            Set ws = HandleNoteExceedingRow34(ws, noteName, noteStartRow, noteRow, TB1Sheet)
            ' Rename the new sheet to N-series format
            ws.Name = "N" & gNoteOrder
        End If
    Else
        ' If no account details were added, remove the note header
        ws.Range(ws.Cells(noteStartRow, 1), ws.Cells(noteRow, 11)).ClearContents
        gNoteOrder = gNoteOrder - 1  ' Decrement the note order if note was not created
    End If
    
    ' Format the note
    FormatNote ws, noteStartRow, noteRow
    
    ' Return the value of noteCreated
    CreateMultiPeriodNoteFromTB1 = noteCreated
End Function

Function CreateCashNoteFromTB1(ws As Worksheet, TB1Sheet As Worksheet) As Boolean
    Dim lastRow As Long
    Dim i As Long
    Dim accountCode As String
    Dim noteCreated As Boolean
    Dim noteRow As Long
    Dim cashAmountCurrent As Double, cashAmountPrevious As Double
    Dim bankAmountCurrent As Double, bankAmountPrevious As Double
    Dim totalAmountCurrent As Double, totalAmountPrevious As Double
    Dim noteStartRow As Long
    Dim years As Variant
    
    ' Get the financial years
    years = GetFinancialYears(ws, True)
    If IsArray(years) Then
        If Left(years(1), 5) = "Error" Then
            MsgBox "Failed to get financial years: " & years(1), vbExclamation
        End If
    Else
        MsgBox "Failed to get financial years", vbExclamation
    End If

    ' Initialize
    noteCreated = False
    cashAmountCurrent = 0: cashAmountPrevious = 0
    bankAmountCurrent = 0: bankAmountPrevious = 0

    ' Find the first empty row after the "EndOfNote" mark
    noteRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    noteStartRow = noteRow

    ' Increment the global note order
    gNoteOrder = gNoteOrder + 1

    ' Create the note header
    ws.Cells(noteRow, 1).Value = gNoteOrder + 2
    ws.Cells(noteRow, 1).HorizontalAlignment = xlCenter
    ws.Cells(noteRow, 2).Value = "เงินสดและรายการเทียบเท่าเงินสด"
    ws.Cells(noteRow, 9).Value = "หน่วย : บาท"
    noteRow = noteRow + 1
    ws.Cells(noteRow, 7).Value = years(1)
    ws.Cells(noteRow, 9).Value = years(2)
    noteRow = noteRow + 1

    ' Calculate amounts
    lastRow = TB1Sheet.Cells(TB1Sheet.Rows.Count, 2).End(xlUp).row
    For i = 2 To lastRow
        accountCode = TB1Sheet.Cells(i, 2).Value
        
        ' Cash (1010-1019)
        If accountCode >= "1010" And accountCode <= "1019" Then
            cashAmountCurrent = cashAmountCurrent + TB1Sheet.Cells(i, 4).Value
            cashAmountPrevious = cashAmountPrevious + TB1Sheet.Cells(i, 3).Value
        End If
        
        ' Bank deposits (1020-1099)
        If accountCode >= "1020" And accountCode <= "1099" Then
            bankAmountCurrent = bankAmountCurrent + TB1Sheet.Cells(i, 4).Value
            bankAmountPrevious = bankAmountPrevious + TB1Sheet.Cells(i, 3).Value
        End If
    Next i

    ' Add cash line if there's any amount
    If cashAmountCurrent <> 0 Or cashAmountPrevious <> 0 Then
        ws.Cells(noteRow, 3).Value = "เงินสด"
        ws.Cells(noteRow, 7).Value = cashAmountCurrent
        ws.Cells(noteRow, 9).Value = cashAmountPrevious
        noteRow = noteRow + 1
        noteCreated = True
    End If

    ' Add bank deposits line if there's any amount
    If bankAmountCurrent <> 0 Or bankAmountPrevious <> 0 Then
        ws.Cells(noteRow, 3).Value = "เงินฝากธนาคาร"
        ws.Cells(noteRow, 7).Value = bankAmountCurrent
        ws.Cells(noteRow, 9).Value = bankAmountPrevious
        noteRow = noteRow + 1
        noteCreated = True
    End If

    ' Add total if note was created
    If noteCreated Then
        totalAmountCurrent = cashAmountCurrent + bankAmountCurrent
        totalAmountPrevious = cashAmountPrevious + bankAmountPrevious
        ws.Cells(noteRow, 3).Value = "รวม"
        With ws.Cells(noteRow, 7)
            .Value = totalAmountCurrent
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlDouble
        End With
        With ws.Cells(noteRow, 9)
            .Value = totalAmountPrevious
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlDouble
        End With
        noteRow = noteRow + 1

        ' Add the "EndOfNote" mark
        ws.Cells(noteRow, 1).Value = "EndOfNote"
        ws.Cells(noteRow, 1).Font.Color = vbWhite
        
        ' Check if note exceeds page limit and create new sheet if needed
        If noteRow > 34 Then
            Set ws = HandleNoteExceedingRow34(ws, "เงินสดและรายการเทียบเท่าเงินสด", noteStartRow, noteRow, TB1Sheet)
            ' Rename the new sheet to N-series format
            ws.Name = "N" & gNoteOrder
        End If
    Else
        ' If no note was created, remove the header
        ws.Range(ws.Cells(noteStartRow, 1), ws.Cells(noteRow, 11)).ClearContents
        gNoteOrder = gNoteOrder - 1
    End If

    ' Format the note
    FormatNote ws, noteStartRow, noteRow

    ' Return success/failure
    CreateCashNoteFromTB1 = noteCreated
End Function
