Function CreateNoteForLandBuildingEquipmentFromTB1(ws As Worksheet, TB1Sheet As Worksheet) As Boolean
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim accountCode As String
    Dim accountName As String
    Dim amountCurrent As Double, amountPrevious As Double
    Dim noteCreated As Boolean
    Dim noteRow As Long
    Dim noteStartRow As Long
    Dim infoSheet As Worksheet
    Dim years As Variant
    Dim assetTotalRow As Long
    Dim accumulatedDepreciationTotalRow As Long
    
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

    ' Find the first empty row after the "EndOfNote" mark
    noteRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    noteStartRow = noteRow

    ' Increment the global note order
    gNoteOrder = gNoteOrder + 1

    ' Create the note header
    ws.Cells(noteRow, 1).Value = gNoteOrder + 2  ' Start from 3
    ws.Cells(noteRow, 1).HorizontalAlignment = xlCenter
    ws.Cells(noteRow, 2).Value = "ที่ดิน อาคารและอุปกรณ์"
    ws.Cells(noteRow, 9).Value = "หน่วย : บาท"
    noteRow = noteRow + 1

    ' Add column headers
    ws.Cells(noteRow, 4).Value = "ณ 31 ธ.ค. " & years(2)
    ws.Cells(noteRow, 6).Value = "ซื้อเพิ่ม"
    ws.Cells(noteRow, 7).Value = "จำหน่ายออก"
    ws.Cells(noteRow, 9).Value = "ณ 31 ธ.ค. " & years(1)
    noteRow = noteRow + 1

    ' Add "ราคาทุนเดิม"
    ws.Cells(noteRow, 3).Value = "ราคาทุนเดิม"
    ws.Cells(noteRow, 3).Font.Bold = True
    noteRow = noteRow + 1

    ' Loop through the TB1 data for assets
    For i = 2 To TB1Sheet.Cells(TB1Sheet.Rows.Count, 2).End(xlUp).row
        accountCode = TB1Sheet.Cells(i, 2).Value
        
        ' Check if the account code is within range and doesn't contain a decimal
        If accountCode >= "1610" And accountCode <= "1659" And InStr(accountCode, ".") = 0 Then
            accountName = TB1Sheet.Cells(i, 1).Value
            amountCurrent = TB1Sheet.Cells(i, 4).Value  ' Current period
            amountPrevious = TB1Sheet.Cells(i, 3).Value  ' Previous period
            
            ' Add the account detail to the note
            ws.Cells(noteRow, 3).Value = accountName
            ws.Cells(noteRow, 4).Value = amountPrevious
            ws.Cells(noteRow, 9).Value = amountCurrent
            
            ' Calculate purchase or sale
            If amountCurrent > amountPrevious Then
                ws.Cells(noteRow, 6).Value = amountCurrent - amountPrevious
            ElseIf amountCurrent < amountPrevious Then
                ws.Cells(noteRow, 7).Value = amountPrevious - amountCurrent
            End If
            
            noteRow = noteRow + 1
            noteCreated = True
        End If
    Next i

    ' Add total row for assets
    If noteCreated Then
        ws.Cells(noteRow, 3).Value = "รวม"
        For j = 4 To 9
            If j <> 5 And j <> 8 Then  ' Skip columns E and H
                With ws.Cells(noteRow, j)
                    .Formula = "=SUM(" & ws.Cells(noteStartRow + 3, j).Address & ":" & ws.Cells(noteRow - 1, j).Address & ")"
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                End With
            End If
        Next j
        assetTotalRow = noteRow
        noteRow = noteRow + 2  ' Add two rows of space
    Else
        ' If no account details were added, remove the note header
        ws.Range(ws.Cells(noteStartRow, 1), ws.Cells(noteRow, 11)).ClearContents
        gNoteOrder = gNoteOrder - 1  ' Decrement the note order if note was not created
        CreateNoteForLandBuildingEquipmentFromTB1 = False
        Exit Function
    End If

    ' Add "ค่าเสื่อมราคาสะสม" header
    ws.Cells(noteRow, 3).Value = "ค่าเสื่อมราคาสะสม"
    ws.Cells(noteRow, 3).Font.Bold = True
    noteRow = noteRow + 1

    ' Loop through the TB1 data for accumulated depreciation
    For i = 2 To TB1Sheet.Cells(TB1Sheet.Rows.Count, 2).End(xlUp).row
        accountCode = TB1Sheet.Cells(i, 2).Value
        
        ' Check if the account code is within range and contains a decimal
        If accountCode >= "1610" And accountCode <= "1659" And InStr(accountCode, ".") > 0 Then
            accountName = TB1Sheet.Cells(i, 1).Value
            amountCurrent = TB1Sheet.Cells(i, 4).Value  ' Current period
            amountPrevious = TB1Sheet.Cells(i, 3).Value  ' Previous period
            
            ' Add the account detail to the note
            ws.Cells(noteRow, 3).Value = accountName
            ws.Cells(noteRow, 4).Value = amountPrevious
            ws.Cells(noteRow, 9).Value = amountCurrent
            
            ' Calculate changes in accumulated depreciation
            ws.Cells(noteRow, 6).Value = amountCurrent - amountPrevious
            
            noteRow = noteRow + 1
        End If
    Next i

    ' Add total row for accumulated depreciation
    ws.Cells(noteRow, 3).Value = "รวม"
    For j = 4 To 9
        If j <> 5 And j <> 8 Then  ' Skip columns E, G, and H
            With ws.Cells(noteRow, j)
                .Formula = "=SUM(" & ws.Cells(assetTotalRow + 3, j).Address & ":" & ws.Cells(noteRow - 1, j).Address & ")"
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
            End With
        End If
    Next j
    
    accumulatedDepreciationTotalRow = noteRow
    noteRow = noteRow + 1

    ' Add "มูลค่าสุทธิ" row
    ws.Cells(noteRow, 3).Value = "มูลค่าสุทธิ"
    ws.Cells(noteRow, 3).Font.Bold = True

    ' Add borders to columns D, F, G, and I
    For Each col In Array(4, 6, 7, 9)
        With ws.Cells(noteRow, col)
            If col = 4 Or col = 9 Then
                .Formula = "=" & ws.Cells(assetTotalRow, col).Address & "-" & ws.Cells(accumulatedDepreciationTotalRow, col).Address
            End If
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlDouble
        End With
    Next col

    noteRow = noteRow + 1

    ' Add "ค่าเสื่อมราคา" row
    ws.Cells(noteRow, 3).Value = "ค่าเสื่อมราคา"
    ws.Cells(noteRow, 3).Font.Bold = True
    ws.Cells(noteRow, 9).Formula = "=" & ws.Cells(accumulatedDepreciationTotalRow, 6).Address & "-" & ws.Cells(accumulatedDepreciationTotalRow, 7).Address
    noteRow = noteRow + 1
    
    ' Add the "EndOfNote" mark
    ws.Cells(noteRow, 1).Value = "EndOfNote"
    ws.Cells(noteRow, 1).Font.Color = vbWhite
    noteRow = noteRow + 1

    ' Check if note exceeds page limit and create new sheet if needed
    If noteRow > 34 Then
        Set ws = HandleNoteExceedingRow34(ws, "ที่ดิน อาคารและอุปกรณ์", noteStartRow, noteRow, TB1Sheet)
        ' Rename the new sheet to N-series format
        ws.Name = "N" & gNoteOrder
    End If

    ' Format the note
    FormatNote ws, noteStartRow, noteRow

    ' Return True to indicate that the note was created
    CreateNoteForLandBuildingEquipmentFromTB1 = True
End Function
