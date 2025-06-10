Function CreateExpensesByNatureNote(ws As Worksheet) As Boolean
    Dim noteRow As Long
    Dim noteStartRow As Long
    Dim infoSheet As Worksheet
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

    ' Find the first empty row after the "EndOfNote" mark
    noteRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    noteStartRow = noteRow

    ' Increment the note order
    gNoteOrder = gNoteOrder + 1

    ' Create the note header
    ws.Cells(noteRow, 1).Value = gNoteOrder + 2  ' Start from 3
    ws.Cells(noteRow, 1).HorizontalAlignment = xlCenter
    ws.Cells(noteRow, 2).Value = "ค่าใช้จ่ายแยกตามลักษณะของค่าใช้จ่าย"
    
    ' Highlight the main title in yellow
    ws.Cells(noteRow, 2).Interior.Color = RGB(255, 255, 0)  ' Yellow highlight
    
    ws.Cells(noteRow, 9).Value = "หน่วย : บาท"
    ws.Cells(noteRow + 1, 7).Value = years(1)
    ws.Cells(noteRow + 1, 9).Value = years(2)
    noteRow = noteRow + 2

    ' Add expense categories
    AddExpenseCategory ws, noteRow, "การเปลี่ยนแปลงในสินค้าสำเร็จรูปและงานระหว่างทำ"
    AddExpenseCategory ws, noteRow, "งานที่ทำโดยกิจการและบันทึกเป็นรายการระหว่างทำ"
    AddExpenseCategory ws, noteRow, "วัตถุดิบและวัสดุสิ้นเปลืองใช้ไป"
    AddExpenseCategory ws, noteRow, "ค่าใช้จ่ายผลประโยชน์พนักงาน"
    AddExpenseCategory ws, noteRow, "ค่าเสื่อมราคาและค่าตัดจำหน่ายราย"
    AddExpenseCategory ws, noteRow, "ค่าใช้จ่ายอื่น"

    ' Add total
    ws.Cells(noteRow, 3).Value = "รวม"
    ws.Cells(noteRow, 3).Font.Bold = True
    noteRow = noteRow + 1

    ' Add the "EndOfNote" mark
    ws.Cells(noteRow, 1).Value = "EndOfNote"
    ws.Cells(noteRow, 1).Font.Color = vbWhite

    ' Check if note exceeds 34 rows
    If noteRow - noteStartRow > 34 Then
        ' Move the note to a new worksheet
        Set ws = HandleNoteExceedingRow34(ws, "ค่าใช้จ่ายแยกตามลักษณะของค่าใช้จ่าย", noteStartRow, noteRow, Nothing)
    End If

    ' Format the note
    FormatNote ws, noteStartRow, noteRow

    CreateExpensesByNatureNote = True
End Function

Sub AddExpenseCategory(ws As Worksheet, ByRef row As Long, categoryName As String)
    ws.Cells(row, 3).Value = categoryName
    row = row + 1
End Sub

Function IsLimitedCompany(targetWorkbook As Workbook) As Boolean
    Dim infoSheet As Worksheet
    Set infoSheet = targetWorkbook.Sheets("Info")
    IsLimitedCompany = (infoSheet.Range("B2").Value = "บริษัทจำกัด")
End Function

Function CreateFinancialApprovalNote(ws As Worksheet) As Boolean
    Dim noteRow As Long
    Dim noteStartRow As Long
    Static gNoteOrder As Integer
    Dim targetWorkbook As Workbook
    
    ' Get the target workbook
    Set targetWorkbook = ws.Parent
    
    ' Check if this is a limited company
    If Not IsLimitedCompany(targetWorkbook) Then
        CreateFinancialApprovalNote = False
        Exit Function
    End If
    
    ' Find the first empty row after the "EndOfNote" mark
    noteRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    noteStartRow = noteRow
    
    ' Increment the note order
    gNoteOrder = gNoteOrder + 1
    
    ' Create the note header
    ws.Cells(noteRow, 1).Value = gNoteOrder + 2
    ws.Cells(noteRow, 1).HorizontalAlignment = xlCenter
    ws.Cells(noteRow, 2).Value = "การอนุมัติงบการเงิน"
    ws.Cells(noteRow, 2).Font.Bold = True
    noteRow = noteRow + 1
    
    ' Add approval text
    ws.Cells(noteRow, 3).Value = "งบการเงินนี้ได้รับการรับรองโดยคณะกรรมการบริหารโดยมติอนุมัติงบการเงิน เมื่อวันที่ ............... ของคณะกรรมการบริษัทแล้ว"
    noteRow = noteRow + 1
    
    ' Add the "EndOfNote" mark
    ws.Cells(noteRow, 1).Value = "EndOfNote"
    ws.Cells(noteRow, 1).Font.Color = vbWhite
    
    CreateFinancialApprovalNote = True
End Function
