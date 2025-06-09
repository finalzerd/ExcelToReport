' Updated utility functions to work with TB1 format
' TB1 Format: Column A = Account Name, Column B = Account Code, Column C = Previous Period, Column D = Current Period

Function CreateHeader(ws As Worksheet, Optional headerType As String = "Notes") As Boolean
    Dim i As Integer
    Dim infoSheet As Worksheet
    Dim companyName As String
    Dim year As String
    Dim headerText As String
    
    ' Get the company name and year from the "Info" sheet
    Set infoSheet = ws.Parent.Sheets("Info")
    companyName = infoSheet.Range("B1").Value
    year = infoSheet.Range("B3").Value
    
    For i = 1 To 3
        ws.Range(ws.Cells(i, 1), ws.Cells(i, 9)).Merge
        
        With ws.Cells(i, 1)
            .Font.Name = "TH Sarabun New"
            .Font.Size = 14
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        Select Case i
            Case 1
                ws.Cells(i, 1).Value = companyName
            Case 2
                Select Case headerType
                    Case "Details"
                        ws.Cells(i, 1).Value = "รายละเอียดประกอบงบการเงิน"
                    Case "Balance Sheet"
                        ws.Cells(i, 1).Value = "งบดุลยะการเงิน"
                    Case "Profit and Loss Statement"
                        ws.Cells(i, 1).Value = "งบกำไรขาดทุน จำแนกค่าใช้จ่ายตามหน้าที่ - แบบขั้นเดียว"
                    Case "Statement of Changes in Equity"
                        ws.Cells(i, 1).Value = "งบการเปลี่ยนแปลงส่วนของผู้ถือหุ้น"
                    Case Else
                        ws.Cells(i, 1).Value = "หมายเหตุประกอบงบการเงิน"
                End Select
            Case 3
                Select Case headerType
                    Case "Balance Sheet"
                        headerText = "ณ วันที่ 31 ธันวาคม " & year
                    Case Else
                        headerText = "สำหรับรอบระยะเวลาบัญชี ตั้งแต่วันที่ 1 มกราคม " & year & " ถึงวันที่ 31 ธันวาคม " & year
                End Select
                ws.Cells(i, 1).Value = headerText
                ws.Range(ws.Cells(i, 1), ws.Cells(i, 9)).Borders(xlEdgeBottom).Weight = xlMedium
        End Select
    Next i
    
    ' Return True to indicate that the header was created successfully
    CreateHeader = True
End Function

Sub FormatWorksheet(ws As Worksheet)
    ' Apply Thai Sarabun font and font size 14 to the worksheet
    ws.Cells.Font.Name = "TH Sarabun New"
    ws.Cells.Font.Size = 14
    
    ' Set number format to use comma style for columns G and I
    ws.Columns("G").NumberFormat = "#,##0.00"
    ws.Columns("I").NumberFormat = "#,##0.00"
    
    ' Adjust column widths
    ws.Columns("A").ColumnWidth = 5
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C:F").ColumnWidth = 15
    ws.Columns("G").ColumnWidth = 20
    ws.Columns("H").ColumnWidth = 15
    ws.Columns("I").ColumnWidth = 20
    
    ' Center align headers
    ws.Range("A1:I4").HorizontalAlignment = xlCenter
    
    ' Right align amount columns
    ws.Columns("G").HorizontalAlignment = xlRight
    ws.Columns("I").HorizontalAlignment = xlRight
End Sub

Sub AddHeaderDetails(ws As Worksheet, row As Long)
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

Function GetAmountFromTB1ByAccountCode(TB1Sheet As Worksheet, accountCode As String, periodColumn As Integer) As Double
    ' periodColumn: 3 = Previous Period (Column C), 4 = Current Period (Column D)
    Dim i As Long
    Dim lastRow As Long
    
    lastRow = TB1Sheet.Cells(TB1Sheet.Rows.Count, 2).End(xlUp).row
    
    For i = 2 To lastRow
        If TB1Sheet.Cells(i, 2).Value = accountCode Then
            GetAmountFromTB1ByAccountCode = TB1Sheet.Cells(i, periodColumn).Value
            Exit Function
        End If
    Next i
    
    GetAmountFromTB1ByAccountCode = 0
End Function

Sub FormatAndAdjustCell(rng As Range)
    With rng
        .WrapText = True
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
    End With
    AdjustMergedCellsHeightInWorksheet ThisWorkbook.Sheets("GIC")
End Sub

Sub AdjustMergedCellsHeightInWorksheet(ws As Worksheet)
    Dim cell As Range
    Dim mergeArea As Range
    Dim textHeight As Double
    Dim rowHeight As Double
    
    For Each cell In ws.usedRange
        If cell.MergeCells Then
            Set mergeArea = cell.mergeArea
            If mergeArea.Cells(1, 1).Address = cell.Address Then
                textHeight = GetTextHeight(cell)
                rowHeight = Application.WorksheetFunction.RoundUp(textHeight / mergeArea.Rows.Count, 0)
                rowHeight = Application.WorksheetFunction.Max(rowHeight, 15) ' Minimum row height
                mergeArea.rowHeight = rowHeight
            End If
        End If
    Next cell
End Sub

Function GetTextHeight(cell As Range) As Double
    Dim textLength As Long
    Dim averageCharWidth As Double
    Dim estimatedWidth As Double
    Dim lineCount As Long
    Dim cellWidth As Double
    
    ' Get the cell width
    cellWidth = cell.mergeArea.Width
    
    ' Get the text length
    textLength = Len(cell.mergeArea.Cells(1, 1).Value)
    
    ' Estimate average character width (adjust this value if needed)
    averageCharWidth = 7 ' This is an approximation, adjust based on your font
    
    ' Estimate the total width of the text
    estimatedWidth = textLength * averageCharWidth
    
    ' Calculate the number of lines
    lineCount = Application.WorksheetFunction.RoundUp(estimatedWidth / cellWidth, 0)
    
    ' Ensure at least one line
    lineCount = Application.WorksheetFunction.Max(lineCount, 1)
    
    ' Calculate the height (15 points per line of text, plus 5 points of padding)
    GetTextHeight = (lineCount * 15) + 5
End Function

Function GetFinancialYears(ws As Worksheet, Optional includePreviousYear As Boolean = False) As Variant
    Dim targetWorkbook As Workbook
    Dim infoSheet As Worksheet
    Dim CurrentYear As String
    Dim PreviousYear As String
    Dim result As Variant
    
    On Error GoTo ErrorHandler
    
    ' Get the parent workbook
    Set targetWorkbook = ws.Parent
    
    ' Get the Info sheet
    Set infoSheet = targetWorkbook.Sheets("Info")
    
    ' Get the current year from cell B3
    CurrentYear = infoSheet.Range("B3").Value
    
    ' Check if current year is empty
    If CurrentYear = "" Then
        If includePreviousYear Then
            ReDim result(1 To 2) As String
            result(1) = "Error: Year not found in cell B3 of Info sheet"
            result(2) = "Error: Cannot calculate previous year"
        Else
            result = "Error: Year not found in cell B3 of Info sheet"
        End If
    Else
        If includePreviousYear Then
            ReDim result(1 To 2) As String
            result(1) = CurrentYear
            result(2) = CStr(CLng(CurrentYear) - 1)
        Else
            result = CurrentYear
        End If
    End If
    
    GetFinancialYears = result
    Exit Function
    
ErrorHandler:
    If includePreviousYear Then
        ReDim result(1 To 2) As String
        result(1) = "Error: " & Err.Description
        result(2) = "Error: Cannot calculate previous year"
    Else
        result = "Error: " & Err.Description
    End If
    GetFinancialYears = result
End Function

Function ContainsAccountCode(uniqueAccountCodes As Collection, accountCode As String) As Boolean
    On Error Resume Next
    uniqueAccountCodes.Item CStr(accountCode)
    ContainsAccountCode = (Err.Number = 0)
    On Error GoTo 0
End Function

Sub FormatNote(ws As Worksheet, startRow As Long, endRow As Long)
    Dim accountingFormat As String
    accountingFormat = "_( * #,##0.00_);_( * (#,##0.00);_( * ""-""??_);_(@_)"

    With ws.Range(ws.Cells(startRow, 1), ws.Cells(endRow, 11))
        .Font.Name = "TH Sarabun New"
        .Font.Size = 14
        
        ' Set accounting format for columns D, F, G, and I if they contain values
        Dim column As Variant
        For Each column In Array("D", "F", "G", "I")
            If WorksheetFunction.CountA(.Columns(column)) > 0 Then
                .Columns(column).NumberFormatLocal = accountingFormat
            End If
        Next column
        
        ' Right align amount columns
        .Columns("D").HorizontalAlignment = xlRight
        .Columns("F").HorizontalAlignment = xlRight
        .Columns("G").HorizontalAlignment = xlRight
        .Columns("I").HorizontalAlignment = xlRight
    End With
    
    ' Bold the note name and total
    ws.Cells(startRow, 2).Font.Bold = True
    ws.Cells(endRow - 1, 3).Font.Bold = True
    
    ' Adjust column widths
    ws.Columns("A").ColumnWidth = 4
    ws.Columns("B").ColumnWidth = 4
    ws.Columns("C").ColumnWidth = 25
    ws.Columns("D").ColumnWidth = 12
    ws.Columns("E").ColumnWidth = 2
    ws.Columns("F").ColumnWidth = 12
    ws.Columns("G").ColumnWidth = 12
    ws.Columns("H").ColumnWidth = 2
    ws.Columns("I").ColumnWidth = 12
    
    ' Format year header row (assuming it's in the second row of the note)
    With ws.Rows(startRow + 1)
        .NumberFormat = "@"  ' Set to text format to prevent number formatting
        .HorizontalAlignment = xlCenter
        .Font.Underline = xlUnderlineStyleSingle
    End With
    
    ' Reapply accounting format to amount columns, excluding the year header
    If endRow > startRow + 2 Then
        Dim reapplyRange As Range
        Set reapplyRange = Union( _
            ws.Range(ws.Cells(startRow + 2, 4), ws.Cells(endRow, 4)), _
            ws.Range(ws.Cells(startRow + 2, 6), ws.Cells(endRow, 6)), _
            ws.Range(ws.Cells(startRow + 2, 7), ws.Cells(endRow, 7)), _
            ws.Range(ws.Cells(startRow + 2, 9), ws.Cells(endRow, 9)) _
        )
        reapplyRange.NumberFormatLocal = accountingFormat
    End If
End Sub

Function HandleNoteExceedingRow34(ws As Worksheet, noteName As String, noteStartRow As Long, noteEndRow As Long, TB1Sheet As Worksheet) As Worksheet
    ' Create a new worksheet for the note
    Dim newWS As Worksheet
    Dim targetWorkbook As Workbook
    
    ' Get the target workbook
    Set targetWorkbook = ws.Parent
    
    ' Create the new worksheet in the target workbook
    Set newWS = targetWorkbook.Sheets.Add(After:=ws)
    ' newWS.Name = noteName
    
    ' Call the CreateHeader function to create the header in the new worksheet
    CreateHeader newWS
    
    ' Find the first empty row after the header in the new worksheet
    Dim newNoteStartRow As Long
    newNoteStartRow = newWS.Cells(newWS.Rows.Count, 1).End(xlUp).row + 1
    
    ' Copy the note content to the new worksheet
    ws.Range(ws.Cells(noteStartRow, 1), ws.Cells(noteEndRow, 11)).Copy
    newWS.Cells(newNoteStartRow, 1).PasteSpecial xlPasteValues
    newWS.Cells(newNoteStartRow, 1).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    
    ' Remove the note from the original worksheet
    ws.Range(ws.Cells(noteStartRow, 1), ws.Cells(noteEndRow, 11)).Delete Shift:=xlUp
    
    ' Apply Thai Sarabun font and font size 14 to the new worksheet
    newWS.Cells.Font.Name = "TH Sarabun New"
    newWS.Cells.Font.Size = 14
    
    ' Set number format to use comma style for columns J and K in the new worksheet
    newWS.Columns("J:K").NumberFormat = "#,##0.00"
    
    ' Return the new worksheet
    Set HandleNoteExceedingRow34 = newWS
End Function

Sub AddGuaranteeNoteToFinancialStatements(targetWorkbook As Workbook)
    Dim ws As Worksheet
    Dim infoSheet As Worksheet
    Dim lastRow As Long
    Dim isLimitedPartnershipFlag As Boolean
    Dim personName As String
    
    ' Check if it's a limited partnership
    isLimitedPartnershipFlag = isLimitedPartnership(targetWorkbook)
    
    ' Get the Info sheet
    Set infoSheet = targetWorkbook.Sheets("Info")
    
    ' Get the person name from Info sheet
    personName = infoSheet.Range("B5").Value
    
    ' Loop through all sheets in the workbook
    For Each ws In targetWorkbook.Worksheets
        ' Check if the sheet is related to Balance Sheet, P&L, or Notes
        If IsFinancialStatementRelated(ws.Name) Then
            ' Find the last row with content
            lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
            
            ' Add the first note
            ws.Cells(lastRow + 3, "B").Value = "หมายเหตุประกอบงบการเงินเป็นส่วนหนึ่งของงบการเงินนี้"
            
            ' Add the second note
            ws.Cells(lastRow + 5, "B").Value = "ข้อรับรองว่าเป็นรายการอันถูกต้องและเป็นความจริง"
            
            ' Add the signature line
            With ws.Cells(lastRow + 7, "C").Resize(1, 5)
                .Merge
                .HorizontalAlignment = xlCenter
                If isLimitedPartnershipFlag Then
                    .Value = "ลงชื่อ ……………………………………………………… หุ้นส่วนผู้จัดการ"
                Else
                    .Value = "ลงชื่อ ……………………………………………………… กรรมการตามอำนาจ"
                End If
            End With
            
            ' Add the person name in parentheses below the signature line
            With ws.Cells(lastRow + 8, "C").Resize(1, 5)
                .Merge
                .HorizontalAlignment = xlCenter
                .Value = "(" & personName & ")"
            End With
            
            ' Format the added text
            With ws.Range(ws.Cells(lastRow + 3, "B"), ws.Cells(lastRow + 8, "G"))
                .Font.Name = "TH Sarabun New"
                .Font.Size = 14
            End With
        End If
    Next ws
End Sub

Function IsFinancialStatementRelated(sheetName As String) As Boolean
    ' Add sheet names that are related to Balance Sheet, P&L, and Notes
    Select Case sheetName
        ' Balance Sheet related
        Case "ABS", "LBS", "MPA", "MPL"
            IsFinancialStatementRelated = True
        ' P&L related
        Case "PL", "PLM"
            IsFinancialStatementRelated = True
        ' Notes related
        Case "APS"
            IsFinancialStatementRelated = True
        ' Add any other relevant sheet names here
        Case Else
            IsFinancialStatementRelated = False
    End Select
End Function

Function isLimitedPartnership(targetWorkbook As Workbook) As Boolean
    Dim infoSheet As Worksheet
    Set infoSheet = targetWorkbook.Sheets("Info")
    ' Check if the company type indicates limited partnership
    isLimitedPartnership = (InStr(1, infoSheet.Range("B2").Value, "ห้างหุ้นส่วนจำกัด") > 0)
End Function

