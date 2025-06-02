Attribute VB_Name = "BasisOfPreparation"
Sub CreateBasisOfPreparation(targetWorkbook As Workbook)
    Dim ws As Worksheet
    Dim basisFilePath As String
    Dim basisWorkbook As Workbook
    Dim basisSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentRow As Long
    Dim mergedRange As Range
    
    ' Set the worksheet
    Set ws = targetWorkbook.Sheets("GIC")
    
    ' Construct the file path
    basisFilePath = targetWorkbook.Path & "\AccountingPolicy\basis_of_preparation.xlsx"
    
    ' Check if the file exists
    If Dir(basisFilePath) = "" Then
        MsgBox "Basis of preparation file not found: " & basisFilePath, vbExclamation
        Exit Sub
    End If
    
    ' Open the basis of preparation workbook
    Set basisWorkbook = Workbooks.Open(basisFilePath)
    Set basisSheet = basisWorkbook.Sheets(1)
    
    ' Find the last row in the basis sheet
    lastRow = basisSheet.Cells(basisSheet.Rows.Count, "A").End(xlUp).row
    
    ' Start at row 13 in the GIC sheet
    currentRow = 13
    
    ' Add the main title and make it bold
    ws.Cells(currentRow, 1).Value = "3"
    ws.Cells(currentRow, 1).HorizontalAlignment = xlCenter
    ws.Cells(currentRow, 2).Value = "เกณฑ์ในการจัดทำและนำเสนองบการเงิน"
    FormatTitleCell ws.Range(ws.Cells(currentRow, 2), ws.Cells(currentRow, 2))
    currentRow = currentRow + 1
    
    ' Iterate through the basis of preparation data
    For i = 2 To lastRow ' Assuming row 1 is header
        If basisSheet.Cells(i, 1).Value <> "" Then
            ' Add content to column B
            ws.Cells(currentRow, 2).Value = basisSheet.Cells(i, 1).Value
            ws.Cells(currentRow, 2).HorizontalAlignment = xlCenter
            ws.Cells(currentRow, 2).VerticalAlignment = xlTop
            
            ' Merge columns C-H
            Set mergedRange = ws.Range(ws.Cells(currentRow, 3), ws.Cells(currentRow, 8))
            mergedRange.Merge
            
            ' Add content to the merged range
            mergedRange.Value = basisSheet.Cells(i, 2).Value
            
            ' Format and adjust cells
            FormatAndAdjustCell ws.Range(ws.Cells(currentRow, 3), ws.Cells(currentRow, 8))
            
            currentRow = currentRow + 1
        Else
            ' Exit loop if we encounter a blank row
            Exit For
        End If
    Next i
    
    ' Close the basis of preparation workbook
    basisWorkbook.Close SaveChanges:=False
    
    ' Adjust column widths
    ws.Columns("A").ColumnWidth = 5
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C:H").ColumnWidth = 15  ' Adjusted for merged cells
End Sub

Sub FormatAndAdjustCell(rng As Range)
    With rng
        .WrapText = True
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
    End With
    
    ' Get the worksheet that contains the range
    Dim ws As Worksheet
    Set ws = rng.Worksheet
    
    ' Get the workbook that contains the worksheet
    Dim targetWorkbook As Workbook
    Set targetWorkbook = ws.Parent
    
    ' Check if "GIC" sheet exists in the target workbook
    Dim gicSheet As Worksheet
    On Error Resume Next
    Set gicSheet = targetWorkbook.Sheets("GIC")
    On Error GoTo 0
    
    If Not gicSheet Is Nothing Then
        AdjustMergedCellsHeight gicSheet
    Else
        ' If "GIC" sheet doesn't exist, adjust the current worksheet
        AdjustMergedCellsHeight ws
    End If
End Sub

Sub FormatTitleCell(rng As Range)
    With rng
        .WrapText = False
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
        .Font.Bold = True
    End With
End Sub

Sub AdjustMergedCellsHeight(ws As Worksheet)
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
