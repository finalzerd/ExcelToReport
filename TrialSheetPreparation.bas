Attribute VB_Name = "TrialSheetPreparation"
Public Sub PrepareTrialSheets()
    Dim targetWorkbookPath As Variant
    Dim targetWorkbook As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    
    ' Choose target workbook
    targetWorkbookPath = Application.GetOpenFilename(FileFilter:="Excel Files (*.xlsx; *.xls; *.xlsm), *.xlsx; *.xls; *.xlsm", title:="Select Target Workbook", MultiSelect:=False)
    
    ' Check if a file was selected
    If TypeName(targetWorkbookPath) = "Boolean" Then
        MsgBox "No workbook selected. Operation cancelled.", vbExclamation
        Exit Sub
    End If
    
    ' Open the selected workbook
    On Error Resume Next
    Set targetWorkbook = Workbooks.Open(targetWorkbookPath)
    If Err.Number <> 0 Then
        MsgBox "Error opening the selected workbook: " & Err.Description, vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Loop through all worksheets
    For Each ws In targetWorkbook.Worksheets
        ' Skip the "Info" sheet
        If ws.Name <> "Info" Then
            Application.StatusBar = "Processing sheet: " & ws.Name
            
            ' Insert new column C
            ws.Columns("C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            
            ' Text to Columns for column A
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
            ws.Range("A1:A" & lastRow).TextToColumns Destination:=ws.Range("A1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar:="("
            
            ' Replace ")" with "" in column B
            ws.Columns("B").Replace What:=")", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
            ' Trim column A
            For i = 1 To lastRow
                ws.Cells(i, 1).Value = Trim(ws.Cells(i, 1).Value)
            Next i
            
            ' AutoFit columns for better visibility
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).column
            ws.Columns("A:" & Split(Cells(1, lastCol).Address, "$")(1)).AutoFit
        End If
    Next ws
    
    Application.StatusBar = False
    
    ' Save and close the workbook
    targetWorkbook.Save
    targetWorkbook.Close
    
    MsgBox "Trial sheets have been prepared successfully.", vbInformation
End Sub

Public Sub MapAccountCodes()
    Dim targetWorkbookPath As String
    Dim targetWorkbook As Workbook
    Dim acCodeWorkbook As Workbook
    Dim newRowCount As Long
    
    ' Select and validate workbooks
    targetWorkbookPath = SelectTargetWorkbook()
    If targetWorkbookPath = "" Then Exit Sub
    
    ' Open workbooks
    If Not OpenWorkbooks(targetWorkbookPath, targetWorkbook, acCodeWorkbook) Then Exit Sub
    
    ' Process the workbooks
    newRowCount = ProcessWorkbooks(targetWorkbook, acCodeWorkbook)
    
    ' Save and close workbooks
    SaveAndCloseWorkbooks targetWorkbook, acCodeWorkbook
    
    ' Show completion message
    ShowCompletionMessage newRowCount
End Sub

Public Sub CreateInfoSheetForExternalWorkbook()
    Dim targetWorkbookPath As Variant
    Dim targetWorkbook As Workbook
    Dim fileFormat As Integer
    
    ' Prompt user for the target workbook path
    targetWorkbookPath = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xlsx; *.xls; *.xlsm),*.xlsx;*.xls;*.xlsm", _
        title:="Select Target Workbook")
    
    If TypeName(targetWorkbookPath) = "Boolean" Then
        MsgBox "No file selected. Operation cancelled.", vbExclamation
        Exit Sub
    End If
    
    ' Determine the file format
    fileFormat = GetFileFormat(CStr(targetWorkbookPath))
    
    ' Open the target workbook
    On Error GoTo ErrorHandler
    Set targetWorkbook = Workbooks.Open(targetWorkbookPath, UpdateLinks:=0, ReadOnly:=False)
    On Error GoTo 0
    
    ' Call the main procedure with the target workbook
    CreateInfoSheet targetWorkbook
    
    ' Save and close the target workbook
    targetWorkbook.Save
    targetWorkbook.Close
    
    MsgBox "Info sheet creation completed successfully.", vbInformation
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred while opening the workbook: " & vbNewLine & Err.Description, vbCritical
    Exit Sub
End Sub

Private Function SelectTargetWorkbook() As String
    Dim targetWorkbookPath As Variant
    Dim targetFolderPath As String
    
    ' Choose target workbook
    targetWorkbookPath = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xlsx; *.xls; *.xlsm), *.xlsx; *.xls; *.xlsm", _
        title:="Select Target Workbook", _
        MultiSelect:=False)
    
    ' Check if a file was selected
    If TypeName(targetWorkbookPath) = "Boolean" Then
        MsgBox "No workbook selected. Operation cancelled.", vbExclamation
        SelectTargetWorkbook = ""
        Exit Function
    End If
    
    ' Get the folder path and validate ACCode.xlsx exists
    targetFolderPath = Left(targetWorkbookPath, InStrRev(targetWorkbookPath, "\"))
    If Dir(targetFolderPath & "ACCode.xlsx") = "" Then
        MsgBox "ACCode.xlsx not found in the same folder as the target workbook.", vbExclamation
        SelectTargetWorkbook = ""
        Exit Function
    End If
    
    SelectTargetWorkbook = targetWorkbookPath
End Function

Private Function OpenWorkbooks(ByVal targetWorkbookPath As String, ByRef targetWorkbook As Workbook, _
                             ByRef acCodeWorkbook As Workbook) As Boolean
    Dim targetFolderPath As String
    
    targetFolderPath = Left(targetWorkbookPath, InStrRev(targetWorkbookPath, "\"))
    
    On Error Resume Next
    Set targetWorkbook = Workbooks.Open(targetWorkbookPath)
    Set acCodeWorkbook = Workbooks.Open(targetFolderPath & "ACCode.xlsx")
    If Err.Number <> 0 Then
        MsgBox "Error opening workbooks: " & Err.Description, vbCritical
        OpenWorkbooks = False
        Exit Function
    End If
    On Error GoTo 0
    
    OpenWorkbooks = True
End Function

Private Function ProcessWorkbooks(ByVal targetWorkbook As Workbook, ByVal acCodeWorkbook As Workbook) As Long
    Dim ws As Worksheet
    Dim acCodeSheet As Worksheet
    Dim newRowCount As Long
    
    Set acCodeSheet = acCodeWorkbook.Sheets(1)
    
    For Each ws In targetWorkbook.Worksheets
        If InStr(1, ws.Name, "Info", vbTextCompare) = 0 Then
            Application.StatusBar = "Processing sheet: " & ws.Name
            ProcessWorksheet ws, acCodeSheet, newRowCount
        End If
    Next ws
    
    Application.StatusBar = False
    ProcessWorkbooks = newRowCount
End Function

Private Sub ProcessWorksheet(ByVal ws As Worksheet, ByVal acCodeSheet As Worksheet, ByRef newRowCount As Long)
    Dim targetLastRow As Long
    Dim acCodeLastRow As Long
    
    targetLastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    acCodeLastRow = acCodeSheet.Cells(acCodeSheet.Rows.Count, "A").End(xlUp).row
    
    ' First pass: Check for duplicates
    CheckForDuplicates ws, targetLastRow
    
    ' Second pass: Process accounts
    ProcessAccounts ws, acCodeSheet, targetLastRow, acCodeLastRow, newRowCount
End Sub

Private Sub CheckForDuplicates(ByVal ws As Worksheet, ByVal targetLastRow As Long)
    Dim i As Long, k As Long
    Dim isDuplicate As Boolean
    
    For i = 2 To targetLastRow
        If Trim(ws.Cells(i, 2).Value) <> "" Then
            isDuplicate = False
            For k = 2 To targetLastRow
                If k <> i And Trim(ws.Cells(k, 2).Value) <> "" Then
                    If ws.Cells(i, 2).Value = ws.Cells(k, 2).Value Then
                        isDuplicate = True
                        HighlightDuplicateRows ws, i, k
                    End If
                End If
            Next k
        End If
    Next i
End Sub

Private Sub ProcessAccounts(ByVal ws As Worksheet, ByVal acCodeSheet As Worksheet, _
                          ByVal targetLastRow As Long, ByRef acCodeLastRow As Long, _
                          ByRef newRowCount As Long)
    Dim i As Long, j As Long
    Dim foundMatch As Boolean
    
    For i = 2 To targetLastRow
        If Trim(ws.Cells(i, 2).Value) <> "" Then
            foundMatch = False
            
            For j = 2 To acCodeLastRow
                If ws.Cells(i, 1).Value = acCodeSheet.Cells(j, 1).Value Then
                    ws.Cells(i, 2).Value = acCodeSheet.Cells(j, 2).Value
                    foundMatch = True
                    Exit For
                End If
            Next j
            
            If Not foundMatch And ws.Cells(i, 1).Interior.Color <> RGB(255, 192, 0) Then
                AddNewAccount ws, acCodeSheet, i, acCodeLastRow, newRowCount
            End If
        End If
    Next i
End Sub

Private Sub HighlightDuplicateRows(ByVal ws As Worksheet, ByVal row1 As Long, ByVal row2 As Long)
    ws.Range(ws.Cells(row1, 1), ws.Cells(row1, 3)).Interior.Color = RGB(255, 192, 0)
    ws.Range(ws.Cells(row2, 1), ws.Cells(row2, 3)).Interior.Color = RGB(255, 192, 0)
End Sub

Private Sub AddNewAccount(ByVal ws As Worksheet, ByVal acCodeSheet As Worksheet, _
                         ByVal rowIndex As Long, ByRef acCodeLastRow As Long, _
                         ByRef newRowCount As Long)
    ' Highlight the row in yellow
    ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, 2)).Interior.Color = RGB(255, 255, 0)
    
    ' Add to ACCode sheet
    acCodeLastRow = acCodeLastRow + 1
    With acCodeSheet
        .Cells(acCodeLastRow, 1).Value = ws.Cells(rowIndex, 1).Value
        .Cells(acCodeLastRow, 2).Value = ws.Cells(rowIndex, 2).Value
        .Range(.Cells(acCodeLastRow, 1), .Cells(acCodeLastRow, 2)).Interior.Color = RGB(255, 255, 0)
    End With
    
    newRowCount = newRowCount + 1
End Sub

Private Sub SaveAndCloseWorkbooks(ByVal targetWorkbook As Workbook, ByVal acCodeWorkbook As Workbook)
    targetWorkbook.Save
    acCodeWorkbook.Save
    acCodeWorkbook.Close
    targetWorkbook.Close
End Sub

Private Sub ShowCompletionMessage(ByVal newRowCount As Long)
    If newRowCount > 0 Then
        MsgBox "Account code mapping completed." & vbNewLine & _
               newRowCount & " new account codes were added to ACCode.xlsx", vbInformation
    Else
        MsgBox "Account code mapping completed. No new account codes were added.", vbInformation
    End If
End Sub




Private Function GetFileFormat(filePath As String) As Integer
    Dim fileExtension As String
    fileExtension = LCase(Right(filePath, 4))
    
    Select Case fileExtension
        Case ".xls"
            GetFileFormat = xlExcel8
        Case "xlsx"
            GetFileFormat = xlOpenXMLWorkbook
        Case "xlsm"
            GetFileFormat = xlOpenXMLWorkbookMacroEnabled
        Case Else
            GetFileFormat = xlWorkbookDefault
    End Select
End Function

Public Sub CreateInfoSheet(targetWorkbook As Workbook)
    Dim ws As Worksheet
    Dim entityNumber As String
    Dim entityFilePath As String
    Dim entityData As Object
    
    On Error GoTo ErrorHandler
    
    ' Check if Info sheet already exists
    On Error Resume Next
    Set ws = targetWorkbook.Sheets("Info")
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo ErrorHandler
    
    ' Create new Info sheet
    Set ws = targetWorkbook.Sheets.Add(Before:=targetWorkbook.Sheets(1))
    ws.Name = "Info"
    
    ' Add headers in column A
    With ws
        .Cells(1, 1).Value = "Company"
        .Cells(2, 1).Value = "สถานะ"
        .Cells(3, 1).Value = "ปีสิ้นสุด"
        .Cells(4, 1).Value = "Entity number"
        .Cells(5, 1).Value = "ผู้ลงนาม"
        .Cells(6, 1).Value = "จำนวนหุ้น"
        .Cells(7, 1).Value = "มูลค่า"
    End With
    
    ' Prompt user for entity number
    entityNumber = InputBox("Please enter the entity number:", "Entity Number Input")
    If entityNumber = "" Then
        MsgBox "Entity number is required.", vbExclamation
        Exit Sub
    End If
    
    ' Construct file path and read CSV
    entityFilePath = targetWorkbook.Path & "\ExtractWebDBD\" & entityNumber & ".csv"
    
    ' Check if file exists
    If Dir(entityFilePath) = "" Then
        MsgBox "Entity file not found: " & entityFilePath, vbExclamation
        Exit Sub
    End If
    
    ' Read CSV data
    Set entityData = ReadCSV(entityFilePath)
    
    ' Fill in values in column B
    With ws
        .Cells(1, 2).Value = entityData("G2") ' Company name from G2
        .Cells(2, 2).Value = entityData("A2") ' Status from A2
        .Cells(3, 2).Value = "2567"          ' Fixed year
        .Cells(4, 2).Value = entityData("H2") ' Entity number from H2
        .Cells(5, 2).Value = entityData("I2") ' Signatory from I2
        .Cells(6, 2).Value = 0               ' Fixed value
        .Cells(7, 2).Value = 0               ' Fixed value
    End With
    
    ' Format the worksheet
    FormatInfoSheet ws
    
    Exit Sub
ErrorHandler:
    MsgBox "An error occurred while creating Info sheet: " & vbNewLine & Err.Description, vbCritical
End Sub

Private Sub FormatInfoSheet(ws As Worksheet)
    ' Apply Thai Sarabun font and font size 14
    ws.Cells.Font.Name = "TH Sarabun New"
    ws.Cells.Font.Size = 14
    
    ' Adjust column widths
    ws.Columns("A").ColumnWidth = 30
    ws.Columns("B").ColumnWidth = 40
    
    ' Add borders
    With ws.Range("A1:B7")
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End With
    
    ' Align text
    ws.Range("A1:A7").HorizontalAlignment = xlLeft
    ws.Range("B1:B7").HorizontalAlignment = xlLeft
    
    ' Add number format for specific cells
    ws.Range("B6:B7").NumberFormat = "#,##0"
End Sub

