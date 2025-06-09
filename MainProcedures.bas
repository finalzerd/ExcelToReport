Public Sub GenerateFinancialStatementForExternalWorkbook()
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
    GenerateFinancialStatement targetWorkbook
    
    ' Save and close the target workbook
    targetWorkbook.Save
    targetWorkbook.Close
    
    MsgBox "Financial statement generation completed successfully.", vbInformation
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred while opening the workbook: " & vbNewLine & Err.Description, vbCritical
    Exit Sub
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

Public Sub GenerateFinancialStatement(targetWorkbook As Workbook)
    Dim ws As Worksheet
    Dim TB1Sheet As Worksheet
    
    ' Set the worksheet object for notes
    On Error Resume Next
    Set ws = targetWorkbook.Sheets("Note1")
    If ws Is Nothing Then
        Set ws = targetWorkbook.Sheets.Add(After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count))
        ws.Name = "N1"  ' Use N-series naming convention
    Else
        ' Rename existing sheet to N-series format
        ws.Name = "N1"
    End If
    On Error GoTo 0
    
    ' Get the TB1 sheet (new unified format)
    On Error GoTo TB1NotFound
    Set TB1Sheet = targetWorkbook.Sheets("TB1")
    On Error GoTo 0
    
    ' Check if TB1 sheet exists and has data
    If TB1Sheet Is Nothing Then
        MsgBox "TB1 sheet not found. Please ensure the TB1 sheet exists with the required data format.", vbExclamation
        Exit Sub
    End If
    
    ' Focus only on note generation for now
    ' Determine if this is single or multiple period based on data structure
    ' For now, assume multiple period (as specified in requirements)
    If CreateHeader(ws) Then
        CreateMultiPeriodNotesFromTB1 ws, TB1Sheet
    Else
        MsgBox "Failed to create header.", vbExclamation
        Exit Sub
    End If
    
    Exit Sub
    
TB1NotFound:
    MsgBox "TB1 sheet not found. This new version requires a TB1 sheet with the unified trial balance format.", vbCritical
    Exit Sub
End Sub

Function GetWorksheetsWithPrefix(wb As Workbook, prefix As String) As Collection
    Dim ws As Worksheet
    Dim matchingSheets As New Collection
    
    For Each ws In wb.Worksheets
        If Left(ws.Name, Len(prefix)) = prefix Then
            matchingSheets.Add ws
        End If
    Next ws
    
    Set GetWorksheetsWithPrefix = matchingSheets
End Function

