Attribute VB_Name = "MainProcedures"
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
    Dim trialBalanceSheets As Collection
    Dim trialPLSheets As Collection
    
    ' Set the worksheet object
    On Error Resume Next
    Set ws = targetWorkbook.Sheets("Note1")
    If ws Is Nothing Then
        Set ws = targetWorkbook.Sheets.Add(After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count))
        ws.Name = "Note1"
    End If
    On Error GoTo 0
    
    ' Get the TrialBalance and TrialPL sheets
    Set trialBalanceSheets = GetWorksheetsWithPrefix(targetWorkbook, "Trial Balance")
    Set trialPLSheets = GetWorksheetsWithPrefix(targetWorkbook, "Trial PL")
    
    ' Check the number of TrialBalance and TrialPL sheets
    If trialBalanceSheets.Count = 1 And trialPLSheets.Count = 1 Then
        ' Single-year statement generation
        If CreateHeader(ws) Then
            CreateSingleYearNotes ws, trialBalanceSheets(1), trialPLSheets(1)
            CreateFirstYearBalanceSheet trialBalanceSheets(1)
            CreateSingleYearProfitLossStatement trialPLSheets(1)
        Else
            MsgBox "Failed to create header.", vbExclamation
            Exit Sub
        End If
    ElseIf trialBalanceSheets.Count = 2 And trialPLSheets.Count = 2 Then
        ' Multi-year statement generation
        If CreateHeader(ws) Then
            CreateMultiYearNotes ws, trialBalanceSheets, trialPLSheets
            CreateMultiPeriodBalanceSheet trialBalanceSheets
            CreateMultiYearProfitLossStatement trialPLSheets
        Else
            MsgBox "Failed to create header.", vbExclamation
            Exit Sub
        End If
    Else
        MsgBox "Invalid number of TrialBalance or TrialPL sheets. Please ensure there are either one or two sheets of each type.", vbExclamation
        Exit Sub
    End If
    
    ' Create Detail sheets for both single and multi-year scenarios
    CreateDetailOne targetWorkbook
    CreateDetailTwo targetWorkbook
    
    ' Generate common components
    CreateGICContent targetWorkbook
    
    ' In the GenerateFinancialStatement function, add this line:
    CreateStatementOfChangesInEquity targetWorkbook
    
    ' Add guarantee note to all relevant financial statement sheets
    AddGuaranteeNoteToFinancialStatements targetWorkbook
    
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
