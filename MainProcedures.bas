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
    Dim response As VbMsgBoxResult
    
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
    
    ' Ask user what they want to generate
    response = MsgBox("What would you like to generate?" & vbNewLine & vbNewLine & _
                     "YES = Financial Notes Only" & vbNewLine & _
                     "NO = Balance Sheet & P&L Only" & vbNewLine & _
                     "CANCEL = Everything (Notes + Balance Sheet + P&L)", _
                     vbYesNoCancel + vbQuestion, "Select Generation Type")
    
    Select Case response
        Case vbYes
            ' Generate Notes Only
            GenerateNotesOnly targetWorkbook, ws, TB1Sheet
            
        Case vbNo
            ' Generate Balance Sheet & P&L Only
            GenerateBalanceSheetAndPLOnly targetWorkbook
            
        Case vbCancel
            ' Generate Everything
            GenerateEverything targetWorkbook, ws, TB1Sheet
            
        Case Else
            MsgBox "No selection made. Operation cancelled.", vbInformation
            Exit Sub
    End Select
    
    Exit Sub
    
TB1NotFound:
    MsgBox "TB1 sheet not found. This new version requires a TB1 sheet with the unified trial balance format.", vbCritical
    Exit Sub
End Sub

Private Sub GenerateNotesOnly(targetWorkbook As Workbook, ws As Worksheet, TB1Sheet As Worksheet)
    ' Focus only on note generation
    If CreateHeader(ws) Then
        CreateMultiPeriodNotesFromTB1 ws, TB1Sheet
        MsgBox "Financial notes generated successfully!", vbInformation
    Else
        MsgBox "Failed to create header.", vbExclamation
    End If
End Sub

Private Sub GenerateBalanceSheetAndPLOnly(targetWorkbook As Workbook)
    Dim userChoice As VbMsgBoxResult
    
    ' Ask user which financial statements to generate
    userChoice = MsgBox("Which financial statements would you like to generate?" & vbNewLine & vbNewLine & _
                       "YES = Balance Sheet Only" & vbNewLine & _
                       "NO = Profit & Loss Only" & vbNewLine & _
                       "CANCEL = Both Balance Sheet & P&L", _
                       vbYesNoCancel + vbQuestion, "Select Financial Statements")
    
    Select Case userChoice
        Case vbYes
            ' Generate Balance Sheet only
            GenerateBalanceSheetFromTB1
            MsgBox "Balance Sheet generated successfully!" & vbNewLine & _
                   "Sheets created: MPA_TB1 (Assets), MPL_TB1 (Liabilities & Equity)", vbInformation
            
        Case vbNo
            ' Generate P&L only
            GenerateProfitLossFromTB1
            MsgBox "Profit & Loss Statement generated successfully!" & vbNewLine & _
                   "Sheet created: PLM_TB1", vbInformation
            
        Case vbCancel
            ' Generate both
            GenerateBalanceSheetFromTB1
            GenerateProfitLossFromTB1
            MsgBox "Balance Sheet and Profit & Loss Statement generated successfully!" & vbNewLine & _
                   "Sheets created: MPA_TB1, MPL_TB1, PLM_TB1", vbInformation
            
        Case Else
            MsgBox "No selection made. Operation cancelled.", vbInformation
    End Select
End Sub

Private Sub GenerateEverything(targetWorkbook As Workbook, ws As Worksheet, TB1Sheet As Worksheet)
    Dim success As Boolean
    success = True
    
    ' Generate Notes
    If CreateHeader(ws) Then
        CreateMultiPeriodNotesFromTB1 ws, TB1Sheet
    Else
        MsgBox "Failed to create header for notes.", vbExclamation
        success = False
    End If
    
    ' Generate Balance Sheet
    On Error Resume Next
    GenerateBalanceSheetFromTB1
    If Err.Number <> 0 Then
        MsgBox "Error generating Balance Sheet: " & Err.Description, vbExclamation
        success = False
        Err.Clear
    End If
    On Error GoTo 0
    
    ' Generate P&L
    On Error Resume Next
    GenerateProfitLossFromTB1
    If Err.Number <> 0 Then
        MsgBox "Error generating Profit & Loss: " & Err.Description, vbExclamation
        success = False
        Err.Clear
    End If
    On Error GoTo 0
    
    If success Then
        MsgBox "Complete financial statements generated successfully!" & vbNewLine & vbNewLine & _
               "Generated:" & vbNewLine & _
               "• Financial Notes (N1)" & vbNewLine & _
               "• Balance Sheet Assets (MPA_TB1)" & vbNewLine & _
               "• Balance Sheet Liabilities & Equity (MPL_TB1)" & vbNewLine & _
               "• Profit & Loss Statement (PLM_TB1)", vbInformation
    Else
        MsgBox "Financial statement generation completed with some errors. Please check the results.", vbExclamation
    End If
End Sub

' TB1-Specific Helper Functions
Function GetWorksheetsWithPrefix(prefix As String) As Collection
    Dim ws As Worksheet
    Dim matchingSheets As New Collection
    Dim wb As Workbook
    
    ' Use the active workbook
    Set wb = ActiveWorkbook
    
    For Each ws In wb.Worksheets
        If Left(ws.Name, Len(prefix)) = prefix Then
            matchingSheets.Add ws
        End If
    Next ws
    
    Set GetWorksheetsWithPrefix = matchingSheets
End Function

' Overloaded version for specific workbook
Function GetWorksheetsWithPrefixInWorkbook(wb As Workbook, prefix As String) As Collection
    Dim ws As Worksheet
    Dim matchingSheets As New Collection
    
    For Each ws In wb.Worksheets
        If Left(ws.Name, Len(prefix)) = prefix Then
            matchingSheets.Add ws
        End If
    Next ws
    
    Set GetWorksheetsWithPrefixInWorkbook = matchingSheets
End Function

' Quick test procedures for individual components
Public Sub TestNotesGeneration()
    Dim ws As Worksheet
    Dim TB1Sheet As Worksheet
    
    On Error GoTo ErrorHandler
    Set TB1Sheet = ActiveWorkbook.Sheets("TB1")
    Set ws = ActiveWorkbook.Sheets("N1")
    
    If CreateHeader(ws) Then
        CreateMultiPeriodNotesFromTB1 ws, TB1Sheet
        MsgBox "Notes generation test completed!", vbInformation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in test: " & Err.Description, vbCritical
End Sub

Public Sub TestBalanceSheetGeneration()
    On Error GoTo ErrorHandler
    GenerateBalanceSheetFromTB1
    MsgBox "Balance Sheet generation test completed!", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in test: " & Err.Description, vbCritical
End Sub

Public Sub TestProfitLossGeneration()
    On Error GoTo ErrorHandler
    GenerateProfitLossFromTB1
    MsgBox "P&L generation test completed!", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in test: " & Err.Description, vbCritical
End Sub

