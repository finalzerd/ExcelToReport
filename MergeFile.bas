Attribute VB_Name = "MergeFile"
Sub MergeFiles()
    On Error Resume Next
    
    Dim sourceWorkbook As Workbook
    Dim mainWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim newSheet As Worksheet
    Dim filePath As String
    Dim sheetCounter As Integer
    Dim totalFiles As Integer
    
    ' Create a new workbook to store merged sheets
    Set mainWorkbook = Application.Workbooks.Add
    
    ' Delete default sheets from the new workbook
    Application.DisplayAlerts = False
    If mainWorkbook.Sheets.Count > 1 Then
        For i = mainWorkbook.Sheets.Count To 2 Step -1
            mainWorkbook.Sheets(i).Delete
        Next i
    End If
    Application.DisplayAlerts = True
    
    ' Ask user if they want to merge 2 or 4 files
    totalFiles = GetFileCount()
    If totalFiles = 0 Then Exit Sub
    
    ' Initialize counter
    sheetCounter = 1
    
    ' Loop to select and merge files
    Do
        On Error Resume Next
        
        ' Show file picker dialog
        filePath = Application.GetOpenFilename( _
            FileFilter:="Excel Files (*.xlsx; *.xls),*.xlsx;*.xls", _
            title:="Select Excel File " & sheetCounter & " of " & totalFiles, _
            MultiSelect:=False)
            
        ' Check if user cancelled
        If filePath = "False" Then
            Exit Do
        End If
        
        ' Check if we've already processed all files
        If sheetCounter > totalFiles Then
            Exit Do
        End If
        
        ' Open selected workbook
        Set sourceWorkbook = Application.Workbooks.Open(filePath)
        If Err.Number <> 0 Then
            MsgBox "Error opening file: " & filePath & vbNewLine & "Error: " & Err.Description, vbCritical
            Err.Clear
            GoTo ContinueLoop
        End If
        
        ' Get the first sheet
        Set sourceSheet = sourceWorkbook.Sheets(1)
        
        ' Add new sheet to main workbook
        If sheetCounter = 1 Then
            Set newSheet = mainWorkbook.Sheets(1)
        Else
            Set newSheet = mainWorkbook.Sheets.Add(After:=mainWorkbook.Sheets(mainWorkbook.Sheets.Count))
        End If
        
        ' Name the sheet according to the pattern and total files
        newSheet.Name = GetSheetName(sheetCounter, totalFiles)
        
        ' Copy contents
        sourceSheet.usedRange.Copy
        newSheet.Range("A1").PasteSpecial xlPasteAll
        
        ' Copy column widths
        On Error Resume Next
        sourceSheet.usedRange.Columns.Copy
        newSheet.usedRange.Columns.PasteSpecial xlPasteColumnWidths
        
        ' Clear clipboard
        Application.CutCopyMode = False
        
        ' Close source workbook
        sourceWorkbook.Close SaveChanges:=False
        
        ' Increment counter
        sheetCounter = sheetCounter + 1
        
ContinueLoop:
        ' Clear any errors before next iteration
        Err.Clear
        Set sourceWorkbook = Nothing
        Set sourceSheet = Nothing
    Loop
    
    ' Check if any files were merged
    If sheetCounter = 1 Then
        MsgBox "No files were selected to merge.", vbInformation
        mainWorkbook.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' Check if correct number of sheets were created
    If sheetCounter - 1 < totalFiles Then
        MsgBox "Warning: Only " & (sheetCounter - 1) & " sheets were created. Expected " & totalFiles & " sheets.", vbExclamation
    End If
    
    ' Save the merged workbook
    Application.DisplayAlerts = True
    MsgBox "Files merged successfully! " & (sheetCounter - 1) & " sheets were created.", vbInformation
    
    ' Clean up
    Set sourceWorkbook = Nothing
    Set mainWorkbook = Nothing
    Set sourceSheet = Nothing
    Set newSheet = Nothing
End Sub

Private Function GetFileCount() As Integer
    Dim response As VbMsgBoxResult
    
    response = MsgBox("Do you want to merge 2 files?" & vbNewLine & _
                     "Click Yes for 2 files (Trial Balance 1 and Trial PL 1)" & vbNewLine & _
                     "Click No for 4 files (including Trial Balance 2 and Trial PL 2)" & vbNewLine & _
                     "Click Cancel to exit", _
                     vbQuestion + vbYesNoCancel, _
                     "Select Number of Files")
    
    Select Case response
        Case vbYes
            GetFileCount = 2
        Case vbNo
            GetFileCount = 4
        Case vbCancel
            GetFileCount = 0
    End Select
End Function

Private Function GetSheetName(counter As Integer, totalFiles As Integer) As String
    If totalFiles = 2 Then
        ' Two-file scenario
        Select Case counter
            Case 1
                GetSheetName = "Trial Balance 1"
            Case 2
                GetSheetName = "Trial PL 1"
        End Select
    Else
        ' Four-file scenario
        Select Case counter
            Case 1
                GetSheetName = "Trial Balance 1"
            Case 2
                GetSheetName = "Trial Balance 2"
            Case 3
                GetSheetName = "Trial PL 1"
            Case 4
                GetSheetName = "Trial PL 2"
        End Select
    End If
End Function

