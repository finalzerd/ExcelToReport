Attribute VB_Name = "AccountingPolicySummary"
Option Explicit

' Main procedure to create accounting policy summary
Public Sub CreateAccountingPolicySummary(targetWorkbook As Workbook)
    Dim ws As Worksheet
    Dim policyWorkbook As Workbook
    Dim policySheet As Worksheet
    Dim trialBalanceSheet As Worksheet
    Dim accountCodes As Collection
    Dim orderNum As Long
    Dim currentRow As Long
    
    On Error GoTo ErrorHandler
    
    ' Create a new sheet for summary of accounting policy
    Set ws = CreateAccountingPolicySheet(targetWorkbook)
    
    ' Create the header
    CreateHeader ws, "Accounting Policy"
    
    ' Add the main topic header
    AddMainTopicHeader ws
    
    ' Open the accounting policy workbook
    Set policyWorkbook = OpenPolicyWorkbook(targetWorkbook.Path)
    Set policySheet = policyWorkbook.Sheets(1)
    
    ' Get account codes from Trial Balance
    Set trialBalanceSheet = targetWorkbook.Sheets(Config.TrialBalanceSheetName)
    Set accountCodes = GetAccountCodes(trialBalanceSheet)
    
    ' Process accounting policies
    currentRow = Config.startRow
    orderNum = 1
    ProcessAccountingPolicies policySheet, ws, accountCodes, orderNum, currentRow
    
    ' Close the policy workbook
    policyWorkbook.Close SaveChanges:=False
    
    ' Format the worksheet
    FormatAccountingPolicySheet ws
    
    ' Set up page for PDF output
    SetupPageForPDF ws
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in CreateAccountingPolicySummary: " & Err.Description, vbCritical
    On Error Resume Next
    If Not policyWorkbook Is Nothing Then
        policyWorkbook.Close SaveChanges:=False
    End If
    On Error GoTo 0
End Sub

' Helper function to create the accounting policy sheet
Private Function CreateAccountingPolicySheet(targetWorkbook As Workbook) As Worksheet
    Dim ws As Worksheet
    Set ws = targetWorkbook.Sheets.Add(After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count))
    ws.Name = Config.AccountingPolicySummarySheetName
    Set CreateAccountingPolicySheet = ws
End Function

' Helper function to add the main topic header
Private Sub AddMainTopicHeader(ws As Worksheet)
    ws.Cells(Config.MainTopicRow, 1).Value = "4"
    ws.Cells(Config.MainTopicRow, 1).HorizontalAlignment = xlCenter
    ws.Cells(Config.MainTopicRow, 2).Value = "สรุปนโยบายการบัญชีที่สำคัญ"
    ws.Cells(Config.MainTopicRow, 2).Font.Bold = True
End Sub

' Helper function to open the policy workbook
Private Function OpenPolicyWorkbook(targetWorkbookPath As String) As Workbook
    Dim policyWorkbookPath As String
    policyWorkbookPath = targetWorkbookPath & Config.PolicyWorkbookRelativePath
    
    If Dir(policyWorkbookPath) = "" Then
        Err.Raise vbObjectError + 1, , "Policy workbook not found: " & policyWorkbookPath
    End If
    
    Set OpenPolicyWorkbook = Workbooks.Open(policyWorkbookPath)
End Function

' Helper function to get account codes from Trial Balance
Private Function GetAccountCodes(trialBalanceSheet As Worksheet) As Collection
    Dim accountCodes As New Collection
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = trialBalanceSheet.Cells(trialBalanceSheet.Rows.Count, 2).End(xlUp).row
    For i = 2 To lastRow
        On Error Resume Next
        accountCodes.Add trialBalanceSheet.Cells(i, 2).Value, CStr(trialBalanceSheet.Cells(i, 2).Value)
        On Error GoTo 0
    Next i
    
    Set GetAccountCodes = accountCodes
End Function

' Helper function to process accounting policies
Private Sub ProcessAccountingPolicies(policySheet As Worksheet, ws As Worksheet, accountCodes As Collection, ByRef orderNum As Long, ByRef currentRow As Long)
    Dim lastRow As Long
    Dim i As Long
    Dim codeRange As String, topic As String, detail1 As String, detail2 As String
    Dim matchFound As Boolean
    
    lastRow = policySheet.Cells(policySheet.Rows.Count, 1).End(xlUp).row
    
    For i = 2 To lastRow ' Start from row 2 to skip header
        codeRange = policySheet.Cells(i, 1).Value
        topic = policySheet.Cells(i, 2).Value
        detail1 = policySheet.Cells(i, 3).Value
        detail2 = policySheet.Cells(i, 4).Value
        
        matchFound = (codeRange = "0") Or IsAnyAccountCodeInRange(accountCodes, codeRange)
        
        If matchFound Then
            AddPolicyToSheet ws, orderNum, currentRow, topic, detail1, detail2
        End If
    Next i
End Sub

' Helper function to check if any account code is in the given range
Private Function IsAnyAccountCodeInRange(accountCodes As Collection, codeRange As String) As Boolean
    Dim code As Variant
    For Each code In accountCodes
        If IsCodeInRange(CStr(code), codeRange) Then
            IsAnyAccountCodeInRange = True
            Exit Function
        End If
    Next code
    IsAnyAccountCodeInRange = False
End Function

' Helper function to add a policy to the sheet
Private Sub AddPolicyToSheet(ws As Worksheet, ByRef orderNum As Long, ByRef currentRow As Long, topic As String, detail1 As String, detail2 As String)
    If topic <> "" Then
        ' New topic
        ws.Cells(currentRow, 2).Value = "4." & orderNum
        ws.Cells(currentRow, 2).HorizontalAlignment = xlCenter
        ws.Cells(currentRow, 3).Value = topic
        ws.Cells(currentRow, 3).Font.Bold = True
        FormatContent ws.Range(ws.Cells(currentRow, 3), ws.Cells(currentRow, 9))
        currentRow = currentRow + 1
        orderNum = orderNum + 1
    End If
    
    If detail2 <> "" Then
        ' Detail1 is sub-header
        ws.Cells(currentRow, 3).Value = detail1
        ws.Cells(currentRow, 3).Font.Italic = True
        ws.Cells(currentRow, 3).Font.Bold = True
        FormatContent ws.Range(ws.Cells(currentRow, 3), ws.Cells(currentRow, 9))
        currentRow = currentRow + 1
        ws.Cells(currentRow, 3).Value = detail2
        FormatContent ws.Range(ws.Cells(currentRow, 3), ws.Cells(currentRow, 9))
    Else
        ' Regular detail
        ws.Cells(currentRow, 3).Value = detail1
        FormatContent ws.Range(ws.Cells(currentRow, 3), ws.Cells(currentRow, 9))
    End If
    currentRow = currentRow + 1
End Sub

' Helper function to format the accounting policy sheet
Private Sub FormatAccountingPolicySheet(ws As Worksheet)
    With ws.Cells
        .Font.Name = "TH Sarabun New"
        .Font.Size = 14
    End With
    
    ' Adjust column widths
    ws.Columns("A").ColumnWidth = 5  ' For order numbers
    ws.Columns("B").ColumnWidth = 7  ' For merged content
    ws.Columns("C:I").ColumnWidth = 11  ' For merged content
    
    ' Adjust row heights for merged cells
    AdjustMergedCellsHeight ws
End Sub

Sub FormatContent(contentRange As Range)
    With contentRange
        .Merge
        .WrapText = True
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
    End With
End Sub

Function IsCodeInRange(code As String, codeRange As String) As Boolean
    Dim rangeParts() As String
    Dim lowerBound As String, upperBound As String
    
    If InStr(codeRange, "-") > 0 Then
        rangeParts = Split(codeRange, "-")
        lowerBound = rangeParts(0)
        upperBound = rangeParts(1)
        IsCodeInRange = (code >= lowerBound And code <= upperBound)
    Else
        IsCodeInRange = (code = codeRange)
    End If
End Function

Sub SetupPageForPDF(ws As Worksheet)
    With ws.PageSetup
        .PrintArea = ws.usedRange.Address
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .Zoom = False
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .CenterHorizontally = True
        .CenterVertically = False
    End With
    
    ' Adjust print titles if needed
    ws.PageSetup.PrintTitleRows = "$1:$4"  ' Assuming header is in rows 1-4
    
    ' Add page numbers in the footer
    ws.PageSetup.RightFooter = "Page &P of &N"
End Sub
