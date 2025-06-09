' This module contains only note-generation related helper functions for TB1 format
' Balance Sheet and P&L generation has been removed to focus on notes only

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

Function GetAccountNameFromTB1(TB1Sheet As Worksheet, accountCode As String) As String
    ' Get account name from TB1 sheet by account code
    Dim i As Long
    Dim lastRow As Long
    
    lastRow = TB1Sheet.Cells(TB1Sheet.Rows.Count, 2).End(xlUp).row
    
    For i = 2 To lastRow
        If TB1Sheet.Cells(i, 2).Value = accountCode Then
            GetAccountNameFromTB1 = TB1Sheet.Cells(i, 1).Value  ' Column A = Account Name
            Exit Function
        End If
    Next i
    
    GetAccountNameFromTB1 = ""
End Function

Function SumAccountsByRange(TB1Sheet As Worksheet, startCode As String, endCode As String, periodColumn As Integer) As Double
    ' Sum all accounts within a range for a specific period
    ' periodColumn: 3 = Previous Period (Column C), 4 = Current Period (Column D)
    Dim i As Long
    Dim lastRow As Long
    Dim accountCode As String
    Dim total As Double
    
    total = 0
    lastRow = TB1Sheet.Cells(TB1Sheet.Rows.Count, 2).End(xlUp).row
    
    For i = 2 To lastRow
        accountCode = TB1Sheet.Cells(i, 2).Value
        If accountCode >= startCode And accountCode <= endCode Then
            total = total + TB1Sheet.Cells(i, periodColumn).Value
        End If
    Next i
    
    SumAccountsByRange = total
End Function

Function ValidateTB1Format(TB1Sheet As Worksheet) As Boolean
    ' Validate that TB1 sheet has the correct format
    Dim headers As Variant
    
    On Error GoTo ValidationError
    
    ' Check if required columns exist and have data
    If TB1Sheet.Cells(1, 1).Value = "" Or TB1Sheet.Cells(1, 2).Value = "" Then
        ValidateTB1Format = False
        MsgBox "TB1 sheet is missing required headers in columns A and B", vbExclamation
        Exit Function
    End If
    
    ' Check if there's data in the sheet
    If TB1Sheet.Cells(TB1Sheet.Rows.Count, 2).End(xlUp).row < 2 Then
        ValidateTB1Format = False
        MsgBox "TB1 sheet appears to be empty", vbExclamation
        Exit Function
    End If
    
    ValidateTB1Format = True
    Exit Function
    
ValidationError:
    ValidateTB1Format = False
    MsgBox "Error validating TB1 format: " & Err.Description, vbCritical
End Function
