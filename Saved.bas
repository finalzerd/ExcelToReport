Attribute VB_Name = "Saved"
Sub dCreateDetailTwo(targetWorkbook As Workbook)
    Dim detailWorksheet As Worksheet
    Dim trialPLSheet As Worksheet
    Dim lastRowPL As Long
    Dim i As Long
    Dim accountCode As String
    Dim accountName As String
    Dim amount As Double
    Dim row As Long
    
    ' Check if the worksheet already exists
    On Error Resume Next
    Set detailWorksheet = targetWorkbook.Sheets("DT2")
    On Error GoTo 0
    
    ' If the worksheet doesn't exist, create a new one
    If detailWorksheet Is Nothing Then
        Set detailWorksheet = targetWorkbook.Sheets.Add(After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count))
        detailWorksheet.Name = "DT2"
    End If
    
    ' Set the worksheet for Trial PL
    Set trialPLSheet = targetWorkbook.Sheets("Trial PL 1")
    
    ' Get the last row of data in the Trial PL sheet
    lastRowPL = trialPLSheet.Cells(trialPLSheet.Rows.Count, 1).End(xlUp).row
    
    ' Create the header for the detail worksheet
    CreateHeader detailWorksheet, "Details"
    
    ' Add "รายละเอียดประกอบที่ 2" below the header
    With detailWorksheet.Range("A4")
        .Value = "รายละเอียดประกอบที่ 2"
        .Font.Bold = True
    End With
    
    ' Add column headers
    detailWorksheet.Range("A5").Value = "ค่าใช้จ่ายในการขายและบริหาร"
    detailWorksheet.Range("G5").Value = "ค่าใช้จ่ายในการขาย"
    detailWorksheet.Range("H5").Value = "ค่าใช้จ่ายในการบริหาร"
    detailWorksheet.Range("I5").Value = "ค่าใช้จ่ายอื่น"
    
    ' Initialize the starting row for account details
    row = 6
    
    ' Loop through the Trial PL data and extract account names and amounts
    For i = 2 To lastRowPL
        accountCode = trialPLSheet.Cells(i, 2).Value
        accountName = trialPLSheet.Cells(i, 1).Value
        amount = trialPLSheet.Cells(i, 6).Value
        
        If accountCode >= "5300" And accountCode <= "5399" Then
            ' Add the account name to column A
            detailWorksheet.Cells(row, 1).Value = accountName
            
            Select Case Left(accountCode, 4)
                Case "5309"
                    ' Add the amount to "ค่าใช้จ่ายในการบริหาร" (column F)
                    detailWorksheet.Cells(row, 8).Value = amount
                Case "5310" To "5319"
                    ' Add the amount to "ค่าใช้จ่ายในการขาย" (column E)
                    detailWorksheet.Cells(row, 7).Value = amount
                Case "5320" To "5350"
                    ' Add the amount to "ค่าใช้จ่ายในการบริหาร" (column F)
                    detailWorksheet.Cells(row, 8).Value = amount
                Case "5351" To "5399"
                    ' Add the amount to "ค่าใช้จ่ายอื่น" (column G)
                    detailWorksheet.Cells(row, 9).Value = amount
            End Select
            
            ' Move to the next row
            row = row + 1
        End If
    Next i
    
    ' Apply Thai Sarabun font and font size 14 to the detail worksheet
    detailWorksheet.Cells.Font.Name = "TH Sarabun New"
    detailWorksheet.Cells.Font.Size = 14
    
    ' Set number format to use comma style for columns E, F, and G in the detail worksheet
    detailWorksheet.Columns("E:G").NumberFormat = "#,##0.00"
End Sub
