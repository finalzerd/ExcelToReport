Sub CreateGeneralInformation(targetWorkbook As Workbook)
    Dim ws As Worksheet
    Dim infoSheet As Worksheet
    Dim entityNumber As String
    Dim entityFilePath As String
    Dim entityData As Object
    Dim mergedRange As Range
    
    ' Check if GIC sheet exists, if not create it
    On Error Resume Next
    Set ws = targetWorkbook.Sheets("GIC")
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' Create new GIC sheet
        Set ws = targetWorkbook.Sheets.Add(After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count))
        ws.Name = "GIC"
    End If
    
    ' Set the Info sheet
    Set infoSheet = targetWorkbook.Sheets("Info")
    
    ' Get the entity number from the Info sheet
    entityNumber = infoSheet.Range("B4").Value
    
    ' Construct the file path
    entityFilePath = targetWorkbook.Path & "\ExtractWebDBD\" & entityNumber & ".csv"
    
    ' Check if the file exists
    If Dir(entityFilePath) = "" Then
        MsgBox "Entity file not found: " & entityFilePath, vbExclamation
        Exit Sub
    End If
    
    ' Read the CSV file
    Set entityData = ReadCSV(entityFilePath)
    
    ' Create the header
    CreateHeader ws, "General Information"
    
    ' Add the main title
    ws.Cells(5, 1).Value = "1"
    ws.Cells(5, 1).HorizontalAlignment = xlCenter
    ws.Cells(5, 2).Value = "ข้อมูลทั่วไป"
    ws.Cells(5, 2).Font.Bold = True
    
    ' Add 1.1 Legal Status
    ws.Cells(7, 3).Value = "1.1"
    ws.Cells(7, 3).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(7, 4), ws.Cells(7, 5)).Merge
    ws.Cells(7, 4).Value = "สถานะทางกฎหมาย"
    ws.Range(ws.Cells(7, 6), ws.Cells(7, 8)).Merge
    ws.Cells(7, 6).Value = "เป็นนิติบุคคลจัดตั้งตามกฎหมายไทย"
    
    ' Add 1.2 Location
    ws.Cells(8, 3).Value = "1.2"
    ws.Cells(8, 3).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(8, 4), ws.Cells(8, 5)).Merge
    ws.Cells(8, 4).Value = "สถานที่ตั้ง"
    Set mergedRange = ws.Range(ws.Cells(8, 6), ws.Cells(8, 8))
    mergedRange.Merge
    mergedRange.Value = entityData("D2")
    FormatAndAdjustCell mergedRange
    
    ' Add 1.3 Business Description
    ws.Cells(9, 3).Value = "1.3"
    ws.Cells(9, 3).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(9, 4), ws.Cells(9, 5)).Merge
    ws.Cells(9, 4).Value = "ลักษณะธุรกิจและการดำเนินงาน"
    Set mergedRange = ws.Range(ws.Cells(9, 6), ws.Cells(9, 8))
    mergedRange.Merge
    mergedRange.Value = entityData("E2")
    FormatAndAdjustCell mergedRange
    
    ' Add Company Status with order number 2
    ws.Cells(11, 1).Value = "2"
    ws.Cells(11, 1).HorizontalAlignment = xlCenter
    ws.Cells(11, 2).Value = "ฐานะการดำเนินงานของบริษัท"
    ws.Cells(11, 2).Font.Bold = True
    
    ' Combine text for company status
    Dim statusText As String
    statusText = entityData("G2") & " ได้จดทะเบียนตามประมวลกฎหมายแพ่งและพาณิชย์เป็นนิติบุคคล ประเภท " & _
                 entityData("A2") & " เมื่อวันที่ " & entityData("B2") & " ทะเบียนเลขที่ " & entityData("H2")
    
    Set mergedRange = ws.Range(ws.Cells(12, 3), ws.Cells(12, 8))
    mergedRange.Merge
    mergedRange.Value = statusText
    FormatAndAdjustCell mergedRange
    
    ' Format the worksheet
    ws.Cells.Font.Name = "TH Sarabun New"
    ws.Cells.Font.Size = 14
    
    ' Adjust column widths
    ws.Columns("A").ColumnWidth = 5
    ws.Columns("B:H").ColumnWidth = 11
End Sub

Function ReadCSV(filePath As String) As Object
    Dim fso As Object
    Dim ts As Object
    Dim line As String
    Dim parts() As String
    Dim dict As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 1)
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Read header
    line = ts.ReadLine
    parts = Split(line, ",")
    
    ' Read data
    line = ts.ReadLine
    parts = Split(line, ",")
    
    ' Populate dictionary
    For i = 0 To UBound(parts)
        dict.Add Chr(65 + i) & "2", parts(i)
    Next i
    
    ts.Close
    Set ReadCSV = dict
End Function

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
        AdjustMergedCellsHeightInWorksheet gicSheet
    Else
        ' If "GIC" sheet doesn't exist, adjust the current worksheet
        AdjustMergedCellsHeightInWorksheet ws
    End If
End Sub

