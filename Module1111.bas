Attribute VB_Name = "Module1111"
Sub ExportAllModules()
    Dim objProject As VBIDE.VBProject
    Dim objComponent As VBIDE.VBComponent
    Dim objFSO As Object
    Dim objFile As Object
    Dim strPath As String
    Dim strFile As String
    Dim strExtension As String
    
    ' Set reference to Microsoft Visual Basic for Applications Extensibility 5.3
    ' Tools -> References -> Microsoft Visual Basic for Applications Extensibility 5.3
    
    ' Create a new FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Set the path to export the modules
    strPath = Application.DefaultFilePath & "\VBAModulesMain\"
    
    ' Create the folder if it doesn't exist
    If Not objFSO.FolderExists(strPath) Then
        objFSO.CreateFolder strPath
    End If
    
    ' Set reference to the current project
    Set objProject = ThisWorkbook.VBProject
    
    ' Loop through all components in the project
    For Each objComponent In objProject.VBComponents
        ' Determine the file extension based on the Type of module
        Select Case objComponent.Type
            Case vbext_ct_ClassModule
                strExtension = ".cls"
            Case vbext_ct_MSForm
                strExtension = ".frm"
            Case vbext_ct_StdModule
                strExtension = ".bas"
            Case vbext_ct_Document
                ' This is a worksheet or workbook object.
                ' Don't export.
                GoTo NextComponent
            Case Else
                ' This is another type
                GoTo NextComponent
        End Select
        
        ' Create the file name
        strFile = strPath & objComponent.Name & strExtension
        
        ' Export the component to a text file
        objComponent.Export strFile
        
        ' Confirm export
        Debug.Print "Exported " & objComponent.Name & " to " & strFile
        
NextComponent:
    Next objComponent
    
    MsgBox "All modules have been exported to: " & strPath, vbInformation
    
    ' Clean up
    Set objFSO = Nothing
    Set objProject = Nothing
End Sub
