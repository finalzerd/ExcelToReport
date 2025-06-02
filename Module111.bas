Attribute VB_Name = "Module111"
Option Explicit

Sub MassImportModules()
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim strFolderPath As String
    Dim strExtension As String
    
    ' Set reference to the FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Prompt user to select the folder containing the modules
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "Select Folder Containing Modules"
        .AllowMultiSelect = False
        If .Show = -1 Then ' If a folder is selected
            strFolderPath = .SelectedItems(1)
        Else
            MsgBox "No folder selected. Import cancelled.", vbExclamation
            Exit Sub
        End If
    End With
    
    ' Set reference to the selected folder
    Set objFolder = objFSO.GetFolder(strFolderPath)
    
    ' Loop through each file in the folder
    For Each objFile In objFolder.Files
        strExtension = LCase(objFSO.GetExtensionName(objFile.Name))
        
        Select Case strExtension
            Case "bas", "cls", "frm"
                ' Import the component
                ThisWorkbook.VBProject.VBComponents.Import objFile.Path
            ' You can add more cases here for other file types if needed
        End Select
    Next objFile
    
    ' Clean up
    Set objFSO = Nothing
    Set objFolder = Nothing
    
    MsgBox "All modules from the selected folder have been imported.", vbInformation
End Sub

