Attribute VB_Name = "ModProjectInOut"
Public Sub ExportModules()
    Dim ExportYN As Boolean
    Dim DlgOpen As FileDialog
    Dim SourceBook As Excel.Workbook
    Dim SourceBookName As String
    Dim ExportFilePath As String
    Dim ExportFileName As String
    Dim VBModule As VBIDE.VBComponent
   
    'open files
    Set DlgOpen = Application.FileDialog(msoFileDialogFolderPicker)
    
     With DlgOpen
        .Title = "Select Export Folder"
        .Show
    End With
        
    ExportFilePath = DlgOpen.SelectedItems(1)

    ''' NOTE: This workbook must be open in Excel.
    SourceBookName = ActiveWorkbook.Name
Debug.Print SourceBookName
    Set SourceBook = Application.Workbooks(SourceBookName)
    
    ExportFilePath = ExportFilePath & "\"
    
    For Each VBModule In SourceBook.VBProject.VBComponents

Debug.Print SourceBook.VBProject.VBComponents.Count
        
        ExportYN = True
        ExportFileName = VBModule.Name

        ''' Concatenate the correct filename for export.
        Select Case VBModule.Type
            Case vbext_ct_ClassModule
                ExportFileName = ExportFileName & ".cls"
            Case vbext_ct_MSForm
                ExportFileName = ExportFileName & ".frm"
            Case vbext_ct_StdModule
                ExportFileName = ExportFileName & ".bas"
            Case vbext_ct_Document
                ExportFileName = ExportFileName & ".cls"
        End Select
        
        If ExportYN Then
            ''' Export the component to a text file.
            VBModule.Export ExportFilePath & ExportFileName
            
            ''' remove it from the project if you want
            If VBModule.Type <> vbext_ct_Document Then
                SourceBook.VBProject.VBComponents.Remove VBModule
            End If
            
        End If
   
    Next VBModule
    
    Set DlgOpen = Nothing

    MsgBox "Export is ready"
End Sub


Public Sub ImportModules()
    Dim wkbTarget As Excel.Workbook
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Dim szTargetWorkbook As String
    Dim szImportPath As String
    Dim ExportFileName As String
    Dim VBModules As VBIDE.VBComponents

    If ActiveWorkbook.Name = ThisWorkbook.Name Then
        MsgBox "Select another destination workbook" & _
        "Not possible to import in this workbook "
        Exit Sub
    End If

    'Get the path to the folder with modules
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Import Folder not exist"
        Exit Sub
    End If

    ''' NOTE: This workbook must be open in Excel.
    szTargetWorkbook = ActiveWorkbook.Name
    Set wkbTarget = Application.Workbooks(szTargetWorkbook)
    
    If wkbTarget.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to Import the code"
    Exit Sub
    End If

    ''' NOTE: Path where the code modules are located.
    szImportPath = FolderWithVBAProjectFiles & "\"
        
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(szImportPath).Files.Count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If

    'Delete all modules/Userforms from the ActiveWorkbook
    Call DeleteVBAModulesAndUserForms

    Set VBModules = wkbTarget.VBProject.VBComponents
    
    ''' Import all the code modules in the specified path
    ''' to the ActiveWorkbook.
    For Each objFile In objFSO.GetFolder(szImportPath).Files
    
        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            VBModules.Import objFile.Path
        End If
        
    Next objFile
    
    MsgBox "Import is ready"
End Sub

Function DeleteVBAModulesAndUserForms()
        Dim VBProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        
        Set VBProj = ActiveWorkbook.VBProject
        
        For Each VBComp In VBProj.VBComponents
            If VBComp.Type = vbext_ct_Document Then
                'Thisworkbook or worksheet module
                'We do nothing
            Else
                VBProj.VBComponents.Remove VBComp
            End If
        Next VBComp
End Function
 

