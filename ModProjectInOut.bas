Attribute VB_Name = "ModProjectInOut"
Public Sub ExportModules()
    Dim ExportYN As Boolean
    Dim DlgOpen As FileDialog
    Dim SourceBook As Excel.Workbook
    Dim SourceBookName As String
    Dim ExportFilePath As String
    Dim EmportFileName As String
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
    Set SourceBook = Application.Workbooks(SourceBookName)
    
    ExportFilePath = ExportFilePath & "\"
    
    For Each VBModule In SourceBook.VBProject.VBComponents
        
        ExportYN = True
        EmportFileName = VBModule.Name

        ''' Concatenate the correct filename for export.
        Select Case VBModule.Type
            Case vbext_ct_ClassModule
                EmportFileName = EmportFileName & ".cls"
            Case vbext_ct_MSForm
                EmportFileName = EmportFileName & ".frm"
            Case vbext_ct_StdModule
                EmportFileName = EmportFileName & ".bas"
            Case vbext_ct_Document
                EmportFileName = EmportFileName & ".cls"
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
'                ExportYN = False
        End Select
        
        If ExportYN Then
            ''' Export the component to a text file.
            VBModule.Export ExportFilePath & EmportFileName
            
        End If
   
    Next VBModule
    
    Set DlgOpen = Nothing

    MsgBox "Export is ready"
End Sub

Public Sub ImportModules()
    Dim TargetBook As Excel.Workbook
    Dim FSO As Scripting.FileSystemObject
    Dim FileObj As Scripting.File
    Dim TargetBookName As String
    Dim DlgOpen As FileDialog
    Dim ImportFilePath As String
    Dim ImportFileName As String
    Dim VBModules As VBIDE.VBComponents

    'open files
    Set DlgOpen = Application.FileDialog(msoFileDialogFolderPicker)
    
     With DlgOpen
        .Title = "Select Import Folder"
        .Show
    End With
        
    ImportFilePath = DlgOpen.SelectedItems(1) & "\"
    ''' NOTE: This workbook must be open in Excel.
    TargetBookName = ActiveWorkbook.Name
    Set TargetBook = Application.Workbooks(TargetBookName)
            
    Set FSO = New Scripting.FileSystemObject
    If FSO.GetFolder(ImportFilePath).Files.Count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If

    Set VBModules = TargetBook.VBProject.VBComponents
    
    ''' Import all the code modules in the specified path
    ''' to the ActiveWorkbook.
    For Each FileObj In FSO.GetFolder(ImportFilePath).Files
    
        If (FSO.GetExtensionName(FileObj.Name) = "cls") Or _
            (FSO.GetExtensionName(FileObj.Name) = "frm") Or _
            (FSO.GetExtensionName(FileObj.Name) = "bas") Then
            VBModules.Import FileObj.Path
        End If
        
    Next FileObj
    
    MsgBox "Import is ready"
End Sub
 
Public Sub RemoveAllModules()
    Dim ExportYN As Boolean
    Dim DlgOpen As FileDialog
    Dim SourceBook As Excel.Workbook
    Dim SourceBookName As String
    Dim ExportFilePath As String
    Dim ImportFileName As String
    Dim VBModule As VBIDE.VBComponent
   
    ''' NOTE: This workbook must be open in Excel.
    SourceBookName = ActiveWorkbook.Name
    Set SourceBook = Application.Workbooks(SourceBookName)
        
    For Each VBModule In SourceBook.VBProject.VBComponents
        
        ''' remove it from the project if you want
        If VBModule.Type <> vbext_ct_Document Then SourceBook.VBProject.VBComponents.Remove VBModule
           
    Next VBModule
    
    Set DlgOpen = Nothing

End Sub



