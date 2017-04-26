Attribute VB_Name = "ModDatabase"
'===============================================================
' Module ModDatabase
' v0,0 - Initial Version
' v0,1 - Improved message box
'---------------------------------------------------------------
' Date - 19 Apr 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModDatabase"

Public DB As DAO.Database
Public MyQueryDef As DAO.QueryDef

' ===============================================================
' SQLQuery
' Queries database with given SQL script
' ---------------------------------------------------------------
Public Function SQLQuery(SQL As String) As Recordset
    Dim RstResults As Recordset
    
    Const StrPROCEDURE As String = "SQLQuery()"

    On Error GoTo ErrorHandler
      
Restart:
    Application.StatusBar = ""

    If DB Is Nothing Then
        Err.Raise NO_DATABASE_FOUND, Description:="Unable to connect to database"
    Else
        If ModErrorHandling.FaultCount1008 > 0 Then ModErrorHandling.FaultCount1008 = 0
    
        Set RstResults = DB.OpenRecordset(SQL, dbOpenDynaset)
        Set SQLQuery = RstResults
    End If
    
    Set RstResults = Nothing
    
Exit Function

ErrorExit:

    Set RstResults = Nothing

    Set SQLQuery = Nothing
    Terminate

Exit Function

ErrorHandler:
    
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        If CustomErrorHandler(Err.Number) Then
            If Not Initialise Then Err.Raise HANDLED_ERROR
            Resume Restart
        Else
            Err.Raise HANDLED_ERROR
        End If
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' DBConnect
' Provides path to database
' ---------------------------------------------------------------
Public Function DBConnect() As Boolean
    Const StrPROCEDURE As String = "DBConnect()"

    On Error GoTo ErrorHandler

    DB_PATH = ShtSettings.Range("DBPath")
    
    If DB_PATH = "" Then
        MsgBox ("No database selected")
        ModDatabase.SelectDB
    Else
        Set DB = OpenDatabase(DB_PATH)
    End If
    DBConnect = True

Exit Function

ErrorExit:

    DBConnect = False

Exit Function

ErrorHandler:

If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function


' ===============================================================
' DBTerminate
' Disconnects and closes down DB connection
' ---------------------------------------------------------------
Public Function DBTerminate() As Boolean
    Const StrPROCEDURE As String = "DBTerminate()"

    On Error GoTo ErrorHandler

    If Not DB Is Nothing Then DB.Close
    Set DB = Nothing

    DBTerminate = True

Exit Function

ErrorExit:

    DBTerminate = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' SelectDB
' Selects DB to connect to
' ---------------------------------------------------------------
Public Function SelectDB() As Boolean
    Const StrPROCEDURE As String = "SelectDB()"

    On Error GoTo ErrorHandler
    Dim DlgOpen As FileDialog
    Dim FileLoc As String
    Dim NoFiles As Integer
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    'open files
    Set DlgOpen = Application.FileDialog(msoFileDialogOpen)
    
     With DlgOpen
        .Filters.Clear
        .Filters.Add "Access Files (*.accdb)", "*.accdb"
        .AllowMultiSelect = False
        .Title = "Connect to Database"
        .Show
    End With
    
    'get no files selected
    NoFiles = DlgOpen.SelectedItems.Count
    
    'exit if no files selected
    If NoFiles = 0 Then
        MsgBox "There was no database selected", vbOKOnly + vbExclamation, "No Files"
        SelectDB = True
        Exit Function
    End If
  
    'add files to array
    For i = 1 To NoFiles
        FileLoc = DlgOpen.SelectedItems(i)
    Next
    
    DB_PATH = FileLoc
    
    Set DlgOpen = Nothing

    SelectDB = True

Exit Function

ErrorExit:

    Set DlgOpen = Nothing
    SelectDB = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' ImportAssetFile
' Imports the Asset file into the database
' ---------------------------------------------------------------
Private Sub ImportAssetFile()
    Dim DlgOpen As FileDialog
    Dim LineInputString As String
    Dim AssetData() As String
    Dim AssetFileLoc As String
    Dim NoFiles As Integer
    Dim AssetFile As Integer
    Dim RstAssets As Recordset
    Dim i As Integer
    
    Const StrPROCEDURE As String = "ImportAssetFile()"

    On Error GoTo ErrorHandler
    Set DlgOpen = Application.FileDialog(msoFileDialogOpen)
    
     With DlgOpen
        .Filters.Clear
        .Filters.Add "CSV Files (*.csv)", "*.csv"
        .AllowMultiSelect = False
        .Title = "Select Spreadsheet of Doom"
        .Show
    End With
    
    'get no files selected
    NoFiles = DlgOpen.SelectedItems.Count
    
    'exit if no files selected
    If NoFiles = 0 Then Err.Raise NO_FILE_SELECTED
  
    AssetFileLoc = DlgOpen.SelectedItems(1)

    AssetFile = FreeFile()
    
    If Dir(AssetFileLoc) = "" Then Err.Raise NO_FILE_SELECTED
    
    'get Asset Recordset
    Set RstAssets = SQLQuery("TblAssets")
    
    Open AssetFileLoc For Input As AssetFile
    
    While Not EOF(AssetFile)
        Line Input #AssetFile, LineInputString
        AssetData = Split(LineInputString, ",")
        i = i + 1
        
        
        
    Wend
    
    Close #AssetFile





    Set RstAssets = Nothing

Exit Sub

ErrorExit:

'    ***CleanUpCode***
    Set RstAssets = Nothing
Exit Sub

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' UpdateDBScript
' Script to update DB
' ---------------------------------------------------------------
Public Sub UpdateDBScript()
    Dim TableDef As DAO.TableDef
    Dim Ind As DAO.Index
    Dim RstTable As Recordset
    Dim i As Integer
    
    Dim Fld As DAO.Field
    
    Initialise
    
    Set TableDef = DB.CreateTableDef("TblDBVersion")
    
    With TableDef
        
        Set Fld = .CreateField("Version", dbText)

        .Fields.Append Fld
        DB.TableDefs.Append TableDef
        
    End With
        
    Set RstTable = SQLQuery("TblDBVersion")
    
    With RstTable
        .AddNew
        .Fields(0) = "v0,31"
        .Update
    End With
    
    Set TableDef = DB.TableDefs("TblAsset")
    
    With TableDef
        
        For Each Ind In .Indexes
            Debug.Print Ind.Name
        
        Next
        
        With .Indexes
            .Delete "PrimaryKey"
            .Delete "ID"
        End With
        .Fields.Delete "ID"
        
        
    End With
    
    Set RstTable = Nothing
    Set TableDef = Nothing
    Set Fld = Nothing

End Sub
