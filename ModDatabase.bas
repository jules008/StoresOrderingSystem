Attribute VB_Name = "ModDatabase"
'===============================================================
' Module ModDatabase
' v0,0 - Initial Version
' v0,1 - Improved message box
' v0,2 - Added GetDBVer function
' v0,33 - Asset Import functionality
' v0,4 - Removed Asset Import functionality to new Module
' v0,5 - Added System Message
' v0,6 - Seperated out Update Message procedure
' v0,7 - Added Release Notes
' v0,8 - Show logged on users
'---------------------------------------------------------------
' Date - 30 Nov 17
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
' UpdateDBScript
' Script to update DB
' ---------------------------------------------------------------
Public Sub UpdateDBScript()
    Dim TableDef As DAO.TableDef
    Dim Ind As DAO.Index
    Dim RstTable As Recordset
    Dim i As Integer
    
    Dim Fld As DAO.Field
    
    DBConnect
        
    Set RstTable = SQLQuery("TblDBVersion")
    
    'check preceding DB Version
    If RstTable.Fields(0) <> "v1,392" Then
        MsgBox "Database needs to be upgraded to v1,393 to continue", vbOKOnly + vbCritical
        Exit Sub
    End If
    
    'Table changes
    DB.Execute "CREATE TABLE TblReports"
    DB.Execute "ALTER TABLE TblReports ADD COLUMN ReportNo Int"
    DB.Execute "ALTER TABLE TblReports ADD COLUMN ReportType Integer"
    DB.Execute "ALTER TABLE TblReports ADD COLUMN ReportName Text"
    DB.Execute "ALTER TABLE TblReports ADD COLUMN DueDate Date"
    DB.Execute "ALTER TABLE TblReports ADD COLUMN Frequency number"
    DB.Execute "INSERT INTO TblReports VALUES (1, 1,'CFS Stock Count Report', '27 Dec 17', 7)"
    DB.Execute "INSERT INTO TblReports VALUES (2, 2,'New Guest User', '', 0)"
    DB.Execute "INSERT INTO TblReports VALUES (3, 2,'Support Query Recieved', '', 0)"
    DB.Execute "INSERT INTO TblReports VALUES (4, 2,'New Order Received', '', 0)"
    
    DB.Execute "CREATE TABLE TblRptsAlerts"
    DB.Execute "ALTER TABLE TblRptsAlerts ADD COLUMN CrewNo Text"
    DB.Execute "ALTER TABLE TblRptsAlerts ADD COLUMN ReportNo Int"
    DB.Execute "ALTER TABLE TblRptsAlerts ADD COLUMN ToCC Text"
    DB.Execute "INSERT INTO TblRptsAlerts VALUES ('3682', 1, 'To')"
    DB.Execute "INSERT INTO TblRptsAlerts VALUES ('5398', 1, 'CC')"
    DB.Execute "INSERT INTO TblRptsAlerts VALUES ('5073', 1, 'CC')"
    DB.Execute "INSERT INTO TblRptsAlerts VALUES ('5398', 2, 'To')"
    DB.Execute "INSERT INTO TblRptsAlerts VALUES ('5196', 2, 'To')"
    DB.Execute "INSERT INTO TblRptsAlerts VALUES ('5398', 3, 'To')"
    DB.Execute "INSERT INTO TblRptsAlerts VALUES ('5073', 3, 'To')"
    DB.Execute "INSERT INTO TblRptsAlerts VALUES ('5196', 3, 'To')"
    DB.Execute "INSERT INTO TblRptsAlerts VALUES ('2506', 3, 'CC')"
    DB.Execute "INSERT INTO TblRptsAlerts VALUES ('5073', 4, 'To')"
    DB.Execute "INSERT INTO TblRptsAlerts VALUES ('5075', 4, 'To')"
    
    'update DB Version
    With RstTable
        .Edit
        .Fields(0) = "v1,393"
        .Update
    End With
    
    UpdateSysMsg
    
    MsgBox "Database successfully updated", vbOKOnly + vbInformation
    
    Set RstTable = Nothing
    Set TableDef = Nothing
    Set Fld = Nothing
    
End Sub
              
' ===============================================================
' UpdateDBScriptUndo
' Script to update DB
' ---------------------------------------------------------------
Public Sub UpdateDBScriptUndo()
    Dim TableDef As DAO.TableDef
    Dim Ind As DAO.Index
    Dim RstTable As Recordset
    Dim i As Integer
        
    Dim Fld As DAO.Field
        
    DBConnect
    
    DB.Execute "DROP TABLE TblReports"
    DB.Execute "DROP TABLE TblRptsAlerts"
     
    Set RstTable = SQLQuery("TblDBVersion")

    With RstTable
        .Edit
        .Fields(0) = "v1,392"
        .Update
    End With
    
    MsgBox "Database reset successfully", vbOKOnly + vbInformation
    
    Set RstTable = Nothing
    Set TableDef = Nothing
    Set Fld = Nothing

End Sub

' ===============================================================
' GetDBVer
' Returns the version of the DB
' ---------------------------------------------------------------
Public Function GetDBVer() As String
    Dim DBVer As Recordset
    
    Const StrPROCEDURE As String = "GetDBVer()"

    On Error GoTo ErrorHandler

    Set DBVer = SQLQuery("TblDBVersion")

    GetDBVer = DBVer.Fields(0)

    Debug.Print DBVer.Fields(0)
    Set DBVer = Nothing
Exit Function

ErrorExit:

    GetDBVer = ""
    
    Set DBVer = Nothing

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' UpdateSysMsg
' Updates the system message and resets read flags
' ---------------------------------------------------------------
Public Sub UpdateSysMsg()
    Dim RstMessage As Recordset
    
    Set RstMessage = SQLQuery("TblMessage")
    
    With RstMessage
        If .RecordCount = 0 Then
            .AddNew
        Else
            .Edit
        End If
        
        .Fields("SystemMessage") = "Version 1.6 - What's New" _
                    & Chr(13) & "(See Release Notes on Support tab for further information)" _
                    & Chr(13) & "" _
                    & Chr(13) & " - Weekly CFS Stock email " _
                    & Chr(13) & "" _
                    & Chr(13) & " - New Reports/Alerts Admin Console " _
        
        .Fields("ReleaseNotes") = "Software Version 1.6" _
                    & Chr(13) & "Database Version 1.393" _
                    & Chr(13) & "Date 22 Dec 17" _
                    & Chr(13) & "" _
                    & Chr(13) & "- 'Weekly CFS Stock Email - The system will now automatically send a  stock status email " _
                    & Chr(13) & "" _
                    & Chr(13) & "- New Reports/Alerts Admin Console - This is accessed from the User Admin Form and allows " _
                    & Chr(13) & "the target addresses (To/CC) for all system alerts and automatic email reports to be managed"
        .Update
    End With
    
    'reset read flags
    DB.Execute "UPDATE TblPerson SET MessageRead = False WHERE MessageRead = True"
    
    Set RstMessage = Nothing

End Sub

' ===============================================================
' ShowUsers
' Show users logged onto system
' ---------------------------------------------------------------
Public Sub ShowUsers()
    Dim RstUsers As Recordset
    
    Set RstUsers = SQLQuery("TblUsers")
    
    With RstUsers
        Debug.Print
        Do While Not .EOF
            Debug.Print "User: " & .Fields(0) & " - Logged on: " & .Fields(1)
            .MoveNext
        Loop
    End With
    
    Set RstUsers = Nothing
End Sub
