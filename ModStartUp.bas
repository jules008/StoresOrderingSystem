Attribute VB_Name = "ModStartUp"
'===============================================================
' Module ModStartUp
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 17 Jan 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModStartUp"

' ===============================================================
' Initialise
' Creates the environment for system start up
' ---------------------------------------------------------------
Public Function Initialise() As Boolean
    Const StrPROCEDURE As String = "Initialise()"
    Dim UserName As String
    
    On Error GoTo ErrorHandler
    
    Terminate
        
    Set CurrentUser = New ClsPerson
    Set Vehicles = New ClsVehicles
    Set Stations = New ClsStations
        
    ShtMain.Unprotect
    
    If Not ReadINIFile Then Err.Raise HANDLED_ERROR
    
    DB_PATH = ShtSettings.Range("DBPath")
    
    If Not ModDatabase.DBConnect Then Err.Raise HANDLED_ERROR
    
    If DEV_MODE Then
        ShtSettings.Visible = xlSheetVisible
    Else
        ShtSettings.Visible = xlSheetHidden
    End If
        
    'build collections
    Vehicles.GetCollection
    Stations.GetCollection
    
    'get username of current user
    If Not ModStartUp.GetUserName Then Err.Raise HANDLED_ERROR
    
    'build styles
    If Not ModUIMenu.BuildStylesMenu Then Err.Raise HANDLED_ERROR
    If Not ModUIMainScreen.BuildStylesMainScreen Then Err.Raise HANDLED_ERROR
    
    'Build menu and backdrop
    If Not ModUIMenu.BuildMenu Then Err.Raise HANDLED_ERROR
    
    If [menuitemno] = "" Then
           
        ModUIMenu.ProcessBtnPress (1)
        [menuitemno] = 1
        
    Else
        ModUIMenu.ProcessBtnPress ([menuitemno])
    End If
        
    ActiveSheet.Range("A1").Select

    ShtMain.Protect
    
    Initialise = True

Exit Function

ErrorExit:

    Set CurrentUser = Nothing
    Set Vehicles = Nothing
    Initialise = False
    
Exit Function

ErrorHandler:
        
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        CustomErrorHandler Err.Number
        Resume Next
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' GetUserName
' gets username from windows, or test user if in test mode
' ---------------------------------------------------------------
Public Function GetUserName() As Boolean
    Dim UserName As String
    Dim CharPos As Integer
    
    Const StrPROCEDURE As String = "GetUserName()"

    On Error GoTo ErrorHandler
    
    If TEST_MODE Then
        If ShtSettings.Range("C15") = True Then
            UserName = ShtSettings.Range("Test_User")
        Else
            UserName = Application.UserName
        End If
    Else
        UserName = Application.UserName
    End If
    
    If UserName = "" Then Err.Raise HANDLED_ERROR

    UserName = Replace(UserName, "'", "")
    UserName = Replace(UserName, "-", "")
    
    CurrentUser.DBGet Trim(UserName)
    
    If CurrentUser.CrewNo = "" Then Err.Raise UNKNOWN_USER

    GetUserName = True

GracefulExit:

Exit Function

ErrorExit:

    GetUserName = False

Exit Function

ErrorHandler:
        
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        CustomErrorHandler Err.Number
        Resume GracefulExit
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' ReadINIFile
' Gets start up variables from ini file
' ---------------------------------------------------------------
Private Function ReadINIFile() As Boolean
    Dim AppFPath As String
    Dim IniFPath As String
    Dim DebugMode As String
    Dim TestMode As String
    Dim OutputMode As String
    Dim EnablePrint As String
    Dim DBPath As String
    Dim SendEmails As String
    Dim DevMode As String
    Dim INIFile As Integer
    Dim TmpFPath As String
    Dim StopFlag As String
    
    Const StrPROCEDURE As String = "ReadINIFile()"

    On Error GoTo ErrorHandler

    AppFPath = ThisWorkbook.Path
    IniFPath = AppFPath & INI_FILE_PATH
    INIFile = FreeFile()
    
    Open IniFPath & INI_FILE For Input As #INIFile
    
    Line Input #INIFile, DebugMode
    Line Input #INIFile, TestMode
    Line Input #INIFile, OutputMode
    Line Input #INIFile, EnablePrint
    Line Input #INIFile, DBPath
    Line Input #INIFile, SendEmails
    Line Input #INIFile, DevMode
    Line Input #INIFile, TmpFPath
    Line Input #INIFile, StopFlag
    
    Close #INIFile
    
    DEBUG_MODE = CBool(DebugMode)
    TEST_MODE = CBool(TestMode)
    OUTPUT_MODE = OutputMode
    ENABLE_PRINT = CBool(EnablePrint)
    ShtSettings.Range("DBPath") = DBPath
    SEND_EMAILS = CBool(SendEmails)
    DEV_MODE = CBool(DevMode)
    TMP_FILE_PATH = TmpFPath
    
    If StopFlag = True Then Stop
    
   ReadINIFile = True

Exit Function

ErrorExit:

'    ***CleanUpCode***
    ReadINIFile = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function
