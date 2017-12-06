Attribute VB_Name = "ModStartUp"
'===============================================================
' Module ModStartUp
' v0,0 - Initial Version
' v0,1 - Added maintenance flag start up option
' v0,2 - Bug fix for maintenance flag
' v0,3 - Hide more sheets plus bug fixes
' v0,4 - Changed start up so always starts Menu 1
' v0,5 - reverted back to restart back to previous menu item
' v0,6 - Stopped the removal of '-' from the user name
' v0,7 - Added DB Version Check
' v0,8 - Update any rogue usernames
' v0,9 - Added System Messages on start up
' v0,10 - Added optional parameter to allow start up to page
' v0,11 - Log On user in DB
' v0,13 - changes to filepth for laptop change
'---------------------------------------------------------------
' Date - 06 Dec 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModStartUp"

' ===============================================================
' Initialise
' Creates the environment for system start up
' ---------------------------------------------------------------
Public Function Initialise(Optional MenuItem As Integer) As Boolean
    Const StrPROCEDURE As String = "Initialise()"
    Dim UserName As String
    
    On Error GoTo ErrorHandler
    
    Terminate
    
    Application.StatusBar = "Initialising....."
    
    Set CurrentUser = New ClsPerson
    Set Vehicles = New ClsVehicles
    Set Stations = New ClsStations
        
    ShtMain.Unprotect
    
    Application.StatusBar = "Reading INI File....."
    
    If Not ReadINIFile Then Err.Raise HANDLED_ERROR
    
    DB_PATH = ShtSettings.Range("DBPath")
    
    Application.StatusBar = "Connecting to DB....."
    
    If Not ModDatabase.DBConnect Then Err.Raise HANDLED_ERROR
    
    Application.StatusBar = "Checking DB Version....."
    
    If ModDatabase.GetDBVer <> DB_VER Then Err.Raise DB_WRONG_VER
    
    If DEV_MODE Then
        ShtSettings.Visible = xlSheetVisible
        ShtLists.Visible = xlSheetVisible
        ShtOrderList.Visible = xlSheetVisible
    
    Else
        ShtSettings.Visible = xlSheetHidden
        ShtLists.Visible = xlSheetHidden
        ShtOrderList.Visible = xlSheetHidden
    End If
        
    Application.StatusBar = "Building Data....."
        
    'build collections
    Vehicles.GetCollection
    Stations.GetCollection
    
    Application.StatusBar = "Finding User....."
    'get username of current user
    If Not ModStartUp.GetUserName Then Err.Raise HANDLED_ERROR
    CurrentUser.LogOn
    
    Application.StatusBar = "Buidling UI....."
    
    'build styles
    If Not ModUIMenu.BuildStylesMenu Then Err.Raise HANDLED_ERROR
    If Not ModUIMainScreen.BuildStylesMainScreen Then Err.Raise HANDLED_ERROR
    
    'Build menu and backdrop
    If Not ModUIMenu.BuildMenu Then Err.Raise HANDLED_ERROR
    
    If [menuitemno] = "" Then
        If MenuItem = 0 Then
            ModUIMenu.ProcessBtnPress (1)
        Else
            ModUIMenu.ProcessBtnPress (MenuItem)
        End If
    Else
        ModUIMenu.ProcessBtnPress ([menuitemno])
    End If
        
    'Show any messages
    If Not MessageCheck Then Err.Raise HANDLED_ERROR
        
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
    
    If Not UpdateUsername Then Err.Raise HANDLED_ERROR
    
    If DEV_MODE Then
       If ShtSettings.Range("C15") = True Then
            UserName = ShtSettings.Range("Test_User")
        Else
            UserName = "Julian Turner"
        End If
    Else
        UserName = Application.UserName
    End If
    
    If UserName = "" Then Err.Raise HANDLED_ERROR

    UserName = Replace(UserName, "'", "")
    
    CurrentUser.DBGet Trim(UserName)
    
    If CurrentUser.CrewNo = "" Then Err.Raise UNKNOWN_USER

GracefulExit:
    
    GetUserName = True

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
    Dim DebugMode As String
    Dim TestMode As String
    Dim OutputMode As String
    Dim EnablePrint As String
    Dim DBPath As String
    Dim SendEmails As String
    Dim DevMode As String
    Dim TmpFPath As String
    Dim INIFile As Integer
    Dim INIFPath As String
    Dim StopFlag As String
    Dim MaintMsg As String
    
    Const StrPROCEDURE As String = "ReadINIFile()"

    On Error GoTo ErrorHandler
    
    AppFPath = ThisWorkbook.Path
    
    If Left(AppFPath, 5) = "https" Then
        INIFPath = "C:\Users\Julian\OneDrive\Dev Area\Dev Environment\System Files\"
    Else
        INIFPath = AppFPath & "\System Files\"
    End If
    
    INIFile = FreeFile()
    
    If Dir(IniFPath & INI_FILE) = "" Then Err.Raise NO_INI_FILE
    
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
    Line Input #INIFile, MaintMsg
    
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
    
    If MaintMsg <> "Online" Then
        MsgBox MaintMsg, vbExclamation, APP_NAME
        Application.DisplayAlerts = False
        ActiveWorkbook.Close
        Application.DisplayAlerts = True
        
    End If
    
    
GracefulExit:
    
    ReadINIFile = True
    Application.DisplayAlerts = True

Exit Function

ErrorExit:

'    ***CleanUpCode***
    ReadINIFile = False
    Application.DisplayAlerts = True

Exit Function

ErrorHandler:
    
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        CustomErrorHandler Err.Number
        Resume ErrorExit
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' MessageCheck
' Checks to see if the user message has been read
' ---------------------------------------------------------------
Public Function MessageCheck() As Boolean
    Dim StrMessage As String
    Dim RstMessage As Recordset
    
    Const StrPROCEDURE As String = "MessageCheck()"

    On Error GoTo ErrorHandler
    
    If CurrentUser Is Nothing Then Err.Raise SYSTEM_RESTART

    If CurrentUser.AccessLvl > BasicLvl_1 Then
        If Not CurrentUser.MessageRead Then
            
            Set RstMessage = SQLQuery("TblMessage")
            
            If RstMessage.RecordCount > 0 Then StrMessage = RstMessage.Fields(0)
            MsgBox StrMessage, vbOKOnly + vbInformation, APP_NAME
            CurrentUser.MessageRead = True
            CurrentUser.DBSave
            
        End If
    End If
    
    Set RstMessage = Nothing
    
    MessageCheck = True

Exit Function

ErrorExit:
    Set RstMessage = Nothing
'    ***CleanUpCode***
    MessageCheck = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' UpdateUsername
' Checks to see whether username needs to be changed and then updates
' ---------------------------------------------------------------
Private Function UpdateUsername() As Boolean
    Const StrPROCEDURE As String = "UpdateUsername()"

    On Error GoTo ErrorHandler

    If Application.Username = "PaulJ Wright" Then Application.Username = "Paul Wright"

    If Application.UserName = "#" Then Application.UserName = "Samuel Hayward"
    
    UpdateUsername = True

Exit Function

ErrorExit:

'    ***CleanUpCode***
    UpdateUsername = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function
