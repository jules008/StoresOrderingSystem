Attribute VB_Name = "ModErrorHandling"
'===============================================================
' Module ModErrorHandling
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 17 Jan 17
'===============================================================

Option Explicit

Public FaultCount1002 As Integer
Public FaultCount1008 As Integer
Private Const StrMODULE As String = "ModErrorHandling"

' ===============================================================
' CentralErrorHandler
' Handles all system errors
' ---------------------------------------------------------------
Public Function CentralErrorHandler( _
            ByVal sModule As String, _
            ByVal sProc As String, _
            Optional ByVal sFile As String, _
            Optional ByVal bEntryPoint As Boolean) As Boolean

    Static sErrMsg As String
    
    Dim iFile As Integer
    Dim lErrNum As Long
    Dim sFullSource As String
    Dim sPath As String
    Dim sLogText As String
    Dim ErrMsgTxt As String
    
    ' Grab the error info before it's cleared by
    ' On Error Resume Next below.
    lErrNum = Err.Number
    
    
    If Len(sErrMsg) = 0 Then sErrMsg = Err.Description
                

    ' We cannot allow errors in the central error handler.
    On Error Resume Next
    
    ' Load the default filename if required.
    If Len(sFile) = 0 Then sFile = ThisWorkbook.Name
    
    ' Get the application directory.
    sPath = ThisWorkbook.Path
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    
    ' Construct the fully-qualified error source name.
    sFullSource = "[" & sFile & "]" & sModule & "." & sProc

    ' Create the error text to be logged.
    ErrMsgTxt = "Sorry, there has been an error.  An Error Log File has been created.  Would " _
                & " like to email this for further investigation?"
        
    sLogText = "  " & sFullSource & ", Error " & _
                        CStr(lErrNum) & ": " & sErrMsg
    
    ' Open the log file, write out the error information and
    ' close the log file.
    If OUTPUT_MODE = "Log" Then
        Dim Response As Integer
        
        iFile = FreeFile()
        Open sPath & FILE_ERROR_LOG For Append As #iFile
        Print #iFile, Format$(Now(), "mm/dd/yy hh:mm:ss"); sLogText
        If bEntryPoint Then Print #iFile,
        Close #iFile
    End If
                
    Debug.Print Format$(Now(), "mm/dd/yy hh:mm:ss"); sLogText
    If bEntryPoint Then Debug.Print
    
    ' Do not display or debug silent errors.
'    If sErrMsg <> SILENT_ERROR Then

    ' Show the error message when we reach the entry point
    ' procedure or immediately if we are in debug mode.
    If bEntryPoint Or DEBUG_MODE Then
        ModLibrary.PerfSettingsOff
        
        If MailSystem Is Nothing Then Set MailSystem = New ClsMailSystem
    
        
        If Not DEV_MODE Then
            Response = MsgBox(ErrMsgTxt, vbYesNo, APP_NAME)
        
            If Response = 6 Then
                With MailSystem
                    .MailItem.To = "Julian Turner"
                    .MailItem.Subject = "Debug Report - " & APP_NAME
                    .MailItem.Importance = olImportanceHigh
                    .MailItem.Attachments.Add sPath & FILE_ERROR_LOG
                    .MailItem.Body = "Please add any further information such " _
                                       & "what you were doing at the time of the error" _
                                       & ", and what candidate were you working on etc "
                    .DisplayEmail
                End With
            End If
            Set MailSystem = Nothing
        End If
        ' Clear the static error message variable once
        ' we've reached the entry point so that we're ready
        ' to handle the next error.
        sErrMsg = vbNullString
    End If
    
    ' The return vale is the debug mode status.
    CentralErrorHandler = DEBUG_MODE
    
'    Else
'        ' If this is a silent error, clear the static error
'        ' message variable when we reach the entry point.
'        If bEntryPoint Then sErrMsg = vbNullString
'        CentralErrorHandler = False
'    End If
    
End Function

' ===============================================================
' CustomErrorHandler
' Handles system custom errors 1000 - 1500
' ---------------------------------------------------------------
Public Function CustomErrorHandler(ErrorCode As Long, Optional Message As String) As Boolean
    
    Const StrPROCEDURE As String = "CustomErrorHandler()"

    On Error GoTo ErrorHandler

    Select Case ErrorCode
        Case UNKNOWN_USER
            
            MsgBox "Sorry, the system does not recognise you.  Please continue with " _
                    & "the order as a guest.  Your name has been forwarded onto the " _
                    & "Administrator so that you can be added to the system"
                    
            
            CurrentUser.AddTempAccount
            
            Set MailSystem = New ClsMailSystem
            
            With MailSystem
                .MailItem.To = "Julian Turner"
                .MailItem.Subject = "Unknown User - " & APP_NAME
                .MailItem.Importance = olImportanceHigh
                .MailItem.Body = "A new user needs to be added to the database - " _
                                & CurrentUser.CrewNo & " " & CurrentUser.UserName
                
                If SEND_EMAILS Then .MailItem.Send
            End With

        Case NO_ITEM_SELECTED
            MsgBox "Please select an item"
            
        Case NO_DATABASE_FOUND
            FaultCount1008 = FaultCount1008 + 1
            Debug.Print "Trying to connect to Database....Attempt " & FaultCount1008
            
            If ModErrorHandling.FaultCount1008 <= 3 Then
            
                Application.DisplayStatusBar = True
                Application.StatusBar = "Trying to connect to Database....Attempt " & FaultCount1008
                Application.Wait (Now + TimeValue("0:00:02"))
                Debug.Print FaultCount1008
            Else
                FaultCount1008 = 0
                Application.StatusBar = "No Database"
                Err.Raise SYSTEM_FAILURE, Description:="Unable to connect to database afer 3 attempts"
                CustomErrorHandler = False
            End If
        
        Case SYSTEM_RESTART
            Debug.Print "system failed - restarting"
            FaultCount1002 = FaultCount1002 + 1

            If ModErrorHandling.FaultCount1002 <= 3 Then
                If Not Initialise Then Err.Raise HANDLED_ERROR
                Application.DisplayStatusBar = True
                Application.StatusBar = "System failed...Restarting Attempt " & FaultCount1002
                Application.Wait (Now + TimeValue("0:00:02"))
            Else
                FaultCount1002 = 0
                Application.StatusBar = "Sysetm Failed"
                Err.Raise SYSTEM_FAILURE, Description:="System restart failed 3 time"
            End If
            
        Case NO_QUANTITY_ENTERED
            MsgBox "Please enter a quantity"
        
        Case NO_SIZE_ENTERED
            MsgBox "Please enter a size"
        
        Case NO_CREW_NO_ENTERED
            MsgBox "Please enter a Brigade No"
            
        Case NUMBERS_ONLY
            MsgBox "Please enter number only"
            
        Case CREWNO_UNRECOGNISED
            MsgBox "The Brigade No is not recognised on the system, please re-enter"
        
        Case NO_VEHICLE_SELECTED
            MsgBox "Please select a vehicle"
        
        Case NO_STATION_SELECTED
            MsgBox "Please select a station"
            
        Case FIELDS_INCOMPLETE
            MsgBox "Please complete all fields"
            
        Case NO_NAMES_SELECTED
            MsgBox "Please select a name"
            
        Case FORM_INPUT_EMPTY
            MsgBox "Please complete all highlighted fields"
            
        Case ACCESS_DENIED
            MsgBox "Sorry you do not have the required Access Level.  " _
                & "Please send a Support Mail if you require access", vbCritical
        Case NO_ORDER_MESSAGE
            MsgBox Message
    End Select
    
    Set MailSystem = Nothing

    CustomErrorHandler = True

Exit Function

ErrorExit:

    Set MailSystem = Nothing
    
    CustomErrorHandler = False

    Exit Function

ErrorHandler:

If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function
