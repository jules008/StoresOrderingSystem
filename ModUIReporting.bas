Attribute VB_Name = "ModUIReporting"
'===============================================================
' Module ModUIReporting
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 02 Jun 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModUIReporting"

' ===============================================================
' BuildReporting
' Builds the display using shapes
' ---------------------------------------------------------------
Public Function BuildReporting() As Boolean
    
    Const StrPROCEDURE As String = "BuildReporting()"

    On Error GoTo ErrorHandler
    
    ModLibrary.PerfSettingsOn
    
    If Not BuildReport1Btn Then Err.Raise HANDLED_ERROR
        
    ModLibrary.PerfSettingsOff
                    
    BuildReporting = True
       
Exit Function

ErrorExit:
    
    ModLibrary.PerfSettingsOff

    BuildReporting = False
    
Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildReport1Btn
' Adds the button to switch order list between open and closed orders
' ---------------------------------------------------------------
Private Function BuildReport1Btn() As Boolean

    Const StrPROCEDURE As String = "BuildReport1Btn()"

    On Error GoTo ErrorHandler

    Set BtnReport1 = New ClsUIMenuItem

    With BtnReport1
        
        .Height = BTN_REPORT_1_HEIGHT
        .Left = BTN_REPORT_1_LEFT
        .Top = BTN_REPORT_1_TOP
        .Width = BTN_REPORT_1_WIDTH
        .Name = "BtnReport1"
        .OnAction = "'ModUIReporting.ProcessBtnPress(12)'"
        .UnSelectStyle = GENERIC_BUTTON
        .Selected = False
        .Text = "All Order Report"
    End With

    MainScreen.Menu.AddItem BtnReport1
    
    BuildReport1Btn = True

Exit Function

ErrorExit:

    BuildReport1Btn = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' ProcessBtnPress
' Receives all button presses and processes
' ---------------------------------------------------------------
Public Function ProcessBtnPress(ButtonNo As EnumBtnNo) As Boolean
    Const StrPROCEDURE As String = "ProcessBtnPress()"

    On Error GoTo ErrorHandler
    
        If MainScreen Is Nothing Then Err.Raise SYSTEM_RESTART
        
Restart:
        Application.StatusBar = ""
        
        Select Case ButtonNo
        
            Case EnumReport1Btn
            
                If Not BtnReport1Sel Then Err.Raise HANDLED_ERROR
                        
        End Select
    
GracefulExit:

    ProcessBtnPress = True

Exit Function

ErrorExit:


    ProcessBtnPress = False

Exit Function

ErrorHandler:
    
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        CustomErrorHandler Err.Number
         If Err.Number = SYSTEM_RESTART Then
            Resume Restart
        Else
            Resume GracefulExit
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
' BtnReport1Sel
' Manages system users
' ---------------------------------------------------------------
Private Function BtnReport1Sel() As Boolean

    Const StrPROCEDURE As String = "BtnReport1Sel()"

    On Error GoTo ErrorHandler

Restart:
    
    Application.StatusBar = ""

    If CurrentUser Is Nothing Then Err.Raise SYSTEM_RESTART
    
    If CurrentUser.AccessLvl < SupervisorLvl_3 Then Err.Raise ACCESS_DENIED

    If Not ModReports.Report1Query Then Err.Raise HANDLED_ERROR
    
GracefulExit:

    BtnReport1Sel = True

Exit Function

ErrorExit:
    
    BtnReport1Sel = False

'    ***CleanUpCode***

Exit Function

ErrorHandler:

    If Err.Number >= 1000 And Err.Number <= 1500 Then
        If Err.Number = ACCESS_DENIED Then
            CustomErrorHandler (Err.Number)
            Resume GracefulExit
        Else
            CustomErrorHandler (Err.Number)
            Resume Restart
        End If
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function


