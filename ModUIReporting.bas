Attribute VB_Name = "ModUIReporting"
'===============================================================
' Module ModUIReporting
' v0,0 - Initial Version
' v0,1 - Updated query 1
'---------------------------------------------------------------
' Date - 19 Jun 17
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
    Dim RstQuery As Recordset
    Dim ColWidths(0 To 14) As Integer
    Dim Headings(0 To 14) As String

    Const StrPROCEDURE As String = "BtnReport1Sel()"

    On Error GoTo ErrorHandler

Restart:
    
    Application.StatusBar = ""

    If CurrentUser Is Nothing Then Err.Raise SYSTEM_RESTART
    
    If CurrentUser.AccessLvl < StoresLvl_2 Then Err.Raise ACCESS_DENIED

    Set RstQuery = ModReports.Report1Query
    
    If RstQuery Is Nothing Then Err.Raise HANDLED_ERROR
    
    'col widths
    ColWidths(0) = 8
    ColWidths(1) = 15
    ColWidths(2) = 25
    ColWidths(3) = 60
    ColWidths(4) = 20
    ColWidths(5) = 20
    ColWidths(6) = 20
    ColWidths(7) = 20
    ColWidths(8) = 20
    ColWidths(9) = 10
    ColWidths(10) = 25
    ColWidths(11) = 25
    ColWidths(12) = 25
    ColWidths(13) = 25
    ColWidths(14) = 25
    
    'headings
    Headings(0) = "Order No"
    Headings(1) = "Order Date"
    Headings(2) = "Ordered By"
    Headings(3) = "Description"
    Headings(4) = "Category 1"
    Headings(5) = "Category 2"
    Headings(6) = "Category 3"
    Headings(7) = "Size 1"
    Headings(8) = "Size 2"
    Headings(9) = "Quantity"
    Headings(10) = "For Person"
    Headings(11) = "For Station"
    Headings(12) = "For Vehicle"
    Headings(13) = "Veh Station"
    Headings(14) = "Request Reason"
    
    
    If Not ModReports.CreateReport(RstQuery, ColWidths, Headings) Then Err.Raise HANDLED_ERROR
    
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


