Attribute VB_Name = "ModUIReporting"
'===============================================================
' Module ModUIReporting
' v0,0 - Initial Version
' v0,1 - Updated query 1
' v0,2 - Addded Report 2
' v0,3 - Removed hard numbering from buttons
' v0,4 - Add cost to Order Report
' v0,5 - Added Report 3
' v0,61 - Added Report Settings Button
'---------------------------------------------------------------
' Date - 19 Dec 17
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
    If Not BuildReport2Btn Then Err.Raise HANDLED_ERROR
    If Not BuildReport3Btn Then Err.Raise HANDLED_ERROR
    If Not BuildSettingsBtn Then Err.Raise HANDLED_ERROR
    
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
        .OnAction = "'ModUIReporting.ProcessBtnPress(" & EnumReport1Btn & ")'"
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
' BuildReport2Btn
' Adds the button to switch order list between open and closed orders
' ---------------------------------------------------------------
Private Function BuildReport2Btn() As Boolean

    Const StrPROCEDURE As String = "BuildReport2Btn()"

    On Error GoTo ErrorHandler

    Set BtnReport2 = New ClsUIMenuItem

    With BtnReport2
        
        .Height = BTN_REPORT_2_HEIGHT
        .Left = BTN_REPORT_2_LEFT
        .Top = BTN_REPORT_2_TOP
        .Width = BTN_REPORT_2_WIDTH
        .Name = "BtnReport2"
        .OnAction = "'ModUIReporting.ProcessBtnPress(" & EnumReport2Btn & ")'"
        .UnSelectStyle = GENERIC_BUTTON
        .Selected = False
        .Text = "Stock Report"
    End With

    MainScreen.Menu.AddItem BtnReport2
    
    BuildReport2Btn = True

Exit Function

ErrorExit:

    BuildReport2Btn = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildReport3Btn
' Button for report 3
' ---------------------------------------------------------------
Private Function BuildReport3Btn() As Boolean

    Const StrPROCEDURE As String = "BuildReport3Btn()"

    On Error GoTo ErrorHandler

    Set BtnReport3 = New ClsUIMenuItem

    With BtnReport3
        
        .Height = BTN_REPORT_3_HEIGHT
        .Left = BTN_REPORT_3_LEFT
        .Top = BTN_REPORT_3_TOP
        .Width = BTN_REPORT_3_WIDTH
        .Name = "BtnReport3"
        .OnAction = "'ModUIReporting.ProcessBtnPress(" & EnumReport3Btn & ")'"
        .UnSelectStyle = GENERIC_BUTTON
        .Selected = False
        .Text = "Non-Return Report"
    End With

    MainScreen.Menu.AddItem BtnReport3
    
    BuildReport3Btn = True

Exit Function

ErrorExit:

    BuildReport3Btn = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildSettingsBtn
' Adds the button to switch order list between open and closed orders
' ---------------------------------------------------------------
Private Function BuildSettingsBtn() As Boolean

    Const StrPROCEDURE As String = "BuildSettingsBtn()"

    On Error GoTo ErrorHandler

    Set BtnRptSettings = New ClsUIMenuItem

    With BtnRptSettings
        
        .Height = BTN_RPT_SETTINGS_HEIGHT
        .Left = BTN_RPT_SETTINGS_LEFT
        .Top = BTN_RPT_SETTINGS_TOP
        .Width = BTN_RPT_SETTINGS_WIDTH
        .Name = "BtnSettings"
        .Icon = ShtMain.Shapes("TEMPLATE - Settings").Duplicate
        .Icon.Left = .Left + 10
        .Icon.Top = .Top + 9
        .Icon.Name = "AlertSettings_Button"
        .Icon.Visible = msoCTrue
        .OnAction = "'ModUIReporting.ProcessBtnPress(" & EnumRptSettings & ")'"
        .UnSelectStyle = GENERIC_BUTTON
        .Selected = False
        .Text = "Alert Settings"
    End With

    MainScreen.Menu.AddItem BtnRptSettings
    
    BuildSettingsBtn = True

Exit Function

ErrorExit:

    BuildSettingsBtn = False

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
                        
            Case EnumReport2Btn
            
                If Not BtnReport2Sel Then Err.Raise HANDLED_ERROR
            
            Case EnumReport3Btn
            
                If Not BtnReport3Sel Then Err.Raise HANDLED_ERROR
                
            Case EnumRptSettings
            
                If Not FrmReportAdmin.ShowForm Then Err.Raise HANDLED_ERROR
          
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
    Dim ColWidths(0 To 15) As Integer
    Dim Headings(0 To 15) As String
    Dim ColFormats(0 To 15) As String

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
    ColWidths(15) = 20
    
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
    Headings(15) = "Total Cost"
    'formats
    ColFormats(0) = "General"
    ColFormats(1) = "General"
    ColFormats(2) = "General"
    ColFormats(3) = "General"
    ColFormats(4) = "General"
    ColFormats(5) = "General"
    ColFormats(6) = "General"
    ColFormats(7) = "General"
    ColFormats(8) = "General"
    ColFormats(9) = "General"
    ColFormats(10) = "General"
    ColFormats(11) = "General"
    ColFormats(12) = "General"
    ColFormats(13) = "General"
    ColFormats(14) = "General"
    ColFormats(15) = "£0.00"
    
    If Not ModReports.CreateReport(RstQuery, ColWidths, Headings, ColFormats) Then Err.Raise HANDLED_ERROR
    
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

' ===============================================================
' BtnReport2Sel
' Manages system users
' ---------------------------------------------------------------
Private Function BtnReport2Sel() As Boolean
    Dim RstQuery As Recordset
    Dim ColWidths(0 To 9) As Integer
    Dim Headings(0 To 9) As String
    Dim ColFormats(0 To 9) As String

    Const StrPROCEDURE As String = "BtnReport2Sel()"

    On Error GoTo ErrorHandler

Restart:
    
    Application.StatusBar = ""

    If CurrentUser Is Nothing Then Err.Raise SYSTEM_RESTART
    
    If CurrentUser.AccessLvl < StoresLvl_2 Then Err.Raise ACCESS_DENIED

    Set RstQuery = ModReports.Report2Query
    
    If RstQuery Is Nothing Then Err.Raise HANDLED_ERROR
    
    'col widths
    ColWidths(0) = 8
    ColWidths(1) = 60
    ColWidths(2) = 10
    ColWidths(3) = 20
    ColWidths(4) = 20
    ColWidths(5) = 20
    ColWidths(6) = 20
    ColWidths(7) = 20
    ColWidths(8) = 20
    ColWidths(9) = 20
    
    'headings
    Headings(0) = "Asset No"
    Headings(1) = "Description"
    Headings(2) = "Quantity"
    Headings(3) = "Category 1"
    Headings(4) = "Category 2"
    Headings(5) = "Category 3"
    Headings(6) = "Size 1"
    Headings(7) = "Size 2"
    Headings(8) = "Item Cost"
    Headings(9) = "Cost of Stock"
    
    'formats
    ColFormats(0) = "General"
    ColFormats(1) = "General"
    ColFormats(2) = "General"
    ColFormats(3) = "General"
    ColFormats(4) = "General"
    ColFormats(5) = "General"
    ColFormats(6) = "General"
    ColFormats(7) = "General"
    ColFormats(8) = "£#,###.00"
    ColFormats(9) = "£#,###.00"
    
    If Not ModReports.CreateReport(RstQuery, ColWidths, Headings, ColFormats) Then Err.Raise HANDLED_ERROR
    
GracefulExit:

    BtnReport2Sel = True

Exit Function

ErrorExit:
    
    BtnReport2Sel = False

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

' ===============================================================
' BtnReport3Sel
' Manages system users
' ---------------------------------------------------------------
Private Function BtnReport3Sel() As Boolean
    Dim RstQuery As Recordset
    Dim ColWidths(0 To 7) As Integer
    Dim Headings(0 To 7) As String
    Dim ColFormats(0 To 7) As String

    Const StrPROCEDURE As String = "BtnReport3Sel()"

    On Error GoTo ErrorHandler

Restart:
    
    Application.StatusBar = ""

    If CurrentUser Is Nothing Then Err.Raise SYSTEM_RESTART
    
    If CurrentUser.AccessLvl < StoresLvl_2 Then Err.Raise ACCESS_DENIED

    Set RstQuery = ModReports.Report3Query
    
    If RstQuery Is Nothing Then Err.Raise HANDLED_ERROR
    
    'col widths
    ColWidths(0) = 8
    ColWidths(1) = 15
    ColWidths(2) = 60
    ColWidths(3) = 15
    ColWidths(4) = 20
    ColWidths(5) = 20
    ColWidths(6) = 20
    ColWidths(7) = 20
    
    'headings
    Headings(0) = "Order No"
    Headings(1) = "Order Date"
    Headings(2) = "Description"
    Headings(3) = "Quantity"
    Headings(4) = "Total Cost"
    Headings(5) = "Station No"
    Headings(6) = "Station Name"
    Headings(7) = "Division"
    
    'formats
    ColFormats(0) = "General"
    ColFormats(1) = "General"
    ColFormats(2) = "General"
    ColFormats(3) = "General"
    ColFormats(4) = "£0.00"
    ColFormats(5) = "General"
    ColFormats(6) = "General"
    ColFormats(7) = "General"
    
    If Not ModReports.CreateReport(RstQuery, ColWidths, Headings, ColFormats) Then Err.Raise HANDLED_ERROR
    
GracefulExit:

    BtnReport3Sel = True

Exit Function

ErrorExit:
    
    BtnReport3Sel = False

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


