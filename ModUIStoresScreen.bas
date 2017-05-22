Attribute VB_Name = "ModUIStoresScreen"
'===============================================================
' Module ModUIStoresScreen
' v0,0 - Initial Version
' v0,1 - Added build Order Switch Button
' v0,2 - Added Remote Order Button
' v0,3 - Changes for disable line item functionality
' v0,4 - Increased Order retrieval performance
' v0,5 - Now passing OnAction as paramater
' v0,6 - Delivery Button and add icons
'---------------------------------------------------------------
' Date - 22 May 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModUIStoresScreen"

' ===============================================================
' BuildUserMangtBtn
' Adds the new order button to the main screen
' ---------------------------------------------------------------
Private Function BuildUserMangtBtn() As Boolean

    Const StrPROCEDURE As String = "BuildUserMangtBtn()"

    On Error GoTo ErrorHandler

    With BtnUserMangt
        
        .Height = BTN_USER_MANGT_HEIGHT
        .Left = BTN_USER_MANGT_LEFT
        .Top = BTN_USER_MANGT_TOP
        .Width = BTN_USER_MANGT_WIDTH
        .Name = "BtnUserMangt"
        .OnAction = "'ModUIStoresScreen.ProcessBtnPress(8)'"
        .UnSelectStyle = GENERIC_BUTTON
        .Selected = False
        .Text = "User Management"
        .Icon = ShtMain.Shapes("TEMPLATE - User").Duplicate
        .Icon.Left = .Left + 10
        .Icon.Top = .Top + 9
        .Icon.Name = "Delivery_Button"
        .Icon.Visible = msoCTrue
    End With

    MainScreen.Menu.AddItem BtnUserMangt
    
    BuildUserMangtBtn = True

Exit Function

ErrorExit:

    BuildUserMangtBtn = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildOrderSwitchBtn
' Adds the button to switch order list between open and closed orders
' ---------------------------------------------------------------
Private Function BuildOrderSwitchBtn() As Boolean

    Const StrPROCEDURE As String = "BuildOrderSwitchBtn()"

    On Error GoTo ErrorHandler

    With BtnOrderSwitch
        
        .Height = BTN_ORDER_SWITCH_HEIGHT
        .Left = BTN_ORDER_SWITCH_LEFT
        .Top = BTN_ORDER_SWITCH_TOP
        .Width = BTN_ORDER_SWITCH_WIDTH
        .Name = "BtnOrderSwitch"
        .OnAction = "'ModUIStoresScreen.ProcessBtnPress(9)'"
        .UnSelectStyle = GENERIC_BUTTON
        .Selected = False
        .Text = "Show Closed Orders"
        .Icon = ShtMain.Shapes("TEMPLATE - Closed Orders").Duplicate
        .Icon.Left = .Left + 10
        .Icon.Top = .Top + 9
        .Icon.Name = "Delivery_Button"
        .Icon.Visible = msoCTrue
    End With

    MainScreen.Menu.AddItem BtnOrderSwitch
    
    BuildOrderSwitchBtn = True

Exit Function

ErrorExit:

    BuildOrderSwitchBtn = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildRemoteOrderBtn
' Adds the new phone order buttomn to the screen
' ---------------------------------------------------------------
Private Function BuildRemoteOrderBtn() As Boolean

    Const StrPROCEDURE As String = "BuildRemoteOrderBtn()"

    On Error GoTo ErrorHandler

    With BtnRemoteOrder
        
        .Height = BTN_REMOTE_ORDER_HEIGHT
        .Left = BTN_REMOTE_ORDER_LEFT
        .Top = BTN_REMOTE_ORDER_TOP
        .Width = BTN_REMOTE_ORDER_WIDTH
        .Name = "BtnRemoteOrder"
        .OnAction = "'ModUIStoresScreen.ProcessBtnPress(10)'"
        .UnSelectStyle = GENERIC_BUTTON
        .Selected = False
        .Text = "New Phone Order"
        .Icon = ShtMain.Shapes("TEMPLATE - Phone").Duplicate
        .Icon.Left = .Left + 10
        .Icon.Top = .Top + 9
        .Icon.Name = "Delivery_Button"
        .Icon.Visible = msoCTrue
    End With

    MainScreen.Menu.AddItem BtnRemoteOrder
    
    BuildRemoteOrderBtn = True

Exit Function

ErrorExit:

    BuildRemoteOrderBtn = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildDeliveryBtn
' Adds the new phone order buttomn to the screen
' ---------------------------------------------------------------
Private Function BuildDeliveryBtn() As Boolean

    Const StrPROCEDURE As String = "BuildDeliveryBtn()"

    On Error GoTo ErrorHandler

    With BtnDelivery
        
        .Height = BTN_DELIVERY_HEIGHT
        .Left = BTN_DELIVERY_LEFT
        .Top = BTN_DELIVERY_TOP
        .Width = BTN_DELIVERY_WIDTH
        .Name = "BtnDelivery"
        .OnAction = "'ModUIStoresScreen.ProcessBtnPress(11)'"
        .UnSelectStyle = GENERIC_BUTTON
        .Selected = False
        .Text = "Add Delivery"
        .Icon = ShtMain.Shapes("TEMPLATE - Delivery").Duplicate
        .Icon.Left = .Left + 10
        .Icon.Top = .Top + 9
        .Icon.Name = "Delivery_Button"
        .Icon.Visible = msoCTrue
    
    End With

    MainScreen.Menu.AddItem BtnDelivery
    
    BuildDeliveryBtn = True

Exit Function

ErrorExit:

    BuildDeliveryBtn = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildStoresFrame1
' Builds first frame on stores page at top of screen
' ---------------------------------------------------------------
Public Function BuildStoresFrame1() As Boolean
    Dim CommentBox As ClsUIDashObj
    Dim CommentBtn As ClsUIMenuItem
    Dim UILineItem As ClsUILineitem
    Dim i As Integer
    
    Const StrPROCEDURE As String = "BuildStoresFrame1()"

    On Error GoTo ErrorHandler

    
    With StoresFrame1
        .Name = "Stores Frame 1"
        MainScreen.Frames.AddItem StoresFrame1
        
        .Top = STORES_FRAME_1_TOP
        .Left = STORES_FRAME_1_LEFT
        .Width = STORES_FRAME_1_WIDTH
        .Height = STORES_FRAME_1_HEIGHT
        .Style = MAIN_FRAME_STYLE
        .EnableHeader = True
        .Visible = True
                

        With .Header
            .Top = .Parent.Top
            .Left = .Parent.Left
            .Width = .Parent.Width
            .Height = HEADER_HEIGHT
            .Name = "Stores 1 Header"
            .Text = "Open Orders"
            .Style = HEADER_STYLE
            .Icon = ShtMain.Shapes("TEMPLATE - Orders").Duplicate
            .Icon.Top = .Parent.Top + HEADER_ICON_TOP
            .Icon.Left = .Parent.Left + .Parent.Width - .Icon.Width - HEADER_ICON_RIGHT
            .Icon.Name = .Parent.Name & " Icon"
            .Icon.Visible = msoCTrue
        End With
    End With
    
    With StoresFrame1.LineItems
        .NoColumns = ORDER_LINEITEM_NOCOLS
        .Top = ORDER_LINEITEM_TOP
        .Left = ORDER_LINEITEM_LEFT
        .Height = ORDER_LINEITEM_HEIGHT
        .Columns = ORDER_LINEITEM_COL_WIDTHS
        .RowOffset = ORDER_LINEITEM_ROWOFFSET
            
    End With
    
    StoresFrame1.ReOrder
    
    Set UILineItem = Nothing
    
    BuildStoresFrame1 = True

Exit Function

ErrorExit:
    Set UILineItem = Nothing

    
    BuildStoresFrame1 = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildStoresScreen
' Builds the display using shapes
' ---------------------------------------------------------------
Public Function BuildStoresScreen() As Boolean
    
    Const StrPROCEDURE As String = "BuildStoresScreen()"

    On Error GoTo ErrorHandler
    
    Set StoresFrame1 = New ClsUIFrame
    Set BtnUserMangt = New ClsUIMenuItem
    Set BtnOrderSwitch = New ClsUIMenuItem
    Set BtnRemoteOrder = New ClsUIMenuItem
    Set BtnDelivery = New ClsUIMenuItem
    
    ModLibrary.PerfSettingsOn
    
    If Not ResetScreen Then Err.Raise HANDLED_ERROR
    If Not BuildStoresFrame1 Then Err.Raise HANDLED_ERROR
    If Not BuildUserMangtBtn Then Err.Raise HANDLED_ERROR
    If Not BuildOrderSwitchBtn Then Err.Raise HANDLED_ERROR
    If Not BuildRemoteOrderBtn Then Err.Raise HANDLED_ERROR
    If Not BuildDeliveryBtn Then Err.Raise HANDLED_ERROR
    If Not RefreshOrderList(False) Then Err.Raise HANDLED_ERROR
    
    ModLibrary.PerfSettingsOff
    
    BuildStoresScreen = True
       
Exit Function

ErrorExit:
    ModLibrary.PerfSettingsOff

    Set StoresFrame1 = Nothing
    Set BtnUserMangt = Nothing
    Set BtnOrderSwitch = Nothing
    Set BtnRemoteOrder = Nothing
    Set BtnDelivery = Nothing
    Terminate
    
    BuildStoresScreen = False
    
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
        
            Case EnumUserMngt
            
                If Not BtnUserManagementSel Then Err.Raise HANDLED_ERROR
                        
            Case EnumOrderSwitch
            
                If Not BtnOrderSwitchSel Then Err.Raise HANDLED_ERROR
        
            Case EnumRemoteOrder
                
                If Not FrmPerson.ShowForm(True) Then Err.Raise HANDLED_ERROR
                
                If Not RefreshOrderList(False) Then Err.Raise HANDLED_ERROR
            
            Case EnumDeliveryBtn
                
                If Not FrmDelivery.ShowForm Then Err.Raise HANDLED_ERROR
                
                If Not RefreshOrderList(False) Then Err.Raise HANDLED_ERROR
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
' BtnUserManagementSel
' Manages system users
' ---------------------------------------------------------------
Private Function BtnUserManagementSel() As Boolean

    Const StrPROCEDURE As String = "BtnUserManagementSel()"

    On Error GoTo ErrorHandler

Restart:
    
    Application.StatusBar = ""

    If CurrentUser Is Nothing Then Err.Raise SYSTEM_RESTART
    
    If CurrentUser.AccessLvl < SupervisorLvl_3 Then Err.Raise ACCESS_DENIED

    If Not FrmUserAdmin.ShowForm Then Err.Raise HANDLED_ERROR
    
    
GracefulExit:

    BtnUserManagementSel = True

Exit Function

ErrorExit:
    
    BtnUserManagementSel = False

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
' RefreshOrderList
' Refreshes the list of open orders from the database
' ---------------------------------------------------------------
Public Function RefreshOrderList(ClosedOrders As Boolean) As Boolean
    Dim OrderNo As Integer
    Dim NoOfItems As String
    Dim ReqBy As String
    Dim Station As String
    Dim AssignedTo As String
    Dim OrderStatus As String
    Dim Orders As ClsOrders
    Dim RstOrder As Recordset
    Dim StrOnAction As String
    Dim Lineitem As ClsUILineitem
    Dim i As Integer
    Dim OnAction As String
    Dim RowTitles() As String

    Const StrPROCEDURE As String = "RefreshOrderList()"
    
    On Error GoTo ErrorHandler
    
    Set Orders = New ClsOrders

    ShtMain.Unprotect
    
    ModLibrary.PerfSettingsOn
    
    With StoresFrame1
        For Each Lineitem In .LineItems
            .LineItems.RemoveItem Lineitem.Name
            Lineitem.ShpLineItem.Delete
            Set Lineitem = Nothing
        Next
        
        ReDim RowTitles(0 To ORDER_LINEITEM_NOCOLS - 1)
        RowTitles = Split(ORDER_LINEITEM_TITLES, ":")

        .LineItems.Style = GENERIC_LINEITEM_HEADER
        
        For i = 0 To ORDER_LINEITEM_NOCOLS - 1
            .LineItems.Text 0, i, RowTitles(i), False
        Next
        
        .LineItems.Style = GENERIC_LINEITEM

    End With
    
    If ClosedOrders Then
        Set RstOrder = Orders.GetClosedOrders
        StoresFrame1.Header.Text = "Closed Orders"
        BtnOrderSwitch.Text = "Show Open Orders"
    Else
        Set RstOrder = Orders.GetOpenOrders
        StoresFrame1.Header.Text = "Open Orders"
        BtnOrderSwitch.Text = "Show Closed Orders"
    End If
    
    StoresFrame1.Height = RstOrder.RecordCount * ORDER_LINEITEM_ROWOFFSET + (ORDER_LINEITEM_TOP * 2)
    
    i = 1
    With RstOrder
        Do While Not .EOF
            With StoresFrame1.LineItems
                If Not IsNull(RstOrder!Order_No) Then OrderNo = RstOrder!Order_No Else OrderNo = 0
                If Not IsNull(RstOrder!No_of_Items) Then NoOfItems = RstOrder!No_of_Items Else NoOfItems = ""
                If Not IsNull(RstOrder!ReqBy) Then ReqBy = RstOrder!ReqBy Else ReqBy = ""
                If Not IsNull(RstOrder!Station) Then Station = RstOrder!Station Else Station = ""
                If Not IsNull(RstOrder!Assigned_To) Then AssignedTo = RstOrder!Assigned_To Else AssignedTo = ""
                If Not IsNull(RstOrder!Status) Then OrderStatus = RstOrder!Status Else OrderStatus = ""
                
                StrOnAction = "'ModUIStoresScreen.OpenOrder(" & OrderNo & ")'"
                
                .Text i, 0, CStr(OrderNo), StrOnAction
                .Text i, 1, NoOfItems, StrOnAction
                .Text i, 2, ReqBy, StrOnAction
                .Text i, 3, Station, StrOnAction
                .Text i, 4, AssignedTo, StrOnAction
                .Text i, 5, OrderStatus, StrOnAction
            End With
            .MoveNext
            i = i + 1
        Loop
        
    End With
    
    ModLibrary.PerfSettingsOff
                
    ShtMain.Protect
    
    RefreshOrderList = True
    
    Set Orders = Nothing
    
Exit Function

ErrorExit:
    Set Orders = Nothing
    
    Terminate
    RefreshOrderList = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' OpenOrder
' Opens the selected order form
' ---------------------------------------------------------------
Private Sub OpenOrder(OrderNo As Integer)
    Const StrPROCEDURE As String = "OpenOrder()"
    
    Dim Order As ClsOrder
    
    On Error GoTo ErrorHandler

    Set Order = New ClsOrder
    
    Order.DBGet OrderNo
    
    If Not FrmDBOrder.ShowForm(Order) Then Err.Raise HANDLED_ERROR
    
    ModLibrary.PerfSettingsOn
    
    If Not ResetScreen Then Err.Raise HANDLED_ERROR
    If Not BuildStoresScreen Then Err.Raise HANDLED_ERROR
    
    ModLibrary.PerfSettingsOff
    
    Set Order = Nothing

Exit Sub

ErrorExit:

    ModLibrary.PerfSettingsOff
    Set Order = Nothing
    Terminate
Exit Sub

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' BtnOrderSwitchSel
' Switches between open orders and closed orders
' ---------------------------------------------------------------
Private Function BtnOrderSwitchSel() As Boolean

    Const StrPROCEDURE As String = "BtnOrderSwitchSel()"

    On Error GoTo ErrorHandler

Restart:
    
    Application.StatusBar = ""

    If CurrentUser Is Nothing Then Err.Raise SYSTEM_RESTART
    
    If StoresFrame1.Header.Text = "Open Orders" Then
        If Not RefreshOrderList(True) Then Err.Raise HANDLED_ERROR
   Else
        If Not RefreshOrderList(False) Then Err.Raise HANDLED_ERROR
    
    End If
    
GracefulExit:

    BtnOrderSwitchSel = True

Exit Function

ErrorExit:
    
    BtnOrderSwitchSel = False

'    ***CleanUpCode***

Exit Function

ErrorHandler:

'    If Err.Number >= 1000 And Err.Number <= 1500 Then
'        If Err.Number = ACCESS_DENIED Then
'            CustomErrorHandler (Err.Number)
'            Resume gracefulexit
'        Else
'            CustomErrorHandler (Err.Number)
'            Resume Restart
'        End If
'    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function




