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
' v0,7 - Moved ResetScreen to main menu
' v0,8 - Data Management Button
' v0,9 - removed hard numbering for buttons
' v0,10 - Added Order Age Column
' v0,11 - Added FindOrder Button
' v0,12 - Only refresh orders not whole page when return from order
' v0,131 - Change Delivery Button to Supplier
' v0,14 - Allow Stores level into Supplier Area
' v0,15 - Add missing PerformSettingsOff
'---------------------------------------------------------------
' Date - 13 Sep 17
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
        .OnAction = "'ModUIStoresScreen.ProcessBtnPress(" & EnumUserMngt & ")'"
        .UnSelectStyle = GENERIC_BUTTON
        .Selected = False
        .Text = "User Management"
        .Icon = ShtMain.Shapes("TEMPLATE - User").Duplicate
        .Icon.Left = .Left + 10
        .Icon.Top = .Top + 9
        .Icon.Name = "Supplier_Button"
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
' BuildManageDataBtn
' Adds the new order button to the main screen
' ---------------------------------------------------------------
Private Function BuildManageDataBtn() As Boolean

    Const StrPROCEDURE As String = "BuildManageDataBtn()"

    On Error GoTo ErrorHandler

    With BtnManageData
        
        .Height = BTN_MANAGE_DATA_HEIGHT
        .Left = BTN_MANAGE_DATA_LEFT
        .Top = BTN_MANAGE_DATA_TOP
        .Width = BTN_MANAGE_DATA_WIDTH
        .Name = "BtnManageData"
        .OnAction = "'ModUIStoresScreen.ProcessBtnPress(" & EnumManageDataBtn & ")'"
        .UnSelectStyle = GENERIC_BUTTON
        .Selected = False
        .Text = "Data Management"
        .Icon = ShtMain.Shapes("TEMPLATE - DataManage").Duplicate
        .Icon.Left = .Left + 10
        .Icon.Top = .Top + 9
        .Icon.Name = "Supplier_Button"
        .Icon.Visible = msoCTrue
    End With

    MainScreen.Menu.AddItem BtnManageData
    
    BuildManageDataBtn = True

Exit Function

ErrorExit:

    BuildManageDataBtn = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildFindOrderBtn
' Adds the new order button to the main screen
' ---------------------------------------------------------------
Private Function BuildFindOrderBtn() As Boolean

    Const StrPROCEDURE As String = "BuildFindOrderBtn()"

    On Error GoTo ErrorHandler

    With BtnFindOrder
        
        .Height = BTN_FIND_ORDER_HEIGHT
        .Left = BTN_FIND_ORDER_LEFT
        .Top = BTN_FIND_ORDER_TOP
        .Width = BTN_FIND_ORDER_WIDTH
        .Name = "BtnFindOrder"
        .OnAction = "'ModUIStoresScreen.ProcessBtnPress(" & EnumFindOrderBtn & ")'"
        .UnSelectStyle = GENERIC_BUTTON
        .Selected = False
        .Text = "Find Order"
        .Icon = ShtMain.Shapes("TEMPLATE - FindOrder").Duplicate
        .Icon.Left = .Left + 10
        .Icon.Top = .Top + 9
        .Icon.Name = "Supplier_Button"
        .Icon.Visible = msoCTrue
    End With

    MainScreen.Menu.AddItem BtnFindOrder
    
    BuildFindOrderBtn = True

Exit Function

ErrorExit:

    BuildFindOrderBtn = False

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
        .OnAction = "'ModUIStoresScreen.ProcessBtnPress(" & EnumOrderSwitch & ")'"
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
        .OnAction = "'ModUIStoresScreen.ProcessBtnPress(" & EnumRemoteOrder & ")'"
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
' BuildSupplierBtn
' Adds the new phone order buttomn to the screen
' ---------------------------------------------------------------
Private Function BuildSupplierBtn() As Boolean

    Const StrPROCEDURE As String = "BuildSupplierBtn()"

    On Error GoTo ErrorHandler

    With BtnSupplier
        
        .Height = BTN_SUPPLIER_HEIGHT
        .Left = BTN_SUPPLIER_LEFT
        .Top = BTN_SUPPLIER_TOP
        .Width = BTN_SUPPLIER_WIDTH
        .Name = "BtnSupplier"
        .OnAction = "'ModUIStoresScreen.ProcessBtnPress(" & EnumSupplierBtn & ")'"
        .UnSelectStyle = GENERIC_BUTTON
        .Selected = False
        .Text = "Suppliers"
        .Icon = ShtMain.Shapes("TEMPLATE - Delivery").Duplicate
        .Icon.Left = .Left + 10
        .Icon.Top = .Top + 9
        .Icon.Name = "Delivery_Button"
        .Icon.Visible = msoCTrue
    
    End With

    MainScreen.Menu.AddItem BtnSupplier
    
    BuildSupplierBtn = True

Exit Function

ErrorExit:

    BuildSupplierBtn = False

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
    Set BtnSupplier = New ClsUIMenuItem
    Set BtnManageData = New ClsUIMenuItem
    Set BtnFindOrder = New ClsUIMenuItem
    
    ModLibrary.PerfSettingsOn
    
    
    If Not BuildStoresFrame1 Then Err.Raise HANDLED_ERROR
    If Not BuildUserMangtBtn Then Err.Raise HANDLED_ERROR
    If Not BuildOrderSwitchBtn Then Err.Raise HANDLED_ERROR
    If Not BuildRemoteOrderBtn Then Err.Raise HANDLED_ERROR
    If Not BuildSupplierBtn Then Err.Raise HANDLED_ERROR
    If Not BuildManageDataBtn Then Err.Raise HANDLED_ERROR
    If Not BuildFindOrderBtn Then Err.Raise HANDLED_ERROR
    
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
    Set BtnSupplier = Nothing
    Set BtnManageData = Nothing
    
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
    Dim OrderNo As String
    
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
                
                MsgBox "Temporarily Disabled", vbInformation
'                If Not FrmPerson.ShowForm(True) Then Err.Raise HANDLED_ERROR
'
'                If Not RefreshOrderList(False) Then Err.Raise HANDLED_ERROR
            
            Case EnumSupplierBtn
                
                If Not BtnSupplierSel Then Err.Raise HANDLED_ERROR
                                
            Case EnumManageDataBtn
                
                If Not FrmDataManagmt.ShowForm Then Err.Raise HANDLED_ERROR
                
            Case EnumFindOrderBtn
                Dim Order As ClsOrder
                
                OrderNo = Application.InputBox("Please enter the Order No", "Order Search")
                
                If Not IsNumeric(OrderNo) Then Err.Raise NUMBERS_ONLY
                
                Set Order = New ClsOrder
                Order.DBGet CInt(OrderNo)
                
                If Order.OrderNo = 0 Then
                    MsgBox "No Order Found", vbExclamation, APP_NAME
                Else
                    If Not FrmDBOrder.ShowForm(Order) Then Err.Raise HANDLED_ERROR
                End If
                
                
        End Select
    
GracefulExit:

    ProcessBtnPress = True
    Set Order = Nothing
Exit Function

ErrorExit:


    ProcessBtnPress = False
    Set Order = Nothing

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
' BtnSupplierSel
' Manages system users
' ---------------------------------------------------------------
Private Function BtnSupplierSel() As Boolean
    Dim Supplier As ClsSupplier
    
    Const StrPROCEDURE As String = "BtnSupplierSel()"

    On Error GoTo ErrorHandler
    
    Set Supplier = New ClsSupplier
    
Restart:
    
    Application.StatusBar = ""

    If CurrentUser Is Nothing Then Err.Raise SYSTEM_RESTART
    
    If CurrentUser.AccessLvl < StoresLvl_2 Then Err.Raise ACCESS_DENIED
        
    If Not FrmSupplierSearch.ShowForm() Then Err.Raise HANDLED_ERROR
    
GracefulExit:
    Set Supplier = Nothing
    BtnSupplierSel = True

Exit Function

ErrorExit:
    
    Set Supplier = Nothing
    BtnSupplierSel = False

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
    Dim OrderAge As Integer
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
                If Not IsNull(RstOrder!Order_Age) Then OrderAge = RstOrder!Order_Age Else OrderAge = 0
                If Not IsNull(RstOrder!No_of_Items) Then NoOfItems = RstOrder!No_of_Items Else NoOfItems = ""
                If Not IsNull(RstOrder!ReqBy) Then ReqBy = RstOrder!ReqBy Else ReqBy = ""
                If Not IsNull(RstOrder!Station) Then Station = RstOrder!Station Else Station = ""
                If Not IsNull(RstOrder!Assigned_To) Then AssignedTo = RstOrder!Assigned_To Else AssignedTo = ""
                If Not IsNull(RstOrder!Status) Then OrderStatus = RstOrder!Status Else OrderStatus = ""
                
                StrOnAction = "'ModUIStoresScreen.OpenOrder(" & OrderNo & ")'"
                
                .Text i, 0, CStr(OrderNo), StrOnAction
                .Text i, 1, CStr(OrderAge), StrOnAction
                .Text i, 2, NoOfItems, StrOnAction
                .Text i, 3, ReqBy, StrOnAction
                .Text i, 4, Station, StrOnAction
                .Text i, 5, AssignedTo, StrOnAction
                .Text i, 6, OrderStatus, StrOnAction
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
    ModLibrary.PerfSettingsOff

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
    
    If StoresFrame1.Header.Text = "Open Orders" Then
        If Not RefreshOrderList(False) Then Err.Raise HANDLED_ERROR
    Else
        If Not RefreshOrderList(True) Then Err.Raise HANDLED_ERROR
    End If
    
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




