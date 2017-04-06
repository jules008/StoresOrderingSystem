Attribute VB_Name = "ModUIMainScreen"
'===============================================================
' Module ModUIMainScreen
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 09 Feb 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModUIMainScreen"

' ===============================================================
' BuildStylesMainScreen
' Builds the UI styles for use on the screen
' ---------------------------------------------------------------
Public Function BuildStylesMainScreen() As Boolean
    Const StrPROCEDURE As String = "BuildStylesMainScreen()"

    On Error GoTo ErrorHandler

    With SCREEN_STYLE
        .BorderWidth = SCREEN_BORDER_WIDTH
        .Fill1 = SCREEN_FILL_1
        .Fill2 = SCREEN_FILL_2
        .Shadow = SCREEN_SHADOW
    End With
    
    With MENUBAR_STYLE
        .BorderWidth = MENUBAR_BORDER_WIDTH
        .Fill1 = MENUBAR_FILL_1
        .Fill2 = MENUBAR_FILL_2
        .Shadow = MENUBAR_SHADOW
    End With
    
    With MENUITEM_UNSET_STYLE
        .BorderWidth = MENUITEM_UNSET_BORDER_WIDTH
        .Fill1 = MENUITEM_UNSET_FILL_1
        .Fill2 = MENUITEM_UNSET_FILL_2
        .Shadow = MENUITEM_UNSET_SHADOW
        .FontStyle = MENUITEM_UNSET_FONT_STYLE
        .FontSize = MENUITEM_UNSET_FONT_SIZE
        .FontColour = MENUITEM_UNSET_FONT_COLOUR
        .FontXJust = MENUITEM_UNSET_FONT_X_JUST
        .FontYJust = MENUITEM_UNSET_FONT_Y_JUST
    End With

    With MENUITEM_SET_STYLE
        .BorderWidth = MENUITEM_SET_BORDER_WIDTH
        .Fill1 = MENUITEM_SET_FILL_1
        .Fill2 = MENUITEM_SET_FILL_2
        .Shadow = MENUITEM_SET_SHADOW
        .FontStyle = MENUITEM_SET_FONT_STYLE
        .FontSize = MENUITEM_SET_FONT_SIZE
        .FontColour = MENUITEM_SET_FONT_COLOUR
        .FontXJust = MENUITEM_SET_FONT_X_JUST
        .FontYJust = MENUITEM_SET_FONT_Y_JUST
    End With
    
    With MAIN_FRAME_STYLE
        .BorderWidth = MAIN_FRAME_BORDER_WIDTH
        .Fill1 = MAIN_FRAME_FILL_1
        .Fill2 = MAIN_FRAME_FILL_2
        .Shadow = MAIN_FRAME_SHADOW
    End With
    
    With HEADER_STYLE
        .BorderWidth = HEADER_BORDER_WIDTH
        .Fill1 = HEADER_FILL_1
        .Fill2 = HEADER_FILL_2
        .Shadow = HEADER_SHADOW
        .FontStyle = HEADER_FONT_STYLE
        .FontSize = HEADER_FONT_SIZE
        .FontBold = HEADER_FONT_BOLD
        .FontColour = HEADER_FONT_COLOUR
        .FontXJust = HEADER_FONT_X_JUST
        .FontYJust = HEADER_FONT_Y_JUST
    End With
    
    With BTN_NEWORDER_STYLE
        .BorderWidth = BTN_NEWORDER_BORDER_WIDTH
        .Fill1 = BTN_NEWORDER_FILL_1
        .Fill2 = BTN_NEWORDER_FILL_2
        .Shadow = BTN_NEWORDER_SHADOW
        .FontStyle = BTN_NEWORDER_FONT_STYLE
        .FontSize = BTN_NEWORDER_FONT_SIZE
        .FontBold = BTN_NEWORDER_FONT_BOLD
        .FontColour = BTN_NEWORDER_FONT_COLOUR
        .FontXJust = BTN_NEWORDER_FONT_X_JUST
        .FontYJust = BTN_NEWORDER_FONT_Y_JUST
    End With
    
    With GENERIC_BUTTON
        .BorderWidth = GENERIC_BUTTON_BORDER_WIDTH
        .Fill1 = GENERIC_BUTTON_FILL_1
        .Fill2 = GENERIC_BUTTON_FILL_2
        .Shadow = GENERIC_BUTTON_SHADOW
        .FontStyle = GENERIC_BUTTON_FONT_STYLE
        .FontSize = GENERIC_BUTTON_FONT_SIZE
        .FontBold = GENERIC_BUTTON_FONT_BOLD
        .FontColour = GENERIC_BUTTON_FONT_COLOUR
        .FontXJust = GENERIC_BUTTON_FONT_X_JUST
        .FontYJust = GENERIC_BUTTON_FONT_Y_JUST
    End With
    
    With GENERIC_LINEITEM
        .BorderWidth = GENERIC_LINEITEM_BORDER_WIDTH
        .Fill1 = GENERIC_LINEITEM_FILL_1
        .Fill2 = GENERIC_LINEITEM_FILL_2
        .Shadow = GENERIC_LINEITEM_SHADOW
        .FontStyle = GENERIC_LINEITEM_FONT_STYLE
        .FontSize = GENERIC_LINEITEM_FONT_SIZE
        .FontBold = GENERIC_LINEITEM_FONT_BOLD
        .FontColour = GENERIC_LINEITEM_FONT_COLOUR
        .FontXJust = GENERIC_LINEITEM_FONT_X_JUST
        .FontYJust = GENERIC_LINEITEM_FONT_Y_JUST
    End With

    With GENERIC_LINEITEM_HEADER
        .BorderWidth = GENERIC_LINEITEM_HEADER_BORDER_WIDTH
        .Fill1 = GENERIC_LINEITEM_HEADER_FILL_1
        .Fill2 = GENERIC_LINEITEM_HEADER_FILL_2
        .Shadow = GENERIC_LINEITEM_HEADER_SHADOW
        .FontStyle = GENERIC_LINEITEM_HEADER_FONT_STYLE
        .FontSize = GENERIC_LINEITEM_HEADER_FONT_SIZE
        .FontBold = GENERIC_LINEITEM_HEADER_FONT_BOLD
        .FontColour = GENERIC_LINEITEM_HEADER_FONT_COLOUR
        .FontXJust = GENERIC_LINEITEM_HEADER_FONT_X_JUST
        .FontYJust = GENERIC_LINEITEM_HEADER_FONT_Y_JUST
    End With
    
    BuildStylesMainScreen = True

Exit Function
    
    
ErrorExit:

    BuildStylesMainScreen = False
    
Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildMainFrame
' Builds main frame at top of screen
' ---------------------------------------------------------------
Private Function BuildMainFrame() As Boolean
    Const StrPROCEDURE As String = "BuildMainFrame()"

    On Error GoTo ErrorHandler

    
    'add main frame
    With MainFrame
        .Name = "Main Frame"
        MainScreen.Frames.AddItem MainFrame
            
        .Top = MAIN_FRAME_TOP
        .Left = MAIN_FRAME_LEFT
        .Width = MAIN_FRAME_WIDTH
        .Height = MAIN_FRAME_HEIGHT
        .Style = MAIN_FRAME_STYLE
        .EnableHeader = True

        With .Header
            .Top = .Parent.Top
            .Left = .Parent.Left
            .Width = .Parent.Width
            .Height = HEADER_HEIGHT
            .Name = "Main Frame Header"
            .Text = "Allocations - " & CurrentUser.Station.Name
            .Style = HEADER_STYLE
            .Icon = ShtMain.Shapes("TEMPLATE - Icon_Alocations").Duplicate
            .Icon.Top = .Parent.Top + HEADER_ICON_TOP
            .Icon.Left = .Parent.Left + .Parent.Width - .Icon.Width - HEADER_ICON_RIGHT
            .Icon.Name = .Parent.Name & " Icon"
            .Icon.Visible = msoCTrue
        End With
    End With
    
    
    BuildMainFrame = True

Exit Function

ErrorExit:

    BuildMainFrame = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildMainScreen
' Builds the display using shapes
' ---------------------------------------------------------------
Public Function BuildMainScreen() As Boolean
    
    Const StrPROCEDURE As String = "BuildMainScreen()"

    On Error GoTo ErrorHandler
    
    Set MainFrame = New ClsUIFrame
    Set LeftFrame = New ClsUIFrame
    Set RightFrame = New ClsUIFrame
    Set BtnNewOrder = New ClsUIMenuItem
    
    If Not BuildMainFrame Then Err.Raise HANDLED_ERROR
    If Not BuildLeftFrame Then Err.Raise HANDLED_ERROR
    If Not BuildRightFrame Then Err.Raise HANDLED_ERROR
    If Not BuildNewOrderBtn Then Err.Raise HANDLED_ERROR
    
    MainScreen.ReOrder
    
    BuildMainScreen = True
       
Exit Function

ErrorExit:

    BuildMainScreen = False
    
Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function


' ===============================================================
' BuildLeftFrame
' Builds Left frame at top of screen
' ---------------------------------------------------------------
Private Function BuildLeftFrame() As Boolean
    Const StrPROCEDURE As String = "BuildLeftFrame()"

    On Error GoTo ErrorHandler

    
    'add Left frame
    With LeftFrame
        .Name = "Left Frame"
        MainScreen.Frames.AddItem LeftFrame
        
        .Top = LEFT_FRAME_TOP
        .Left = LEFT_FRAME_LEFT
        .Width = LEFT_FRAME_WIDTH
        .Height = LEFT_FRAME_HEIGHT
        .Style = MAIN_FRAME_STYLE
        .EnableHeader = True

        With .Header
            .Top = .Parent.Top
            .Left = .Parent.Left
            .Width = .Parent.Width
            .Height = HEADER_HEIGHT
            .Name = "Left Frame Header"
            .Text = "Recent Orders"
            .Style = HEADER_STYLE
            .Icon = ShtMain.Shapes("TEMPLATE - Icon_Left_Frame").Duplicate
            .Icon.Top = .Parent.Top + HEADER_ICON_TOP
            .Icon.Left = .Parent.Left + .Parent.Width - .Icon.Width - HEADER_ICON_RIGHT
            .Icon.Name = .Parent.Name & " Icon"
            .Icon.Visible = msoCTrue
        End With
    End With

    
    BuildLeftFrame = True

Exit Function

ErrorExit:

    BuildLeftFrame = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildRightFrame
' Builds Right frame at top of screen
' ---------------------------------------------------------------
Private Function BuildRightFrame() As Boolean
    Const StrPROCEDURE As String = "BuildRightFrame()"

    On Error GoTo ErrorHandler

    
    'add Right frame
    With RightFrame
        .Name = "Right Frame"
        MainScreen.Frames.AddItem RightFrame
        
        .Top = RIGHT_FRAME_TOP
        .Left = RIGHT_FRAME_LEFT
        .Width = RIGHT_FRAME_WIDTH
        .Height = RIGHT_FRAME_HEIGHT
        .Style = MAIN_FRAME_STYLE
        .EnableHeader = True

        With .Header
            .Top = .Parent.Top
            .Left = .Parent.Left
            .Width = .Parent.Width
            .Height = HEADER_HEIGHT
            .Name = "Right Frame Header"
            .Text = "My Orders"
            .Style = HEADER_STYLE
            .Icon = ShtMain.Shapes("TEMPLATE - Icon_Right_Frame").Duplicate
            .Icon.Top = .Parent.Top + HEADER_ICON_TOP
            .Icon.Left = .Parent.Left + .Parent.Width - .Icon.Width - HEADER_ICON_RIGHT
            .Icon.Name = .Parent.Name & " Icon"
            .Icon.Visible = msoCTrue
        End With
    End With

    
    BuildRightFrame = True

Exit Function

ErrorExit:

    BuildRightFrame = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildNewOrderBtn
' Adds the new order button to the main screen
' ---------------------------------------------------------------
Private Function BuildNewOrderBtn() As Boolean

    Const StrPROCEDURE As String = "BuildNewOrderBtn()"

    On Error GoTo ErrorHandler

    With BtnNewOrder
        .Height = BTN_NEWORDER_HEIGHT
        .Left = BTN_NEWORDER_LEFT
        .Top = BTN_NEWORDER_TOP
        .Width = BTN_NEWORDER_WIDTH
        .Name = "New Order Button"
        .OnAction = "'moduimenu.ProcessBtnPress(6)'"
        .UnSelectStyle = BTN_NEWORDER_STYLE
        .Selected = False
        .Text = "New Order    "
        .Icon = ShtMain.Shapes("TEMPLATE - Icon_New_Order").Duplicate
        .Icon.Left = .Left + 290
        .Icon.Top = .Top + 16
        .Icon.Name = "New_Order_Button"
        .Icon.Visible = msoCTrue
    End With
    
    MainScreen.Menu.AddItem BtnNewOrder
    
    BuildNewOrderBtn = True

Exit Function

ErrorExit:

    BuildNewOrderBtn = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' DestroyMainScreen
' Builds the display using shapes
' ---------------------------------------------------------------
Public Function DestroyMainScreen() As Boolean
    Dim Frame As ClsUIFrame
    
    Const StrPROCEDURE As String = "DestroyMainScreen()"

    On Error GoTo ErrorHandler
    
    Set Frame = New ClsUIFrame
    
    For Each Frame In MainScreen.Frames
        If Frame.Name <> "MenuBar" Then
            MainScreen.Frames.RemoveItem Frame.Name
        End If
    Next
        
    If Not MainFrame Is Nothing Then MainFrame.Visible = False
    If Not LeftFrame Is Nothing Then LeftFrame.Visible = False
    If Not RightFrame Is Nothing Then RightFrame.Visible = False
    If Not BtnNewOrder Is Nothing Then BtnNewOrder.Visible = False
    
    Set MainFrame = Nothing
    Set LeftFrame = Nothing
    Set RightFrame = Nothing
    Set BtnNewOrder = Nothing
    
    Set Frame = Nothing
    
    DestroyMainScreen = True
       
Exit Function

ErrorExit:

    Set Frame = Nothing
    
    DestroyMainScreen = False
    
Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

