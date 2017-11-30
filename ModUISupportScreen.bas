Attribute VB_Name = "ModUISupportScreen"
'===============================================================
' Module ModUISupportScreen
' v0,0 - Initial Version
' v0,1 - improved message box
' v0,2 - Fix Error 287 by opening Outlook if closed
' v0,3 - 287 issue, tried new Outlook detector
' v0,4 - Add shane and emma to support message
' v0,5 - Removed hard numbering for buttons
' v0,6 - Add Julia Whitfield as cc to support query email
' v0,7 - Added Release notes
' v0,8 - Centralised mail messages
'---------------------------------------------------------------
' Date - 30 Nov 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModUISupportScreen"

' ===============================================================
' BuildSupportFrame1
' Builds first frame on support page at top of screen
' ---------------------------------------------------------------
Public Function BuildSupportFrame1() As Boolean
    Dim CommentBox As ClsUIDashObj
    Dim CommentBtn As ClsUIMenuItem
    
    Const StrPROCEDURE As String = "BuildSupportFrame1()"

    On Error GoTo ErrorHandler

    Set CommentBox = New ClsUIDashObj
    Set CommentBtn = New ClsUIMenuItem

    With SupportFrame1
        .Name = "Support Frame 1"
        MainScreen.Frames.AddItem SupportFrame1
        
        .Top = SUPPORT_FRAME_1_TOP
        .Left = SUPPORT_FRAME_1_LEFT
        .Width = SUPPORT_FRAME_1_WIDTH
        .Height = SUPPORT_FRAME_1_HEIGHT
        .Style = MAIN_FRAME_STYLE
        .EnableHeader = True
        .Visible = True
                

        With .Header
            .Top = .Parent.Top
            .Left = .Parent.Left
            .Width = .Parent.Width
            .Height = HEADER_HEIGHT
            .Name = "Support 1 Header"
            .Text = "Support Message"
            .Style = HEADER_STYLE
            .Icon = ShtMain.Shapes("TEMPLATE - Message").Duplicate
            .Icon.Top = .Parent.Top + HEADER_ICON_TOP
            .Icon.Left = .Parent.Left + .Parent.Width - .Icon.Width - HEADER_ICON_RIGHT
            .Icon.Name = .Parent.Name & " Icon"
            .Icon.Visible = msoCTrue
        End With
    End With
    
    With CommentBox
        .Name = "CommentBox"
        .ShpDashObj.Delete
        .ShpDashObj = ShtMain.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 10, 10, 10)
        SupportFrame1.DashObs.AddItem CommentBox
        .Top = COMMENT_BOX_TOP
        .Left = COMMENT_BOX_LEFT
        .Width = COMMENT_BOX_WIDTH
        .Height = COMMENT_BOX_HEIGHT
        .Locked = False
    End With
    
    With CommentBtn
        .Name = "CommentBtn"
        SupportFrame1.Menu.AddItem CommentBtn
        .UnSelectStyle = GENERIC_BUTTON
        .Selected = False
        .Top = COMMENT_BTN_TOP
        .Left = COMMENT_BTN_LEFT
        .Height = COMMENT_BTN_HEIGHT
        .Width = COMMENT_BTN_WIDTH
        .Text = "Send Message"
        .OnAction = "'ModUISupportScreen.ProcessBtnPress(" & EnumSupportMsg & ")'"
    End With
    
    SupportFrame1.ReOrder
    
    Set CommentBox = Nothing
    Set CommentBtn = Nothing
    
    BuildSupportFrame1 = True

Exit Function

ErrorExit:

    Set CommentBox = Nothing
    Set CommentBtn = Nothing
    
    BuildSupportFrame1 = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildSupportFrame2
' Builds first frame on support page at top of screen
' ---------------------------------------------------------------
Public Function BuildSupportFrame2() As Boolean
    Dim TxtReleaseNotes As ClsUIDashObj
    Dim CommentBtn As ClsUIMenuItem
    Dim RstReleaseNotes As Recordset
    Dim StrReleaseNotes As String
    
    Const StrPROCEDURE As String = "BuildSupportFrame2()"

    On Error GoTo ErrorHandler

    Set TxtReleaseNotes = New ClsUIDashObj
    
    Set RstReleaseNotes = SQLQuery("TblMessage")
    
    If RstReleaseNotes.RecordCount > 0 Then StrReleaseNotes = RstReleaseNotes.Fields(1)
    
    With SupportFrame2
        .Name = "Support Frame 2"
        MainScreen.Frames.AddItem SupportFrame2
        
        .Top = SUPPORT_FRAME_2_TOP
        .Left = SUPPORT_FRAME_2_LEFT
        .Width = SUPPORT_FRAME_2_WIDTH
        .Height = SUPPORT_FRAME_2_HEIGHT
        .Style = MAIN_FRAME_STYLE
        .EnableHeader = True
        .Visible = True
                

        With .Header
            .Top = .Parent.Top
            .Left = .Parent.Left
            .Width = .Parent.Width
            .Height = HEADER_HEIGHT
            .Name = "Support 2 Header"
            .Text = "Latest Release Notes"
            .Style = HEADER_STYLE
            .Icon = ShtMain.Shapes("TEMPLATE - Icon_Document").Duplicate
            .Icon.Top = .Parent.Top + HEADER_ICON_TOP
            .Icon.Left = .Parent.Left + .Parent.Width - .Icon.Width - HEADER_ICON_RIGHT
            .Icon.Name = .Parent.Name & " Icon"
            .Icon.Visible = msoCTrue
        End With
    End With
    
    With TxtReleaseNotes
        .Name = "TxtReleaseNotes"
        .ShpDashObj.Delete
        .ShpDashObj = ShtMain.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 10, 10, 10)
        SupportFrame2.DashObs.AddItem TxtReleaseNotes
        .Top = RELEASE_NOTES_TOP
        .Left = RELEASE_NOTES_LEFT
        .Width = RELEASE_NOTES_WIDTH
        .Height = RELEASE_NOTES_HEIGHT
        .Style = TRANSPARENT_TEXT_BOX
        .Locked = True
        .Text = StrReleaseNotes
    End With
        
    SupportFrame2.ReOrder
    
    Set TxtReleaseNotes = Nothing
    Set CommentBtn = Nothing
    Set RstReleaseNotes = Nothing
    
    BuildSupportFrame2 = True

Exit Function

ErrorExit:

    Set TxtReleaseNotes = Nothing
    Set CommentBtn = Nothing
    Set RstReleaseNotes = Nothing
    
    BuildSupportFrame2 = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildSupportScreen
' Builds the display using shapes
' ---------------------------------------------------------------
Public Function BuildSupportScreen() As Boolean
    
    Const StrPROCEDURE As String = "BuildSupportScreen()"

    On Error GoTo ErrorHandler
    
    Set SupportFrame1 = New ClsUIFrame
    Set SupportFrame2 = New ClsUIFrame
    
    If Not BuildSupportFrame1 Then Err.Raise HANDLED_ERROR
    If Not BuildSupportFrame2 Then Err.Raise HANDLED_ERROR
    
    BuildSupportScreen = True
       
Exit Function

ErrorExit:

    BuildSupportScreen = False
    
Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' DestroySupportScreen
' Builds the display using shapes
' ---------------------------------------------------------------
Public Function DestroySupportScreen() As Boolean
    Dim Frame As ClsUIFrame
    
    Const StrPROCEDURE As String = "DestroySupportScreen()"

    On Error GoTo ErrorHandler
    
    Set Frame = New ClsUIFrame
    
    For Each Frame In MainScreen.Frames
        If Frame.Name <> "MenuBar" Then
            MainScreen.Frames.RemoveItem Frame.Name
        End If
    Next

    If Not SupportFrame1 Is Nothing Then SupportFrame1.Visible = False
    
    Set SupportFrame1 = Nothing
    
    DestroySupportScreen = True
       
Exit Function

ErrorExit:

    DestroySupportScreen = False
    
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
        
            Case EnumSupportMsg
            
                If Not BtnFeedbackSend Then Err.Raise HANDLED_ERROR
                
                MsgBox "Thank you for your message", vbOKOnly + vbInformation, APP_NAME
                        
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
' BtnFeedbackSend
' Sends feedback
' ---------------------------------------------------------------
Private Function BtnFeedbackSend() As Boolean

    Const StrPROCEDURE As String = "BtnFeedbackSend()"

    On Error GoTo ErrorHandler

    If Not ModReports.SendEmailReports("Stores IT Project - Query received from " & CurrentUser.UserName, SupportFrame1.DashObs("CommentBox").Text, EnumSupportQueryRecieved) Then Err.Raise HANDLED_ERROR
    
    SupportFrame1.DashObs("CommentBox").Text = ""
        
    BtnFeedbackSend = True
    
Exit Function

ErrorExit:
    BtnFeedbackSend = False
    
    Terminate

Exit Function

ErrorHandler:


    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function


