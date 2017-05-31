Attribute VB_Name = "ModUISupportScreen"
'===============================================================
' Module ModUISupportScreen
' v0,0 - Initial Version
' v0,1 - improved message box
' v0,2 - Fix Error 287 by opening Outlook if closed
' v0,3 - 287 issue, tried new Outlook detector
' v0,4 - Add shane and emma to support message
'---------------------------------------------------------------
' Date - 31 May 17
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
        .Width = SUPPORT_FRAME_1_HEIGHT
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
        .OnAction = "'ModUISupportScreen.ProcessBtnPress(EnumSupportMsg)'"
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
' BuildSupportScreen
' Builds the display using shapes
' ---------------------------------------------------------------
Public Function BuildSupportScreen() As Boolean
    
    Const StrPROCEDURE As String = "BuildSupportScreen()"

    On Error GoTo ErrorHandler
    
    Set SupportFrame1 = New ClsUIFrame
    
    If Not BuildSupportFrame1 Then Err.Raise HANDLED_ERROR
    
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
                
                MsgBox "Thank you for your feedback", vbOKOnly + vbInformation, APP_NAME
                        
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

    If Not ModLibrary.OutlookRunning Then
        Shell "Outlook.exe"
    End If

    Set MailSystem = New ClsMailSystem
    
    With MailSystem.MailItem
        .To = "Julian Turner; Emma Morton; Shane Redhead"
        .Subject = "Stores IT Project - Feedback received from " & Application.UserName
        .Body = SupportFrame1.DashObs("CommentBox").Text
        If SEND_EMAILS Then .Send
    End With
    
    SupportFrame1.DashObs("CommentBox").Text = ""
    Set MailSystem = Nothing
    
    BtnFeedbackSend = True
    
Exit Function

ErrorExit:
    BtnFeedbackSend = False
    
    Terminate

    Set MailSystem = Nothing

Exit Function

ErrorHandler:


    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function


