Attribute VB_Name = "ModUIMyProfile"
'===============================================================
' Module ModUIMyProfile
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 10 Oct 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModUIMyProfile"

' ===============================================================
' BuildMyProfileFrame1
' Builds first frame on My Profile page at top of screen
' ---------------------------------------------------------------
Public Function BuildMyProfileFrame1() As Boolean
    Dim TxtCrewNo As ClsUIDashObj
    Dim TxtForeName As ClsUIDashObj
    Dim TxtSurname As ClsUIDashObj
    Dim TxtUsername As ClsUIDashObj
    Dim TxtRole As ClsUIDashObj
    Dim TxtRankGrade As ClsUIDashObj
    Dim TxtLocation As ClsUIDashObj
    Dim TxtWatch As ClsUIDashObj
    Dim TxtAccessLvl As ClsUIDashObj
    Dim LblCrewNo As ClsUIDashObj
    Dim LblForeName As ClsUIDashObj
    Dim LblSurname As ClsUIDashObj
    Dim LblUsername As ClsUIDashObj
    Dim LblRole As ClsUIDashObj
    Dim LblRankGrade As ClsUIDashObj
    Dim LblLocation As ClsUIDashObj
    Dim LblWatch As ClsUIDashObj
    Dim LblAccessLvl As ClsUIDashObj
    Dim BtnUpdate As ClsUIMenuItem
    
    Const StrPROCEDURE As String = "BuildMyProfileFrame1()"

    On Error GoTo ErrorHandler

    Set TxtCrewNo = New ClsUIDashObj
    Set TxtForeName = New ClsUIDashObj
    Set TxtSurname = New ClsUIDashObj
    Set TxtUsername = New ClsUIDashObj
    Set TxtRole = New ClsUIDashObj
    Set TxtRankGrade = New ClsUIDashObj
    Set TxtLocation = New ClsUIDashObj
    Set TxtWatch = New ClsUIDashObj
    Set TxtAccessLvl = New ClsUIDashObj
    Set LblCrewNo = New ClsUIDashObj
    Set LblForeName = New ClsUIDashObj
    Set LblSurname = New ClsUIDashObj
    Set LblUsername = New ClsUIDashObj
    Set LblRole = New ClsUIDashObj
    Set LblRankGrade = New ClsUIDashObj
    Set LblLocation = New ClsUIDashObj
    Set LblWatch = New ClsUIDashObj
    Set LblAccessLvl = New ClsUIDashObj
    
    Set BtnUpdate = New ClsUIMenuItem

    With MyProfileFrame1
        .Name = "My Profile Frame 1"
        MainScreen.Frames.AddItem MyProfileFrame1
        
        .Top = MY_PROFILE_1_TOP
        .Left = MY_PROFILE_1_LEFT
        .Width = MY_PROFILE_1_WIDTH
        .Height = MY_PROFILE_1_HEIGHT
        .Style = MAIN_FRAME_STYLE
        .EnableHeader = True
        .Visible = True
                

        With .Header
            .Top = .Parent.Top
            .Left = .Parent.Left
            .Width = .Parent.Width
            .Height = HEADER_HEIGHT
            .Name = "My Profile 1 Header"
            .Text = "My Profile"
            .Style = HEADER_STYLE
            .Icon = ShtMain.Shapes("TEMPLATE - Icon_Head").Duplicate
            .Icon.Top = .Parent.Top + HEADER_ICON_TOP
            .Icon.Left = .Parent.Left + .Parent.Width - .Icon.Width - HEADER_ICON_RIGHT
            .Icon.Name = .Parent.Name & " Icon"
            .Icon.Visible = msoCTrue
        End With
    End With
    
    '--------------------------------------------------------------------------------
    'Crew No
    '--------------------------------------------------------------------------------
    With LblCrewNo
        .Name = "LblCrewNo"
        .ShpDashObj.Delete
        .ShpDashObj = ShtMain.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 10, 10, 10)
        MyProfileFrame1.DashObs.AddItem LblCrewNo
        .Top = MY_PROFILE_LBLCREWNO_TOP
        .Left = MY_PROFILE_LBLCREWNO_LEFT
        .Width = MY_PROFILE_LABEL_WIDTH
        .Height = MY_PROFILE_LABEL_HEIGHT
        .Style = GENERIC_LABEL
        .Text = "Crew No"
        .Locked = True
    End With
    
    With TxtCrewNo
        .Name = "TxtCrewNo"
        .ShpDashObj.Delete
        .ShpDashObj = ShtMain.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 10, 10, 10)
        MyProfileFrame1.DashObs.AddItem TxtCrewNo
        .Top = MY_PROFILE_TXTCREWNO_TOP
        .Left = MY_PROFILE_TXTCREWNO_LEFT
        .Width = MY_PROFILE_TEXTBOX_WIDTH
        .Height = MY_PROFILE_TEXTBOX_HEIGHT
        .Locked = False
    End With
    
    '--------------------------------------------------------------------------------
    'Username
    '--------------------------------------------------------------------------------
    With LblUsername
        .Name = "LblUsername"
        .ShpDashObj.Delete
        .ShpDashObj = ShtMain.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 10, 10, 10)
        MyProfileFrame1.DashObs.AddItem LblUsername
        .Top = MY_PROFILE_LBLUSERNAME_TOP
        .Left = MY_PROFILE_LBLUSERNAME_LEFT
        .Width = MY_PROFILE_LABEL_WIDTH
        .Height = MY_PROFILE_LABEL_HEIGHT
        .Style = GENERIC_LABEL
        .Text = "Username"
        .Locked = True
    End With
    
    With TxtUsername
        .Name = "TxtUsername"
        .ShpDashObj.Delete
        .ShpDashObj = ShtMain.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 10, 10, 10)
        MyProfileFrame1.DashObs.AddItem TxtUsername
        .Top = MY_PROFILE_TXTUSERNAME_TOP
        .Left = MY_PROFILE_TXTUSERNAME_LEFT
        .Width = MY_PROFILE_TEXTBOX_WIDTH
        .Height = MY_PROFILE_TEXTBOX_HEIGHT
        .Locked = False
    End With
    
    '--------------------------------------------------------------------------------
    'Forename
    '--------------------------------------------------------------------------------
    With LblForeName
        .Name = "LblForename"
        .ShpDashObj.Delete
        .ShpDashObj = ShtMain.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 10, 10, 10)
        MyProfileFrame1.DashObs.AddItem LblForeName
        .Top = MY_PROFILE_LBLFORENAME_TOP
        .Left = MY_PROFILE_LBLFORENAME_LEFT
        .Width = MY_PROFILE_LABEL_WIDTH
        .Height = MY_PROFILE_LABEL_HEIGHT
        .Style = GENERIC_LABEL
        .Text = "Forename"
        .Locked = True
    End With
    
    With TxtForeName
        .Name = "TxtForeName"
        .ShpDashObj.Delete
        .ShpDashObj = ShtMain.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 10, 10, 10)
        MyProfileFrame1.DashObs.AddItem TxtForeName
        .Top = MY_PROFILE_TXTFORENAME_TOP
        .Left = MY_PROFILE_TXTFORENAME_LEFT
        .Width = MY_PROFILE_TEXTBOX_WIDTH
        .Height = MY_PROFILE_TEXTBOX_HEIGHT
        .Locked = False
    End With
    
    '--------------------------------------------------------------------------------
    'Surname
    '--------------------------------------------------------------------------------
    With LblSurname
        .Name = "LblSurname"
        .ShpDashObj.Delete
        .ShpDashObj = ShtMain.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 10, 10, 10)
        MyProfileFrame1.DashObs.AddItem LblSurname
        .Top = MY_PROFILE_LBLSURNAME_TOP
        .Left = MY_PROFILE_LBLSURNAME_LEFT
        .Width = MY_PROFILE_LABEL_WIDTH
        .Height = MY_PROFILE_LABEL_HEIGHT
        .Style = GENERIC_LABEL
        .Text = "Surname"
        .Locked = True
    End With
    
    With TxtSurname
        .Name = "TxtSurname"
        .ShpDashObj.Delete
        .ShpDashObj = ShtMain.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 10, 10, 10)
        MyProfileFrame1.DashObs.AddItem TxtSurname
        .Top = MY_PROFILE_TXTSURNAME_TOP
        .Left = MY_PROFILE_TXTSURNAME_LEFT
        .Width = MY_PROFILE_TEXTBOX_WIDTH
        .Height = MY_PROFILE_TEXTBOX_HEIGHT
        .Locked = False
    End With
    
    '--------------------------------------------------------------------------------
    'Role
    '--------------------------------------------------------------------------------
    With LblRole
        .Name = "LblRole"
        .ShpDashObj.Delete
        .ShpDashObj = ShtMain.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 10, 10, 10)
        MyProfileFrame1.DashObs.AddItem LblRole
        .Top = MY_PROFILE_LBLROLE_TOP
        .Left = MY_PROFILE_LBLROLE_LEFT
        .Width = MY_PROFILE_LABEL_WIDTH
        .Height = MY_PROFILE_LABEL_HEIGHT
        .Style = GENERIC_LABEL
        .Text = "Role"
        .Locked = True
    End With
    
    With TxtRole
        .Name = "TxtRole"
        .ShpDashObj.Delete
        .ShpDashObj = ShtMain.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 10, 10, 10)
        MyProfileFrame1.DashObs.AddItem TxtRole
        .Top = MY_PROFILE_TXTROLE_TOP
        .Left = MY_PROFILE_TXTROLE_LEFT
        .Width = MY_PROFILE_TEXTBOX_WIDTH
        .Height = MY_PROFILE_TEXTBOX_HEIGHT
        .Locked = False
    End With
    
    '--------------------------------------------------------------------------------
    'RankGrade
    '--------------------------------------------------------------------------------
    With LblRankGrade
        .Name = "LblRankGrade"
        .ShpDashObj.Delete
        .ShpDashObj = ShtMain.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 10, 10, 10)
        MyProfileFrame1.DashObs.AddItem LblRankGrade
        .Top = MY_PROFILE_LBLRANKGRADE_TOP
        .Left = MY_PROFILE_LBLRANKGRADE_LEFT
        .Width = MY_PROFILE_LABEL_WIDTH
        .Height = MY_PROFILE_LABEL_HEIGHT
        .Style = GENERIC_LABEL
        .Text = "Rank / Grade"
        .Locked = True
    End With
    
    With TxtRankGrade
        .Name = "TxtRankGrade"
        .ShpDashObj.Delete
        .ShpDashObj = ShtMain.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 10, 10, 10)
        MyProfileFrame1.DashObs.AddItem TxtRankGrade
        .Top = MY_PROFILE_TXTRANKGRADE_TOP
        .Left = MY_PROFILE_TXTRANKGRADE_LEFT
        .Width = MY_PROFILE_TEXTBOX_WIDTH
        .Height = MY_PROFILE_TEXTBOX_HEIGHT
        .Locked = False
    End With
    
    '--------------------------------------------------------------------------------
    'Location
    '--------------------------------------------------------------------------------
    With LblLocation
        .Name = "LblLocation"
        .ShpDashObj.Delete
        .ShpDashObj = ShtMain.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 10, 10, 10)
        MyProfileFrame1.DashObs.AddItem LblLocation
        .Top = MY_PROFILE_LBLLOCATION_TOP
        .Left = MY_PROFILE_LBLLOCATION_LEFT
        .Width = MY_PROFILE_LABEL_WIDTH
        .Height = MY_PROFILE_LABEL_HEIGHT
        .Style = GENERIC_LABEL
        .Text = "Location"
        .Locked = True
    End With
    
    With TxtLocation
        .Name = "TxtLocation"
        .ShpDashObj.Delete
        .ShpDashObj = ShtMain.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 10, 10, 10)
        MyProfileFrame1.DashObs.AddItem TxtLocation
        .Top = MY_PROFILE_TXTLOCATION_TOP
        .Left = MY_PROFILE_TXTLOCATION_LEFT
        .Width = MY_PROFILE_TEXTBOX_WIDTH
        .Height = MY_PROFILE_TEXTBOX_HEIGHT
        .Locked = False
    End With
        
    '--------------------------------------------------------------------------------
    'Watch
    '--------------------------------------------------------------------------------
    With LblWatch
        .Name = "LblWatch"
        .ShpDashObj.Delete
        .ShpDashObj = ShtMain.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 10, 10, 10)
        MyProfileFrame1.DashObs.AddItem LblWatch
        .Top = MY_PROFILE_LBLWATCH_TOP
        .Left = MY_PROFILE_LBLWATCH_LEFT
        .Width = MY_PROFILE_LABEL_WIDTH
        .Height = MY_PROFILE_LABEL_HEIGHT
        .Style = GENERIC_LABEL
        .Text = "Watch"
        .Locked = True
    End With
    
    With TxtWatch
        .Name = "TxtWatch"
        .ShpDashObj.Delete
        .ShpDashObj = ShtMain.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 10, 10, 10)
        MyProfileFrame1.DashObs.AddItem TxtWatch
        .Top = MY_PROFILE_TXTWATCH_TOP
        .Left = MY_PROFILE_TXTWATCH_LEFT
        .Width = MY_PROFILE_TEXTBOX_WIDTH
        .Height = MY_PROFILE_TEXTBOX_HEIGHT
        .Locked = False
    End With
      
    '--------------------------------------------------------------------------------
    'AccessLvl
    '--------------------------------------------------------------------------------
    With LblAccessLvl
        .Name = "LblAccessLvl"
        .ShpDashObj.Delete
        .ShpDashObj = ShtMain.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 10, 10, 10)
        MyProfileFrame1.DashObs.AddItem LblAccessLvl
        .Top = MY_PROFILE_LBLACCESSLVL_TOP
        .Left = MY_PROFILE_LBLACCESSLVL_LEFT
        .Width = MY_PROFILE_LABEL_WIDTH
        .Height = MY_PROFILE_LABEL_HEIGHT
        .Style = GENERIC_LABEL
        .Text = "Access Level"
        .Locked = True
    End With
    
    With TxtAccessLvl
        .Name = "TxtAccessLvl"
        .ShpDashObj.Delete
        .ShpDashObj = ShtMain.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 10, 10, 10)
        MyProfileFrame1.DashObs.AddItem TxtAccessLvl
        .Top = MY_PROFILE_TXTACCESSLVL_TOP
        .Left = MY_PROFILE_TXTACCESSLVL_LEFT
        .Width = MY_PROFILE_TEXTBOX_WIDTH
        .Height = MY_PROFILE_TEXTBOX_HEIGHT
        .Locked = False
    End With
            
    '--------------------------------------------------------------------------------
    'Update Button
    '--------------------------------------------------------------------------------
    With BtnUpdate
        .Name = "BtnUpdate"
        MyProfileFrame1.Menu.AddItem BtnUpdate
        .UnSelectStyle = GENERIC_BUTTON
        .Selected = False
        .Top = MY_PROFILE_BTNUPDATE_TOP
        .Left = MY_PROFILE_BTNUPDATE_LEFT
        .Height = MY_PROFILE_BUTTON_HEIGHT
        .Width = MY_PROFILE_BUTTON_WIDTH
        .Text = "Update"
        .OnAction = "'ModUISupportScreen.ProcessBtnPress(" & EnumSupportMsg & ")'"
    End With
    
    MyProfileFrame1.ReOrder
    
    
    BuildMyProfileFrame1 = True

Exit Function

ErrorExit:

    Set TxtCrewNo = Nothing
    Set TxtForeName = Nothing
    Set TxtSurname = Nothing
    Set TxtUsername = Nothing
    Set TxtRole = Nothing
    Set TxtRankGrade = Nothing
    Set TxtLocation = Nothing
    Set TxtWatch = Nothing
    Set TxtAccessLvl = Nothing
    Set LblCrewNo = Nothing
    Set LblForeName = Nothing
    Set LblSurname = Nothing
    Set LblUsername = Nothing
    Set LblRole = Nothing
    Set LblRankGrade = Nothing
    Set LblLocation = Nothing
    Set LblWatch = Nothing
    Set LblAccessLvl = Nothing
    Set BtnUpdate = Nothing
    
    BuildMyProfileFrame1 = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildProfileScreen
' Builds the display using shapes
' ---------------------------------------------------------------
Public Function BuildProfileScreen() As Boolean
    
    Const StrPROCEDURE As String = "BuildProfileScreen()"

    On Error GoTo ErrorHandler
    
    Set MyProfileFrame1 = New ClsUIFrame
    
    If Not BuildMyProfileFrame1 Then Err.Raise HANDLED_ERROR
    
    BuildProfileScreen = True
       
Exit Function

ErrorExit:

    BuildProfileScreen = False
    
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
            
'                If Not BtnFeedbackSend Then Err.Raise HANDLED_ERROR
                
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




