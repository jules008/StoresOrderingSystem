VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmUserAdmin 
   Caption         =   "Action Plan"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8010
   OleObjectBlob   =   "FrmUserAdmin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmUserAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 20 Mar 17
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmUserAdmin"

Private SelectedUser As ClsPerson

' ===============================================================
' ShowForm
' Initial entry point to form
' ---------------------------------------------------------------
Public Function ShowForm() As Boolean
    
   Const StrPROCEDURE As String = "ShowForm()"
   
   On Error GoTo ErrorHandler
   
    If Not ResetForm Then Err.Raise HANDLED_ERROR
    Show

    ShowForm = True
Exit Function

ErrorExit:
    ShowForm = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BtnClose_Click
' Event for page close button
' ---------------------------------------------------------------
Private Sub BtnClose_Click()

    On Error Resume Next
    
    FormTerminate
    
End Sub

' ===============================================================
' BtnDel_Click
' Deletes user
' ---------------------------------------------------------------
Private Sub BtnDel_Click()
    Dim Response As Integer
    Dim SelUser As Integer
    Dim UserName As String
    
    Const StrPROCEDURE As String = "BtnDel_Click()"
    
    On Error GoTo ErrorHandler
        
    SelUser = LstAccessList.ListIndex
    
    If SelUser <> -1 Then
        UserName = LstAccessList.List(SelUser, 0)
        Response = MsgBox("Are you sure you want to remove " _
                            & UserName & " from the system? ", 36)
    
        If Response = 6 Then SelectedUser.DBDelete FullDelete:=True
        
        If Not ShtLists.RefreshNameList Then Err.Raise HANDLED_ERROR

        If Not ResetForm Then Err.Raise HANDLED_ERROR
    End If

Exit Sub

ErrorExit:

'    ***CleanUpCode***

Exit Sub

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' BtnAdd_Click
' Clears form for new user
' ---------------------------------------------------------------
Private Sub BtnAdd_Click()
    Const StrPROCEDURE As String = "BtnAdd_Click()"

    On Error GoTo ErrorHandler

    If Not ResetForm Then Err.Raise HANDLED_ERROR

Exit Sub

ErrorExit:

'    ***CleanUpCode***

Exit Sub

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' BtnShowMail_Click
' Shows users with mail alert selected
' ---------------------------------------------------------------
Private Sub BtnShowMail_Click()
    Dim Persons As ClsPersons
    Dim RstPersons As Recordset
    Dim i As Integer
    
    Const StrPROCEDURE As String = "BtnShowMail_Click()"

    On Error GoTo ErrorHandler
    
    Set Persons = New ClsPersons
    Set RstPersons = Persons.GetMailAlertUsers
    
    With LstAccessList
        .Clear
        
        For i = 0 To RstPersons.RecordCount - 1
            .AddItem
            .List(i, 0) = RstPersons!CrewNo
            .List(i, 1) = RstPersons!UserName
            RstPersons.MoveNext
        Next
    
    End With



    Set Persons = Nothing
    Set RstPersons = Nothing
Exit Sub

ErrorExit:
    Set Persons = Nothing
    Set RstPersons = Nothing
'    ***CleanUpCode***

Exit Sub

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' BtnUpdate_Click
' updates users details
' ---------------------------------------------------------------
Private Sub BtnUpdate_Click()
    Const StrPROCEDURE As String = "BtnUpdate_Click()"
    
    On Error GoTo ErrorHandler

    Select Case ValidateForm
    
        Case Is = FunctionalError
            Err.Raise HANDLED_ERROR
        
        Case Is = FormOK
        
            With SelectedUser
                .CrewNo = TxtCrewNo
                .AccessLvl = CmoAccessLvl.ListIndex
                .Forename = TxtForeName
                .Role = TxtRole
                .Surname = TxtSurname
                .RankGrade = TxtRankGrade
                .Station.DBGet CmoStation.ListIndex
                .Watch = TxtWatch
                .MailAlert = ChkMailAlert
                .DBSave
            End With
            MsgBox "User updated successfully.  If any change have been made " _
                    & "to your own profile, please restart the system for these changes " _
                     & "to take effect"
            
            If Not ShtLists.RefreshNameList Then Err.Raise HANDLED_ERROR

    End Select
Exit Sub

ErrorExit:

Exit Sub

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' LstAccessList_Click
' Selects user that is clicked in the list
' ---------------------------------------------------------------
Private Sub LstAccessList_Click()
    Const StrPROCEDURE As String = "LstAccessList_Click()"

    On Error GoTo ErrorHandler

    If Not RefreshUserDetails Then Err.Raise HANDLED_ERROR

Exit Sub

ErrorExit:

'    ***CleanUpCode***

Exit Sub

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' TxtCrewNo_Change
' Textbox change
' ---------------------------------------------------------------
Private Sub TxtCrewNo_Change()
    TxtCrewNo.BackColor = COLOUR_3
End Sub

' ===============================================================
' TxtForeName_Change
' Textbox change
' ---------------------------------------------------------------
Private Sub TxtForeName_Change()
    TxtForeName.BackColor = COLOUR_3
End Sub

' ===============================================================
' TxtRankGrade_Change
' Textbox change
' ---------------------------------------------------------------
Private Sub TxtRankGrade_Change()
    TxtRankGrade.BackColor = COLOUR_3
End Sub

' ===============================================================
' TxtRole_Change
' Textbox change
' ---------------------------------------------------------------
Private Sub TxtRole_Change()
    TxtRole.BackColor = COLOUR_3
End Sub

' ===============================================================
' TxtSearch_Change
' Entry for search string
' ---------------------------------------------------------------
Private Sub TxtSearch_Change()
    Const StrPROCEDURE As String = "TxtSearch_Change()"

    On Error GoTo ErrorHandler

    Dim ListResults As String

    On Error GoTo ErrorHandler
    
    If TxtSearch = "" Then
        If Not ResetForm Then Err.Raise HANDLED_ERROR
    End If
    
    With LstAccessList
        If .ListIndex <> -1 Then ListResults = .List(.ListIndex)
    
        'if the search box has been changed since being updated by the results box, clear the result box
        If ListResults <> TxtSearch.Value Then .ListIndex = -1
        
        'if the results box has been clicked, add the selected result to the search box
        If .ListIndex = -1 Then
        
            'if no results selected, populate with new results
            If Len(TxtSearch.Value) > 1 Then
                If Not GetSearchItems(TxtSearch.Value) Then Err.Raise HANDLED_ERROR
            Else
                .Clear
            End If
        End If
    End With

Exit Sub

ErrorExit:

'    ***CleanUpCode***

Exit Sub

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' ValidateForm
' Ensures the form is filled out correctly before moving on
' ---------------------------------------------------------------
Private Function ValidateForm() As EnumFormValidation
    Const StrPROCEDURE As String = "ValidateForm()"

    On Error GoTo ErrorHandler


    With TxtCrewNo
        If Trim(.Value) = "" Then
            .BackColor = COLOUR_6
            ValidateForm = ValidationError
        End If
    End With
    
    If Not IsNumeric(TxtCrewNo) Then Err.Raise CREWNO_UNRECOGNISED

    With TxtRankGrade
        If Trim(.Value) = "" Then
            .BackColor = COLOUR_6
            ValidateForm = ValidationError
        End If
    End With
    
    With TxtForeName
        If Trim(.Value) = "" Then
            .BackColor = COLOUR_6
            ValidateForm = ValidationError
        End If
    End With
    
    With TxtSurname
        If Trim(.Value) = "" Then
            .BackColor = COLOUR_6
            ValidateForm = ValidationError
        End If
    End With
    
    With CmoStation
        If .ListIndex = -1 Then
            .BackColor = COLOUR_6
            ValidateForm = ValidationError
        End If
    End With
    
    With TxtRole
        If Trim(.Value) = "" Then
            .BackColor = COLOUR_6
            ValidateForm = ValidationError
        End If
    End With
    
    With CmoAccessLvl
        If .ListIndex = -1 Then
            .BackColor = COLOUR_6
            ValidateForm = ValidationError
        End If
    End With
    
    If ValidateForm = ValidationError Then
        Err.Raise FORM_INPUT_EMPTY
    Else
        ValidateForm = FormOK
    End If

Exit Function

ValidationError:


Exit Function

ErrorExit:

    ValidateForm = FunctionalError
    FormTerminate
    Terminate

Exit Function

ErrorHandler:
    
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        CustomErrorHandler Err.Number
        Resume ValidationError
    End If

If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' ResetForm
' Resets form fields
' ---------------------------------------------------------------
Private Function ResetForm() As Boolean
    Const StrPROCEDURE As String = "ResetForm()"

    On Error GoTo ErrorHandler


    CmoAccessLvl = ""
    CmoStation = ""
    TxtCrewNo = ""
    TxtForeName = ""
    TxtSurname = ""
    TxtSearch = ""
    TxtRankGrade = ""
    TxtRole = ""


    ResetForm = True

Exit Function

ErrorExit:

'    ***CleanUpCode***
    ResetForm = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function


' ===============================================================
' RefreshUserList
' Lists all users in system
' ---------------------------------------------------------------
Public Function RefreshUserList() As Boolean
    Const StrPROCEDURE As String = "RefreshUserList()"

    Dim RstUserList As Recordset
    Dim i As Integer
    
    On Error GoTo ErrorHandler

'   Set RstUserList = GetAccessList
    
    LstAccessList.Clear
    
    If Not RstUserList Is Nothing Then
        With RstUserList
            Do While Not .EOF
                    
                LstAccessList.AddItem
                LstAccessList.List(i, 0) = RstUserList!UserName
                .MoveNext
                i = i + 1
            Loop
        End With
    End If
    Set RstUserList = Nothing

    RefreshUserList = True
Exit Function

ErrorExit:
    Set RstUserList = Nothing
    RefreshUserList = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' GetSearchItems
' Gets items from the name list that match Txtsearch box
' ---------------------------------------------------------------
Private Function GetSearchItems(StrSearch As String) As Boolean
    Const StrPROCEDURE As String = "GetSearchItems()"

    On Error GoTo ErrorHandler

    Dim ListLength As Integer
    Dim RngResult As Range
    Dim RngItems As Range
    Dim RngFirstResult As Range
    Dim i As Integer
    
    'get length of item list
    ListLength = Application.WorksheetFunction.CountA(ShtLists.Range("C:C"))
    
    If IsNumeric(StrSearch) Then
    
        Set RngItems = ShtLists.Range("C1:C" & ListLength)
    Else
        Set RngItems = ShtLists.Range("D1:D" & ListLength)
    
    End If
        
    Set RngResult = RngItems.Find(StrSearch)
    Set RngFirstResult = RngResult
    
    LstAccessList.Clear
    'search item list and populate results.  Stop before looping back to start
    If Not RngResult Is Nothing Then
    
        i = 0
        Do
            Set RngResult = RngItems.FindNext(RngResult)
                With LstAccessList
                    .AddItem
                    If IsNumeric(StrSearch) Then
                        .List(i, 0) = RngResult.Value
                        .List(i, 1) = RngResult.Offset(0, 1)
                    Else
                        .List(i, 1) = RngResult.Value
                        .List(i, 0) = RngResult.Offset(0, -1)
                    End If
                    i = i + 1
            End With
        Loop While RngResult <> 0 And RngResult.Address <> RngFirstResult.Address
    End If

    GetSearchItems = True
    
    Set RngItems = Nothing
    Set RngResult = Nothing
    Set RngFirstResult = Nothing
    
Exit Function

ErrorExit:

    Set RngItems = Nothing
    Set RngResult = Nothing
    Set RngFirstResult = Nothing
    
    FormTerminate
    Terminate

    GetSearchItems = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' RefreshUserDetails
' Adds details from selected user
' ---------------------------------------------------------------
Private Function RefreshUserDetails() As Boolean
    Dim ListSelection As Integer
    Dim CrewNo As String
    
    Const StrPROCEDURE As String = "RefreshUserDetails()"

    On Error GoTo ErrorHandler

    ListSelection = LstAccessList.ListIndex
    
    If ListSelection = -1 Then
        TxtCrewNo = ""
        TxtForeName = ""
        CmoAccessLvl = ""
        CmoStation = ""
        TxtRankGrade = ""
        TxtRole = ""
        TxtSurname = ""
    Else
        CrewNo = LstAccessList.List(ListSelection, 0)
        
        SelectedUser.DBGet CrewNo
        
        With SelectedUser
            TxtCrewNo = .CrewNo
            TxtForeName = .Forename
            CmoAccessLvl.ListIndex = .AccessLvl
            CmoStation.ListIndex = .Station.StationID
            TxtRankGrade = .RankGrade
            TxtRole = .Role
            TxtWatch = .Watch
            TxtSurname = .Surname
            ChkMailAlert = .MailAlert
        End With
    
    End If

    RefreshUserDetails = True

Exit Function

ErrorExit:

'    ***CleanUpCode***
    RefreshUserDetails = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' TxtSurname_Change
' Textbox change
' ---------------------------------------------------------------
Private Sub TxtSurname_Change()
    TxtSurname.BackColor = COLOUR_3
End Sub

' ===============================================================
' UserForm_Terminate
' Automatic Terminate event that triggers custom Terminate
' ---------------------------------------------------------------
Private Sub UserForm_Terminate()

    On Error Resume Next
    
    FormTerminate
    
End Sub

' ===============================================================
' FormInitialise
' initialises controls on form at start up
' ---------------------------------------------------------------
Private Function FormInitialise() As Boolean
    Dim RstAccessLevels As Recordset
    Dim i As Integer
    
    Const StrPROCEDURE As String = "FormInitialise()"

    On Error GoTo ErrorHandler

Restart:

    Application.StatusBar = ""
    
    If Vehicles Is Nothing Then Err.Raise SYSTEM_RESTART, Description:="Object Model Failed, system restarting"
    
    If Stations Is Nothing Then Err.Raise SYSTEM_RESTART, Description:="Object Model Failed, system restarting"
    
    If CurrentUser Is Nothing Then Err.Raise SYSTEM_RESTART, Description:="Object Model Failed, system restarting"
        
    If ModErrorHandling.FaultCount1002 > 0 Then ModErrorHandling.FaultCount1002 = 0
    
    If Not ShtLists.RefreshNameList Then Err.Raise HANDLED_ERROR
    
    Set SelectedUser = New ClsPerson
    
    With LstHeadings
        .AddItem
        .List(0, 0) = "No"
        .List(0, 1) = "Name"
    End With
    
    With CmoStation
        For i = 1 To Stations.Count
            .AddItem Stations(i).Name
        Next
    End With
    
    With CmoAccessLvl
        Set RstAccessLevels = ModSecurity.GetAccessLevelList
        
        For i = 1 To RstAccessLevels.RecordCount
            .AddItem RstAccessLevels!AccessLevel
            RstAccessLevels.MoveNext
        Next
    End With
    
    FormInitialise = True

Exit Function

ErrorExit:
    
    FormTerminate
    Terminate
    
    FormInitialise = False

Exit Function

ErrorHandler:
        
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        CustomErrorHandler Err.Number
        Resume Restart
    End If
    
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' UserForm_Initialize
' Automatic initialise event that triggers custom Initialise
' ---------------------------------------------------------------
Private Sub UserForm_Initialize()

    On Error Resume Next
    
    FormInitialise
    
End Sub

' ===============================================================
' FormTerminate
' Terminates the form gracefully
' ---------------------------------------------------------------
Private Function FormTerminate() As Boolean

    On Error Resume Next

    Set SelectedUser = Nothing
    
    Unload Me

End Function
