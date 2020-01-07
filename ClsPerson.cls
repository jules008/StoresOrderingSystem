VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsPerson
' v0,0 - Initial Version
' v0,1 - Changes to Username functionality
' v0,2 - Make Username read/write
' v0,3 - Added Message Read Flag
' v0,41 - Log on and Log Off from DB
' v0,5 - Fixed bug preventing Alford being selected
' v0,6 - Changed LastNo to Long to prevent overflows
'---------------------------------------------------------------
' Date - 07 Jan 20
'===============================================================
' Methods
'---------------------------------------------------------------
' DBGet - Gets class from Database
' DBSave - Saves class to Database
' DBDelete(Optional FullDelete As Boolean) - Marks record as deleted or fully deletes
' GetUserName - Gets Username from CrewNo
' AddTempAccount - adds a temporary account in the database if user not recognised
'===============================================================

Option Explicit
Private pDeleted As Date
Private pCrewNo As String
Private pForename As String
Private pSurname As String
Private pUsername As String
Private pRankGrade As String
Private pWatch As String
Private pMailAlert As Boolean
Private pRole As String
Private pAccessLvl As EnumAccessLvl
Private pStation As ClsStation
Private pVehicle As ClsVehicle
Private pMessageRead As Boolean
Private pParent As Long

'---------------------------------------------------------------
Public Property Get Deleted() As Date
    Deleted = pDeleted
End Property

Public Property Let Deleted(ByVal vNewValue As Date)
    pDeleted = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get CrewNo() As String
    CrewNo = pCrewNo
End Property

Public Property Let CrewNo(ByVal vNewValue As String)
    pCrewNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Forename() As String
    Forename = pForename
End Property

Public Property Let Forename(ByVal vNewValue As String)
    pForename = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Surname() As String
    Surname = pSurname
End Property

Public Property Let Surname(ByVal vNewValue As String)
    pSurname = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get UserName() As String
    UserName = pUsername
End Property

Public Property Let UserName(ByVal vNewValue As String)
    pUsername = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get RankGrade() As String
    RankGrade = pRankGrade
End Property

Public Property Let RankGrade(ByVal vNewValue As String)
    pRankGrade = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Watch() As String
    Watch = pWatch
End Property

Public Property Let Watch(ByVal vNewValue As String)
    pWatch = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get MailAlert() As Boolean
    MailAlert = pMailAlert
End Property

Public Property Let MailAlert(ByVal vNewValue As Boolean)
    pMailAlert = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Role() As String
    Role = pRole
End Property

Public Property Let Role(ByVal vNewValue As String)
    pRole = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get AccessLvl() As EnumAccessLvl
    AccessLvl = pAccessLvl
End Property

Public Property Let AccessLvl(ByVal vNewValue As EnumAccessLvl)
    pAccessLvl = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Station() As ClsStation
    Set Station = pStation
End Property

Public Property Let Station(ByVal vNewValue As ClsStation)
    Set pStation = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Vehicle() As ClsVehicle
    Set Vehicle = pVehicle
End Property

Public Property Let Vehicle(ByVal vNewValue As ClsVehicle)
    Set pVehicle = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get MessageRead() As Boolean
   MessageRead = pMessageRead
End Property

Public Property Let MessageRead(ByVal vNewValue As Boolean)
    pMessageRead = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Parent() As ClsOrder
    If pParent <> 0 Then
        Set Parent = GetParentFromPtr(pParent)
    End If
End Property

Friend Function SetParent(ByVal Ptr As Long) As Boolean
    pParent = Ptr
End Function

' ===============================================================
' Method DBGet
' Gets class from Database
'---------------------------------------------------------------
Public Sub DBGet(UserNameOrCrewNo As String)
    Dim RstPerson As Recordset
    Dim UserName As String

    If Not IsNumeric(UserNameOrCrewNo) Then
        Set RstPerson = ModDatabase.SQLQuery("SELECT * FROM TblPerson WHERE Username = '" & UserNameOrCrewNo & "' AND Deleted IS NULL")
    Else
        Set RstPerson = ModDatabase.SQLQuery("SELECT * FROM TblPerson WHERE CrewNo = '" & UserNameOrCrewNo & "' AND Deleted IS NULL")
    End If
    
    With RstPerson
        If .RecordCount > 0 Then
            If Not IsNull(!Deleted) Then pDeleted = !Deleted
            If Not IsNull(!CrewNo) Then pCrewNo = !CrewNo
            If Not IsNull(!Forename) Then pForename = !Forename
            If Not IsNull(!Surname) Then pSurname = !Surname
            If Not IsNull(!UserName) Then pUsername = !UserName
            If Not IsNull(!RankGrade) Then pRankGrade = !RankGrade
            If Not IsNull(!Watch) Then pWatch = !Watch
            If Not IsNull(!Role) Then pRole = !Role
            If Not IsNull(!MailAlert) Then pMailAlert = !MailAlert
            If Not IsNull(!AccessLvl) Then pAccessLvl = !AccessLvl
            If Not IsNull(!StationID) Then pStation.DBGet (!StationID)
            If Not IsNull(!MessageRead) Then pMessageRead = !MessageRead
            
            If Not IsNull(!VehicleID) Then
                Set pVehicle = New ClsVehicle
                pVehicle.DBGet (!VehicleID)
            End If
            
        End If
    End With
Set RstPerson = Nothing
End Sub

' ===============================================================
' Method DBSave
' Saves class to Database
'---------------------------------------------------------------
Public Sub DBSave()
    Dim RstPerson As Recordset
    Dim RstMaxNo As Recordset
    Dim LastNo As Long

    Set RstPerson = ModDatabase.SQLQuery("SELECT * FROM TblPerson WHERE crewno = '" & pCrewNo & "' AND Deleted IS NULL")

    
    With RstPerson
        If .RecordCount = 0 Then
            .AddNew
        Else
            .Edit
        End If
        If pCrewNo <> "" Then !CrewNo = pCrewNo
        If pForename <> "" Then !Forename = pForename
        If pSurname <> "" Then !Surname = pSurname
        If pUsername <> "" Then !UserName = pUsername Else !UserName = pForename & " " & pSurname
        If pRankGrade <> "" Then !RankGrade = pRankGrade
        If pRole <> "" Then !Role = pRole
        If pWatch <> "" Then !Watch = pWatch
        !MailAlert = pMailAlert
        !MessageRead = pMessageRead
        If pAccessLvl <> 0 Then !AccessLvl = pAccessLvl
        !StationID = pStation.StationID
        .Update

    End With
    Set RstPerson = Nothing
    Set RstMaxNo = Nothing
End Sub

' ===============================================================
' Method DBDelete(Optional FullDelete As Boolean)
' Marks record as deleted or fully deletes
'---------------------------------------------------------------
Public Sub DBDelete(Optional FullDelete As Boolean)
    Dim RstPerson As Recordset
    Dim i As Integer

    Set RstPerson = ModDatabase.SQLQuery("SELECT * FROM TblPerson WHERE crewno = '" & pCrewNo & "' AND Deleted IS NULL")
    With RstPerson
        For i = .RecordCount To 1 Step -1
            If FullDelete Then
                .Delete
                .MoveNext
            Else
                .Edit
                !Deleted = Now
                .Update
            End If
        Next
    End With

    Set RstPerson = Nothing
End Sub

' ===============================================================
' Method AddTempAccount
' adds a temporary account in the database if user not recognised
'---------------------------------------------------------------
Public Sub AddTempAccount()
    Dim LocName As String
    Dim Persons As ClsPersons
    
    Dim i As Integer
    
    Set Persons = New ClsPersons
    
    If TEST_MODE Then
        If ShtSettings.Range("C15") = True Then
            pUsername = ShtSettings.Range("Test_User")
        Else
            pUsername = Application.UserName
        End If
    Else
        pUsername = Application.UserName
    End If
    
    If pUsername = "" Then Err.Raise NO_USER_AVAILABLE, Description:="The system cannot find a username from the application"

    For i = 9999 To 1 Step -1
        pCrewNo = i
        LocName = Persons.ReturnUserName(CStr(i))
        If LocName = "" Then Exit For
    Next
    
    pAccessLvl = BasicLvl_1
    
    Set Persons = Nothing
    
End Sub

' ===============================================================
' Method LogOn
' Adds the user to the logged on table in DB
' ---------------------------------------------------------------
Public Sub LogOn()
    Dim RstUser As Recordset
    
    On Error GoTo ErrorHandler
    
    Set RstUser = SQLQuery("TblUsers")
    
    If RstUser.RecordCount = 0 Then
        DB.Execute "INSERT INTO TblUsers VALUES ('" & pUsername & "', '" & Now & "')"
    End If
Exit Sub

ErrorHandler:
    Err.Raise HANDLED_ERROR, , "Unable to add user to user list"
    Exit Sub
End Sub

' ===============================================================
' Method LogUserOff
' Removes the user from the logged on table in DB
' ---------------------------------------------------------------
Public Sub LogUserOff()

    On Error GoTo ErrorHandler
    
    DB.Execute "DELETE FROM TblUsers WHERE CurrentUsers = '" & pUsername & "'"
Exit Sub

ErrorHandler:
    Err.Raise HANDLED_ERROR, , "Unable to delete user from user list"
    Exit Sub
End Sub

' ===============================================================
' Method GetParentFromPtr
' Private routine to copy memory address of parent class
' ---------------------------------------------------------------
Private Function GetParentFromPtr(ByVal Ptr As Long) As ClsOrder
    Dim tmp As ClsOrder

    CopyMemory tmp, Ptr, 4
    Set GetParentFromPtr = tmp
    CopyMemory tmp, 0&, 4
End Function

' ===============================================================
Private Sub Class_Initialize()
    Set pStation = New ClsStation

End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pStation = Nothing
End Sub

'---------------------------------------------------------------


