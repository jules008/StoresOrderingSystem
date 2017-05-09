VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsPersons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsPersons
' v0,0 - Initial Version
' v0,1 - change GetMailAlertUsers
' v0,2 - Return guest accounts function
'---------------------------------------------------------------
' Date - 09 May 17
'===============================================================
' Methods
'---------------------------------------------------------------
' FindItem - Finds Item from index
' AddItem - Add item to Collection
' RemoveItem - Deletes item from collection
' ItemCount - Returns number of items in collection
' GetCollection - Gets all items in collection
' DeleteCollection - Deletes all items in collection
' GetMailAlertUsers - Returns users that have mail alert selected
' ReturnUserName - Returns username from crew no
' GetGuestAccts - Returns Guest accounts from the system
'===============================================================

Option Explicit
Private pPersons As Collection

'---------------------------------------------------------------

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pPersons.[_NewEnum]
End Function

' ===============================================================
' Method FindItem
' Finds Item from index
'---------------------------------------------------------------

Public Function FindItem(CrewNo As Variant) As ClsPerson
Attribute FindItem.VB_UserMemId = 0
    On Error Resume Next
    Set FindItem = pPersons.Item(CrewNo)
End Function

' ===============================================================
' Method AddItem
' Add item to Collection
'---------------------------------------------------------------

Public Sub AddItem(Person As ClsPerson)
    Person.SetParent ObjPtr(Me)
    pPersons.Add Person, Key:=CStr(Person.CrewNo)
End Sub

' ===============================================================
' Method RemoveItem
' Deletes item from collection
'---------------------------------------------------------------

Public Sub RemoveItem(CrewNo As Variant)
    pPersons.Remove CrewNo
End Sub

' ===============================================================
' Method ItemCount
' Returns number of items in collection
'---------------------------------------------------------------

Public Function Count() As Integer
    Count = pPersons.Count
End Function

' ===============================================================
' Method GetCollection
' Gets all items in collection
'---------------------------------------------------------------

Public Sub GetCollection()
    Dim RstPerson As Recordset
    Dim Person As ClsPerson
    Dim i As Integer

    Set RstPerson = ModDatabase.SQLQuery("SELECT * FROM TblPerson WHERE Deleted IS NULL")
    With RstPerson
        .MoveLast
        .MoveFirst
        For i = 1 To .RecordCount
            Set Person = New ClsPerson
            Person.DBGet !CrewNo
            Me.AddItem Person
            .MoveNext
        Next
    End With
End Sub

' ===============================================================
' Method DeleteCollection
' Deletes all items in collection
'---------------------------------------------------------------

Public Sub DeleteCollection()
    Dim Person As ClsPerson
    For Each Person In pPersons
        pPersons.Remove CStr(Person.CrewNo)
        Person.DBDelete
    Next
End Sub

' ===============================================================
' Method GetMailAlertUsers
' Returns users that have mail alert selected
'---------------------------------------------------------------
Public Function GetMailAlertUsers() As Recordset
    Dim RstPerson As Recordset
    Dim Person As ClsPerson
    Dim i As Integer

    Set RstPerson = ModDatabase.SQLQuery("SELECT * FROM TblPerson WHERE mailalert = True")
    
    With RstPerson
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            Set GetMailAlertUsers = RstPerson
        Else
            Set GetMailAlertUsers = Nothing
        End If
    End With
End Function

' ===============================================================
' Method GetGuestAccts
' Returns Guest accounts from the system
'---------------------------------------------------------------
Public Function GetGuestAccts() As Recordset
    Dim RstPerson As Recordset
    Dim Person As ClsPerson
    Dim i As Integer

    Set RstPerson = ModDatabase.SQLQuery("SELECT * FROM TblPerson WHERE LEN (crewno) =4 AND LEFT (crewno, 2) = '99'")
    
    With RstPerson
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            Set GetGuestAccts = RstPerson
        Else
            Set GetGuestAccts = Nothing
        End If
    End With
End Function

' ===============================================================
' Method ReturnUserName
' Returns username from crew no
'---------------------------------------------------------------
Public Function ReturnUserName(CrewNo As String) As String
    Dim RstPerson As Recordset
    Dim Person As ClsPerson
    Dim i As Integer

    Set RstPerson = ModDatabase.SQLQuery("SELECT Username FROM TblPerson WHERE CrewNo = '" & CrewNo & "'")
    With RstPerson
        If .RecordCount > 0 Then ReturnUserName = .Fields(0)
    End With
End Function

' ===============================================================
Private Sub Class_Initialize()
    Set pPersons = New Collection

End Sub

Private Sub Class_Terminate()
    Set pPersons = Nothing

End Sub
'---------------------------------------------------------------


