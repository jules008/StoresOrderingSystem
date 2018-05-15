VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsStations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsStations
' v0,0 - Initial Version
' v0,1 - Sort Get Station by StationNo
'---------------------------------------------------------------
' Date - 14 May 18
'===============================================================
' Methods
'---------------------------------------------------------------
' FindItem - Finds Item from index
' AddItem - Add item to Collection
' RemoveItem - Deletes item from collection
' ItemCount - Returns number of items in collection
' GetCollection - Gets all items in collection
' DeleteCollection - Deletes all items in collection
'===============================================================

Option Explicit
Private pStations As Collection

'---------------------------------------------------------------

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pStations.[_NewEnum]
End Function

' ===============================================================
' Method FindItem
' Finds Item from index
'---------------------------------------------------------------

Public Function FindItem(StationID As Variant) As ClsStation
Attribute FindItem.VB_UserMemId = 0
    On Error Resume Next
    Set FindItem = pStations.Item(StationID)
End Function

' ===============================================================
' Method AddItem
' Add item to Collection
'---------------------------------------------------------------

Public Sub AddItem(Station As ClsStation)
    pStations.Add Station, Key:=CStr(Station.StationID)
End Sub

' ===============================================================
' Method RemoveItem
' Deletes item from collection
'---------------------------------------------------------------

Public Sub RemoveItem(StationNo As Variant)
    pStations.Remove StationNo
End Sub

' ===============================================================
' Method ItemCount
' Returns number of items in collection
'---------------------------------------------------------------

Public Function Count() As Integer
    Count = pStations.Count
End Function

' ===============================================================
' Method GetCollection
' Gets all items in collection
'---------------------------------------------------------------

Public Sub GetCollection()
    Dim RstStation As Recordset
    Dim Station As ClsStation
    Dim i As Integer

    Set RstStation = ModDatabase.SQLQuery("SELECT * FROM TblStation ORDER BY Name ASC")
    With RstStation
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                Set Station = New ClsStation
                Station.DBGet !StationID
                Me.AddItem Station
                .MoveNext
            Loop
        End If
    End With
End Sub

' ===============================================================
Private Sub Class_Initialize()
    Set pStations = New Collection

End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pStations = Nothing

End Sub

'---------------------------------------------------------------
