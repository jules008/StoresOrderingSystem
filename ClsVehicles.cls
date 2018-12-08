VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsVehicles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsVehicles
' v0,0 - Initial Version
' v0,1 - Added GetVehicleType 
'---------------------------------------------------------------
' Date - 08 Dec 18
'===============================================================
' Methods
'---------------------------------------------------------------
' FindItem - Finds Item from index
' AddItem - Add item to Collection
' RemoveItem - Deletes item from collection
' ItemCount - Returns number of items in collection
' GetCollection - Gets all items in collection
' DeleteCollection - Deletes all items in collection
' GetVehicleTypes - Returns a list of vehicle types in a recordset
'===============================================================

Option Explicit
Private pVehicles As Collection
Private pParent As Long

'---------------------------------------------------------------
Public Property Get Parent() As ClsStation
    If pParent <> 0 Then
        Set Parent = GetParentFromPtr(pParent)
    End If
End Property

Friend Function SetParent(ByVal Ptr As Long) As Boolean
    pParent = Ptr
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pVehicles.[_NewEnum]
End Function

' ===============================================================
' Method FindItem
' Finds Item from index
'---------------------------------------------------------------

Public Function FindItem(VehNo As Variant) As ClsVehicle
Attribute FindItem.VB_UserMemId = 0
    On Error Resume Next
    Set FindItem = pVehicles.Item(VehNo)
End Function

' ===============================================================
' Method AddItem
' Add item to Collection
'---------------------------------------------------------------

Public Sub AddItem(Vehicle As ClsVehicle)
    Vehicle.SetParent ObjPtr(Me.Parent)
    pVehicles.Add Vehicle, Key:=CStr(Vehicle.VehNo)
End Sub

' ===============================================================
' Method RemoveItem
' Deletes item from collection
'---------------------------------------------------------------

Public Sub RemoveItem(VehNo As Variant)
    pVehicles.Remove VehNo
End Sub

' ===============================================================
' Method ItemCount
' Returns number of items in collection
'---------------------------------------------------------------

Public Function Count() As Integer
    Count = pVehicles.Count
End Function

' ===============================================================
' Method GetCollection
' Gets all items in collection
'---------------------------------------------------------------

Public Sub GetCollection(Optional StationNo As EnumStationID)
    Dim RstVehicle As Recordset
    Dim StrSearch As String
    Dim Vehicle As ClsVehicle
    Dim i As Integer
    
    If StationNo <> 0 Then
        StrSearch = "SELECT * FROM TblVehicle WHERE StationID = " & StationNo & _
                    " ORDER BY callsign ASC"
    Else
        StrSearch = "SELECT * FROM TblVehicle ORDER BY callsign ASC"
    End If
    
    Set RstVehicle = ModDatabase.SQLQuery(StrSearch)
    
    With RstVehicle
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                Set Vehicle = New ClsVehicle
                Vehicle.DBGet (!VehNo)
                Me.AddItem Vehicle
                .MoveNext
            Loop
        End If
    End With
    Set RstVehicle = Nothing
End Sub

' ===============================================================
' Method DeleteCollection
' Deletes all items in collection
'---------------------------------------------------------------
Public Sub DeleteCollection()
    Dim Vehicle As ClsVehicle
    For Each Vehicle In pVehicles
        pVehicles.Remove CStr(Vehicle.VehNo)
    Next
End Sub

' ===============================================================
' Method GetVehicleTypes
' Returns a list of vehicle types in a recordset
'---------------------------------------------------------------
Public Function GetVehicleTypes() As Recordset
    Dim RstVehTypes As Recordset
    
    Set RstVehTypes = ModDatabase.SQLQuery("TblVehicleType")
    
    Set GetVehicleTypes = RstVehTypes
    
    Set RstVehTypes = Nothing
    
End Function

' ===============================================================
' Method GetVehicleType
' Returns a single vehicle type as a string
'---------------------------------------------------------------
Public Function GetVehicleType(VehTypeNo) As String
    Dim RstVehType As Recordset
    
    Set RstVehType = ModDatabase.SQLQuery("SELECT VehicleType FROM TblVehicleType WHERE VehicleTypeNo = " & VehTypeNo)
    
    GetVehicleType = RstVehType.Fields(0)
    
    Set RstVehType = Nothing
    
End Function

' ===============================================================
' Method GetParentFromPtr
' Private routine to copy memory address of parent class
' ---------------------------------------------------------------
Private Function GetParentFromPtr(ByVal Ptr As Long) As ClsStation
    Dim tmp As ClsStation

    CopyMemory tmp, Ptr, 4
    Set GetParentFromPtr = tmp
    CopyMemory tmp, 0&, 4
End Function

' ===============================================================
Private Sub Class_Initialize()
    Set pVehicles = New Collection
End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pVehicles = Nothing

End Sub

'---------------------------------------------------------------

