VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsStation
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 22 Feb 17
'===============================================================
' Methods
'---------------------------------------------------------------
' DBGet - Gets class from Database
' DBSave - Saves class to Database
' DBDelete(Optional FullDelete As Boolean) - Marks record as deleted or fully deletes
'===============================================================

Option Explicit
Private pStationID As EnumStationID
Private pStationNo As String
Private pName As String
Private pAddress As String
Private pStationType As EnumStnType
Private pVehicles As ClsVehicles

'---------------------------------------------------------------
Public Property Get StationID() As EnumStationID
    StationID = pStationID
End Property

Public Property Let StationID(ByVal VnewValue As EnumStationID)
    pStationID = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get StationNo() As String
    StationNo = pStationNo
End Property

Public Property Let StationNo(ByVal VnewValue As String)
    pStationNo = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(ByVal VnewValue As String)
    pName = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get Address() As String
    Address = pAddress
End Property

Public Property Let Address(ByVal VnewValue As String)
    pAddress = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get StationType() As EnumStnType
    StationType = pStationType
End Property

Public Property Let StationType(ByVal VnewValue As EnumStnType)
    pStationType = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get Vehicles() As ClsVehicles
    Set Vehicles = pVehicles
End Property

Public Property Let Vehicles(ByVal VnewValue As ClsVehicles)
    Set pVehicles = VnewValue
End Property

'---------------------------------------------------------------
' ===============================================================
' Method DBGet
' Gets class from Database
'---------------------------------------------------------------
Public Sub DBGet(Index As Integer)
    Dim RstStation As Recordset

    Set RstStation = ModDatabase.SQLQuery("SELECT * FROM TblStation WHERE StationID = " & Index)
    With RstStation
        If .RecordCount > 0 Then
            If Not IsNull(!StationID) Then pStationID = !StationID
            If Not IsNull(!StationNo) Then pStationNo = !StationNo
            If Not IsNull(!Name) Then pName = !Name
            If Not IsNull(!Address) Then pAddress = !Address
            If Not IsNull(!StationType) Then pStationType = !StationType
            
            pVehicles.DeleteCollection
            pVehicles.GetCollection (StationID)
        End If
    End With
Set RstStation = Nothing
End Sub

' ===============================================================
Private Sub Class_Initialize()
    Set pVehicles = New ClsVehicles
    
    pVehicles.SetParent ObjPtr(Me)
End Sub

Private Sub Class_Terminate()
    pVehicles.SetParent 0
    
    Set pVehicles = Nothing
End Sub
' ===============================================================

