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
' V0,1 - Add Station Active Field
'---------------------------------------------------------------
' Date - 10 May 18
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
Private pStnActive As Boolean
Private pVehicles As ClsVehicles

'---------------------------------------------------------------
Public Property Get StationID() As EnumStationID
    StationID = pStationID
End Property

Public Property Let StationID(ByVal vNewValue As EnumStationID)
    pStationID = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get StationNo() As String
    StationNo = pStationNo
End Property

Public Property Let StationNo(ByVal vNewValue As String)
    pStationNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(ByVal vNewValue As String)
    pName = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Address() As String
    Address = pAddress
End Property

Public Property Let Address(ByVal vNewValue As String)
    pAddress = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get StationType() As EnumStnType
    StationType = pStationType
End Property

Public Property Let StationType(ByVal vNewValue As EnumStnType)
    pStationType = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get StnActive() As Boolean
    StnActive = pStnActive
End Property

Public Property Let StnActive(ByVal vNewValue As Boolean)
    pStnActive = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Vehicles() As ClsVehicles
    Set Vehicles = pVehicles
End Property

Public Property Let Vehicles(ByVal vNewValue As ClsVehicles)
    Set pVehicles = vNewValue
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
            If Not IsNull(!StnActive) Then pStnActive = !StnActive

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

