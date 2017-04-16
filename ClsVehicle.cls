VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsVehicle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsVehicle
' v0,0 - Initial Version
' v0,1 - Added Station ID
'---------------------------------------------------------------
' Date - 16 Apr 17
'===============================================================
' Methods
'---------------------------------------------------------------
' DBGet - Gets class from Database
' DBSave - Saves class to Database
' DBDelete(Optional FullDelete As Boolean) - Marks record as deleted or fully deletes
' GetVehicleTypeString - Returns the type of vehicle as a string from the input enum
'===============================================================

Option Explicit
Private pVehNo As Integer
Private pVehReg As String
Private pVehType As EnumVehType
Private pCallSign As String
Private pVehicleMake As String
Private pRegYear As Integer
Private pFleetNo As String
Private pDisposalDate As Date
Private pStationID As String
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

'Public Function NewEnum() As IUnknown
'    Set NewEnum = pVehicle.[_newenum]
'End Function

'---------------------------------------------------------------
Public Property Get VehNo() As Integer
    VehNo = pVehNo
End Property

Public Property Let VehNo(ByVal VnewValue As Integer)
    pVehNo = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get VehReg() As String
    VehReg = pVehReg
End Property

Public Property Let VehReg(ByVal VnewValue As String)
    pVehReg = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get VehType() As EnumVehType
    VehType = pVehType
End Property

Public Property Let VehType(ByVal VnewValue As EnumVehType)
    pVehType = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get CallSign() As String
    CallSign = pCallSign
End Property

Public Property Let CallSign(ByVal VnewValue As String)
    pCallSign = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get VehicleMake() As String
    VehicleMake = pVehicleMake
End Property

Public Property Let VehicleMake(ByVal VnewValue As String)
    pVehicleMake = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get RegYear() As Integer
    RegYear = pRegYear
End Property

Public Property Let RegYear(ByVal VnewValue As Integer)
    pRegYear = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get FleetNo() As String
    FleetNo = pFleetNo
End Property

Public Property Let FleetNo(ByVal VnewValue As String)
    pFleetNo = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get DisposalDate() As Date
    DisposalDate = pDisposalDate
End Property

Public Property Let DisposalDate(ByVal VnewValue As Date)
    pDisposalDate = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get StationID() As String
    StationID = pStationID
End Property

Public Property Let StationID(ByVal VnewValue As String)
    pStationID = VnewValue
End Property

'---------------------------------------------------------------
' ===============================================================
' Method DBGet
' Gets class from Database
'---------------------------------------------------------------
Public Sub DBGet(VehReg As String)
    Dim RstVehicle As Recordset

    Set RstVehicle = ModDatabase.SQLQuery("SELECT * FROM TblVehicle WHERE Vehno = " & VehReg)
    With RstVehicle
        If .RecordCount > 0 Then
            If Not IsNull(!VehNo) Then pVehNo = !VehNo
            If Not IsNull(!VehReg) Then pVehReg = !VehReg
            If Not IsNull(!VehType) Then pVehType = !VehType
            If Not IsNull(!CallSign) Then pCallSign = !CallSign
            If Not IsNull(!VehicleMake) Then pVehicleMake = !VehicleMake
            If Not IsNull(!RegYear) Then pRegYear = !RegYear
            If Not IsNull(!FleetNo) Then pFleetNo = !FleetNo
            If Not IsNull(!DisposalDate) Then pDisposalDate = !DisposalDate
            If Not IsNull(!StationID) Then pStationID = !StationID
            
        End If
    End With
Set RstVehicle = Nothing
End Sub

' ===============================================================
' Method DBSave
' Saves class to Database
'---------------------------------------------------------------
Public Sub DBSave()
    Dim RstVehicle As Recordset
    Dim RstMaxNo As Recordset
    Dim LastNo As Integer

    Set RstVehicle = ModDatabase.SQLQuery("SELECT * FROM TblVehicle WHERE VehReg = " & pVehReg & " AND Deleted IS NULL")
    Set RstMaxNo = ModDatabase.SQLQuery("SELECT MAX(VehReg) FROM TblVehicle ")

    If RstMaxNo.Fields(0).Value <> 0 Then
        LastNo = RstMaxNo.Fields(0).Value
    Else
        LastNo = 0
    End If

    With RstVehicle
        If .RecordCount = 0 Then
            .addnew
            pVehNo = LastNo + 1
        Else
            .Edit
        End If
        !VehNo = pVehNo
        !VehReg = pVehReg
        !VehType = pVehType
        !CallSign = pCallSign
        !VehicleMake = pVehicleMake
        !RegYear = pRegYear
        !FleetNo = pFleetNo
        .Update
        
    End With
    Set RstVehicle = Nothing
    Set RstMaxNo = Nothing
End Sub

' ===============================================================
' Method DBDelete(Optional FullDelete As Boolean)
' Marks record as deleted or fully deletes
'---------------------------------------------------------------
Public Sub DBDelete(Optional FullDelete As Boolean)
    Dim RstVehicle As Recordset
    Dim i As Integer

    Set RstVehicle = ModDatabase.SQLQuery("SELECT * FROM TblVehicle WHERE Vehreg = '" & pVehReg & "'")
    With RstVehicle
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

    Set RstVehicle = Nothing
End Sub

' ===============================================================
' Method GetVehicleTypeString
' Returns the type of vehicle as a string from the input enum
' ---------------------------------------------------------------
Public Function GetVehicleTypeString() As String
    Dim RstVehicle As Recordset

    Set RstVehicle = ModDatabase.SQLQuery("SELECT VehicleType FROM TblVehicleType WHERE VehicleTypeNo = " & pVehType)
    With RstVehicle
        If .RecordCount > 0 Then
            GetVehicleTypeString = .Fields(0)
        End If
    End With
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
