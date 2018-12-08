VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsLineItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsLineitem
' v0,0 - Initial Version
' v0,1 - Save Asset with Line item
' v0,2 - Force Null if allocation type is not station
' v0,3 - Removed Deleted filter from DBSave and cascaded delete to LossReports
'---------------------------------------------------------------
' Date - 02 Oct 17
'===============================================================
' Methods
'---------------------------------------------------------------
' DBGet - Gets class from Database
' DBSave - Saves class to Database
' DBDelete(Optional FullDelete As Boolean) - Marks record as deleted or fully deletes
' ReturnLineitemStatus - Returns string value from enumLineitemStatus
'===============================================================

Option Explicit
Private pDeleted As Date
Private pLineitemNo As Integer
Private pAsset As ClsAsset
Private pLossReport As ClsLossReport
Private pQuantity As Integer
Private pReqReason As EnumReqReason
Private pReturnReqd As Boolean
Private pForStation As ClsStation
Private pForVehicle As ClsVehicle
Private pForPerson As ClsPerson
Private pStatus As EnumLineitemStatus
Private pitemsReturned As Boolean
Private pItemsDelivered As Boolean
Private pItemsIssued As Boolean
Private pOnHoldReason As String
Private pParent As Long

'---------------------------------------------------------------
Public Property Get Deleted() As Date
    Deleted = pDeleted
End Property

Public Property Let Deleted(ByVal vNewValue As Date)
    pDeleted = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get LineitemNo() As Integer
    LineitemNo = pLineitemNo
End Property

Public Property Let LineitemNo(ByVal vNewValue As Integer)
    pLineitemNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Asset() As ClsAsset
    Set Asset = pAsset
End Property

Public Property Let Asset(ByVal vNewValue As ClsAsset)
    Set pAsset = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get LossReport() As ClsLossReport
    Set LossReport = pLossReport
End Property

Public Property Let LossReport(ByVal vNewValue As ClsLossReport)
    Set pLossReport = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Quantity() As Integer
    Quantity = pQuantity
End Property

Public Property Let Quantity(ByVal vNewValue As Integer)
    pQuantity = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get ReqReason() As EnumReqReason
    ReqReason = pReqReason
End Property

Public Property Let ReqReason(ByVal vNewValue As EnumReqReason)
    pReqReason = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get ReturnReqd() As Boolean
    ReturnReqd = pReturnReqd
End Property

Public Property Let ReturnReqd(ByVal vNewValue As Boolean)
    pReturnReqd = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get ForStation() As ClsStation
    Set ForStation = pForStation
End Property

Public Property Let ForStation(ByVal vNewValue As ClsStation)
    Set pForStation = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get ForVehicle() As ClsVehicle
    Set ForVehicle = pForVehicle
End Property

Public Property Let ForVehicle(ByVal vNewValue As ClsVehicle)
    Set pForVehicle = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get ForPerson() As ClsPerson
    Set ForPerson = pForPerson
End Property

Public Property Let ForPerson(ByVal vNewValue As ClsPerson)
    Set pForPerson = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Status() As EnumLineitemStatus
    Status = pStatus
End Property

Public Property Let Status(ByVal vNewValue As EnumLineitemStatus)
    pStatus = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get itemsReturned() As Boolean
    itemsReturned = pitemsReturned
End Property

Public Property Let itemsReturned(ByVal vNewValue As Boolean)
    pitemsReturned = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get ItemsDelivered() As Boolean
    ItemsDelivered = pItemsDelivered
End Property

Public Property Let ItemsDelivered(ByVal vNewValue As Boolean)
    pItemsDelivered = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get ItemsIssued() As Boolean
    ItemsIssued = pItemsIssued
End Property

Public Property Let ItemsIssued(ByVal vNewValue As Boolean)
    pItemsIssued = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get OnHoldReason() As String
    OnHoldReason = pOnHoldReason
End Property

Public Property Let OnHoldReason(ByVal vNewValue As String)
    pOnHoldReason = vNewValue
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
Public Sub DBGet()
    Dim RstLineitem As Recordset

    Set RstLineitem = ModDatabase.SQLQuery("SELECT * FROM TblLineitem WHERE LineitemNo = " & pLineitemNo & " AND Deleted IS NULL")
    With RstLineitem
        If .RecordCount > 0 Then
            If Not IsNull(!Deleted) Then pDeleted = !Deleted
            If Not IsNull(!LineitemNo) Then pLineitemNo = !LineitemNo
            If Not IsNull(!Quantity) Then pQuantity = !Quantity
            If Not IsNull(!ReqReason) Then pReqReason = !ReqReason
            If Not IsNull(!ReturnReqd) Then pReturnReqd = !ReturnReqd
            If Not IsNull(!Status) Then pStatus = !Status
            If Not IsNull(!itemsReturned) Then pitemsReturned = !itemsReturned
            If Not IsNull(!ItemsDelivered) Then pItemsDelivered = !ItemsDelivered
            If Not IsNull(!ItemsIssued) Then pItemsIssued = !ItemsIssued
            If Not IsNull(!OnHoldReason) Then pOnHoldReason = !OnHoldReason

            If !LossReportID <> 0 Then pLossReport.DBGet !LossReportID
            If !ForStationID <> 0 Then pForStation.DBGet !ForStationID
            If !ForVehicleID <> 0 Then pForVehicle.DBGet !ForVehicleID
            If !ForPersonID <> 0 Then pForPerson.DBGet !ForPersonID
            If !AssetID <> 0 Then pAsset.DBGet !AssetID
        End If
    End With
Set RstLineitem = Nothing
End Sub

' ===============================================================
' Method DBSave
' Saves class to Database
'---------------------------------------------------------------
Public Sub DBSave()
    Dim RstLineitem As Recordset
    Dim RstMaxNo As Recordset
    Dim LastNo As Integer

    Set RstLineitem = ModDatabase.SQLQuery("SELECT * FROM TblLineitem WHERE LineitemNo = " & pLineitemNo)
    Set RstMaxNo = ModDatabase.SQLQuery("SELECT MAX(LineitemNo) FROM TblLineitem ")

    If RstMaxNo.Fields(0).Value <> 0 Then
        LastNo = RstMaxNo.Fields(0).Value
    Else
        LastNo = 0
    End If

    With RstLineitem
        If .RecordCount = 0 Then
            .AddNew
            pLineitemNo = LastNo + 1
        Else
            .Edit
        End If
        
        If pLossReport.LossReportNo <> 0 Then pLossReport.DBSave
        
        If Not Parent Is Nothing Then !OrderNo = Parent.OrderNo
        !LineitemNo = pLineitemNo
        !AssetID = pAsset.AssetNo
        !LossReportID = pLossReport.LossReportNo
        !Quantity = pQuantity
        !ReqReason = pReqReason
        !ReturnReqd = pReturnReqd
        !ForVehicleID = pForVehicle.VehNo
        !ForPersonID = pForPerson.CrewNo
        !Status = pStatus
        !itemsReturned = pitemsReturned
        !ItemsDelivered = pItemsDelivered
        !ItemsIssued = pItemsIssued
        !OnHoldReason = pOnHoldReason
        
        If pAsset.AllocationType = Station Then
            !ForStationID = pForStation.StationID
        Else
            !ForStationID = Null
        End If
        .Update
   
        pAsset.DBSave
        
    End With
    Set RstLineitem = Nothing
    Set RstMaxNo = Nothing
End Sub

' ===============================================================
' Method DBDelete(Optional FullDelete As Boolean)
' Marks record as deleted or fully deletes
'---------------------------------------------------------------
Public Sub DBDelete(Optional FullDelete As Boolean)
    Dim RstLineitem As Recordset
    
    Dim i As Integer

    Set RstLineitem = ModDatabase.SQLQuery("SELECT * FROM TblLineitem WHERE LineitemNo = " & pLineitemNo & " AND Deleted IS NULL")
    With RstLineitem
        For i = .RecordCount To 1 Step -1
            If FullDelete Then
                .Delete
                .MoveNext
                LossReport.DBDelete True
            Else
                .Edit
                !Deleted = Now
                .Update
                LossReport.DBDelete
            End If
        Next
    End With

    Set RstLineitem = Nothing
End Sub

' ===============================================================
' ReturnLineitemStatus
' Returns string value from enumLineitemStatus
' ---------------------------------------------------------------
Public Function ReturnLineitemStatus() As String
    Const StrPROCEDURE As String = "ReturnLineitemStatus()"
    Dim RstLineitemStatus As Recordset
    
    Set RstLineitemStatus = ModDatabase.SQLQuery("SELECT LineitemStatus FROM TblLineitemStatus WHERE LineitemStatusNo = " & pStatus)
    
    With RstLineitemStatus
        If .RecordCount > 0 Then ReturnLineitemStatus = !LineitemStatus
    
    End With
    
    Set RstLineitemStatus = Nothing

End Function

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
    Set pAsset = New ClsAsset
    Set pLossReport = New ClsLossReport
    Set pForStation = New ClsStation
    Set pForVehicle = New ClsVehicle
    Set pForPerson = New ClsPerson

    pLossReport.SetParent ObjPtr(Me)
End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
    pLossReport.SetParent 0
    
    Set pAsset = Nothing
    Set pLossReport = Nothing
    Set pForStation = Nothing
    Set pForVehicle = Nothing
    Set pForPerson = Nothing

End Sub

'---------------------------------------------------------------
