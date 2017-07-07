VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsDelivery
' v0,0 - Initial Version
' v0,1 - Change SupplierName to SupplierID
'---------------------------------------------------------------
' Date - 07 Jul 17
'===============================================================
' Methods
'---------------------------------------------------------------
' DBGet - Gets class from Database
' DBSave - Saves class to Database
' DBDelete() - fully deletes
'===============================================================

Option Explicit
Private pDeleted As Date
Private pDeliveryNo As Integer
Private pSupplierID As Integer
Private pDeliveryDate As Date
Private pAssetNo As Integer
Private pAssetDescr As String
Private pQuantity As Integer

'---------------------------------------------------------------
Public Property Get Deleted() As Date
    Deleted = pDeleted
End Property

Public Property Let Deleted(ByVal vNewValue As Date)
    pDeleted = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get DeliveryNo() As Integer
    DeliveryNo = pDeliveryNo
End Property

Public Property Let DeliveryNo(ByVal vNewValue As Integer)
    pDeliveryNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get SupplierID() As Integer
    SupplierID = pSupplierID
End Property

Public Property Let SupplierID(ByVal vNewValue As Integer)
    pSupplierID = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get DeliveryDate() As Date
    DeliveryDate = pDeliveryDate
End Property

Public Property Let DeliveryDate(ByVal vNewValue As Date)
    pDeliveryDate = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get AssetNo() As Integer
    AssetNo = pAssetNo
End Property

Public Property Let AssetNo(ByVal vNewValue As Integer)
    pAssetNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Quantity() As Integer
    Quantity = pQuantity
End Property

Public Property Let Quantity(ByVal vNewValue As Integer)
    pQuantity = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get AssetDescr() As String
    AssetDescr = pAssetDescr
End Property

Public Property Let AssetDescr(ByVal vNewValue As String)
    pAssetDescr = vNewValue
End Property

'---------------------------------------------------------------
' ===============================================================
' Method DBGet
' Gets class from Database
'---------------------------------------------------------------
Public Sub DBGet()
    Dim RstDelivery As Recordset

    Set RstDelivery = ModDatabase.SQLQuery("SELECT * FROM TblDelivery WHERE DeliveryNo = " & pDeliveryNo)
    With RstDelivery
        If .RecordCount > 0 Then
            If Not IsNull(!DeliveryNo) Then pDeliveryNo = !DeliveryNo
            If Not IsNull(!SupplierID) Then pSupplierID = !SupplierID
            If Not IsNull(!DeliveryDate) Then pDeliveryDate = !DeliveryDate
            If Not IsNull(!AssetNo) Then pAssetNo = !AssetNo
            If Not IsNull(!Quantity) Then pQuantity = !Quantity

        End If
    End With
Set RstDelivery = Nothing
End Sub

' ===============================================================
' Method DBSave
' Saves class to Database
'---------------------------------------------------------------
Public Sub DBSave()
    Dim RstDelivery As Recordset
    Dim RstMaxNo As Recordset
    Dim LastNo As Integer

    Set RstDelivery = ModDatabase.SQLQuery("SELECT * FROM TblDelivery WHERE DeliveryNo = " & pDeliveryNo)
    Set RstMaxNo = ModDatabase.SQLQuery("SELECT MAX(DeliveryNo) FROM TblDelivery")

    If RstMaxNo.Fields(0).Value <> 0 Then
        LastNo = RstMaxNo.Fields(0).Value
    Else
        LastNo = 0
    End If

    With RstDelivery
        If .RecordCount = 0 Then
            .addnew
            pDeliveryNo = LastNo + 1
        Else
            .Edit
        End If
        !DeliveryNo = pDeliveryNo
        !SupplierID = pSupplierID
        !DeliveryDate = pDeliveryDate
        !AssetNo = pAssetNo
        !Quantity = pQuantity
        .Update

    End With
    Set RstDelivery = Nothing
    Set RstMaxNo = Nothing
End Sub

' ===============================================================
' Method DBDelete()
' fully deletes
'---------------------------------------------------------------
Public Sub DBDelete(Optional FullDelete As Boolean)
    Dim RstDelivery As Recordset
    Dim i As Integer

    Set RstDelivery = ModDatabase.SQLQuery("SELECT * FROM TblDelivery WHERE DeliveryNo = " & pDeliveryNo)
    With RstDelivery
        For i = .RecordCount To 1 Step -1
            .Delete
            .MoveNext
        Next
    End With

    Set RstDelivery = Nothing
End Sub

