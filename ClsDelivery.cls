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
' V0,2 - Re-Write for Supplier Functionality
'---------------------------------------------------------------
' Date - 10 Jul 17
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
Private pSupplier As ClsSupplier
Private pDeliveryDate As Date
Private pAsset As ClsAsset
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
Public Property Get Supplier() As ClsSupplier
    Set Supplier = pSupplier
End Property

Public Property Let Supplier(ByVal vNewValue As ClsSupplier)
    Set pSupplier = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get DeliveryDate() As Date
    DeliveryDate = pDeliveryDate
End Property

Public Property Let DeliveryDate(ByVal vNewValue As Date)
    pDeliveryDate = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Asset() As ClsAsset
    Set Asset = pAsset
End Property

Public Property Let Asset(ByVal vNewValue As ClsAsset)
    Set pAsset = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Quantity() As Integer
    Quantity = pQuantity
End Property

Public Property Let Quantity(ByVal vNewValue As Integer)
    pQuantity = vNewValue
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
            If Not IsNull(!DeliveryDate) Then pDeliveryDate = !DeliveryDate
            If Not IsNull(!Quantity) Then pQuantity = !Quantity
            
            pAsset.DBGet !AssetNo
            pSupplier.DBGet !SupplierID
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
            .AddNew
            pDeliveryNo = LastNo + 1
        Else
            .Edit
        End If
        !DeliveryNo = pDeliveryNo
        !SupplierID = pSupplier.SupplierID
        !DeliveryDate = pDeliveryDate
        !AssetNo = pAsset.AssetNo
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

' ===============================================================
Private Sub Class_Initialize()
    Set pAsset = New ClsAsset
    Set pSupplier = New ClsSupplier
End Sub

Private Sub Class_Terminate()
    Set pAsset = Nothing
    Set pSupplier = Nothing
End Sub
' ===============================================================

