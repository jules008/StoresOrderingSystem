VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsDeliveries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsDeliveries
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 22 May 17
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
Private pDeliveries As Collection

'---------------------------------------------------------------

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pDeliveries.[_newenum]
End Function

' ===============================================================
' Method FindItem
' Finds Item from index
'---------------------------------------------------------------

Public Function FindItem(DeliveryNo As Variant) As ClsDelivery
Attribute FindItem.VB_UserMemId = 0
    On Error Resume Next
    Set FindItem = pDeliveries.Item(DeliveryNo)
End Function

' ===============================================================
' Method AddItem
' Add item to Collection
'---------------------------------------------------------------

Public Sub AddItem(Delivery As ClsDelivery)
    pDeliveries.Add Delivery, Key:=CStr(Delivery.DeliveryNo)
End Sub

' ===============================================================
' Method RemoveItem
' Deletes item from collection
'---------------------------------------------------------------

Public Sub RemoveItem(DeliveryNo As Variant)
    pDeliveries.Remove DeliveryNo
End Sub

' ===============================================================
' Method ItemCount
' Returns number of items in collection
'---------------------------------------------------------------

Public Function Count() As Integer
    Count = pDeliveries.Count
End Function

' ===============================================================
' Method GetCollection
' Gets all items in collection
'---------------------------------------------------------------

Public Sub GetCollection()
    Dim RstDelivery As Recordset
    Dim Delivery As ClsDelivery
    Dim i As Integer

    Set RstDelivery = ModDatabase.SQLQuery("SELECT * FROM TblDelivery")
    With RstDelivery
        .MoveLast
        .MoveFirst
        For i = 1 To .recordcount
            Set Delivery = New ClsDelivery
            Delivery.DeliveryNo = !DeliveryNo
            Delivery.DBGet
            Me.AddItem Delivery
            .Movenext
        Next
    End With
End Sub

' ===============================================================
' Method DeleteCollection
' Deletes all items in collection
'---------------------------------------------------------------

Public Sub DeleteCollection()
    Dim Delivery As ClsDelivery
    For Each Delivery In pDeliveries
        pDeliveries.Remove CStr(Delivery.DeliveryNo)
        Delivery.DBDelete
    Next
End Sub


' ===============================================================
Private Sub Class_Initialize()
    Set pDeliveries = New Collection

End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pDeliveries = Nothing

End Sub

'---------------------------------------------------------------
