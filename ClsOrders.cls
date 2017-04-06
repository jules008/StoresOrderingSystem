VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsOrders
' v0,0 - Initial Version
' v0,1 - Add Get Closed Orders Method
'---------------------------------------------------------------
' Date - 06 Apr 17
'===============================================================
' Methods
'---------------------------------------------------------------
' FindItem - Finds Item from index
' AddItem - Add item to Collection
' RemoveItem - Deletes item from collection
' ItemCount - Returns number of items in collection
' GetCollection - Gets all items in collection
' DeleteCollection - Deletes all items in collection
' GetOpenOrders - Gets all orders that are not closed
' GetClosedOrders - Gets all orders that are closed
'===============================================================

Option Explicit
Private pOrders As Collection

'---------------------------------------------------------------

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pOrders.[_NewEnum]
End Function

' ===============================================================
' Method FindItem
' Finds Item from index
'---------------------------------------------------------------

Public Function FindItem(OrderNo As Variant) As ClsOrder
Attribute FindItem.VB_UserMemId = 0
    On Error Resume Next
    Set FindItem = pOrders.Item(OrderNo)
End Function

' ===============================================================
' Method AddItem
' Add item to Collection
'---------------------------------------------------------------

Public Sub AddItem(Order As ClsOrder)
    pOrders.Add Order, Key:=CStr(Order.OrderNo)
End Sub

' ===============================================================
' Method RemoveItem
' Deletes item from collection
'---------------------------------------------------------------

Public Sub RemoveItem(OrderNo As Variant)
    pOrders.Remove OrderNo
End Sub

' ===============================================================
' Method ItemCount
' Returns number of items in collection
'---------------------------------------------------------------

Public Function Count() As Integer
    Count = pOrders.Count
End Function

' ===============================================================
' Method GetOpenOrders
' Gets all orders that are not closed
'---------------------------------------------------------------

Public Sub GetOpenOrders()
    Dim RstOrder As Recordset
    Dim Order As ClsOrder
    Dim OrderStatus As EnumOrderStatus
    
    Dim i As Integer
    
    OrderStatus = orderClosed
    
    Set RstOrder = ModDatabase.SQLQuery("SELECT * FROM TblOrder WHERE " _
                        & "status <> " & OrderStatus & " AND Deleted IS NULL")
    With RstOrder
        If .RecordCount > 0 Then
            Do While Not .EOF
                Set Order = New ClsOrder
                Order.DBGet !OrderNo
                Me.AddItem Order
                .MoveNext
            Loop
        End If
    End With
End Sub

' ===============================================================
' Method GetClosedOrders
' Gets all orders that are closed
'---------------------------------------------------------------
Public Sub GetClosedOrders()
    Dim RstOrder As Recordset
    Dim Order As ClsOrder
    Dim OrderStatus As EnumOrderStatus
    
    Dim i As Integer
    
    OrderStatus = orderClosed
    
    Set RstOrder = ModDatabase.SQLQuery("SELECT * FROM TblOrder WHERE " _
                        & "status = " & OrderStatus & " AND Deleted IS NULL")
    With RstOrder
        If .RecordCount > 0 Then
            Do While Not .EOF
                Set Order = New ClsOrder
                Order.DBGet !OrderNo
                Me.AddItem Order
                .MoveNext
            Loop
        End If
    End With
End Sub

' ===============================================================
' Method DeleteCollection
' Deletes all items in collection
'---------------------------------------------------------------

Public Sub DeleteCollection()
    Dim Order As ClsOrder
    For Each Order In pOrders
        pOrders.Remove CStr(Order.OrderNo)
        Order.DBDelete
    Next
End Sub


' ===============================================================
Private Sub Class_Initialize()
    Set pOrders = New Collection

End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pOrders = Nothing

End Sub

'---------------------------------------------------------------
