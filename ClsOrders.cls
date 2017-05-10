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
' v0,2 - Improved Order retrieval performance
' v0,3 - limited closed orders to last 14 days
' v0,4 - New function FinOrders
' v0,5 - Added MyOrders function
'---------------------------------------------------------------
' Date - 10 May 17
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
' FindOrders - Finds open orders from Crew No
' MyOrders - Query for my orders screen
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

Public Function GetOpenOrders() As Recordset
    Dim RstOrder As Recordset
    Dim Order As ClsOrder
    Dim OrderStatus As EnumOrderStatus
    
    Dim i As Integer
    
    OrderStatus = OrderClosed
    
    Set RstOrder = ModDatabase.SQLQuery("SELECT " _
                                        & "TblOrder.OrderNo AS Order_No, " _
                                        & "Count(TblLineItem.LineItemNo) AS No_of_Items, " _
                                        & "TblPerson.Username AS ReqBy, " _
                                        & "TblStation.Name AS Station, " _
                                        & "TblPerson1.Username AS Assigned_To, " _
                                        & "TblOrderStatus.OrderStatus AS Status " _
                                        & "FROM " _
                                        & "((((TblOrder " _
                                        & "INNER JOIN TblPerson ON TblPerson.CrewNo = TblOrder.RequestorID) " _
                                        & "INNER JOIN TblStation ON TblStation.StationID = TblPerson.StationID) " _
                                        & "INNER JOIN TblOrderStatus ON TblOrder.Status = TblOrderStatus.OrderStatusNo) " _
                                        & "LEFT JOIN TblPerson TblPerson1 ON TblPerson1.CrewNo = TblOrder.AssignedToID) " _
                                        & "LEFT JOIN TblLineItem ON TblLineItem.OrderNo = TblOrder.OrderNo " _
                                        & "WHERE " _
                                        & "TblOrder.Status < 4 " _
                                        & "GROUP BY " _
                                        & "TblOrder.OrderNo, " _
                                        & "TblPerson.Username, " _
                                        & "TblStation.Name, " _
                                        & "TblPerson1.Username, " _
                                        & "TblOrderStatus.OrderStatus " _
                                        & "ORDER BY " _
                                        & "TblOrder.OrderNo")
    Set GetOpenOrders = RstOrder
End Function

' ===============================================================
' Method GetClosedOrders
' Gets all orders that are closed
'---------------------------------------------------------------
Public Function GetClosedOrders() As Recordset
    Dim RstOrder As Recordset
    Dim Order As ClsOrder
    Dim OrderStatus As EnumOrderStatus
    
    Dim i As Integer
    
    OrderStatus = OrderClosed
    
    Set RstOrder = ModDatabase.SQLQuery("SELECT " _
                                        & "TblOrder.OrderNo AS Order_No, " _
                                        & "Count(TblLineItem.LineItemNo) AS No_of_Items, " _
                                        & "TblPerson.Username AS ReqBy, " _
                                        & "TblStation.Name AS Station, " _
                                        & "TblPerson1.Username AS Assigned_To, " _
                                        & "TblOrderStatus.OrderStatus AS Status " _
                                        & "FROM " _
                                        & "((((TblOrder " _
                                        & "INNER JOIN TblPerson ON TblPerson.CrewNo = TblOrder.RequestorID) " _
                                        & "INNER JOIN TblStation ON TblStation.StationID = TblPerson.StationID) " _
                                        & "INNER JOIN TblOrderStatus ON TblOrder.Status = TblOrderStatus.OrderStatusNo) " _
                                        & "LEFT JOIN TblPerson TblPerson1 ON TblPerson1.CrewNo = TblOrder.AssignedToID) " _
                                        & "LEFT JOIN TblLineItem ON TblLineItem.OrderNo = TblOrder.OrderNo " _
                                        & "WHERE " _
                                        & "TblOrder.Status = 4 " _
                                        & "AND DateDiff('d', TblOrder.OrderDate, Now()) < 14 " _
                                        & "GROUP BY " _
                                        & "TblOrder.OrderNo, " _
                                        & "TblPerson.Username, " _
                                        & "TblStation.Name, " _
                                        & "TblPerson1.Username, " _
                                        & "TblOrderStatus.OrderStatus " _
                                        & "ORDER BY TblOrder.OrderNo")
    Set GetClosedOrders = RstOrder
End Function

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
' FindOrders
' Finds open orders from Crew No
' ---------------------------------------------------------------
Public Function FindOrders(CrewNo As String) As Recordset
    Dim RstOrder As Recordset
    Dim i As Integer

    Set RstOrder = ModDatabase.SQLQuery("SELECT * FROM TblOrder WHERE RequestorID = '" & CrewNo _
                                        & "' AND Status < 4" _
                                        & " AND Deleted IS NULL")
    Set FindOrders = RstOrder

    Set RstOrder = Nothing
End Function

' ===============================================================
' MyOrders
' Query for my orders screen
' ---------------------------------------------------------------
Public Function MyOrders() As Recordset
    Dim RstOrder As Recordset
    Dim i As Integer

    Set RstOrder = ModDatabase.SQLQuery("SELECT " _
                                          & "TblOrder.OrderNo AS [Order_No], " _
                                          & "Count(TblLineItem.LineItemNo) AS [No_of_Items], " _
                                          & "TblPerson.Username AS [Assigned_To], " _
                                          & "TblOrderStatus.OrderStatus AS Status " _
                                        & "From " _
                                          & "((TblOrder " _
                                          & "LEFT JOIN TblLineItem ON TblOrder.OrderNo = TblLineItem.OrderNo) " _
                                          & "LEFT JOIN TblOrderStatus ON TblOrder.Status = TblOrderStatus.OrderStatusNo) " _
                                          & "LEFT JOIN TblPerson ON TblOrder.AssignedToID = TblPerson.CrewNo " _
                                        & "WHERE " _
                                          & "TblOrder.Status < 4 " _
                                          & "AND CrewNo = '" & CurrentUser.CrewNo & " '" _
                                        & "Group BY " _
                                          & "TblOrder.OrderNo, " _
                                          & "TblOrder.Status, " _
                                          & "TblPerson.Username, " _
                                          & "TblOrderStatus.OrderStatus " _
                                        & "Order BY " _
                                          & "TblOrder.Status ")
  
    Set MyOrders = RstOrder

    Set RstOrder = Nothing
End Function

' ===============================================================
Private Sub Class_Initialize()
    Set pOrders = New Collection

End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pOrders = Nothing

End Sub

'---------------------------------------------------------------
