VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsSuppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsSuppliers
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 07 Jul 17
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
Private pSuppliers As Collection

'---------------------------------------------------------------

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pSuppliers.[_newenum]
End Function

' ===============================================================
' Method FindItem
' Finds Item from index
'---------------------------------------------------------------

Public Function FindItem(SupplierID As Variant) As ClsSupplier
Attribute FindItem.VB_UserMemId = 0
    On Error Resume Next
    Set FindItem = pSuppliers.Item(SupplierID)
End Function

' ===============================================================
' Method AddItem
' Add item to Collection
'---------------------------------------------------------------

Public Sub AddItem(Supplier As ClsSupplier)
    Supplier.SetParent ObjPtr(Me)
    pSuppliers.Add Supplier, Key:=CStr(Supplier.SupplierID)
End Sub

' ===============================================================
' Method RemoveItem
' Deletes item from collection
'---------------------------------------------------------------

Public Sub RemoveItem(SupplierID As Variant)
    pSuppliers.Remove SupplierID
End Sub

' ===============================================================
' Method ItemCount
' Returns number of items in collection
'---------------------------------------------------------------

Public Function Count() As Integer
    Count = pSuppliers.Count
End Function

' ===============================================================
' Method GetCollection
' Gets all items in collection
'---------------------------------------------------------------

Public Sub GetCollection()
    Dim RstSupplier As Recordset
    Dim Supplier As ClsSupplier
    Dim i As Integer

    Set RstSupplier = ModDatabase.SQLQuery("TblSupplier")
    With RstSupplier
        .MoveLast
        .MoveFirst
        For i = 1 To .recordcount
            Set Supplier = New ClsSupplier
            Supplier.DBGet !SupplierID
            Me.AddItem Supplier
            .Movenext
        Next
    End With
End Sub

' ===============================================================
' Method DeleteCollection
' Deletes all items in collection
'---------------------------------------------------------------

Public Sub DeleteCollection()
    Dim Supplier As ClsSupplier
    For Each Supplier In pSuppliers
        pSuppliers.Remove CStr(Supplier.SupplierID)
        Supplier.DBDelete
    Next
End Sub


' ===============================================================
Private Sub Class_Initialize()
    Set pSuppliers = New Collection

End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pSuppliers = Nothing

End Sub

'---------------------------------------------------------------
