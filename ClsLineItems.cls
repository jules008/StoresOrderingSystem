VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsLineItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsLineitems
' v0,0 - Initial Version
' v0,1 - Clean up unfinished Lineitems when closing class
' v0,2 - Em Fix - disable clean up
'---------------------------------------------------------------
' Date - 03 Oct 17
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
Private pLineitems As Collection
Private pParent As Long

'---------------------------------------------------------------
Public Property Get Parent() As ClsOrder
    If pParent <> 0 Then
        Set Parent = GetParentFromPtr(pParent)
    End If
End Property

Friend Function SetParent(ByVal Ptr As Long) As Boolean
    pParent = Ptr
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pLineitems.[_NewEnum]
End Function

' ===============================================================
' Method FindItem
' Finds Item from index
'---------------------------------------------------------------

Public Function FindItem(LineitemNo As Variant) As ClsLineItem
Attribute FindItem.VB_UserMemId = 0
    On Error Resume Next
    Set FindItem = pLineitems(LineitemNo)
End Function

' ===============================================================
' Method AddItem
' Add item to Collection
'---------------------------------------------------------------

Public Sub AddItem(Lineitem As ClsLineItem)
    Lineitem.SetParent ObjPtr(Me.Parent)
    pLineitems.Add Lineitem, Key:=CStr(Lineitem.LineitemNo)
End Sub

' ===============================================================
' Method RemoveItem
' Deletes item from collection
'---------------------------------------------------------------

Public Sub RemoveItem(LineitemNo As Variant)
    pLineitems.Remove LineitemNo
End Sub

' ===============================================================
' Method ItemCount
' Returns number of items in collection
'---------------------------------------------------------------

Public Function Count() As Integer
    Count = pLineitems.Count
End Function

' ===============================================================
' Method GetCollection
' Gets all items in collection
'---------------------------------------------------------------

Public Sub GetCollection()
    Dim RstLineitem As Recordset
    Dim Lineitem As ClsLineItem
    Dim i As Integer

    Set RstLineitem = ModDatabase.SQLQuery("SELECT * FROM TblLineitem WHERE OrderNo = " & Parent.OrderNo _
                                            & " AND deleted IS NULL")
    With RstLineitem
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                Set Lineitem = New ClsLineItem
                Lineitem.LineitemNo = !LineitemNo
                Lineitem.DBGet
                Me.AddItem Lineitem
                .MoveNext
            Loop
        End If
    End With
    Set RstLineitem = Nothing
End Sub

' ===============================================================
' Method DeleteCollection
' Deletes all items in collection
'---------------------------------------------------------------

Public Sub DeleteCollection()
    Dim Lineitem As ClsLineItem
    For Each Lineitem In pLineitems
        pLineitems.Remove CStr(Lineitem.LineitemNo)
        Lineitem.DBDelete
    Next
End Sub

' ===============================================================
' Method Cleanup
' Cleans up incomplete Lineitems
'---------------------------------------------------------------
Public Sub CleanUp()
    Dim RstLineitems As Recordset

    Set RstLineitems = ModDatabase.SQLQuery("SELECT * FROM TblLineitem WHERE OrderNo = 0 OR OrderNo IS NULL")
    
    With RstLineitems
        If .RecordCount > 0 Then
            Do While Not .EOF
'                Debug.Print !LineitemNo
                .Delete
                .MoveNext
            Loop
        End If
    End With
End Sub

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
    Set pLineitems = New Collection

End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
    'CleanUp
    Set pLineitems = Nothing

End Sub

'---------------------------------------------------------------
