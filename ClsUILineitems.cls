VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsUILineitems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class LINEITEMs
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 03 Apr 17
'===============================================================
' Methods
'---------------------------------------------------------------
' FindItem - Finds Item from index
' AddItem - Add item to Collection
' RemoveItem - Deletes item from collection
' ItemCount - Returns number of items in collection
' ReOrder - re-orders the shapes so that the LineItems sit on top of the screen
' Text - Adds text to the line item
'===============================================================

Option Explicit
Private pNoColumns As Integer
Private pColumns() As Integer
Private PUILineItems As Collection
Private pTop As Integer
Private pLeft As Integer
Private pHeight As Integer
Private pWidth As Integer
Private pStyle As TypeStyle
Private pRowOffset As Integer
Private pParent As Long

'---------------------------------------------------------------
Public Property Get NoColumns() As Integer
    NoColumns = pNoColumns
End Property

Public Property Let NoColumns(ByVal VnewValue As Integer)
    pNoColumns = VnewValue
    ReDim pColumns(0 To NoColumns - 1)
End Property

'---------------------------------------------------------------
Public Property Get ColumnWidth(i As Integer) As Integer
    ColumnWidth = pColumns(i)
End Property

Public Property Let Columns(VnewValue As String)
    Dim i As Integer
    Dim ColArray() As String
    
    ColArray() = Split(VnewValue, ":")
    
    For i = 0 To pNoColumns - 1
        pColumns(i) = CInt(ColArray(i))
    Next
End Property

'---------------------------------------------------------------
Public Property Get Top() As Integer
    Top = pTop - Parent.Top
End Property

Public Property Let Top(ByVal VnewValue As Integer)
    pTop = VnewValue + Parent.Top
End Property

'---------------------------------------------------------------
Public Property Get Left() As Integer
    Left = pLeft
End Property

Public Property Let Left(ByVal VnewValue As Integer)
    pLeft = VnewValue + Parent.Left
End Property

'---------------------------------------------------------------
Public Property Get Height() As Integer
    Height = pHeight
End Property

Public Property Let Height(ByVal VnewValue As Integer)
    pHeight = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get Width() As Integer
    Width = pShpLineItem.Width
End Property

Public Property Let Width(ByVal VnewValue As Integer)
    pWidth = VnewValue
    pShpLineItem.Width = pWidth
End Property

'---------------------------------------------------------------
Public Property Get RowOffset() As Integer
    RowOffset = pRowOffset
End Property

Public Property Let RowOffset(ByVal VnewValue As Integer)
    pRowOffset = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get Style() As TypeStyle
    Style = pStyle
End Property

Public Property Let Style(VnewValue As TypeStyle)
    pStyle = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get Parent() As ClsUIFrame
    If pParent <> 0 Then
        Set Parent = GetParentFromPtr(pParent)
    End If
End Property

Friend Function SetParent(ByVal Ptr As Long) As Boolean
    pParent = Ptr
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = PUILineItems.[_NewEnum]
End Function

' ===============================================================
' Method FindItem
' Finds Item from index
'---------------------------------------------------------------

Public Function FindItem(Name As Variant) As ClsUILineitem
Attribute FindItem.VB_UserMemId = 0
    On Error Resume Next
    Set FindItem = PUILineItems.Item(Name)
End Function

' ===============================================================
' Method AddItem
' Add item to Collection
'---------------------------------------------------------------

Public Sub AddItem(UILineItem As ClsUILineitem)
    UILineItem.SetParent ObjPtr(Me)
    PUILineItems.Add UILineItem, Key:=CStr(UILineItem.Name)
End Sub

' ===============================================================
' Method RemoveItem
' Deletes item from collection
'---------------------------------------------------------------

Public Sub RemoveItem(Name As Variant)
    PUILineItems.Remove Name
End Sub

' ===============================================================
' Method ItemCount
' Returns number of items in collection
'---------------------------------------------------------------

Public Function Count() As Integer
    Count = PUILineItems.Count
End Function

' ===============================================================
' Method Text
' Adds text to the line item
'---------------------------------------------------------------
Public Sub Text(Row As Integer, Col As Integer, ItemText As String, Optional Index As Integer)
    Dim DashObj As ClsUIDashObj
    Dim UILineItem As ClsUILineitem
    Dim LeftOffset As Integer
    Dim i As Integer
    
    Do While i <> Col
        LeftOffset = LeftOffset + pColumns(i)
        i = i + 1
    Loop
    
    Set UILineItem = New ClsUILineitem
    
    With UILineItem
        .Style = pStyle
        .Left = pLeft + LeftOffset
        .Width = pColumns(Col)
        .Height = pHeight
        .Top = pTop + (Row * pRowOffset)
        .Text = ItemText
        .OnAction = "'ModUIStoresScreen.OpenOrder(" & Index & ")'"
    End With
    Me.AddItem UILineItem
End Sub

' ===============================================================
' Method ReOrder
' re-orders the shapes so that the LineItems sit on top of the screen
'---------------------------------------------------------------
Public Sub ReOrder()
    Dim LineItem As ClsUILineitem
        
    For Each LineItem In PUILineItems
        LineItem.ShpLineItem.ZOrder msoBringToFront
        LineItem.ReOrder
    Next
    
End Sub

' ===============================================================
' Method GetParentFromPtr
' Private routine to copy memory address of parent class
' ---------------------------------------------------------------
Private Function GetParentFromPtr(ByVal Ptr As Long) As ClsUIScreen
    Dim tmp As ClsUIScreen

    CopyMemory tmp, Ptr, 4
    Set GetParentFromPtr = tmp
    CopyMemory tmp, 0&, 4
End Function

' ===============================================================
Private Sub Class_Initialize()
    Set PUILineItems = New Collection

End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
    
    Set PUILineItems = Nothing

End Sub

'---------------------------------------------------------------
