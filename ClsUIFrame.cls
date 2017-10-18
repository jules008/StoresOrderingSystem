VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsUIFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsUIFrame
' v0,0 - Initial Version
' v0,1 - Add Dropdowns to frame
'---------------------------------------------------------------
' Date - 18 Oct 17
'===============================================================
' Methods
'---------------------------------------------------------------
' ReOrder - re-orders the shapes so the objects sit in the right order
'===============================================================

Option Explicit
Private pTop As Integer
Private pLeft As Integer
Private pHeight As Integer
Private pWidth As Integer
Private pName As String
Private pStyle As TypeStyle
Private pDashObjs As ClsUIDashObjs
Private pDropDowns As ClsUIDropDowns
Private pMenu As ClsUIMenu
Private pEnableHeader As Boolean
Private pHeader As ClsUIHeader
Private pLineItems As ClsUILineitems
Private pShpFrame As Shape
Private pVisible As Boolean
Private pParent As Long

'---------------------------------------------------------------
Public Property Get Top() As Integer
    Top = pTop
End Property

Public Property Let Top(ByVal VnewValue As Integer)

    pTop = VnewValue + Parent.Top
    pShpFrame.Top = pTop
End Property

'---------------------------------------------------------------
Public Property Get Left() As Integer
    Left = pShpFrame.Left
End Property

Public Property Let Left(ByVal VnewValue As Integer)
    pLeft = VnewValue
    pShpFrame.Left = pLeft
End Property

'---------------------------------------------------------------
Public Property Get Height() As Integer
    Height = pShpFrame.Height
End Property

Public Property Let Height(ByVal VnewValue As Integer)
    pHeight = VnewValue
    pShpFrame.Height = pHeight
End Property

'---------------------------------------------------------------
Public Property Get Width() As Integer
    Width = pShpFrame.Width
End Property

Public Property Let Width(ByVal VnewValue As Integer)
    pWidth = VnewValue
    pShpFrame.Width = pWidth
End Property

'---------------------------------------------------------------
Public Property Get Name() As String
    Name = pShpFrame.Name
End Property

Public Property Let Name(ByVal VnewValue As String)
    pName = VnewValue
    pShpFrame.Name = pName
End Property

'---------------------------------------------------------------
Public Property Get DashObs() As ClsUIDashObjs
    Set DashObs = pDashObjs
End Property

Public Property Let DashObs(ByVal VnewValue As ClsUIDashObjs)
    Set pDashObjs = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get DropDowns() As ClsUIDropDowns
    Set DropDowns = pDropDowns
End Property

Public Property Let DropDowns(ByVal vNewValue As ClsUIDropDowns)
    Set pDropDowns = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Menu() As ClsUIMenu
    Set Menu = pMenu
End Property

Public Property Let Menu(ByVal VnewValue As ClsUIMenu)
    Set pMenu = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get EnableHeader() As Boolean
    EnableHeader = pEnableHeader
End Property

Public Property Let EnableHeader(ByVal VnewValue As Boolean)
    pEnableHeader = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get Header() As ClsUIHeader
    Set Header = pHeader
End Property

Public Property Let Header(ByVal VnewValue As ClsUIHeader)
    Set pHeader = VnewValue
    pHeader.SetParent ObjPtr(Me)
End Property

'---------------------------------------------------------------
Public Property Get LineItems() As ClsUILineitems
    Set LineItems = pLineItems
End Property

Public Property Let LineItems(ByVal VnewValue As ClsUILineitems)
    Set pLineItems = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get Visible() As Boolean
    Visible = pShpFrame.Visible
End Property

Public Property Let Visible(ByVal VnewValue As Boolean)
    Dim DashObj As ClsUIDashObj
    
    pVisible = VnewValue
    
    pShpFrame.Visible = pVisible
    
    For Each DashObj In pDashObjs
        DashObj.Visible = pVisible
    Next
    
    If pEnableHeader Then pHeader.Visible = pVisible
    
    pMenu.Visible = pVisible
    
End Property

'---------------------------------------------------------------
Public Property Get ShpFrame() As Shape
    Set ShpFrame = pShpFrame
End Property

Public Property Let ShpFrame(ByVal VnewValue As Shape)
    Set pShpFrame = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get Parent() As ClsUIScreen
    If pParent <> 0 Then
        Set Parent = GetParentFromPtr(pParent)
    End If
End Property

Friend Function SetParent(ByVal Ptr As Long) As Boolean
    pParent = Ptr
End Function

'---------------------------------------------------------------
Public Property Get Style() As TypeStyle
    Style = pStyle
End Property

Public Property Let Style(VnewValue As TypeStyle)
    pStyle = VnewValue
    
    With pShpFrame.Fill
        .TwoColorGradient msoGradientHorizontal, 1
        .ForeColor.RGB = VnewValue.Fill1
        .BackColor.RGB = VnewValue.Fill2
    End With
    
    With pShpFrame.Line
        .ForeColor.RGB = VnewValue.BorderColour
        .Weight = VnewValue.BorderWidth
        If .Weight = 0 Then .Visible = msoFalse Else .Visible = msoCTrue
    End With
    
    With pShpFrame.TextFrame
        .Characters.Font.Bold = VnewValue.FontBold
        .Characters.Font.Name = VnewValue.FontStyle
        .Characters.Font.Color = VnewValue.FontColour
        .Characters.Font.Size = VnewValue.FontSize
        If VnewValue.FontXJust <> 0 Then .HorizontalAlignment = VnewValue.FontXJust
    End With
    
    With pShpFrame.Shadow
        If VnewValue.Shadow = 0 Then
            .Visible = msoFalse
        Else
            .Type = VnewValue.Shadow
            .Visible = msoCTrue
        End If
    End With
End Property

' ===============================================================
' Method ReOrder
' re-orders the shapes so that the frames sit on top of the screen
'---------------------------------------------------------------
Public Sub ReOrder()
    Dim DashObj As ClsUIDashObj
    Dim DropDown As ClsUIDropDown
    Dim UILineItem As ClsUILineitem
    
    pShpFrame.ZOrder msoBringToFront
    pHeader.ReOrder
    
    For Each DashObj In pDashObjs
        DashObj.ShpDashObj.ZOrder msoBringToFront
    Next
    
    For Each DropDown In pDropDowns
        DropDown.ReOrder
    Next
    
    For Each UILineItem In pLineItems
        UILineItem.ShpLineItem.ZOrder msoBringToFront
    Next
    pMenu.ReOrder
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
    Set pMenu = New ClsUIMenu
    Set pHeader = New ClsUIHeader
    Set pDashObjs = New ClsUIDashObjs
    Set pDropDowns = New ClsUIDropDowns
    Set pLineItems = New ClsUILineitems

    Set pShpFrame = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 0, 10, 10)
    
    pMenu.SetParent ObjPtr(Me)
    pHeader.SetParent ObjPtr(Me)
    pDashObjs.SetParent ObjPtr(Me)
    pDropDowns.SetParent ObjPtr(Me)
    pLineItems.SetParent ObjPtr(Me)
End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
    pDashObjs.SetParent 0
    pDropDowns.SetParent 0
    pMenu.SetParent 0
    pHeader.SetParent 0
    pLineItems.SetParent 0
    
    Set pDashObjs = Nothing
    Set pDropDowns = Nothing
    Set pMenu = Nothing
    Set pHeader = Nothing
    Set pLineItems = Nothing

End Sub

'---------------------------------------------------------------
