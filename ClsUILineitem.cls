VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsUILineitem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsUiLineItem
' v0,0 - Initial Version
' v0,1 - Remove dashobjs
'---------------------------------------------------------------
' Date - 06 Apr 17
'===============================================================
' Methods
'---------------------------------------------------------------
' ReOrder - re-orders the shapes so the objects sit in the right order
' Text - Adds text to the line item
'===============================================================

Option Explicit
Private pTop As Integer
Private pHeight As Integer
Private pLeft As Integer
Private pWidth As Integer
Private pText As String
Private pStyle As TypeStyle
Private pName As String
Private pMenu As ClsUIMenu
Private pOnAction As String
Private pShpLineItem As Shape
Private pVisible As Boolean
Private pParent As Long

'---------------------------------------------------------------
Public Property Get Top() As Integer
    Top = pShpLineItem.Top
End Property

Public Property Let Top(ByVal VnewValue As Integer)
    pTop = VnewValue
    pShpLineItem.Top = pTop
End Property

'---------------------------------------------------------------
Public Property Get Left() As Integer
    Left = pShpLineItem.Left
End Property

Public Property Let Left(ByVal VnewValue As Integer)
    pLeft = VnewValue
    pShpLineItem.Left = pLeft
End Property

'---------------------------------------------------------------
Public Property Get Height() As Integer
    Height = pShpLineItem.Height
End Property

Public Property Let Height(ByVal VnewValue As Integer)
    pHeight = VnewValue
    pShpLineItem.Height = pHeight
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
Public Property Get Text() As String
    Text = pShpLineItem.TextFrame.Characters.Text
End Property

Public Property Let Text(ByVal VnewValue As String)
    pText = VnewValue
    pShpLineItem.TextFrame.Characters.Text = pText
End Property

'---------------------------------------------------------------
Public Property Get Name() As String
    Name = pShpLineItem.Name
End Property

Public Property Let Name(ByVal VnewValue As String)
    pName = VnewValue
    pShpLineItem.Name = pName
End Property

'---------------------------------------------------------------
Public Property Get Menu() As ClsUIMenu
    Set Menu = pMenu
End Property

Public Property Let Menu(ByVal VnewValue As ClsUIMenu)
    Set pMenu = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get OnAction() As String
    OnAction = pShpLineItem.OnAction
End Property

Public Property Let OnAction(ByVal VnewValue As String)
    pOnAction = VnewValue
    pShpLineItem.OnAction = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get Visible() As Boolean
    Visible = pShpLineItem.Visible
End Property

Public Property Let Visible(ByVal VnewValue As Boolean)
    Dim DashObj As ClsUIDashObj
    
    pVisible = VnewValue
    
    pShpLineItem.Visible = pVisible
       
End Property

'---------------------------------------------------------------
Public Property Get ShpLineItem() As Shape
    Set ShpLineItem = pShpLineItem
End Property

Public Property Let ShpLineItem(ByVal VnewValue As Shape)
    Set pShpLineItem = VnewValue
End Property

Public Property Let Style(VnewValue As TypeStyle)
    pStyle = VnewValue
    
    With pShpLineItem.Fill
        .TwoColorGradient msoGradientHorizontal, 1
        .ForeColor.RGB = VnewValue.Fill1
        .BackColor.RGB = VnewValue.Fill2
    End With
    
    With pShpLineItem.Line
        .ForeColor.RGB = VnewValue.BorderColour
        .Weight = VnewValue.BorderWidth
        If .Weight = 0 Then .Visible = msoFalse Else .Visible = msoCTrue
    End With
    
    With pShpLineItem.TextFrame
        .Characters.Font.Bold = VnewValue.FontBold
        .Characters.Font.Name = VnewValue.FontStyle
        .Characters.Font.Color = VnewValue.FontColour
        .Characters.Font.Size = VnewValue.FontSize
        If VnewValue.FontXJust <> 0 Then .HorizontalAlignment = VnewValue.FontXJust
    End With
    
    With pShpLineItem.Shadow
        If VnewValue.Shadow = 0 Then
            .Visible = msoFalse
        Else
            .Type = VnewValue.Shadow
            .Visible = msoCTrue
        End If
    End With
End Property

'---------------------------------------------------------------
Public Property Get Parent() As ClsUILineitems
    If pParent <> 0 Then
        Set Parent = GetParentFromPtr(pParent)
    End If
End Property

Friend Function SetParent(ByVal Ptr As Long) As Boolean
    pParent = Ptr
End Function


' ===============================================================
' Method ReOrder
' re-orders the shapes so that the LineItems sit on top of the screen
'---------------------------------------------------------------
Public Sub ReOrder()
    Dim DashObj As ClsUIDashObj
    
    pShpLineItem.ZOrder msoBringToFront
    
    pMenu.ReOrder
End Sub

' ===============================================================
' Method GetParentFromPtr
' Private routine to copy memory address of parent class
' ---------------------------------------------------------------
Private Function GetParentFromPtr(ByVal Ptr As Long) As ClsUILineitem
    Dim tmp As ClsUILineitem

    CopyMemory tmp, Ptr, 4
    Set GetParentFromPtr = tmp
    CopyMemory tmp, 0&, 4
End Function

' ===============================================================
Private Sub Class_Initialize()

    Set pShpLineItem = ShtMain.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 10, 10)
End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
    

End Sub

'---------------------------------------------------------------
