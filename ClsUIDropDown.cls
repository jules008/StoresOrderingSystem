VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsUIDropDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsUIDropDown
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 18 Oct 17
'===============================================================
' Methods
'---------------------------------------------------------------
' ReOrder - re-orders the shapes so that the frames sit on top of the screen
'===============================================================

Option Explicit
Private pTop As Integer
Private pLeft As Integer
Private pHeight As Integer
Private pWidth As Integer
Private pName As String
Private pText As String
Private pStyle As TypeStyle
Private pIcon As Shape
Private pShpTextBox As Shape
Private pVisible As Boolean
Private pParent As Long

'---------------------------------------------------------------
Public Property Get Top() As Integer
    Top = pShpTextBox.Top
End Property

Public Property Let Top(ByVal vNewValue As Integer)
    Dim ParentOffset As Integer
    
    pTop = vNewValue
    
    If Not Parent Is Nothing Then ParentOffset = Parent.Top
    
    pShpTextBox.Top = pTop + ParentOffset
End Property

'---------------------------------------------------------------
Public Property Get Left() As Integer
    Left = pShpTextBox.Left
End Property

Public Property Let Left(ByVal vNewValue As Integer)
    Dim ParentOffset As Integer
    
    pLeft = vNewValue
    
    If Not Parent Is Nothing Then ParentOffset = Parent.Left
    
    pShpTextBox.Left = pLeft + ParentOffset
End Property

'---------------------------------------------------------------
Public Property Get Height() As Integer
    Height = pShpTextBox.Height
End Property

Public Property Let Height(ByVal vNewValue As Integer)
    pHeight = vNewValue
    pShpTextBox.Height = pHeight
End Property

'---------------------------------------------------------------
Public Property Get Width() As Integer
    Width = pShpTextBox.Width
End Property

Public Property Let Width(ByVal vNewValue As Integer)
    pWidth = vNewValue
    pShpTextBox.Width = pWidth
End Property

'---------------------------------------------------------------
Public Property Get Name() As String
    Name = pShpTextBox.Name
End Property

Public Property Let Name(ByVal vNewValue As String)
    pName = vNewValue
    pShpTextBox.Name = pName
End Property

'---------------------------------------------------------------
Public Property Get Text() As String
    Text = pShpTextBox.TextFrame.Characters.Caption
End Property

Public Property Let Text(ByVal vNewValue As String)
    pText = vNewValue
    pShpTextBox.TextFrame.Characters.Caption = pText
End Property

'---------------------------------------------------------------
Public Property Get Style() As TypeStyle
    Style = pStyle
End Property

Public Property Let Style(vNewValue As TypeStyle)
    pStyle = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get ShpTextBox() As Shape
    Set ShpTextBox = pShpTextBox
End Property

Public Property Let ShpTextBox(ByVal vNewValue As Shape)
    Set pShpTextBox = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Icon() As Shape
    Set Icon = pIcon
End Property

Public Property Let Icon(ByVal vNewValue As Shape)
    Set pIcon = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Visible() As Boolean
    Visible = pVisible
End Property

Public Property Let Visible(ByVal vNewValue As Boolean)
    pVisible = vNewValue
    pShpTextBox.Visible = vNewValue
    pIcon.Visible = vNewValue
    
End Property

' ===============================================================
' Method ReOrder
' re-orders the shapes so that the frames sit on top of the screen
'---------------------------------------------------------------
Public Sub ReOrder()
    pShpTextBox.ZOrder msoBringToFront
    pIcon.ZOrder msoBringToFront
End Sub

'---------------------------------------------------------------
Public Property Get Parent() As ClsUIMenu
    If pParent <> 0 Then
        Set Parent = GetParentFromPtr(pParent)
    End If
End Property

Friend Function SetParent(ByVal Ptr As Long) As Boolean
    pParent = Ptr
End Function

' ===============================================================
' Method GetParentFromPtr
' Private routine to copy memory address of parent class
' ---------------------------------------------------------------
Private Function GetParentFromPtr(ByVal Ptr As Long) As ClsUIMenu
    Dim tmp As ClsUIMenu

    CopyMemory tmp, Ptr, 4
    Set GetParentFromPtr = tmp
    CopyMemory tmp, 0&, 4
End Function

' ===============================================================
Private Sub Class_Initialize()
    Set pShpTextBox = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 0, 10, 10)
    Set pIcon = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 0, 10, 10)
End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
End Sub

'---------------------------------------------------------------
