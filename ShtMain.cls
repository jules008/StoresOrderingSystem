VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 27 Feb 17
'===============================================================
Option Explicit

Private Const StrMODULE As String = "ShtMain"

' ===============================================================
' Worksheet_Activate
' Start up processing
' ---------------------------------------------------------------
Private Sub Worksheet_Activate()
    Const StrPROCEDURE As String = "Worksheet_Activate()"

    On Error GoTo ErrorHandler

Restart:

    Application.StatusBar = ""
    If MainScreen Is Nothing Then Err.Raise SYSTEM_RESTART
    
Exit Sub

ErrorExit:

    Terminate
    
Exit Sub

ErrorHandler:

    If Err.Number >= 1000 And Err.Number <= 1500 Then
        If CustomErrorHandler(Err.Number) Then
            Resume Restart
        End If
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' ClearOrderList
' Clears list of orders on stores page
' ---------------------------------------------------------------
Public Sub ClearOrderList()
    Dim BtnShape As Shape
    Dim i As Integer
    
    Const StrPROCEDURE As String = "ClearOrderList()"

    On Error Resume Next

    Me.Range("D6:J200").ClearContents
    
    For i = ShtMain.Shapes.Count To 1 Step -1
        
        If Not ShtMain.Shapes(i) Is Nothing Then Set BtnShape = ShtMain.Shapes(i)
        If Left(BtnShape.Name, 8) = "BtnOrder" Then BtnShape.Delete

    Next
                
    Set BtnShape = Nothing
    
End Sub

' ===============================================================
' OpenOrder
' Opens the selected order form
' ---------------------------------------------------------------
Private Sub OpenOrder(OrderNo As Integer)
    Const StrPROCEDURE As String = "OpenOrder()"
    
    Dim Order As ClsOrder
    
    On Error GoTo ErrorHandler

    Set Order = New ClsOrder
    
    Order.DBGet OrderNo
    
    If Not FrmDBOrder.ShowForm(Order) Then Err.Raise HANDLED_ERROR
    
    Set Order = Nothing

Exit Sub

ErrorExit:

    Set Order = Nothing
    Terminate
Exit Sub

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub


