Attribute VB_Name = "Main"
Option Explicit
Dim CurrentOrder As ClsOrder
Private Const StrMODULE As String = "Main"


' ===============================================================
' Main
'
' ---------------------------------------------------------------
Public Sub Main()
    Const StrPROCEDURE As String = "Main()"

    On Error GoTo ErrorHandler

    
    If CurrentUser Is Nothing Then
        Err.Raise SYSTEM_RESTART, Description:="Object Model Failed, system restarting"
    Else
        
        If Not FrmTextSearch.ShowForm Then Err.Raise HANDLED_ERROR
        
    End If
GracefulExit:

Exit Sub

ErrorExit:

Set CurrentOrder = Nothing
Set CurrentUser = Nothing

Initialise

Exit Sub

ErrorHandler:

    
If Err.Number >= 1000 And Err.Number <= 1500 Then
    CustomErrorHandler Err.Number
    Resume
End If


If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

