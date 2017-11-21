VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================
' v0,0 - Initial version
' v0,1 - Rest menu item store on start up
' v0,2 - trigger terminate processing and Log Off from DB
'---------------------------------------------------------------
' Date - 14 Nov 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ThisWorkbook"

' ===============================================================
' Workbook_BeforeClose
' Triggers closedown
' ---------------------------------------------------------------
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Const StrPROCEDURE As String = "Workbook_BeforeClose()"

    On Error GoTo ErrorHandler

    If Not CloseDown Then Err.Raise HANDLED_ERROR

Exit Sub

ErrorExit:

'    ***CleanUpCode***

Exit Sub

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub
' ===============================================================
' CloseDown
' Carries out closedown procedures
' ---------------------------------------------------------------
Private Function CloseDown() As Boolean
    Const StrPROCEDURE As String = "CloseDown()"

    On Error GoTo ErrorHandler

    If Not ModCloseDown.Terminate Then Err.Raise HANDLED_ERROR

    CloseDown = True

Exit Function

ErrorExit:

'    ***CleanUpCode***
    CloseDown = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' Workbook_Open
' Start up processing
' ---------------------------------------------------------------
Private Sub Workbook_Open()
    Const StrPROCEDURE As String = "Workbook_Open()"

    On Error GoTo ErrorHandler
    
    ShtMain.Unprotect
    [menuitemno] = ""
    ShtMain.Protect
    
    If MainScreen Is Nothing Then
        ShtMain.Activate
        If Not ModStartUp.Initialise Then Err.Raise HANDLED_ERROR
    End If
Exit Sub

ErrorExit:

'    ***CleanUpCode***

Exit Sub

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub





