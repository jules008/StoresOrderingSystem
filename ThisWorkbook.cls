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
'---------------------------------------------------------------
' Date - 06 Apr 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ThisWorkbook"

' ===============================================================
' Workbook_Open
' Start up processing
' ---------------------------------------------------------------
Private Sub Workbook_Open()
    Const StrPROCEDURE As String = "Workbook_Open()"

    On Error GoTo ErrorHandler
    
    ShtMain.Unprotect
    [menuitemno] = 1
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





