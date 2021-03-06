VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtOrderList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================
' Module ShtOrderList
' v0,0 - Initial Version
' v0,1 - Reinforced sheet reference
' v0,2 - Bug fix for Shane's printing error
'---------------------------------------------------------------
' Date - 15 May 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ShtOrderList"


' ===============================================================
' ClearForm
' Clears all orders from form
' ---------------------------------------------------------------
Public Function ClearForm() As Boolean
    
    Const StrPROCEDURE As String = "ClearForm()"

    On Error Resume Next

    ShtOrderList.Range("RngClear").ClearContents

    ClearForm = True
    
    On Error GoTo 0

Exit Function

ErrorExit:

    
    ClearForm = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function


