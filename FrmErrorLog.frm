VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmErrorLog 
   Caption         =   "Category Search"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9750
   OleObjectBlob   =   "FrmErrorLog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmErrorLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 09 Jun 17
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmErrorLog"

Private ErrorLog() As String
Private WarningLog() As String

' ===============================================================
' ShowForm
' Initial entry point to form
' ---------------------------------------------------------------
Public Function ShowForm(ByRef LocErrorLog() As String, ByRef LocWarningLog() As String) As Boolean
    
    Const StrPROCEDURE As String = "ShowForm()"
    
    On Error GoTo ErrorHandler
    
    ErrorLog = LocErrorLog
    WarningLog = LocWarningLog
    
    If Not PopulateForm Then Err.Raise HANDLED_ERROR
    
    Show
    ShowForm = True

Exit Function

ErrorExit:

    FormTerminate
    Terminate
    ShowForm = False

Exit Function

ErrorHandler:
    
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        CustomErrorHandler Err.Number
        Resume Next
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' PopulateForm
' Populates form if asset already found
' ---------------------------------------------------------------
Private Function PopulateForm() As Boolean
    Dim i As Integer
    
    Const StrPROCEDURE As String = "PopulateForm()"

    On Error GoTo ErrorHandler
    
    With LstErrors
        For i = 0 To UBound(ErrorLog)
            .AddItem ErrorLog(i)
        Next
    End With
    
    With LstWarnings
        For i = 0 To UBound(WarningLog)
            .AddItem WarningLog(i)
        Next
    End With
    
    PopulateForm = True

Exit Function

ErrorExit:

    PopulateForm = False
    FormTerminate
    Terminate

Exit Function

ErrorHandler:
    
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BtnClose_Click
' Event for page close button
' ---------------------------------------------------------------
Private Sub BtnClose_Click()

    On Error Resume Next
    
    FormTerminate
    
End Sub

' ===============================================================
' UserForm_Initialize
' Automatic initialise event that triggers custom Initialise
' ---------------------------------------------------------------
Private Sub UserForm_Initialize()

    On Error Resume Next
    
    FormInitialise
    
End Sub

' ===============================================================
' UserForm_Terminate
' Automatic Terminate event that triggers custom Terminate
' ---------------------------------------------------------------
Private Sub UserForm_Terminate()

    On Error Resume Next
    
    FormTerminate
    
End Sub

' ===============================================================
' FormInitialise
' Custom initialise form to run start up actions for form
' ---------------------------------------------------------------
Public Function FormInitialise() As Boolean
    Const StrPROCEDURE As String = "FormInitialise()"
        
    On Error GoTo ErrorHandler
    
    LstErrors.Clear
    LstWarnings.Clear
    
    FormInitialise = True
    
Exit Function

ErrorExit:
    
    FormTerminate
    Terminate
    
    FormInitialise = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If

End Function

' ===============================================================
' FormTerminate
' Custom Terminate form to run close down actions for form
' ---------------------------------------------------------------
Public Sub FormTerminate()
    On Error Resume Next
    
    Unload Me
    
End Sub



