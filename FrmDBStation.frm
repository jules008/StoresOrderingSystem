VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmDBStation 
   Caption         =   "Personnel Details"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7290
   OleObjectBlob   =   "FrmDBStation.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmDBStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 20 Mar 17
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmDBStation"

Private Station As ClsStation

' ===============================================================
' ShowForm
' Initial entry point to form
' ---------------------------------------------------------------
Public Function ShowForm(Optional LocStation As ClsStation) As Boolean
    
    Const StrPROCEDURE As String = "ShowForm()"
    
    On Error GoTo ErrorHandler
    
    If LocStation Is Nothing Then
        Err.Raise NO_LINE_ITEM, Description:="Station unavailable"
    Else
        Set Station = LocStation
        If Not PopulateForm Then Err.Raise HANDLED_ERROR
    End If
    
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
' Populates form controls
' ---------------------------------------------------------------
Private Function PopulateForm() As Boolean
    Dim Address() As String
    Dim i As Integer
    
    Const StrPROCEDURE As String = "PopulateForm()"

    On Error GoTo ErrorHandler
            
    With Station
        Address() = Split(.Address, ",")
        
        For i = LBound(Address) To UBound(Address)
            Me.Controls("TxtAddress" & i) = Address(i)
        Next
        
        TxtName = .Name
        TxtStationNo = .StationNo
        TxtStationType = .StationType
        
    End With
    
    PopulateForm = True

Exit Function

ErrorExit:
    
        
    PopulateForm = False
    FormTerminate
    Terminate

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' FormInitialise
' initialises controls on form at start up
' ---------------------------------------------------------------
Private Function FormInitialise() As Boolean
    Const StrPROCEDURE As String = "FormInitialise()"

    On Error GoTo ErrorHandler

    FormInitialise = True

Exit Function

ErrorExit:

    FormTerminate
    Terminate
    
    FormInitialise = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' FormTerminate
' Terminates the form gracefully
' ---------------------------------------------------------------
Private Function FormTerminate() As Boolean

    On Error Resume Next

    Set Station = Nothing
    Unload Me

End Function


' ===============================================================
' BtnClose_Click
' Closes the form
' ---------------------------------------------------------------
Private Sub BtnClose_Click()

    On Error Resume Next

    FormTerminate

End Sub


