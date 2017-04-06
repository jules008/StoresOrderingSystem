VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmDBVehicle 
   Caption         =   "Personnel Details"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6645
   OleObjectBlob   =   "FrmDBVehicle.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmDBVehicle"
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

Private Const StrMODULE As String = "FrmDBVehicle"

Private Vehicle As ClsVehicle

' ===============================================================
' ShowForm
' Initial entry point to form
' ---------------------------------------------------------------
Public Function ShowForm(Optional LocVehicle As ClsVehicle) As Boolean
    
    Const StrPROCEDURE As String = "ShowForm()"
    
    On Error GoTo ErrorHandler
    
    If LocVehicle Is Nothing Then
        Err.Raise NO_LINE_ITEM, Description:="Vehicle unavailable"
    Else
        Set Vehicle = LocVehicle
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
    
    Const StrPROCEDURE As String = "PopulateForm()"

    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    With Vehicle
        TxtCallsign = .CallSign
        TxtDisposalDate = .DisposalDate
        TxtFleetNo = .FleetNo
        TxtRegYear = .RegYear
        TxtVehicleMake = .VehicleMake
        TxtVehReg = .VehReg
        TxtVehType = .VehType
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

    Set Vehicle = Nothing
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


Private Sub CommandButton1_Click()

End Sub

' ===============================================================
' BtnViewLocation_Click
' View the location details
' ---------------------------------------------------------------
Private Sub BtnViewLocation_Click()
    Const StrPROCEDURE As String = "BtnViewLocation_Click()"

    On Error GoTo ErrorHandler

'    If Not FrmDBStation.ShowForm(Vehicle.Station) Then Err.Raise HANDLED_ERROR

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
