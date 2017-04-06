VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmDBLossReport 
   Caption         =   "Loss Report"
   ClientHeight    =   9960
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12570
   OleObjectBlob   =   "FrmDBLossReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmDBLossReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 11 Mar 17
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmDBLossReport"

Private LossReport As ClsLossReport

' ===============================================================
' ShowForm
' Initial entry point to form
' ---------------------------------------------------------------
Public Function ShowForm(Optional LocLossReport As ClsLossReport) As Boolean
    
    Const StrPROCEDURE As String = "ShowForm()"
    
    On Error GoTo ErrorHandler
    
    If LocLossReport Is Nothing Then
        Err.Raise NO_LOSS_REPORT, Description:="LossReport unavailable"
    Else
        Set LossReport = LocLossReport
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
    Dim LineItem As ClsLineItem
    
    On Error GoTo ErrorHandler

    With LossReport
        TxtActionsTaken = .ActionsTaken
        TxtAdditInfo = .AdditInfo
        TxtRequestor = .Parent.Parent.Requestor.UserName
        TxtIncNo = .IncNo
        TxtLossReportNo = .LossReportNo
        TxtOpsSupportAction = .OpsSupportAction
        TxtOrderNo = .Parent.Parent.OrderNo
        TxtReportDate = .ReportDate
        TxtReportingOfficer = .ReportingOfficer.UserName
        ChkAuthorised = .Authorised
        ChkRejected = .Rejected
        CmoReason.ListIndex = .Parent.ReqReason
        CmoStatus.ListIndex = .Status
    
    
    
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

    With CmoReason
        .Clear
        .AddItem "Used / Consumed"
        .AddItem "Lost"
        .AddItem "Stolen"
        .AddItem "Damaged Incident / Training"
        .AddItem "Damaged Other"
        .AddItem "Malfunction"
        .AddItem "New Issue"
    
    End With
    
    With CmoStatus
        .Clear
        .AddItem "Open"
        .AddItem "Assigned"
        .AddItem "On Hold"
        .AddItem "Approved"
        .AddItem "Rejected"
    
    End With
    
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

    Set LossReport = Nothing
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

Private Sub BtnCancel_Click()

End Sub

' ===============================================================
' CmoStatus_Change
' Selects new status of loss report
' ---------------------------------------------------------------
Private Sub CmoStatus_Change()
    Const StrPROCEDURE As String = "CmoStatus_Change()"

    On Error GoTo ErrorHandler

    With LossReport
        If CmoStatus <> -1 Then
            .Status = CmoStatus.ListIndex
        End If
    End With
    
    If Not ProcessStatus Then Err.Raise HANDLED_ERROR
    
    If Not PopulateForm Then Err.Raise HANDLED_ERROR

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
' ChkAuthorised_Click
' Authorises Loss Report
' ---------------------------------------------------------------
Private Sub ChkAuthorised_Click()
    Const StrPROCEDURE As String = "ChkAuthorised_Click()"

    On Error GoTo ErrorHandler

    With LossReport
        If ChkAuthorised Then
            .Authorised = True
            .Rejected = False
        Else
            .Authorised = False
        End If
    End With
    
    If Not ProcessStatus Then Err.Raise HANDLED_ERROR
    
    If Not PopulateForm Then Err.Raise HANDLED_ERROR

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
' ChkRejected_Click
' Authorises Loss Report
' ---------------------------------------------------------------
Private Sub ChkRejected_Click()
    Const StrPROCEDURE As String = "ChkRejected_Click()"

    On Error GoTo ErrorHandler

    With LossReport
        If ChkRejected Then
            .Rejected = True
            .Authorised = False
        Else
            .Rejected = False
        End If
        .DBSave
    End With
        
    If Not ProcessStatus Then Err.Raise HANDLED_ERROR
    
    If Not PopulateForm Then Err.Raise HANDLED_ERROR

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
' ProcessStatus
' Status processing for Line Item
' ---------------------------------------------------------------
Private Function ProcessStatus() As Boolean
    
    Const StrPROCEDURE As String = "ProcessStatus()"

    On Error GoTo ErrorHandler

    With LossReport
        If .ReportingOfficer.CrewNo = "" Then
            .Status = RepOpen
        Else
            .Status = RepAssigned
        End If
        
        If .Authorised Then .Status = RepApproved
        If .Rejected Then .Status = RepRejected
        .DBSave
    End With
    
    ProcessStatus = True

Exit Function

ErrorExit:

'    ***CleanUpCode***
    ProcessStatus = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

