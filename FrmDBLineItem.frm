VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmDBLineItem 
   Caption         =   "Loss Report"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12030
   OleObjectBlob   =   "FrmDBLineItem.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmDBLineItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
' v0,0 - Initial version
' v0,1 - Auto assign
'---------------------------------------------------------------
' Date - 07 Apr 17
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmDBLineItem"

Private Lineitem As ClsLineItem

' ===============================================================
' ShowForm
' Initial entry point to form
' ---------------------------------------------------------------
Public Function ShowForm(Optional LocLineItem As ClsLineItem) As Boolean
    
    Const StrPROCEDURE As String = "ShowForm()"
    
    On Error GoTo ErrorHandler
    
    If LocLineItem Is Nothing Then
        Err.Raise NO_LINE_ITEM, Description:="LineItem unavailable"
    Else
        Set Lineitem = LocLineItem
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
    
    With Lineitem
        TxtAsset = .Asset.Description
        TxtOnHoldReason = .OnHoldReason
        ChkReturned = .itemsReturned
        ChkIssued = .ItemsIssued
        TxtLineItemNo = .LineItemNo
        TxtOrderDate = .Parent.OrderDate
        TxtOrderNo = .Parent.OrderNo
        TxtQuantity = .Quantity
        TxtRequestedBy = .Parent.Requestor.UserName
        ChkDelivered = .ItemsDelivered
        ChkReturnReqd = .ReturnReqd
        CmoReqReason.ListIndex = .ReqReason
        CmoStatus.ListIndex = .Status
        
        If .LossReport.LossReportNo <> 0 Then
            ChkLossReport = True
        Else
            ChkLossReport = False
        End If
        
        With TxtLossRepStatus
            Select Case Lineitem.LossReport.Status
                Case Is = 0
                    .Value = "Open"
                Case Is = 1
                    .Value = "Assigned"
                Case Is = 2
                    .Value = "On Hold"
                Case Is = 3
                    .Value = "Approved"
                Case Is = 4
                    .Value = "Rejected"
            End Select
        
        End With
        
        Select Case .Asset.AllocationType
            Case Is = Person
                TxtItemFor = .ForPerson.UserName
            Case Is = Station
                
                With .ForStation
                    TxtItemFor = .StationNo & " " & .Name
                End With
            
            Case Is = Vehicle
                With .ForVehicle
                    TxtItemFor = .CallSign & " " & .VehicleMake & " (" & .VehReg & ")"
                End With
        End Select
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

    With CmoStatus
        .Clear
        .AddItem "Open"
        .AddItem "On Hold"
        .AddItem "Issued"
        .AddItem "Delivered"
        .AddItem "Complete"
    End With
    
    With CmoReqReason
        .Clear
        .AddItem "Used / Consumed"
        .AddItem "Lost"
        .AddItem "Stolen"
        .AddItem "Damaged Incident / Training"
        .AddItem "Damaged Other"
        .AddItem "Malfunction"
        .AddItem "New Issue"
    
    End With
    
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

    Set Lineitem = Nothing
    Unload Me

End Function

Private Sub BtnCancel_Click()

End Sub

' ===============================================================
' BtnAsset_Click
' Shows the asset details
' ---------------------------------------------------------------
Private Sub BtnAsset_Click()
    Dim LineItemNo As Integer
    
    Const StrPROCEDURE As String = "BtnAsset_Click()"
        
    On Error GoTo ErrorHandler
    
    If Lineitem Is Nothing Then Err.Raise NO_LINE_ITEM, Description:="No LineItem available"
    
    If Lineitem.Asset Is Nothing Then Err.Raise NO_ASSET_ON_ORDER, Description:="No asset on Order"
    
    If Not FrmDBAsset.ShowForm(Lineitem.Asset) Then Err.Raise HANDLED_ERROR

GracefulExit:

    Set Lineitem = Nothing

Exit Sub

ErrorExit:

    Set Lineitem = Nothing
    FormTerminate
    Terminate

Exit Sub

ErrorHandler:
    
        If Err.Number >= 1000 And Err.Number <= 1500 Then
        CustomErrorHandler Err.Number
        Resume GracefulExit
    End If
    
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' BtnClose_Click
' Closes the form
' ---------------------------------------------------------------
Private Sub BtnClose_Click()

    On Error Resume Next

    FormTerminate

End Sub

Private Sub CommandButton3_Click()

End Sub

Private Sub CommandButton1_Click()

End Sub

' ===============================================================
' BtnIssue_Click
' Issues line item
' ---------------------------------------------------------------
Private Sub BtnIssue_Click()
    Const StrPROCEDURE As String = "BtnIssue_Click()"

    On Error GoTo ErrorHandler

    With Lineitem
        .Status = LineIssued
        ChkIssued = True
        .DBSave
    End With

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
' BtnPutOnHold_Click
' Puts lineitem on hold
' ---------------------------------------------------------------
Private Sub BtnPutOnHold_Click()
    Const StrPROCEDURE As String = "BtnPutOnHold_Click()"

    On Error GoTo ErrorHandler

    With Lineitem
        .Parent.AssignedTo = CurrentUser
        .Status = LineOnHold
        .DBSave
    End With

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

Private Sub BtnUpdate_Click()

End Sub

' ===============================================================
' BtnViewDelivery_Click
' View the delivery details
' ---------------------------------------------------------------
Private Sub BtnViewDelivery_Click()
    Const StrPROCEDURE As String = "BtnViewDelivery_Click()"

    On Error GoTo ErrorHandler

    With Lineitem
        Select Case .Asset.AllocationType
            Case Is = Person
                If Not FrmDBPerson.ShowForm(.ForPerson) Then Err.Raise HANDLED_ERROR
            
            Case Is = Station
                If Not FrmDBStation.ShowForm(.ForStation) Then Err.Raise HANDLED_ERROR
            
            Case Is = Vehicle
                If Not FrmDBVehicle.ShowForm(.ForVehicle) Then Err.Raise HANDLED_ERROR
        End Select
    End With




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
' BtnViewLossRep_Click
' View Loss Report
' ---------------------------------------------------------------
Private Sub BtnViewLossRep_Click()
    Const StrPROCEDURE As String = "BtnViewLossRep_Click()"

    On Error GoTo ErrorHandler

    If ChkLossReport Then
        If Not FrmDBLossReport.ShowForm(Lineitem.LossReport) Then Err.Raise HANDLED_ERROR
    
        If Not ProcessStatus Then Err.Raise HANDLED_ERROR
        
        If Not PopulateForm Then Err.Raise HANDLED_ERROR
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
' ===============================================================
' BtnViewRequestor_Click
' Views the person who requested the order
' ---------------------------------------------------------------
Private Sub BtnViewRequestor_Click()
    Const StrPROCEDURE As String = "BtnViewRequestor_Click()"

    On Error GoTo ErrorHandler

    If Not FrmDBPerson.ShowForm(Lineitem.Parent.Requestor) Then Err.Raise HANDLED_ERROR

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

Private Sub CommandButton2_Click()

End Sub

' ===============================================================
' ChkDelivered_Click
' Mark item as delivered
' ---------------------------------------------------------------
Private Sub ChkDelivered_Click()
    Const StrPROCEDURE As String = "ChkDelivered_Click()"

    On Error GoTo ErrorHandler
    
    With Lineitem
        If ChkDelivered Then .ItemsDelivered = True Else .ItemsDelivered = False
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
' ChkIssued_Click
' Mark item as Issued
' ---------------------------------------------------------------
Private Sub ChkIssued_Click()
    Const StrPROCEDURE As String = "ChkIssued_Click()"

    On Error GoTo ErrorHandler

    With Lineitem
        If ChkIssued Then
            .ItemsIssued = True
            .Parent.AssignedTo = CurrentUser
        Else
            .ItemsIssued = False
            .ItemsDelivered = False
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
' ChkReturned_Click
' Mark item as delivered
' ---------------------------------------------------------------
Private Sub ChkReturned_Click()
    Const StrPROCEDURE As String = "ChkReturned_Click()"

    On Error GoTo ErrorHandler

    With Lineitem
        If ChkReturned Then .itemsReturned = True Else .itemsReturned = False
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
' CmoReqReason_Change
' Change of Request Reason
' ---------------------------------------------------------------
Private Sub CmoReqReason_Change()
    Const StrPROCEDURE As String = "CmoReqReason_Change()"

    On Error GoTo ErrorHandler

    With Lineitem
        .ReqReason = CmoReqReason.ListIndex
        .DBSave
    End With
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
' TxtOnHoldReason_Change
' Captures changes to on hold reason
' ---------------------------------------------------------------
Private Sub TxtOnHoldReason_Change()
    Const StrPROCEDURE As String = "TxtOnHoldReason_Change()"

    On Error GoTo ErrorHandler

    With Lineitem
        .Parent.AssignedTo = CurrentUser
        .OnHoldReason = TxtOnHoldReason
        .DBSave
    End With

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
' ProcessStatus
' Status processing for Line Item
' ---------------------------------------------------------------
Private Function ProcessStatus() As Boolean
    Dim LossReport As Boolean
    
    Const StrPROCEDURE As String = "ProcessStatus()"

    On Error GoTo ErrorHandler
    
    With Lineitem
    
        If .LossReport.LossReportNo = 0 Then
            LossReport = False
        Else
            LossReport = True
        End If
        
        If ChkIssued Then
            If ChkDelivered Then
                If .ReturnReqd Or LossReport Then
                    If .ReturnReqd Then
                        If Not .itemsReturned Then
                            .Status = LineOnHold
                            .OnHoldReason = "Waiting for Return Items"
                        Else
                            .Status = LineComplete
                        End If
                    End If
                    
                    If LossReport Then
                        If Not .LossReport.Authorised Then
                            .Status = LineOnHold
                            .OnHoldReason = "Loss Report not Authorised"
                        Else
                            .Status = LineComplete
                        End If
                    End If
                Else
                    .Status = LineComplete
                End If
            Else
                .Status = LineIssued
            End If
        Else
            .Status = LineOpen
        End If
        
        
        If .Status <> LineOnHold Then .OnHoldReason = ""
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

' ===============================================================
' UserForm_Terminate
' Automatic Terminate event that triggers custom Terminate
' ---------------------------------------------------------------
Private Sub UserForm_Terminate()

    On Error Resume Next
    
    FormTerminate
End Sub

