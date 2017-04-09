VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmLossReport 
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7560
   OleObjectBlob   =   "FrmLossReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmLossReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 06 Mar 17
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmLossReport"

Private Lineitem As ClsLineItem

' ===============================================================
' ShowForm
' Initial entry point to form
' ---------------------------------------------------------------
Public Function ShowForm(Optional LocLineItem As ClsLineItem) As Boolean
    
    Const StrPROCEDURE As String = "ShowForm()"
    
    On Error GoTo ErrorHandler
    
    If LocLineItem Is Nothing Then
        Err.Raise NO_LINE_ITEM, Description:="No Line Item Passed to ShowForm Function"
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
    Dim i As Integer
    Dim x As Integer
    Dim AllowedRns() As String
    
    Const StrPROCEDURE As String = "PopulateForm()"

    On Error GoTo ErrorHandler

    AllowedRns = Split(Lineitem.Asset.AllowedOrderReasons, ":")
    
    x = 0
    With CmoReason
        .Clear
        If AllowedRns(0) = 1 Then
            .AddItem
            .List(x, 0) = "Used / Consumed"
            .List(x, 1) = 0
            x = x + 1
        End If
        
        If AllowedRns(1) = 1 Then
            .AddItem
            .List(x, 0) = "Lost"
            .List(x, 1) = 1
            x = x + 1
        End If
        
        If AllowedRns(2) = 1 Then
            .AddItem
            .List(x, 0) = "Stolen"
            .List(x, 1) = 2
            x = x + 1
        End If
        
        If AllowedRns(3) = 1 Then
            .AddItem
            .List(x, 0) = "Damaged Op / Training"
            .List(x, 1) = 3
            x = x + 1
        End If
    
        If AllowedRns(4) = 1 Then
            .AddItem
            .List(x, 0) = "Damaged Other"
            .List(x, 1) = 4
            x = x + 1
        End If
    
        If AllowedRns(5) = 1 Then
            .AddItem
            .List(x, 0) = "Malfunction"
            .List(x, 1) = 5
            x = x + 1
        End If
    

        If AllowedRns(6) = 1 Then
            .AddItem
            .List(x, 0) = "New Issue"
            .List(x, 1) = 6
            x = x + 1
        End If
    End With
                    
    With Lineitem
        If .ReqReason <> 0 Then
            For i = 0 To CmoReason.ListCount - 1
                If CmoReason.List(i, 1) = .ReqReason Then
                    CmoReason.ListIndex = i
                End If
            Next
            If .LossReport.ActionsTaken <> "" Then TxtComments1 = .LossReport.ActionsTaken
            If .LossReport.IncNo <> "" Then TxtCrimeNo = .LossReport.IncNo
        End If
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
' FormTerminate
' Terminates the form gracefully
' ---------------------------------------------------------------
Private Function FormTerminate() As Boolean

    On Error Resume Next

    Set Lineitem = Nothing
    Unload Me

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
' BtnNext_Click
' Moves onto next form
' ---------------------------------------------------------------
Private Sub BtnNext_Click()
    Dim StrUserName As String
    
    Const StrPROCEDURE As String = "BtnNext_Click()"

    On Error GoTo ErrorHandler

        Select Case ValidateForm
    
            Case Is = FunctionalError
                Err.Raise HANDLED_ERROR
            
            Case Is = FormOK

                If Lineitem Is Nothing Then Err.Raise SYSTEM_FAILURE, Description:="No LineItem Available"
                                
                If Lineitem.LossReport Is Nothing Then Err.Raise NO_LOSS_REPORT, Description:="Loss Report is missing"
                                     
                If Not BuildLossReport Then Err.Raise HANDLED_ERROR
                
                If Not FrmOrder.AddLineItem(Lineitem) Then Err.Raise HANDLED_ERROR
                
                Unload Me
        
        End Select
        

Exit Sub

ErrorExit:

    FormTerminate
    Terminate

Exit Sub

ErrorHandler:
    
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        CustomErrorHandler Err.Number
        Resume Next
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' BtnPrev_Click
' Back to previous screen event
' ---------------------------------------------------------------
Private Sub BtnPrev_Click()
    Const StrPROCEDURE As String = "BtnPrev_Click()"

    On Error GoTo ErrorHandler

    If Not SelectPrevForm Then Err.Raise HANDLED_ERROR
    
Exit Sub

ErrorExit:
    FormTerminate
    Terminate
Exit Sub

ErrorHandler:

    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' CmoReason_Change
' Selection of reason for request
' ---------------------------------------------------------------
Private Sub CmoReason_Change()
    
    Const StrPROCEDURE As String = "CmoReason_Change()"

    Dim Reason As EnumReqReason
    
    On Error GoTo ErrorHandler
    
    CmoReason.BackColor = COLOUR_3
    
    If CmoReason.ListIndex <> -1 Then
        
        With CmoReason
            Reason = .List(.ListIndex, 1)
        End With
        
        Select Case Reason
            
            Case Is = DamagedOpTraining
                Lineitem.ReturnReqd = True
                TxtCrimeNo.Visible = False
                TxtComments1.Visible = True
                TxtCrimeNo.Value = ""
                TxtComments1.Value = ""
                LblText2.Visible = True
                LblText2.Caption = "Please describe damage to the item"
                LblText3.Visible = False
                LblText3.Caption = ""
                
            Case Is = DamagedOther
                Lineitem.ReturnReqd = True
                TxtCrimeNo.Visible = False
                TxtComments1.Visible = True
                TxtCrimeNo.Value = ""
                TxtComments1.Value = ""
                LblText2.Visible = True
                LblText2.Caption = "Please describe damage to the item"
                LblText3.Visible = False
                LblText3.Caption = ""
            
            Case Is = lost
                Lineitem.ReturnReqd = False
                TxtCrimeNo.Visible = False
                TxtComments1.Visible = True
                TxtCrimeNo.Value = ""
                TxtComments1.Value = ""
                LblText2.Visible = True
                LblText2.Caption = "What actions have you taken to find the item"
                LblText3.Visible = False
                LblText3.Caption = ""
           
            Case Is = Malfunction
                Lineitem.ReturnReqd = True
                TxtCrimeNo.Visible = False
                TxtComments1.Visible = True
                TxtCrimeNo.Value = ""
                TxtComments1.Value = ""
                LblText2.Visible = True
                LblText2.Caption = "How has the item malfunctioned?"
                LblText3.Visible = False
                LblText3.Caption = ""
          
            Case Is = NewIssue
                Lineitem.ReturnReqd = False
                TxtCrimeNo.Visible = False
                TxtComments1.Visible = False
                TxtCrimeNo.Value = ""
                TxtComments1.Value = ""
                LblText2.Visible = False
                LblText2.Caption = ""
                LblText3.Visible = False
                LblText3.Caption = ""
            
            Case Is = Stolen
                Lineitem.ReturnReqd = False
                TxtCrimeNo.Visible = True
                TxtComments1.Visible = True
                TxtCrimeNo.Value = ""
                TxtComments1.Value = ""
                LblText2.Visible = True
                LblText2.Caption = "Please provide any information about the theft and provide " _
                                    & "the incident number or crime no"
                LblText3.Visible = True
          
            Case Is = UsedConsumed
                Lineitem.ReturnReqd = False
                TxtComments1.Visible = False
                TxtCrimeNo.Value = ""
                TxtComments1.Value = ""
                LblText2.Visible = False
                LblText2.Caption = ""
                LblText3.Visible = False
                LblText3.Caption = ""
       
        End Select

    End If
Exit Sub

ErrorExit:

    FormTerminate
    Terminate

Exit Sub

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' TxtComments1_Change
' event for change of comments box
' ---------------------------------------------------------------
Private Sub TxtComments1_Change()
    TxtComments1.BackColor = COLOUR_3
End Sub

' ===============================================================
' TxtCrimeNo_Change
' event for change of crime no
' ---------------------------------------------------------------
Private Sub TxtCrimeNo_Change()
    TxtCrimeNo.BackColor = COLOUR_3
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
' initialises controls on form at start up
' ---------------------------------------------------------------
Private Function FormInitialise() As Boolean
    Const StrPROCEDURE As String = "FormInitialise()"
    Dim i As Integer
    
    On Error GoTo ErrorHandler

    With LblText1
        .Visible = True
    End With
    
    
    LblText2.Visible = False
    LblText3.Visible = False
    TxtComments1.Visible = False
    TxtCrimeNo.Visible = False

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
' ValidateForm
' Ensures the form is filled out correctly before moving on
' ---------------------------------------------------------------
Private Function ValidateForm() As EnumFormValidation
    Dim Reason As EnumReqReason
    
    Const StrPROCEDURE As String = "ValidateForm()"
 
    On Error GoTo ErrorHandler

    With CmoReason
        If .ListIndex = -1 Then
            .BackColor = COLOUR_6
            ValidateForm = ValidationError
        End If
    End With
    
    With TxtCrimeNo
        If .Visible = True And Trim(.Value) = "" Then
            .BackColor = COLOUR_6
            ValidateForm = ValidationError
        End If
    End With

    With TxtComments1
        If .Visible = True And Trim(.Value) = "" Then
            .BackColor = COLOUR_6
            ValidateForm = ValidationError
        End If
    End With
            
    If ValidateForm = ValidationError Then
        Err.Raise FORM_INPUT_EMPTY
    Else
        ValidateForm = FormOK
    End If

Exit Function

ValidationError:

    ValidateForm = ValidationError

Exit Function

ErrorExit:

    ValidateForm = FunctionalError
    FormTerminate
    Terminate

Exit Function

ErrorHandler:
    
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        CustomErrorHandler Err.Number
        Resume ValidationError
    End If

If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function
' ===============================================================
' SelectPrevForm
' Selects the Prev form to display from Allocation type
' ---------------------------------------------------------------
Private Function SelectPrevForm() As Boolean
    Const StrPROCEDURE As String = "SelectPrevForm()"

    On Error GoTo ErrorHandler
    
    If Lineitem Is Nothing Then Err.Raise NO_LINE_ITEM, Description:=" No Line Item available"

    If Lineitem.Asset Is Nothing Then Err.Raise NO_ASSET_ON_ORDER, Description:="No Asset is on the order"
    
    Select Case Lineitem.Asset.AllocationType
        Case Is = Person
            
            Hide
            If Not FrmPerson.ShowForm(Lineitem) Then Err.Raise HANDLED_ERROR
            Unload Me
            
            Hide
        Case Is = Vehicle
            Hide
            If Not FrmVehicle.ShowForm(Lineitem) Then Err.Raise HANDLED_ERROR
            Unload Me
            
        Case Is = Station
            If Not FrmStation.ShowForm(Lineitem) Then Err.Raise HANDLED_ERROR
            Unload Me
    
    End Select

    SelectPrevForm = True

Exit Function

ErrorExit:

    FormTerminate
    Terminate
    SelectPrevForm = False
Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildLossReport
' Adds data to loss report
' ---------------------------------------------------------------
Private Function BuildLossReport() As Boolean
    Const StrPROCEDURE As String = "BuildLossReport()"

    Dim Reason As EnumReqReason
    
    On Error GoTo ErrorHandler

    With CmoReason
        Reason = .List(.ListIndex, 1)
    End With
    
    With Lineitem
        .ReqReason = Reason
    End With
    
    With Lineitem.LossReport
        Select Case Reason
        
            Case Is = DamagedOpTraining
                
                .AdditInfo = TxtComments1.Visible
                .Status = RepOpen
                .DBSave
            
            Case Is = DamagedOther
            
                .AdditInfo = TxtComments1.Visible
                .Status = RepOpen
                .DBSave
            
            Case Is = lost
                .ActionsTaken = TxtComments1.Value
                .Status = RepOpen
                .DBSave
           
            Case Is = Malfunction
                .AdditInfo = TxtComments1.Visible
                .Status = RepOpen
                .DBSave
          
            Case Is = Stolen
                .Theft = True
                .ActionsTaken = TxtComments1.Value
                .Status = RepOpen
                .IncNo = TxtCrimeNo
                .DBSave
                  
        End Select
        .ReportDate = Format(Now, "dd/mm/yy")
    End With

    BuildLossReport = True

Exit Function

ErrorExit:

    FormTerminate
    Terminate
    BuildLossReport = False
    
    

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

