VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmVehicle 
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10380
   OleObjectBlob   =   "FrmVehicle.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmVehicle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
' v0,0 - Initial version
' v0,1 - Changes for Phone Order Functionality
'---------------------------------------------------------------
' Date - 16 Apr 17
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmVehicle"

Private Lineitem As ClsLineItem

' ===============================================================
' ShowForm
' Initial entry point to form
' ---------------------------------------------------------------
Public Function ShowForm(Optional LocLineItem As ClsLineItem) As Boolean
    
    Const StrPROCEDURE As String = "ShowForm()"
    
    On Error GoTo ErrorHandler
    
    If LocLineItem Is Nothing Then
        Err.Raise NO_LINE_ITEM, Description:="No LineItem Passed to ShowForm Function"
    Else
        Set Lineitem = LocLineItem
        If Not PopulateForm Then Err.Raise HANDLED_ERROR
    End If
    Debug.Print Me.Visible
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

    Dim Vehicle As ClsVehicle
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    With LblText1
        .Visible = True
    End With
    
    With LstVehicles
        .Clear
        If Not CurrentUser.Vehicle Is Nothing Then
    
            .AddItem
            .List(0, 0) = CurrentUser.Vehicle.VehNo
            .List(0, 1) = CurrentUser.Vehicle.VehReg
            .List(0, 2) = CurrentUser.Vehicle.CallSign
            .List(0, 3) = CurrentUser.Vehicle.VehicleMake
            .List(0, 4) = CurrentUser.Vehicle.GetVehicleTypeString
            i = i + 1
        Else
        
            For Each Vehicle In Lineitem.Parent.Requestor.Station.Vehicles
                If Vehicle.CallSign <> "" Then
                    .AddItem
                    .List(i, 0) = Vehicle.VehNo
                    .List(i, 1) = Vehicle.VehReg
                    .List(i, 2) = Vehicle.CallSign
                    .List(i, 3) = Vehicle.VehicleMake
                    .List(i, 4) = Vehicle.GetVehicleTypeString
                    i = i + 1
                End If
            Next
        End If
            
        .AddItem
        .List(i, 1) = "Other...."
        
        If .ListCount > 1 Then .ListIndex = 0
                
    End With
        
    LblText2.Visible = False
    CmoVehicleTypes.Visible = False
    LstOtherVehs.Visible = False
    
    Set Vehicle = Nothing
    PopulateForm = True

Exit Function

ErrorExit:
    Set Vehicle = Nothing
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
                                
                If Lineitem.ForVehicle Is Nothing Then Err.Raise NO_VEHICLE_SELECTED
                
                Hide
                If Not FrmLossReport.ShowForm(Lineitem) Then Err.Raise HANDLED_ERROR
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

    Hide
    If Not FrmCatSearch.ShowForm(Lineitem) Then Err.Raise HANDLED_ERROR
    Unload Me
    
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
' CmoVehicleTypes_Change
' Event processing vehicle type list
' ---------------------------------------------------------------
Private Sub CmoVehicleTypes_Change()
    Dim SelVehType As EnumVehType
    Dim Vehicle As ClsVehicle
    Dim i As Integer

    Const StrPROCEDURE As String = "CmoVehicleTypes_Change()"
      
    On Error GoTo ErrorHandler
    
    CmoVehicleTypes.BackColor = COLOUR_3
    LstOtherVehs.BackColor = COLOUR_3
    
    If CmoVehicleTypes <> "" Then
        LstOtherVehs.Visible = True
        LblText3.Visible = True
            
        
        With CmoVehicleTypes
            If .ListIndex <> -1 Then SelVehType = .List(.ListIndex, 1)
        End With
        
        i = 0
        With LstOtherVehs
            
            .Clear
            For Each Vehicle In Vehicles
                
                If Vehicle.VehType = SelVehType Then
                    .AddItem
                    .List(i, 0) = Vehicle.VehNo
                    .List(i, 1) = Vehicle.VehReg
                    .List(i, 2) = Vehicle.CallSign
                    .List(i, 3) = Vehicle.VehicleMake
                    .List(i, 4) = Vehicle.GetVehicleTypeString
                    i = i + 1
                End If
            Next
        End With
    
        Set Vehicle = Nothing
    End If
Exit Sub

ErrorExit:

    FormTerminate
    Terminate
    Set Vehicle = Nothing

Exit Sub

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' LstOtherVehs_Click
' Selects vehicle off other list
' ---------------------------------------------------------------
Private Sub LstOtherVehs_Click()
    Const StrPROCEDURE As String = "LstOtherVehs_Click()"

    On Error GoTo ErrorHandler
        
    With LstOtherVehs
        .BackColor = COLOUR_3
        Lineitem.ForVehicle = Vehicles(.List(.ListIndex, 0))
    End With

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
' LstVehicles_Click
' Event processing for vehicle list
' ---------------------------------------------------------------
Private Sub LstVehicles_Click()
    
    Const StrPROCEDURE As String = "LstVehicles_Click()"
    
    Dim VehNotShown As Boolean
    Dim RstVehTypes As Recordset
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    Set RstVehTypes = Vehicles.GetVehicleTypes
    
    With LstVehicles
        If .List(.ListIndex, 1) = "Other...." Then
            LblText2.Visible = True
            CmoVehicleTypes.Visible = True
            VehNotShown = True
            Lineitem.ForVehicle = Nothing
        Else
            LblText2.Visible = False
            CmoVehicleTypes.Visible = False
            LstOtherVehs.Visible = False
            LblText3.Visible = False
            VehNotShown = False
            
            Lineitem.ForVehicle = Vehicles(.List(.ListIndex, 0))
            
        End If
    
    End With

    If VehNotShown Then
    
        With CmoVehicleTypes
            
            i = 0
            RstVehTypes.MoveFirst
            Do Until RstVehTypes.EOF
                
                .AddItem
                .List(i, 0) = RstVehTypes!VehicleType
                .List(i, 1) = RstVehTypes!VehicleTypeNo
                RstVehTypes.MoveNext
                i = i + 1
            Loop
        End With
        
        
    End If
    

    Set RstVehTypes = Nothing

Exit Sub

ErrorExit:

    FormTerminate
    Terminate
    Set RstVehTypes = Nothing
    
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
' FormInitialise
' initialises controls on form at start up
' ---------------------------------------------------------------
Private Function FormInitialise() As Boolean
    Const StrPROCEDURE As String = "FormInitialise()"

    On Error GoTo ErrorHandler

    LblText1.Visible = False
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
    Const StrPROCEDURE As String = "ValidateForm()"

    On Error GoTo ErrorHandler
    
    With CmoVehicleTypes
        If .Visible = True And .ListIndex = -1 Then
            .BackColor = COLOUR_6
            ValidateForm = ValidationError
        End If
    End With
    
    With LstOtherVehs
        If .Visible = True And .ListIndex = -1 Then
            .BackColor = COLOUR_6
            ValidateForm = ValidationError
        End If
    End With

    If ValidateForm = ValidationError Then
        Err.Raise FORM_INPUT_EMPTY
    Else
        ValidateForm = FormOK
    End If

GracefulExit:
Exit Function

ErrorExit:

    ValidateForm = FunctionalError
    FormTerminate
    Terminate

Exit Function

ErrorHandler:
    
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        CustomErrorHandler Err.Number
        Resume GracefulExit
    End If

If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function



