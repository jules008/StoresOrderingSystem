VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmDeliveryList 
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14520
   OleObjectBlob   =   "FrmDeliveryList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmDeliveryList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 19 Jul 17
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmDeliveryList"

Private Supplier As ClsSupplier
Private Deliveries As ClsDeliveries

' ===============================================================
' ShowForm
' Initial entry point to form
' ---------------------------------------------------------------
Public Function ShowForm(LocSupplier As ClsSupplier) As Boolean
    
    Const StrPROCEDURE As String = "ShowForm()"
    
    On Error GoTo ErrorHandler
                
    Set Supplier = LocSupplier
    
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
' Populates form controls
' ---------------------------------------------------------------
Private Function PopulateForm() As Boolean
    Dim Delivery As ClsDelivery
    
    Const StrPROCEDURE As String = "PopulateForm()"
    
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    Set Deliveries = Nothing
    
    Set Deliveries = New ClsDeliveries
    
    Deliveries.GetCollection Supplier.SupplierID
    
    i = 0
    LstDeliveries.Clear
    For Each Delivery In Deliveries
    
        With LstDeliveries
            .AddItem
            .List(i, 0) = Delivery.DeliveryNo
            .List(i, 1) = Format(Delivery.DeliveryDate, "dd/m/yy")
            .List(i, 2) = Delivery.Asset.Description
            .List(i, 3) = Delivery.Quantity
            .List(i, 4) = Delivery.Asset.Size1
            .List(i, 5) = Delivery.Asset.Size2
            i = i + 1
        End With
    Next
    
    Set Delivery = Nothing
    
    PopulateForm = True

Exit Function

ErrorExit:
    
    Set Delivery = Nothing
        
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
    
    Set Supplier = Nothing
    Set Deliveries = Nothing
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
' BtnViewDelivery_Click
' View selected Delivery
' ---------------------------------------------------------------
Private Sub BtnViewDelivery_Click()
    Dim LocDeliveries As ClsDeliveries
    Dim DeliveryDate As Date

    Const StrPROCEDURE As String = "BtnViewDelivery_Click()"

    On Error GoTo ErrorHandler

    With LstDeliveries
        If .ListCount = 0 Then Exit Sub
        If .ListIndex = -1 Then Err.Raise NO_ITEM_SELECTED
        
        DeliveryDate = .List(.ListIndex, 1)
    End With
    
    Set LocDeliveries = New ClsDeliveries
    
    LocDeliveries.GetCollection Supplier.SupplierID, DeliveryDate
    
    If Not FrmDelivery.ShowForm(LocDeliveries) Then Err.Raise HANDLED_ERROR
    

GracefulExit:

    Set LocDeliveries = Nothing

Exit Sub

ErrorExit:
    Set LocDeliveries = Nothing
'    ***CleanUpCode***

Exit Sub

ErrorHandler:
    
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        CustomErrorHandler Err.Number
        Resume GracefulExit:
    End If


    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
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

    Set Deliveries = New ClsDeliveries
    
    With LstHeading
        .AddItem
        .List(0, 0) = "Delivery No"
        .List(0, 1) = "Date"
        .List(0, 2) = "Description"
        .List(0, 3) = "Quantity"
        .List(0, 4) = "Size 1"
        .List(0, 5) = "Size 2"
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
' ValidateForm
' Ensures the form is filled out correctly before moving on
' ---------------------------------------------------------------
Private Function ValidateForm() As EnumFormValidation
    Const StrPROCEDURE As String = "ValidateForm()"

    On Error GoTo ErrorHandler

    If ValidateForm <> ValidationError Then
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
        Resume ValidationError:
    End If

If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

