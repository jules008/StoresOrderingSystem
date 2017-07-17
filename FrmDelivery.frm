VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmDelivery 
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12795
   OleObjectBlob   =   "FrmDelivery.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'===============================================================
' v0,0 - Initial version
' v0,1 - Bug fix - Changing size1 does not update size 2 choices
' v0,2 - Re-write for Supplier functionality
' v0,3 - Adapted for Supplier Functionality
'---------------------------------------------------------------
' Date - 17 Jun 17
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmDeliveries"

Private Deliveries As ClsDeliveries

' ===============================================================
' ShowForm
' Initial entry point to form
' ---------------------------------------------------------------
Public Function ShowForm(LocDeliveries As ClsDeliveries) As Boolean
    
    Const StrPROCEDURE As String = "ShowForm()"
    
    On Error GoTo ErrorHandler
               
    If LocDeliveries Is Nothing Then Err.Raise HANDLED_ERROR, , "No Deliveries"
    
    Set Deliveries = LocDeliveries
    
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
' FormTerminate
' Terminates the form gracefully
' ---------------------------------------------------------------
Private Function FormTerminate() As Boolean

    On Error Resume Next
    
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

    With LstHeading
        .AddItem
        .List(0, 0) = "Delivery No"
        .List(0, 1) = "Description"
        .List(0, 2) = "Quantity"
        .List(0, 3) = "Size 1"
        .List(0, 4) = "Size 2"
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
' PopulateForm
' Populates form controls
' ---------------------------------------------------------------
Private Function PopulateForm() As Boolean
    Dim Delivery As ClsDelivery
    
    Const StrPROCEDURE As String = "PopulateForm()"

    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    With LstDeliveries
        .Clear
        
        i = 0
        For Each Delivery In Deliveries
        
            .AddItem
            .List(i, 0) = Delivery.DeliveryNo
            .List(i, 1) = Delivery.Asset.Description
            .List(i, 2) = Delivery.Quantity
            .List(i, 3) = Delivery.Asset.Size1
            .List(i, 4) = Delivery.Asset.Size2
            i = i + 1
            TxtDate = Format(Delivery.DeliveryDate, "dd mmm yy")
        Next
    End With
    
    
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


