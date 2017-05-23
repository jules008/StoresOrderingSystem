VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmDeliveryCxs 
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9870
   OleObjectBlob   =   "FrmDeliveryCxs.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmDeliveryCxs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 23 May 17
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmDeliveryCxs"

Private Deliveries As ClsDeliveries

' ===============================================================
' ShowForm
' Initial entry point to form
' ---------------------------------------------------------------
Public Function ShowForm(LocDeliveries As ClsDeliveries) As Boolean
    
    Const StrPROCEDURE As String = "ShowForm()"
    
    On Error GoTo ErrorHandler
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
' BtnAbandon_Click
' Close form without changing
' ---------------------------------------------------------------
Private Sub BtnAbandon_Click()
    FormTerminate
End Sub

' ===============================================================
' BtnApply_Click
' Apply stock changes to DB
' ---------------------------------------------------------------
Private Sub BtnApply_Click()
    Dim Delivery As ClsDelivery
    Dim Asset As ClsAsset
    
    Const StrPROCEDURE As String = "BtnApply_Click()"

    On Error GoTo ErrorHandler

    Set Asset = New ClsAsset
    Set Delivery = New ClsDelivery
    
    For Each Delivery In Deliveries
        With Asset
        .DBGet Delivery.AssetNo
        .QtyInStock = .QtyInStock + Delivery.Quantity
        .DBSave
        End With
    Next

    MsgBox "Stock has been updated successfully", vbInformation, APP_NAME
    Unload Me
    Set Asset = Nothing
    Set Delivery = Nothing
Exit Sub

ErrorExit:
    Set Asset = Nothing
    Set Delivery = Nothing
    
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
' FormInitialise
' initialises controls on form at start up
' ---------------------------------------------------------------
Private Function FormInitialise() As Boolean
    Const StrPROCEDURE As String = "FormInitialise()"

    On Error GoTo ErrorHandler

    Set Deliveries = New ClsDeliveries
    
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
' Populate form details
' ---------------------------------------------------------------
Private Function PopulateForm() As Boolean
    Dim Asset As ClsAsset
    Dim Delivery As ClsDelivery
    
    Const StrPROCEDURE As String = "PopulateForm()"

    On Error GoTo ErrorHandler

    Set Asset = New ClsAsset
    Set Delivery = New ClsDelivery
    
    For Each Delivery In Deliveries
    
        Asset.DBGet Delivery.AssetNo
        
        If Asset Is Nothing Then Err.Raise HANDLED_ERROR
        
        With LstStockCxs
            .AddItem "Item " & Left(Asset.Description, 30) & " stock will change from " _
                            & Asset.QtyInStock & " to " & Asset.QtyInStock + Delivery.Quantity
    
        End With
    Next

    PopulateForm = True
    Set Asset = Nothing
    Set Delivery = Nothing
    
Exit Function

ErrorExit:
    Set Asset = Nothing
    Set Delivery = Nothing
'    ***CleanUpCode***
    PopulateForm = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function
