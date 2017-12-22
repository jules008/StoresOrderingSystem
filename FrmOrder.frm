VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmOrder 
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13230
   OleObjectBlob   =   "FrmOrder.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




'===============================================================
' v0,0 - Initial version
' v0,1 - bug fix SendEmailAlerts
' v0,2 - changes for Phone Order functionality
' v0,3 - Improved message boxes
' v0,4 - Fix Error 287 by opening Outlook if it is closed
' v0,5 - Added checks before removing line items
' v0,6 - 287 issue, tried different Outlook detector
' v0,7 - Do not save order if no lineitems
' v0,81 - Centralised the mail messages
'---------------------------------------------------------------
' Date - 19 Dec 17
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmOrder"

Private Order As ClsOrder

' ===============================================================
' ShowForm
' Initial entry point to form
' ---------------------------------------------------------------
Public Function ShowForm(Optional LocOrder As ClsOrder) As Boolean
    
    Const StrPROCEDURE As String = "ShowForm()"
    
    On Error GoTo ErrorHandler
    
    If LocOrder Is Nothing Then
        Set Order = New ClsOrder
        Order.Requestor = CurrentUser
    Else
        Set Order = LocOrder
    End If
    
    With Order
        .OrderDate = Format(Now, "dd/mm/yy")
        .Status = OrderOpen
    End With
    
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
    
    Const StrPROCEDURE As String = "PopulateForm()"

    Dim i As Integer
    Dim Lineitem As ClsLineItem
    
    On Error GoTo ErrorHandler
    
    Set Lineitem = New ClsLineItem
        
    With LstItems
        .Clear
        
        i = 0
        For Each Lineitem In Order.LineItems
            .AddItem
            .List(i, 0) = Lineitem.LineItemNo
            .List(i, 1) = i + 1
            .List(i, 2) = Lineitem.Asset.Description
            .List(i, 3) = Lineitem.Quantity
            .List(i, 4) = Lineitem.Asset.Size1
            .List(i, 5) = Lineitem.Asset.Size2
            If Lineitem.ReturnReqd = True Then .List(i, 6) = "Yes" Else .List(i, 6) = "No"
            If Lineitem.LossReport.LossReportNo <> 0 Then .List(i, 7) = "Yes" Else .List(i, 7) = "No"
            i = i + 1
        Next
    End With
    
    If Order.OrderNo = 0 Then TxtOrderNo = "New" Else TxtOrderNo = Order.OrderNo
    
    Set Lineitem = Nothing
    
    PopulateForm = True

Exit Function

ErrorExit:
    
    Set Lineitem = Nothing
        
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

    Set Order = Nothing
    Unload Me

End Function

' ===============================================================
' BtnCatSearch_Click
' Click event for new category search
' ---------------------------------------------------------------
Private Sub BtnCatSearch_Click()
    Dim Lineitem As ClsLineItem
    
    Const StrPROCEDURE As String = "BtnCatSearch_Click()"

    On Error GoTo ErrorHandler
    
    Set Lineitem = New ClsLineItem
    
    Lineitem.DBSave
    
    Order.LineItems.AddItem Lineitem
    
    If Not FrmCatSearch.ShowForm(Lineitem) Then Err.Raise HANDLED_ERROR

    Set Lineitem = Nothing

Exit Sub

ErrorExit:
    
    Set Lineitem = Nothing
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
' BtnClearAll_Click
' Clears order form
' ---------------------------------------------------------------
Private Sub BtnClearAll_Click()
    Const StrPROCEDURE As String = "BtnClearAll_Click()"

    On Error GoTo ErrorHandler

    With LstItems
        .Clear
    
    End With

    Set Order = Nothing
    Set Order = New ClsOrder



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
' BtnClose_Click
' Event for page close button
' ---------------------------------------------------------------
Private Sub BtnClose_Click()

    On Error Resume Next
    
    FormTerminate
    
End Sub

' ===============================================================
' BtnEditItem_Click
' Edits selected item in list
' ---------------------------------------------------------------
Private Sub BtnEditItem_Click()
    Const StrPROCEDURE As String = "BtnEditItem_Click()"

    Dim LineItemNo As Integer
    
    On Error GoTo ErrorHandler
    
    With LstItems
        If .ListIndex <> -1 Then
            LineItemNo = .List(.ListIndex, 0)
            
            If Not FrmCatSearch.ShowForm(Order.LineItems(CStr(LineItemNo))) Then Err.Raise HANDLED_ERROR
        End If
    End With

Exit Sub

ErrorExit:



Exit Sub

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' BtnPrintRec_Click
' Prints order receipt
' ---------------------------------------------------------------
Private Sub BtnPrintRec_Click()
    Const StrPROCEDURE As String = "BtnPrintRec_Click()"

    On Error GoTo ErrorHandler
    
    Order.DBSave
    
    If Not ModPrint.PrintOrderReceipt(Order) Then Err.Raise HANDLED_ERROR

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
' BtnRemove_Click
' Removes selected lineitem
' ---------------------------------------------------------------
Private Sub BtnRemove_Click()
    Const StrPROCEDURE As String = "BtnRemove_Click()"
    
    Dim LineItemNo As Integer
    
    On Error GoTo ErrorHandler

    If LstItems.ListCount = 0 Then Exit Sub
    If LstItems.ListIndex = -1 Then Err.Raise NO_ITEM_SELECTED

    With LstItems
        LineItemNo = .List(.ListIndex, 0)
        
        With Order
            .LineItems(CStr(LineItemNo)).DBDelete True
            .LineItems.RemoveItem CStr(LineItemNo)
        End With
        
        .RemoveItem (.ListIndex)
        
    End With
    
GracefulExit:

Exit Sub

ErrorExit:
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
' BtnRemoveAll_Click
' Removes all items from list
' ---------------------------------------------------------------
Private Sub BtnRemoveAll_Click()
    Const StrPROCEDURE As String = "BtnRemoveAll_Click()"
    
    Dim i As Integer
    Dim LineItemNo As Integer
    
    On Error GoTo ErrorHandler

    With LstItems
        For i = (.ListCount - 1) To 0 Step -1
        
            LineItemNo = .List(i, 0)
            
            With Order
                .LineItems(CStr(LineItemNo)).DBDelete
                .LineItems.RemoveItem CStr(LineItemNo)
            End With
            .RemoveItem i
        Next
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
' BtnSubmit_Click
' Submits order
' ---------------------------------------------------------------
Private Sub BtnSubmit_Click()
    Const StrPROCEDURE As String = "BtnSubmit_Click()"

    On Error GoTo ErrorHandler

    If Order Is Nothing Then Err.Raise NO_ORDER, Description:="System failure, no Order"
    
    If Order.OrderNo = 0 And Order.LineItems.Count <> 0 Then
        Order.DBSave
        TxtOrderNo = Order.OrderNo
                
        If Order.OrderNo <> 0 Then
            MsgBox "Thank you, your order has been submitted successfully", vbOKOnly + vbInformation, APP_NAME
            
            If Not SendEmailAlerts Then Err.Raise HANDLED_ERROR
            
        Else
            MsgBox "Sorry, there has been an error, Please contact Stores", vbCritical, APP_NAME
        End If
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
' BtnTextSearch_Click
' Click event for new text search
' ---------------------------------------------------------------
Private Sub BtnTextSearch_Click()
    Dim Lineitem As ClsLineItem
    
    Const StrPROCEDURE As String = "BtnTextSearch_Click()"

    On Error GoTo ErrorHandler
    
    Set Lineitem = New ClsLineItem
    Lineitem.DBSave
    
    Order.LineItems.AddItem Lineitem
    
    If Not FrmTextSearch.ShowForm(Lineitem) Then Err.Raise HANDLED_ERROR

    Set Lineitem = Nothing
    
Exit Sub

ErrorExit:

    Set Lineitem = Nothing
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
        .Clear
        .AddItem
        .List(0, 0) = ""
        .List(0, 1) = "Item No"
        .List(0, 2) = "Description"
        .List(0, 3) = "Quantity"
        .List(0, 4) = "Size 1"
        .List(0, 5) = "Size 2"
        .List(0, 6) = "Return?"
        .List(0, 7) = "Loss Report?"
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

    Dim Reason As EnumReqReason
    
    On Error GoTo ErrorHandler
    
    
    ValidateForm = FormOK

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
' AddLineItem
' Adds lineitem to active order
' ---------------------------------------------------------------
Public Function AddLineItem(Lineitem As ClsLineItem) As Boolean
    Const StrPROCEDURE As String = "AddLineItem()"

    On Error GoTo ErrorHandler

    If Order Is Nothing Then Err.Raise NO_ORDER, Description:="Cannot find active Order"
    
    Lineitem.DBSave
    
    If Not PopulateForm Then Err.Raise HANDLED_ERROR
    
    AddLineItem = True

Exit Function

ErrorExit:

    AddLineItem = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' SendEmailAlerts
' Sends alerts to selected users
' ---------------------------------------------------------------
Private Function SendEmailAlerts() As Boolean
    
    Const StrPROCEDURE As String = "SendEmailAlerts()"

    On Error GoTo ErrorHandler

    If Not ModReports.SendEmailReports("New Order Alert", "A new Stores Order has been received from " & Order.Requestor.UserName, EnumNewOrderReceived) Then Err.Raise HANDLED_ERROR
    
    SendEmailAlerts = True

Exit Function

ErrorExit:

'    ***CleanUpCode***

    SendEmailAlerts = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function


