VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmDBOrder 
   Caption         =   "F402"
   ClientHeight    =   10260
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16380
   OleObjectBlob   =   "FrmDBOrder.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmDBOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'===============================================================
' v0,0 - Initial version
' v0,1 - Auto assign Order and improved printing
' v0,2 - Add Delete Order functionality
' v0,3 - Fix Order Form double click issue
' v0,4 - Restrict view for level 1
' v0,5 - Add option to print order to PDF
' v0,6 - Add Print Date and removed PDF button
' v0,7 - Change delete to use .DBDelete
' v0,8 - Resized for laptop screens
'---------------------------------------------------------------
' Date - 26 Apr 18
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmDBOrder"

Private Order As ClsOrder

' ===============================================================
' ShowForm
' Initial entry point to form
' ---------------------------------------------------------------
Public Function ShowForm(Optional LocOrder As ClsOrder) As Boolean
    
    Const StrPROCEDURE As String = "ShowForm()"
    
    On Error GoTo ErrorHandler
    
    If LocOrder Is Nothing Then
        Err.Raise NO_ORDER, Description:="Order not available"
    Else
        Set Order = LocOrder
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
    Dim Lineitem As ClsLineItem
    
    On Error GoTo ErrorHandler
    
    Set Lineitem = New ClsLineItem

    If Not ProcessStatus Then Err.Raise HANDLED_ERROR

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
        .List(0, 8) = "Status"
    End With
    
    With LstItems
        .Clear
        
        i = 0
        For Each Lineitem In Order.Lineitems
            .AddItem
            .List(i, 0) = Lineitem.LineitemNo
            .List(i, 1) = i + 1
            .List(i, 2) = Lineitem.Asset.Description
            .List(i, 3) = Lineitem.Quantity
            .List(i, 4) = Lineitem.Asset.Size1
            .List(i, 5) = Lineitem.Asset.Size2
            If Lineitem.ReturnReqd = True Then .List(i, 6) = "Yes" Else .List(i, 6) = "No"
            If Lineitem.LossReport.LossReportNo <> 0 Then .List(i, 7) = "Yes" Else .List(i, 6) = "No"
            .List(i, 8) = Lineitem.ReturnLineitemStatus
            i = i + 1
        Next
    End With
    
    If Order.Requestor Is Nothing Then Err.Raise NO_REQUESTOR, Description:="No requestor available"
    
    With Order.Requestor
        TxtCrewNo = .CrewNo
        TxtName = .UserName
        TxtRole = .RankGrade
        TxtWatch = .Watch
        TxtWorkplace = .Station.Name
    End With
        
    If Order Is Nothing Then Err.Raise NO_ORDER, Description:="System failure, no Order"
    
    With Order
        TxtOrderNo = .OrderNo
        TxtOrderDate = Format(.OrderDate, "dd/mm/yy")
        CmoStatus.ListIndex = .Status
        TxtAssignedTo = .AssignedTo.UserName
        TxtOrderNote = .OrderNote
        
        If Not .PrintedDate = 0 Then TxtPrinted = "Printed on " & .PrintedDate Else TxtPrinted = ""
    
    End With
    
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
    Const StrPROCEDURE As String = "FormTerminate()"

    On Error GoTo ErrorHandler

    Set Order = Nothing
    
    Unload Me
    

    FormTerminate = True

Exit Function

ErrorExit:

    FormTerminate = False
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
' BtnAssignToMe_Click
' Assign Order to me
' ---------------------------------------------------------------
Private Sub BtnAssignToMe_Click()
    Const StrPROCEDURE As String = "BtnAssignToMe_Click()"

    On Error GoTo ErrorHandler

    With Order
        .Status = OrderAssigned
        .AssignedTo = CurrentUser
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
' BtnClose_Click
' Closes the form
' ---------------------------------------------------------------
Private Sub BtnClose_Click()
    Const StrPROCEDURE As String = "BtnClose_Click()"

    On Error GoTo ErrorHandler

    If Not FormTerminate Then Err.Raise HANDLED_ERROR

Exit Sub

ErrorExit:

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
' BtnCloseOrder_Click
' Checks status and closes order
' ---------------------------------------------------------------
Private Sub BtnCloseOrder_Click()
    Const StrPROCEDURE As String = "BtnCloseOrder_Click()"

    On Error GoTo ErrorHandler

    With Order
        If .Status = OrderIssued Or .Status = OrderClosed Then
            .Status = OrderClosed
            .DBSave
        Else
            MsgBox "Unable to close Order until all Items are completed", vbOKOnly + vbInformation, APP_NAME
        End If
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
' BtnDelete_Click
' Marks order as deleted
' ---------------------------------------------------------------
Private Sub BtnDelete_Click()
    Dim Response As String
    
    Const StrPROCEDURE As String = "BtnDelete_Click()"

    On Error GoTo ErrorHandler

    Response = MsgBox("Are you sure you want to mark the Order as deleted?", vbYesNo + vbDefaultButton2 + vbExclamation, APP_NAME)

    If Response = 6 Then
        With Order
            .Status = OrderDeleted
            .DBSave
            .DBDelete
            If Not ProcessStatus Then Err.Raise HANDLED_ERROR
            If Not PopulateForm Then Err.Raise HANDLED_ERROR
        End With
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
' BtnDeleteItem_Click
' Deletes line item
' ---------------------------------------------------------------
Private Sub BtnDeleteItem_Click()
    Dim Response As Integer
    Dim Lineitem As ClsLineItem
    
    Const StrPROCEDURE As String = "BtnDeleteItem_Click()"

    On Error GoTo ErrorHandler

    Set Lineitem = New ClsLineItem
    
    If LstItems.ListIndex = -1 Then Err.Raise NO_ITEM_SELECTED
    
    Response = MsgBox("Are you sure you want to delete the line?", vbYesNoCancel + vbExclamation, APP_NAME)
    
    If Response <> 6 Then Exit Sub
    
    Set Lineitem = Order.Lineitems(LstItems.ListIndex + 1)

    If Lineitem Is Nothing Then Err.Raise NO_LINE_ITEM, Description:="No Lineitem available"
    
    Lineitem.Parent.Lineitems.RemoveItem CStr(Lineitem.LineitemNo)
    
    Lineitem.DBDelete
    
    If Not PopulateForm Then Err.Raise HANDLED_ERROR

    
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
' BtnLineitem_Click
' opens the selected line item
' ---------------------------------------------------------------
Private Sub BtnLineitem_Click()
    Dim Lineitem As ClsLineItem
    
    Const StrPROCEDURE As String = "BtnLineitem_Click()"
        
    On Error GoTo ErrorHandler

    Set Lineitem = New ClsLineItem
    
    If LstItems.ListIndex = -1 Then Err.Raise NO_ITEM_SELECTED
    
    Set Lineitem = Order.Lineitems(LstItems.ListIndex + 1)

    If Lineitem Is Nothing Then Err.Raise NO_LINE_ITEM, Description:="No Lineitem available"
    
    If Not FrmDBLineItem.ShowForm(Lineitem) Then Err.Raise HANDLED_ERROR
    
    If Not PopulateForm Then Err.Raise HANDLED_ERROR

    
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
' BtnPrint_Click
' Prints order receipt
' ---------------------------------------------------------------
Private Sub BtnPrint_Click()
    Const StrPROCEDURE As String = "BtnPrint_Click()"

    On Error GoTo ErrorHandler
    
    Order.PrintedDate = Now
    Order.DBSave
    
    If Not ModPrint.PrintOrderList(Order, True) Then Err.Raise HANDLED_ERROR
    
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
' CmoStatus_Change
' Changes order status
' ---------------------------------------------------------------
Private Sub CmoStatus_Change()
    Const StrPROCEDURE As String = "CmoStatus_Change()"

    On Error GoTo ErrorHandler
    
    If Order Is Nothing Then Err.Raise NO_ORDER, Description:="System failure, no Order"

    With Order
        If CmoStatus.ListIndex <> -1 Then
            If CmoStatus.ListIndex = 0 Then .AssignedTo = New ClsPerson
            .Status = CmoStatus.ListIndex
            .DBSave
            
            If Not PopulateForm Then Err.Raise HANDLED_ERROR
        End If
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
' TxtOrderNote_Change
' Saves notes to Order Class when updated
' ---------------------------------------------------------------
Private Sub TxtOrderNote_Change()
    Const StrPROCEDURE As String = "TxtOrderNote_Change()"

    On Error GoTo ErrorHandler

    Order.OrderNote = TxtOrderNote
    Order.DBSave

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
' Terminates the form gracefully
' ---------------------------------------------------------------
Private Sub UserForm_Terminate()
    Const StrPROCEDURE As String = "UserForm_Terminate()"

    On Error GoTo ErrorHandler

    If Not FormTerminate Then Err.Raise HANDLED_ERROR

Exit Sub

ErrorExit:

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
' FormInitialise
' initialises controls on form at start up
' ---------------------------------------------------------------
Private Function FormInitialise() As Boolean
    Const StrPROCEDURE As String = "FormInitialise()"
    
    On Error GoTo ErrorHandler
    
    With CmoStatus
        .Clear
        .AddItem "Open"
        .AddItem "Assigned"
        .AddItem "On Hold"
        .AddItem "Issued"
        .AddItem "Closed"
        .AddItem "Deleted"
    End With
    
    If CurrentUser.AccessLvl < StoresLvl_2 Then
        BtnAssignToMe.Visible = False
        BtnDeleteItem.Visible = False
        BtnDelete.Visible = False
        BtnCloseOrder.Visible = False
        BtnPrint.Visible = False
        BtnDeleteItem = False
        CmoStatus.Enabled = False
    End If
    
    LstItems.SetFocus
    FormInitialise = True

Exit Function

ErrorExit:
    
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
' ProcessStatus
' Status processing for order
' ---------------------------------------------------------------
Private Function ProcessStatus() As Boolean
    Dim Lineitem As ClsLineItem
    Dim NoOpen As Integer
    Dim NoOnHold As Integer
    Dim NoIssued As Integer
    Dim NoDelivered As Integer
    Dim NoComplete As Integer
    Dim NoTotal As Integer
    Dim Status As EnumOrderStatus
    
    Const StrPROCEDURE As String = "ProcessStatus()"

    On Error GoTo ErrorHandler

    Set Lineitem = New ClsLineItem
    
    NoTotal = Order.Lineitems.Count
    
    If Order.Status <> OrderClosed And Order.Status <> OrderDeleted Then
        For Each Lineitem In Order.Lineitems
            With Lineitem
                Select Case .Status
                    Case Is = LineComplete
                        NoComplete = NoComplete + 1
                    Case Is = LineDelivered
                        NoDelivered = NoDelivered + 1
                    Case Is = LineIssued
                        NoIssued = NoIssued + 1
                    Case Is = LineOnHold
                        NoOnHold = NoOnHold + 1
                    Case Is = LineOpen
                        NoOpen = NoOpen + 1
                End Select
            End With
            
        Next
        If Order.AssignedTo.CrewNo = "" Then
            Status = OrderOpen
        Else
            Status = OrderAssigned
        End If
        
        If NoOnHold > 0 Then
            If NoOpen = 0 Then Status = OrderOnHold
        End If
        
        If NoIssued + NoComplete + NoDelivered = NoTotal Then Status = OrderIssued
    
        With Order
            .Status = Status
            .DBSave
        End With
    
    End If
    
    ProcessStatus = True

    Set Lineitem = Nothing
Exit Function

ErrorExit:

    Set Lineitem = Nothing
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
