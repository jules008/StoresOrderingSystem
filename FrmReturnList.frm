VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmReturnList 
   Caption         =   "Return Items"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9915
   OleObjectBlob   =   "FrmReturnList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmReturnList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
' v0,0 - Initial version
' v0,1 - Return Multiple items
'---------------------------------------------------------------
' Date - 09 Jan 19
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmReturnList"

Public ReturnOrder As ClsOrder

' ===============================================================
' ShowForm
' Initial entry point to form
' ---------------------------------------------------------------
Public Function ShowForm() As Boolean
    
    Const StrPROCEDURE As String = "ShowForm()"
    
    On Error GoTo ErrorHandler
       
    Show
    ShowForm = True
    
Exit Function

ErrorExit:
    
    Unload Me
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
    
    Set ReturnOrder = Nothing
    Unload Me

End Function

' ===============================================================
' BtnAdd_Click
' Adds item to delivery
' ---------------------------------------------------------------
Private Sub BtnAdd_Click()
    Dim AssetNo As Integer
    Dim i As Integer
    
    Const StrPROCEDURE As String = "BtnAdd_Click()"

    On Error GoTo ErrorHandler

    If Not FrmReturnStock.ShowForm Then Err.Raise HANDLED_ERROR
    
    If Not PopulateForm Then Err.Raise HANDLED_ERROR
    
GracefulExit:

Exit Sub

ErrorExit:
        

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
' Event for page close button
' ---------------------------------------------------------------
Private Sub BtnClose_Click()

    On Error Resume Next
        
    FormTerminate
    
End Sub

' ===============================================================
' BtnRemove_Click
' Removes selected Lineitem
' ---------------------------------------------------------------
Private Sub BtnRemove_Click()
    Dim ItemNo As Integer
    
    Const StrPROCEDURE As String = "BtnRemove_Click()"
    
    On Error GoTo ErrorHandler

    If LstReturnList.ListCount = 0 Then Exit Sub

    If LstReturnList.ListIndex = -1 Then Err.Raise NO_ITEM_SELECTED
        
    With LstReturnList
        ItemNo = .List(.ListIndex, 0)
        ReturnOrder.Lineitems(CStr(ItemNo)).DBDelete
        ReturnOrder.Lineitems.RemoveItem CStr(ItemNo)
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

    Set ReturnOrder = New ClsOrder
    
    With LstHeading
        .AddItem
        .List(0, 0) = "Order No"
        .List(0, 1) = "From"
        .List(0, 2) = "Asset"
        .List(0, 3) = "Qty"
    End With
    
    FormInitialise = True

Exit Function

ErrorExit:

    FormTerminate
    
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
' BtnReturn_Click
' Moves onto next form
' ---------------------------------------------------------------
Private Sub BtnReturn_Click()
    Dim StrUserName As String
    
    Const StrPROCEDURE As String = "BtnReturn_Click()"

    On Error GoTo ErrorHandler

        Select Case ValidateForm
    
            Case Is = FunctionalError
                Err.Raise HANDLED_ERROR
            
            Case Is = FormOK
                
                With RetOrder
                    .OrderDate = Now
                    .Requestor = CurrentUser
                    .Lineitems(1).Quantity = 0 - TxtQty
                    .Lineitems(1).ReqReason = ItemReturn
                    .Status = OrderClosed
                    .DBSave
                End With
                
                MsgBox "Return has been successfully processed", vbOKCancel + vbInformation, APP_NAME
                Hide
                Unload Me
                 
        End Select
        
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
' ProcessReturn
' Creates a negative order to return stock to Stores
' ---------------------------------------------------------------
Private Function ProcessReturn() As Boolean
    Dim RetOrder As ClsOrder
    Dim RetLineItem As ClsLineItem
    Dim Assets As ClsAssets
    Dim AssetNo As Integer
    
    Const StrPROCEDURE As String = "ProcessReturn()"

    On Error GoTo ErrorHandler

    With RetLineItem
        .Asset = Assets.FindItem(AssetNo)
    End With
    
    With RetOrder
        
    
    
    End With



    Set RetOrder = Nothing
    Set RetLineItem = Nothing
    Set Assets = Nothing
    
    ProcessReturn = True


Exit Function

ErrorExit:
    
    Set RetOrder = Nothing
    Set RetLineItem = Nothing
    Set Assets = Nothing
    
    '***CleanUpCode***
    ProcessReturn = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' PopulateForm
' Lists return items
' ---------------------------------------------------------------
Private Function PopulateForm() As Boolean
    Dim LineItem As ClsLineItem
    Dim i As Integer
    
    Const StrPROCEDURE As String = "PopulateForm()"

    On Error GoTo ErrorHandler

    i = 0
    With LstReturnList
        .Clear
        For Each LineItem In ReturnOrder.Lineitems
            .AddItem
            .List(i, 0) = ReturnOrder.OrderNo
            If LineItem.Asset.AllocationType = Person Then .List(i, 1) = LineItem.ForPerson.UserName
            If LineItem.Asset.AllocationType = Station Then .List(i, 1) = LineItem.ForStation.Name
            If LineItem.Asset.AllocationType = Vehicle Then .List(i, 1) = LineItem.ForVehicle.VehReg
            .List(i, 2) = LineItem.Asset.Description
            .List(i, 3) = LineItem.Quantity
            i = i + 1

        Next
    End With

    PopulateForm = True


Exit Function

ErrorExit:

    '***CleanUpCode***
    PopulateForm = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function
