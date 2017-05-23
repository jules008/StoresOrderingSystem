VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmDelivery 
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9525
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
'---------------------------------------------------------------
' Date - 23 May 17
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmDelivery"

Private Deliveries As ClsDeliveries
Private Asset As ClsAsset

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
    Set Asset = Nothing
    Unload Me

End Function

' ===============================================================
' BtnAdd_Click
' Adds item to delivery
' ---------------------------------------------------------------
Private Sub BtnAdd_Click()
    Dim Delivery As ClsDelivery
    Dim Assets As ClsAssets
    Dim Asset As ClsAsset
    Dim AssetNo As Integer
    Dim i As Integer
    
    Dim Validation As EnumFormValidation
    
    Const StrPROCEDURE As String = "BtnAdd_Click()"

    On Error GoTo ErrorHandler

    Validation = ValidateForm
    
    Select Case Validation
        Case Is = FunctionalError
            Err.Raise HANDLED_ERROR
        Case Is = ValidationError
            Err.Raise FIELDS_INCOMPLETE
    End Select
    
    Set Asset = New ClsAsset
    Set Assets = New ClsAssets
    Set Delivery = New ClsDelivery
    
    AssetNo = Assets.FindAssetNo(TxtSearch, CmoSize1, CmoSize2)
    
    If AssetNo = 0 Then Err.Raise NO_ASSET_FOUND
    
    Asset.DBGet AssetNo
    
    With Delivery
        .AssetNo = Asset.AssetNo
        .AssetDescr = Asset.Description
        .SupplierName = TxtSupplier
        .DeliveryDate = TxtDate
        .Quantity = TxtQty
        .DBSave
    End With

    With LstOrderItems
        i = .ListCount
        .AddItem
        .List(i, 0) = Delivery.DeliveryNo
        .List(i, 1) = Delivery.AssetDescr
        .List(i, 2) = Asset.Size1
        .List(i, 3) = Asset.Size2
        .List(i, 4) = Delivery.Quantity
    End With

    Deliveries.AddItem Delivery
    
    If Not ClearSearch Then Err.Raise HANDLED_ERROR
    
GracefulExit:
    Set Delivery = Nothing
    Set Assets = Nothing
    Set Asset = Nothing
Exit Sub

ErrorExit:
    Set Delivery = Nothing
    Set Assets = Nothing
    Set Asset = Nothing

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
' BtnApply_Click
' Process Delivery
' ---------------------------------------------------------------
Private Sub BtnApply_Click()
    ProcessDelivery
End Sub

' ===============================================================
' BtnClear_Click
' Clears search
' ---------------------------------------------------------------
Private Sub BtnClear_Click()
    TxtSearch = ""
    TxtQty = ""
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
' BtnDatePicker_Click
' Shows date picker form
' ---------------------------------------------------------------
Private Sub BtnDatePicker_Click()
    FrmDatePicker.Show
    TxtDate = Format(FrmDatePicker.Tag, "dd/mm/yy")
    TxtQty.BackColor = COLOUR_3

End Sub

' ===============================================================
' BtnRemove_Click
' Removes selected lineitem
' ---------------------------------------------------------------
Private Sub BtnRemove_Click()
    Dim ItemNo As Integer
    
    Const StrPROCEDURE As String = "BtnRemove_Click()"
    
    On Error GoTo ErrorHandler

    If LstOrderItems.ListCount = 0 Then Exit Sub

    If LstOrderItems.ListIndex = -1 Then Err.Raise NO_ITEM_SELECTED
        
    With LstOrderItems
        ItemNo = .List(.ListIndex, 0)
        Deliveries(CStr(ItemNo)).DBDelete
        Deliveries.RemoveItem CStr(ItemNo)
        
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
' TxtDate_Change
' Change event for date txt box
' ---------------------------------------------------------------
Private Sub TxtDate_Change()
    TxtDate.BackColor = COLOUR_3

End Sub

' ===============================================================
' TxtQty_Change
' Change event for quantity txt box
' ---------------------------------------------------------------
Private Sub TxtQty_Change()
    TxtQty.BackColor = COLOUR_3

End Sub

' ===============================================================
' TxtSearch_Change
' Entry for search string
' ---------------------------------------------------------------
Private Sub TxtSearch_Change()
    Const StrPROCEDURE As String = "TxtSearch_Change()"

    On Error GoTo ErrorHandler

    Dim ListResults As String

    On Error GoTo ErrorHandler
    
    TxtSearch.BackColor = COLOUR_3
    LstAssets.BackColor = COLOUR_3

    With LstAssets
        If .ListIndex <> -1 Then ListResults = .List(.ListIndex)
    
        'if the search box has been changed since being updated by the results box, clear the result box
        If ListResults <> TxtSearch.Value Then .ListIndex = -1
        
        'if the results box has been clicked, add the selected result to the search box
        If .ListIndex = -1 Then
        
            'if no results selected, populate with new results
            If Len(TxtSearch.Value) > 1 Then
                If Not GetSearchItems(TxtSearch.Value) Then Err.Raise HANDLED_ERROR
            Else
                .Clear
            End If
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
' TxtSupplier_Change
' Change event for supplier txt box
' ---------------------------------------------------------------
Private Sub TxtSupplier_Change()
    TxtSupplier.BackColor = COLOUR_3

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
' ClearSearch
' Clears form ready for new item search
' ---------------------------------------------------------------
Private Function ClearSearch() As Boolean
    Const StrPROCEDURE As String = "ClearSearch()"

    On Error GoTo ErrorHandler

    TxtSearch = ""
    TxtQty = ""
    CmoSize1 = ""
    CmoSize2 = ""
    CmoSize1.Visible = False
    CmoSize2.Visible = False
    LblSize1.Visible = False
    LblSize2.Visible = False

    ClearSearch = True

Exit Function

ErrorExit:

    ClearSearch = False

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

    Set Deliveries = New ClsDeliveries

    If Not ClearSearch Then Err.Raise HANDLED_ERROR
    
    If Not ShtLists.RefreshAssetList Then Err.Raise HANDLED_ERROR

    With LstHeading
        .AddItem
        .List(0, 0) = "Asset No"
        .List(0, 1) = "Asset"
        .List(0, 2) = "Size 1"
        .List(0, 3) = "Size 2"
        .List(0, 4) = "Qty"
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

    With TxtSearch
        If .Value = "" Then
            .BackColor = COLOUR_6
            ValidateForm = ValidationError
        End If
    End With
    
    With TxtDate
        If .Value = "" Then
            .BackColor = COLOUR_6
            ValidateForm = ValidationError
        End If
    End With
    
    With LstAssets
        If .ListIndex = -1 Then
            .BackColor = COLOUR_6
            ValidateForm = ValidationError
        End If
    End With
    
    With TxtSupplier
        If .Value = "" Then
            .BackColor = COLOUR_6
            ValidateForm = ValidationError
        End If
    End With

    With TxtQty
        If .Value = "" Then
            .BackColor = COLOUR_6
            ValidateForm = ValidationError
        End If
    End With
    
    With CmoSize1
        If .Visible = True And .Value = "" Then
            .BackColor = COLOUR_6
            ValidateForm = ValidationError
        End If
    End With
    
    With CmoSize2
        If .Visible = True And .Value = "" Then
            .BackColor = COLOUR_6
            ValidateForm = ValidationError
        End If
    End With

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

' ===============================================================
' GetSearchItems
' Gets items from the name list that match Txtsearch box
' ---------------------------------------------------------------
Private Function GetSearchItems(StrSearch As String) As Boolean
    Const StrPROCEDURE As String = "GetSearchItems()"

    On Error GoTo ErrorHandler

    Dim ListLength As Integer
    Dim RngResult As Range
    Dim RngItems As Range
    Dim RngFirstResult As Range
    Dim i As Integer
    
    'get length of item list
    ListLength = Application.WorksheetFunction.CountA(ShtLists.Range("A:A"))
    
    Set RngItems = ShtLists.Range("A1:A" & ListLength)
            
    Set RngResult = RngItems.Find(StrSearch)
    Set RngFirstResult = RngResult
    
    LstAssets.Clear
    'search item list and populate results.  Stop before looping back to start
    If Not RngResult Is Nothing Then
        Do
            Set RngResult = RngItems.FindNext(RngResult)
            LstAssets.AddItem RngResult
        Loop While RngResult <> 0 And RngResult.Address <> RngFirstResult.Address
    End If

    GetSearchItems = True
    
    Set RngItems = Nothing
    Set RngResult = Nothing
    Set RngFirstResult = Nothing
    
Exit Function

ErrorExit:

    Set RngItems = Nothing
    Set RngResult = Nothing
    Set RngFirstResult = Nothing
    
    FormTerminate
    Terminate

    GetSearchItems = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' LstAssets_Click
' Gets items from the name list that match Txtsearch box
' ---------------------------------------------------------------
Private Sub LstAssets_Click()

    On Error Resume Next

    LstAssets.BackColor = COLOUR_3
    
    With LstAssets
        Me.TxtSearch.Value = .List(.ListIndex)
        
    End With
                
    If Not ItemChange Then Err.Raise HANDLED_ERROR
                
    TxtQty.SetFocus
End Sub

' ===============================================================
' CmoSize1_Change
' Event when entry changes
' ---------------------------------------------------------------
Private Sub CmoSize1_Change()
    Dim Assets As ClsAssets
    Dim StrSize2Arry() As String
    Dim i As Integer
    
    Const StrPROCEDURE As String = "CmoSize1_Change()"

    On Error GoTo ErrorHandler
   
    Set Assets = New ClsAssets
    
    CmoSize1.BackColor = COLOUR_3
    
    If CmoSize1 = "" Then CmoSize2 = ""

    StrSize2Arry() = Assets.GetSizeLists(TxtSearch, 2)

    If UBound(StrSize2Arry) <> LBound(StrSize2Arry) Then
        LblSize2.Visible = True
        CmoSize2.Visible = True
        
        CmoSize2.Clear
        For i = LBound(StrSize2Arry) To UBound(StrSize2Arry)
            CmoSize2.AddItem StrSize2Arry(i)
        Next
    End If

    Set Assets = Nothing
Exit Sub

ErrorExit:

    Set Assets = Nothing
    FormTerminate
    Terminate

Exit Sub

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE, True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' CmoSize2_Change
' Event when entry changes
' ---------------------------------------------------------------
Private Sub CmoSize2_Change()
    Dim Assets As ClsAssets
    
    Const StrPROCEDURE As String = "CmoSize1_Change()"

    On Error GoTo ErrorHandler
   
    Set Assets = New ClsAssets
    
    CmoSize2.BackColor = COLOUR_3
    
    Set Assets = Nothing

Exit Sub

ErrorExit:

    Set Assets = Nothing
    FormTerminate
    Terminate

Exit Sub

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE, True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' ItemChange
' Event when item changes
' ---------------------------------------------------------------
Private Function ItemChange() As Boolean
    Dim LocAsset As ClsAsset
    Dim AssetNo As Integer
    Dim Assets As ClsAssets
    Dim StrSize1Arry() As String
    Dim StrSize2Arry() As String
    Dim i As Integer
    
    Const StrPROCEDURE As String = "ItemChange()"

    On Error GoTo ErrorHandler
        
    TxtSearch.BackColor = COLOUR_3
    
    If TxtSearch <> "" Then
        
        Set Assets = New ClsAssets
        
        Set LocAsset = New ClsAsset
        StrSize1Arry() = Assets.GetSizeLists(TxtSearch, 1)
        StrSize2Arry() = Assets.GetSizeLists(TxtSearch, 2)
        
        CmoSize1.Clear
        CmoSize2.Clear
        CmoSize1.Visible = False
        CmoSize2.Visible = False
        LblSize1.Visible = False
        LblSize2.Visible = False
        TxtQty = ""
        
        If UBound(StrSize1Arry) <> LBound(StrSize1Arry) Then
            LblSize1.Visible = True
            CmoSize1.Visible = True
            
            For i = LBound(StrSize1Arry) To UBound(StrSize1Arry)
                CmoSize1.AddItem StrSize1Arry(i)
            Next
        End If
        
        If UBound(StrSize2Arry) <> LBound(StrSize2Arry) Then
            LblSize2.Visible = True
            CmoSize2.Visible = True
            
            For i = LBound(StrSize2Arry) To UBound(StrSize2Arry)
                CmoSize2.AddItem StrSize2Arry(i)
            Next
        End If
        
        LocAsset.DBGet (Assets.FindAssetNo(TxtSearch, CmoSize1, CmoSize2))
        
        Set LocAsset = Nothing
        Set Assets = Nothing
    End If
    
    ItemChange = True
    
GracefulExit:

Exit Function

ErrorExit:

    ItemChange = False
    
    Set LocAsset = Nothing
    Set Assets = Nothing
    
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

' ===============================================================
' ProcessDelivery
' Adds delivery items to stock
' ---------------------------------------------------------------
Private Function ProcessDelivery() As Boolean
    Dim Response As Integer
    
    Const StrPROCEDURE As String = "ProcessDelivery()"

    On Error GoTo ErrorHandler

    If Deliveries Is Nothing Then Exit Function
    
    If Deliveries.Count = 0 Then Exit Function
    
        Response = MsgBox("Do you want to process the delivery and alter stock details? " & Chr(13) & Chr(13) _
                            & "Yes - Alter stock" & Chr(13) _
                            & "No - Save delivery record only, no stock changes", vbInformation + vbYesNo, APP_NAME)
    
        If Response = 6 Then

            If Not FrmDeliveryCxs.ShowForm(Deliveries) Then Err.Raise HANDLED_ERROR
            Set Deliveries = Nothing
            LstOrderItems.Clear
        Else
            MsgBox "Delivery added to database", vbInformation
            Unload Me
            
        End If

    ProcessDelivery = True

Exit Function

ErrorExit:

'    ***CleanUpCode***
    ProcessDelivery = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

