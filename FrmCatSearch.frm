VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmCatSearch 
   Caption         =   "Category Search"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6705
   OleObjectBlob   =   "FrmCatSearch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmCatSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===============================================================
' v0,0 - Initial version
' v0,1 - Changes for Phone Order Button
' v0,2 - Bug fix for Phone Order
' v0,3 - update purchase unit when size changed
' v0,4 - add validation to prevent quantity of 0
'---------------------------------------------------------------
' Date - 03 May 17
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmCatSearch"

Private Lineitem As ClsLineItem

' ===============================================================
' ShowForm
' Initial entry point to form
' ---------------------------------------------------------------
Public Function ShowForm(Optional LocLineItem As ClsLineItem) As Boolean
    
    Const StrPROCEDURE As String = "ShowForm()"
    
    On Error GoTo ErrorHandler
    
    If LocLineItem Is Nothing Then
        Set Lineitem = New ClsLineItem
        
        With Lineitem
            .Status = OrderOpen
        End With
        
        BtnPrev.Enabled = False
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
' ValidateForm
' Ensures the form is filled out correctly before moving on
' ---------------------------------------------------------------
Private Function ValidateForm() As EnumFormValidation
    Const StrPROCEDURE As String = "ValidateForm()"

    On Error GoTo ErrorHandler

    With CmoItem
        If .Value = "" Then
            .BackColor = COLOUR_6
            ValidateForm = ValidationError
        End If
    End With
    
    With CmoQuantity
        If .Value = "" Then
            .BackColor = COLOUR_6
            ValidateForm = ValidationError
        End If
        
        If .Value = 0 Then
            .BackColor = COLOUR_6
            ValidateForm = ValidationError
            Err.Raise NO_QUANTITY_ENTERED
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
' PopulateForm
' Populates form if asset already found
' ---------------------------------------------------------------
Private Function PopulateForm() As Boolean
    Const StrPROCEDURE As String = "PopulateForm()"

    On Error GoTo ErrorHandler
    
    If Lineitem.Asset Is Nothing Then Err.Raise NO_ASSET_ON_ORDER, Description:="No Asset on Order"
    
    With Lineitem.Asset
        CmoCategory1 = .Category1
        CmoCategory2 = .Category2
        CmoCategory3 = .Category3
        CmoItem = .Description
        CmoSize1 = .Size1
        CmoSize2 = .Size2
    End With
    
    With Lineitem
        CmoQuantity = .Quantity
    End With
    PopulateForm = True

Exit Function

ErrorExit:

    PopulateForm = False
    FormTerminate
    Terminate

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
' BtnClosePage_Click
' Event for page close button
' ---------------------------------------------------------------
Private Sub BtnClosePage_Click()

    On Error Resume Next
    
    FormTerminate
    
End Sub

' ===============================================================
' BtnNext_Click
' saves order and moves onto next form
' ---------------------------------------------------------------
Private Sub BtnNext_Click()
    Dim Assets As ClsAssets
    
    Const StrPROCEDURE As String = "BtnNext_Click()"
    
    On Error GoTo ErrorHandler
    
    Set Assets = New ClsAssets
    
    Select Case ValidateForm

        Case Is = FunctionalError
            Err.Raise HANDLED_ERROR

        Case Is = FormOK
            
            If Lineitem Is Nothing Then
                Err.Raise NO_LINE_ITEM, Description:=" No Line Item available"
            Else
                With Lineitem
                    .Asset.DBGet (Assets.FindAssetNo(CmoItem, CmoSize1, CmoSize2))
                    
                    If .Asset.AssetNo = 0 Then
                        Err.Raise NO_ASSET_ON_ORDER, Description:="No Asset found"
                    Else
                        If .Asset.NoOrderMessage <> "" Then Err.Raise NO_ORDER_MESSAGE
                        
                        .Quantity = CmoQuantity
                        
                        If Not SelectNextForm Then Err.Raise HANDLED_ERROR
                        
                    End If
                End With
            End If
    End Select

GracefulExit:

    Set Assets = Nothing

Exit Sub
    
ErrorExit:
    
    Set Assets = Nothing
    FormTerminate
    Terminate
    
Exit Sub

ErrorHandler:
    
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        Dim ErrMessage As String
        If Err.Number = NO_ORDER_MESSAGE Then ErrMessage = Lineitem.Asset.NoOrderMessage
        CustomErrorHandler Err.Number, ErrMessage
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
' BtnPrev_Click
' Back to previous screen event
' ---------------------------------------------------------------
Private Sub BtnPrev_Click()
    Const StrPROCEDURE As String = "BtnPrev_Click()"

    On Error GoTo ErrorHandler

    Hide
    If Not FrmTextSearch.ShowForm(Lineitem) Then Err.Raise HANDLED_ERROR
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
' CmoCategory1_Click
' Event handling when first category box selected
' ---------------------------------------------------------------
Private Sub CmoCategory1_Click()
    Dim Assets As ClsAssets
    
    Const StrPROCEDURE As String = "CmoCategory1_Click()"

    On Error GoTo ErrorHandler

    Dim RstCategoryList As Recordset
    
    If CmoCategory1 <> "" Then
        Set Assets = New ClsAssets
        
        CmoCategory2.Clear
        CmoCategory3.Clear
        CmoItem.Clear
        LblSize1.Visible = False
        LblSize2.Visible = False
        
        With CmoSize1
            .Clear
            .Value = ""
            .Visible = False
        End With
        
        With CmoSize2
            .Clear
            .Value = ""
            .Visible = False
        End With
        
        CmoQuantity.Value = ""
        TxtPurchaseUnit = ""
        
        Set RstCategoryList = Assets.GetCategoryLists(CmoCategory1.Value)
        
        If RstCategoryList Is Nothing Then
            CmoCategory2.Visible = False
            LblCategory2.Visible = False
            CmoCategory3.Visible = False
            LblCategory3.Visible = False
        Else
            With RstCategoryList
                .MoveFirst
                Do While Not .EOF
                    CmoCategory2.Visible = True
                    LblCategory2.Visible = True
                    If Not IsNull(.Fields(0)) Then CmoCategory2.AddItem .Fields(0)
                    .MoveNext
                Loop
                
            End With
        End If
        CmoCategory3.Visible = False
        LblCategory3.Visible = False
        
        If Not UpdateItemList Then Err.Raise HANDLED_ERROR
        
        Set RstCategoryList = Nothing
        Set Assets = Nothing
    End If
Exit Sub

ErrorExit:

    Set RstCategoryList = Nothing
    Set Assets = Nothing
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
' CmoCategory2_Click
' Event handling when second category box selected
' ---------------------------------------------------------------
Private Sub CmoCategory2_Click()
    Dim Assets As ClsAssets
    Dim RstCategoryList As Recordset
    
    Const StrPROCEDURE As String = "CmoCategory2_Click()"

    On Error GoTo ErrorHandler

    If CmoCategory2 <> "" Then
    
        Set Assets = New ClsAssets
        
        Set RstCategoryList = Assets.GetCategoryLists(CmoCategory1.Value, CmoCategory2.Value)
        
        CmoCategory3.Clear
        CmoItem.Clear
        LblSize1.Visible = False
        LblSize2.Visible = False
        
        With CmoSize1
            .Clear
            .Value = ""
            .Visible = False
        End With
        
        With CmoSize2
            .Clear
            .Value = ""
            .Visible = False
        End With
        
        CmoQuantity = ""
        TxtPurchaseUnit = ""
        
        If RstCategoryList Is Nothing Then
            CmoCategory3.Visible = False
            LblCategory3.Visible = False
        Else
            With RstCategoryList
                .MoveFirst
                Do While Not .EOF
                    CmoCategory3.Visible = True
                    LblCategory3.Visible = True
                    If Not IsNull(.Fields(0)) Then CmoCategory3.AddItem .Fields(0)
                    .MoveNext
                Loop
            End With
        End If
        
        If Not UpdateItemList Then Err.Raise HANDLED_ERROR
                
        Set RstCategoryList = Nothing
        Set Assets = Nothing
    End If
Exit Sub

ErrorExit:

    Set RstCategoryList = Nothing
    Set Assets = Nothing
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
' CmoCategory3_Click
' Event handling when second category box selected
' ---------------------------------------------------------------
Private Sub CmoCategory3_Click()
    Dim Assets As ClsAssets
    
    Const StrPROCEDURE As String = "CmoCategory3_Click()"

    On Error GoTo ErrorHandler

    If CmoCategory3 <> "" Then
    
        Set Assets = New ClsAssets
        
        CmoItem.Clear
        LblSize1.Visible = False
        LblSize2.Visible = False
       
        With CmoSize1
            .Clear
            .Value = ""
            .Visible = False
        End With
        
        With CmoSize2
            .Clear
            .Value = ""
            .Visible = False
        End With
        
        CmoQuantity = ""
        TxtPurchaseUnit = ""
               
        If Not UpdateItemList Then Err.Raise HANDLED_ERROR
                
        Set Assets = Nothing
    End If
Exit Sub

ErrorExit:

    Set Assets = Nothing
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
' CmoItem_Change
' Event when item changes
' ---------------------------------------------------------------
Private Sub CmoItem_Change()
    Dim LocAsset As ClsAsset
    Dim AssetNo As Integer
    Dim Assets As ClsAssets
    Dim StrSize1Arry() As String
    Dim StrSize2Arry() As String
    Dim i As Integer
    
    Const StrPROCEDURE As String = "CmoItem_Change()"

    On Error GoTo ErrorHandler
        
    CmoItem.BackColor = COLOUR_3
    
    If CmoItem <> "" Then
        
        If Lineitem Is Nothing Then Err.Raise SYSTEM_RESTART
        
        Set Assets = New ClsAssets
        
        Set LocAsset = New ClsAsset
        StrSize1Arry() = Assets.GetSizeLists(CmoItem, 1)
        StrSize2Arry() = Assets.GetSizeLists(CmoItem, 2)
        
        CmoSize1.Clear
        CmoSize2.Clear
        CmoSize1.Visible = False
        CmoSize2.Visible = False
        LblSize1.Visible = False
        LblSize2.Visible = False
        CmoQuantity = ""
        TxtPurchaseUnit = ""
        
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
        
        LocAsset.DBGet (Assets.FindAssetNo(CmoItem, CmoSize1, CmoSize2))
        
        If Not UpdateStockMessage(LocAsset) Then Err.Raise HANDLED_ERROR
        
        TxtPurchaseUnit = LocAsset.PurchaseUnit
    
        Set LocAsset = Nothing
        Set Assets = Nothing
    End If

GracefulExit:

Exit Sub

ErrorExit:

    Set LocAsset = Nothing
    Set Assets = Nothing
    
    FormTerminate
    Terminate

Exit Sub

ErrorHandler:
    
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        CustomErrorHandler Err.Number
        Resume GracefulExit
    End If


    If CentralErrorHandler(StrMODULE, StrPROCEDURE, True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' CmoQuantity_Change
' Event when quantity changes
' ---------------------------------------------------------------
Private Sub CmoQuantity_Change()
    CmoQuantity.BackColor = COLOUR_3
End Sub

' ===============================================================
' CmoSize1_Change
' Event when entry changes
' ---------------------------------------------------------------
Private Sub CmoSize1_Change()
    Dim Assets As ClsAssets
    
    Const StrPROCEDURE As String = "CmoSize1_Change()"

    On Error GoTo ErrorHandler
   
    Set Assets = New ClsAssets
    
    CmoSize1.BackColor = COLOUR_3
    
    Lineitem.Asset.DBGet (Assets.FindAssetNo(CmoItem, CmoSize1, CmoSize2))
    
    If Not UpdateStockMessage(Lineitem.Asset) Then Err.Raise HANDLED_ERROR
    
    TxtPurchaseUnit = Lineitem.Asset.PurchaseUnit
    
    If CmoSize1 = "" Then CmoSize2 = ""

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
    
    Lineitem.Asset.DBGet (Assets.FindAssetNo(CmoItem, CmoSize1, CmoSize2))
    
    If Not UpdateStockMessage(Lineitem.Asset) Then Err.Raise HANDLED_ERROR
    
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
' Custom initialise form to run start up actions for form
' ---------------------------------------------------------------
Public Function FormInitialise() As Boolean
    Const StrPROCEDURE As String = "FormInitialise()"
    
    Dim RstCategoryList As Recordset
    Dim Assets As ClsAssets
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    Set Assets = New ClsAssets
    
    'refresh asset list
    If Not ShtLists.RefreshAssetList Then Err.Raise HANDLED_ERROR
    
    'refresh category lists
    Set RstCategoryList = Assets.GetCategoryLists
    
    CmoCategory1.Clear
    
    If RstCategoryList Is Nothing Then Err.Raise NO_RECORDSET_RETURNED, Description:="No Recordset returned from Database search"
    
    With RstCategoryList
        .MoveFirst
        Do While Not .EOF
            CmoCategory1.AddItem .Fields(0)
            .MoveNext
        Loop
    End With
    
    If Not UpdateItemList Then Err.Raise HANDLED_ERROR
    
    CmoCategory2.Visible = False
    CmoCategory3.Visible = False
    LblCategory2.Visible = False
    LblCategory3.Visible = False
    CmoSize1.Visible = False
    CmoSize2.Visible = False
    LblSize1.Visible = False
    LblSize2.Visible = False
    LblStockMessage.Visible = False
    
    CmoQuantity.Clear
    For i = 1 To 50
        CmoQuantity.AddItem i
    Next
    
    CmoQuantity.ListIndex = 0
    
    CmoQuantity.Value = 1
    
    FormInitialise = True
    
    Set Assets = Nothing
    Set RstCategoryList = Nothing
Exit Function

ErrorExit:
    
    Set Assets = Nothing
    Set RstCategoryList = Nothing
    FormTerminate
    Terminate
    
    FormInitialise = False

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
' FormTerminate
' Custom Terminate form to run close down actions for form
' ---------------------------------------------------------------
Public Sub FormTerminate()
    On Error Resume Next
    
    Set Lineitem = Nothing
    Unload Me
    
End Sub

' ===============================================================
' UpdateItemList
' Updates the item list from search with category list entries
' ---------------------------------------------------------------
Private Function UpdateItemList() As Boolean
    Dim RstItemList As Recordset
    Dim Assets As ClsAssets
    
    Const StrPROCEDURE As String = "UpdateItemList()"
    
    On Error GoTo ErrorHandler
    
    'refresh item list
    Set Assets = New ClsAssets
    Set RstItemList = Assets.GetItemList(CmoCategory1, CmoCategory2, CmoCategory3)
    
    If RstItemList Is Nothing Then Err.Raise NO_RECORDSET_RETURNED, Description:="No Recordset returned from Database search"
    
    With RstItemList
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                CmoItem.AddItem .Fields(0)
                .MoveNext
            Loop
        End If
    End With

    Set Assets = Nothing
    Set RstItemList = Nothing
    
    UpdateItemList = True

Exit Function

ErrorExit:

    Set Assets = Nothing
    Set RstItemList = Nothing
    
    FormTerminate
    Terminate
    UpdateItemList = False

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
' UpdateStockMessage
' Tells the user whether item is in stock
' ---------------------------------------------------------------
Private Function UpdateStockMessage(Asset As ClsAsset) As Boolean
    Const StrPROCEDURE As String = "UpdateStockMessage()"

    On Error GoTo ErrorHandler
    
    LblStockMessage.Visible = True
    
    With CmoSize1
        If .Visible = True And .Value = "" Then LblStockMessage.Visible = False
    End With
    
    With CmoSize2
        If .Visible = True And .Value = "" Then LblStockMessage.Visible = False
    End With
    
    With LblStockMessage
        Select Case Asset.QtyInStock
            Case Is = 0
                .Caption = "No Stock"
                .BackColor = COLOUR_7
                .BorderColor = COLOUR_1
                .ForeColor = COLOUR_8
                UpdateStockMessage = True
                Exit Function
            Case Is < Lineitem.Asset.OrderLevel
                .Caption = "Low Stock"
                .BackColor = COLOUR_11
                .BorderColor = COLOUR_1
                .ForeColor = COLOUR_1
                UpdateStockMessage = True
                Exit Function
            Case Else
                .Caption = "In Stock"
                .BackColor = COLOUR_10
                .BorderColor = COLOUR_1
                .ForeColor = COLOUR_1
        
        End Select
    End With


    UpdateStockMessage = True

Exit Function

ErrorExit:

    UpdateStockMessage = False
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
' SelectNextForm
' Selects the next form to display from Allocation type
' ---------------------------------------------------------------
Private Function SelectNextForm() As Boolean
    Const StrPROCEDURE As String = "SelectNextForm()"

    On Error GoTo ErrorHandler
    
    If Lineitem Is Nothing Then Err.Raise NO_LINE_ITEM, Description:=" No Line Item available"
    
    Hide
    
    Select Case Lineitem.Asset.AllocationType
        Case Is = Person
            If Not FrmPerson.ShowForm(False, Lineitem) Then Err.Raise HANDLED_ERROR
            Unload Me
        Case Is = Vehicle
            If Not FrmVehicle.ShowForm(Lineitem) Then Err.Raise HANDLED_ERROR
            Unload Me
        Case Else
            If Not FrmStation.ShowForm(Lineitem) Then Err.Raise HANDLED_ERROR
            Unload Me
    End Select

    Unload Me
    
    SelectNextForm = True

Exit Function

ErrorExit:

    SelectNextForm = False
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
