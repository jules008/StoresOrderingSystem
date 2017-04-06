VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmDBAsset 
   Caption         =   "Asset"
   ClientHeight    =   11700
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13050
   OleObjectBlob   =   "FrmDBAsset.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmDBAsset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 10 Mar 17
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmDBAsset"

Dim Asset As ClsAsset

' ===============================================================
' ShowForm
' Initial entry point to form
' ---------------------------------------------------------------
Public Function ShowForm(Optional LocAsset As ClsAsset) As Boolean
    
    Const StrPROCEDURE As String = "ShowForm()"
    
    On Error GoTo ErrorHandler
    
    If LocAsset Is Nothing Then
        Err.Raise NO_ASSET_ON_ORDER, Description:="No Asset has been passed to form"
    Else
        Set Asset = LocAsset
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
' Populates form after start up
' ---------------------------------------------------------------
Private Function PopulateForm() As Boolean
    Dim Keywords() As String
    Dim i As Integer
    Dim AllowedReasons() As String

    Const StrPROCEDURE As String = "PopulateForm()"

    On Error GoTo ErrorHandler
        
    With Asset
        TxtAdditInfo = .AdditInfo
        TxtAssetNo = .AssetNo
        TxtBrand = .Brand
        TxtCategory1 = .Category1
        TxtCategory2 = .Category2
        TxtCategory3 = .Category3
        TxtCost = .cost
        TxtDescription = .Description
        TxtLeadTime = .LeadTime
        CmoLocation = .Location.Name
        TxtMaxAmount = .MaxAmount
        TxtMinAmount = .MinAmount
        TxtNoOrderMessage = .NoOrderMessage
        TxtOrderLevel = .OrderLevel
        TxtPurchaseUnit = .PurchaseUnit
        TxtQtyInStock = .QtyInStock
        TxtSize1 = .Size1
        TxtSize2 = .Size2
        TxtStatus = .ReturnAssetStatus
'        TxtSupplier1 = .Supplier1
'        TxtSupplier2 = .Supplier2
        CmoAllocationType.ListIndex = .AllocationType
    End With

    With TxtStatus
        Select Case Asset.Status
            Case Is = 0
                .BackColor = COLOUR_10
                .ForeColor = COLOUR_1
            Case Is = 1
                .BackColor = COLOUR_11
                .ForeColor = COLOUR_1
            Case Else
                .BackColor = COLOUR_7
                .ForeColor = COLOUR_3
        End Select
    
    End With
    
    With LstKeywords
        Keywords = Split(Asset.Keywords, ",")
        
        For i = LBound(Keywords) To UBound(Keywords)
            .AddItem Keywords(i)
        Next
        
    End With
    
    AllowedReasons = Split(Asset.AllowedOrderReasons, ":")
    
    If AllowedReasons(0) = "1" Then ChkOrder0.Value = True Else ChkOrder0.Value = False
    If AllowedReasons(1) = "1" Then ChkOrder1.Value = True Else ChkOrder1.Value = False
    If AllowedReasons(2) = "1" Then ChkOrder2.Value = True Else ChkOrder2.Value = False
    If AllowedReasons(3) = "1" Then ChkOrder3.Value = True Else ChkOrder3.Value = False
    If AllowedReasons(4) = "1" Then ChkOrder4.Value = True Else ChkOrder4.Value = False
    If AllowedReasons(5) = "1" Then ChkOrder5.Value = True Else ChkOrder5.Value = False
    If AllowedReasons(6) = "1" Then ChkOrder6.Value = True Else ChkOrder6.Value = False
    
    If Not UpdateStockGauge Then Err.Raise HANDLED_ERROR
    
    PopulateForm = True

Exit Function

ErrorExit:

    FormTerminate
    Terminate

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

' ===============================================================
' FormInitialise
' initialises controls on form at start up
' ---------------------------------------------------------------
Private Function FormInitialise() As Boolean
    Const StrPROCEDURE As String = "FormInitialise()"

    On Error GoTo ErrorHandler
    
    With CmoAllocationType
        .AddItem "Person"
        .AddItem "Vehicle"
        .AddItem "Station"
    
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
' FormTerminate
' Terminates the form gracefully
' ---------------------------------------------------------------
Private Function FormTerminate() As Boolean

    On Error Resume Next

    Set Asset = Nothing
    Unload Me

End Function

' ===============================================================
' UpdateStockGauge
' Updates gauge for stock level
' ---------------------------------------------------------------
Private Function UpdateStockGauge() As Boolean
    Dim Level As Single
    
    Const StrPROCEDURE As String = "UpdateStockGauge()"

    On Error GoTo ErrorHandler
    
    With Asset
        Level = .QtyInStock / .MaxAmount

        GaugeLvl.Height = Gauge.Height * Level
        GaugeLvl.Top = Gauge.Top + Gauge.Height - GaugeLvl.Height
        
        TxtStockPercent = Format(Level * 100, "0")
    End With

    With GaugeLvl
        Select Case Asset.Status
            Case Is = 0
                .BackColor = COLOUR_10
                .ForeColor = COLOUR_1
            Case Is = 1
                .BackColor = COLOUR_11
                .ForeColor = COLOUR_1
            Case Else
                .BackColor = COLOUR_7
                .ForeColor = COLOUR_3
        End Select
    
    End With

    UpdateStockGauge = True

Exit Function

ErrorExit:

    FormTerminate
    Terminate
    UpdateStockGauge = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BtnAddKeyWord_Click
' Adds keyword to list
' ---------------------------------------------------------------
Private Sub BtnAddKeyWord_Click()
    Dim Keyword As String
    
    Const StrPROCEDURE As String = "BtnAddKeyWord_Click()"
    
    On Error GoTo ErrorHandler

    Keyword = Trim(Application.InputBox("Enter your new search keyword"))
    
    If Keyword = "" Then Err.Raise FORM_INPUT_EMPTY
    
    LstKeywords.AddItem Keyword
    
    If Not UpdateChanges Then Err.Raise HANDLED_ERROR

GracefulExit:


Exit Sub

ErrorExit:

'    ***CleanUpCode***

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
' Closes the form
' ---------------------------------------------------------------
Private Sub BtnClose_Click()

    On Error Resume Next

    FormTerminate

End Sub

' ===============================================================
' BtnDelKeyWord_Click
' Deletes keyword from list
' ---------------------------------------------------------------
Private Sub BtnDelKeyWord_Click()
    Const StrPROCEDURE As String = "BtnDelKeyWord_Click()"

    On Error GoTo ErrorHandler

    With LstKeywords
        If .ListIndex = -1 Then Err.Raise NO_ITEM_SELECTED
        .RemoveItem (.ListIndex)
    End With

    If Not UpdateChanges Then Err.Raise HANDLED_ERROR


GracefulExit:


Exit Sub

ErrorExit:

'    ***CleanUpCode***

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
' BtnEditKeyWord_Click
' Edits keyword in list
' ---------------------------------------------------------------
Private Sub BtnEditKeyWord_Click()
    Dim Keyword As String
    
    Const StrPROCEDURE As String = "BtnEditKeyWord_Click()"

    On Error GoTo ErrorHandler

    
    With LstKeywords
        
        If .ListIndex = -1 Then Err.Raise NO_ITEM_SELECTED
        Keyword = .List(.ListIndex)
        
        Keyword = Application.InputBox("Please amend the keyword and press OK", "Edit Search Keyword", Keyword)
        
        .List(.ListIndex) = Keyword
    End With
    
    If Not UpdateChanges Then Err.Raise HANDLED_ERROR


GracefulExit:


Exit Sub

ErrorExit:
'
'    ***CleanUpCode***

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
' BtnUpdate_Click
' Updates any changes to the database
' ---------------------------------------------------------------
Private Sub BtnUpdate_Click()
    Const StrPROCEDURE As String = "BtnUpdate_Click()"

    On Error GoTo ErrorHandler

    If Not UpdateChanges Then Err.Raise HANDLED_ERROR

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
' UpdateChanges
' Updates any changes to the DB
' ---------------------------------------------------------------
Private Function UpdateChanges() As Boolean
    Dim StrKeywords As String
    Dim StrAllowedReasons As String
    Dim LocReason As String
    Dim i As Integer
    
    Const StrPROCEDURE As String = "UpdateChanges()"

    On Error GoTo ErrorHandler
    
    
    
    'Allowed reasons
    For i = 0 To 5
        If Me.Controls("ChkOrder" & i) = True Then LocReason = "1" Else LocReason = "0"
        
        StrAllowedReasons = StrAllowedReasons & LocReason & ":"
    Next
    StrAllowedReasons = StrAllowedReasons & LocReason
    
    
    'Keywords
    With LstKeywords
        For i = 0 To .ListCount - 2
            StrKeywords = StrKeywords & .List(i) & ","
        Next
        StrKeywords = StrKeywords & .List(i)
    End With
    
    'update asset
    With Asset
        .Description = TxtDescription
        .AllowedOrderReasons = StrAllowedReasons
        .Keywords = StrKeywords
        .DBSave
    End With

    UpdateChanges = True

Exit Function

ErrorExit:

'    ***CleanUpCode***
    UpdateChanges = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

Private Sub UserForm_Initialize()
    On Error Resume Next
    
    FormInitialise
    
End Sub
