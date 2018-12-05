VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmReturnStock 
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8445
   OleObjectBlob   =   "FrmReturnStock.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmReturnStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 04 Dec 18
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmReturnStock"

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

    Unload Me

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
' BtnClose_Click
' Closes form
' ---------------------------------------------------------------
Private Sub BtnClose_Click()
    Dim StrUserName As String
    
    Const StrPROCEDURE As String = "BtnClose_Click()"

    On Error GoTo ErrorHandler

    Hide
    Unload Me
                       
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
    
    Dim i As Integer
    Dim Station As ClsStation
    
    On Error GoTo ErrorHandler
    
    With LstStations
        .Clear
        .Visible = True
        i = 0
        For Each Station In Stations
            If Station.StnActive Then
                .AddItem
                .List(i, 0) = Station.StationID
                .List(i, 1) = Station.StationNo
                .List(i, 2) = Station.Name
                i = i + 1
            End If
        Next
    End With
    
    If Not ClearSearch Then Err.Raise HANDLED_ERROR

    If Not ShtLists.RefreshAssetList Then Err.Raise HANDLED_ERROR

    Set Station = Nothing
    
    FormInitialise = True

Exit Function

ErrorExit:
    
    Set Station = Nothing

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
            
    With LstAssets
        If .ListIndex = -1 Then
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
    
    With LstStations
        If .ListIndex = -1 Then
            .BackColor = COLOUR_6
            ValidateForm = ValidationError
        End If
    End With
                        
    If ValidateForm = ValidationError Then
        Err.Raise FORM_INPUT_EMPTY
    Else
        ValidateForm = FormOK
    End If
    
    ValidateForm = FormOK

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

    StrSize2Arry() = Assets.GetSizeLists(TxtSearch, 2, CmoSize1)

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
        StrSize2Arry() = Assets.GetSizeLists(TxtSearch, 2, CmoSize1)
        
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

    Set RetOrder = New ClsOrder
    Set RetLineItem = New ClsLineItem
    Set Assets = New ClsAssets
    
    If Assets = Nothing Then Err.Raise HANDLED_ERROR, , "No Asset Collection"
    If LstAssets.ListIndex = -1 Then Err.Raise HANDLED_ERROR, , "No Asset Selected"
    
    AssetNo = Assets.FindAssetNo(LstAssets.List(.ListIndex))
    Assets.GetCollection
    
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

