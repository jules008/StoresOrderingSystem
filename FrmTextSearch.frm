VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmTextSearch 
   Caption         =   "Text Search"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7740
   OleObjectBlob   =   "FrmTextSearch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmTextSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 22 Feb 17
'===============================================================
' Methods
'---------------------------------------------------------------
' ShowForm - Initial entry point to form
' ResetForm - Resets the form
' BtnClose_Click - Close button event
' BtnNext_Click - saves order and moves onto next form
' TxtSearch_Change - detects when the search box changes and then re-runs the search
' FormInitialise - Custom initialise form to run start up actions for form
' FormTerminate - Custom Terminate form to run close down actions for form
' GetSearchItems - Gets items from the asset list that match Txtsearch box
' LstResults_Click - Gets items from the asset list that match Txtsearch box
' ValidateForm - Ensures the form is filled out correctly before moving on

'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmTextSearch"

Private LineItem As ClsLineItem

' ===============================================================
' ShowForm
' Initial entry point to form
' ---------------------------------------------------------------
Public Function ShowForm(Optional LocLineItem As ClsLineItem) As Boolean
    
    Const StrPROCEDURE As String = "ShowForm()"
    
    On Error GoTo ErrorHandler
    
    ResetForm
    
    If LocLineItem Is Nothing Then
        Set LineItem = New ClsLineItem
    Else
        Set LineItem = LocLineItem
        TxtSearch = LineItem.Asset.Description
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
' ResetForm
' Resets the form
' ---------------------------------------------------------------
Private Sub ResetForm()

    On Error Resume Next
    
    LstResults = ""
    TxtSearch = ""

End Sub

' ===============================================================
' BtnClose_Click
' Close button event
' ---------------------------------------------------------------
Private Sub BtnClose_Click()

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
            
            If LineItem Is Nothing Then Err.Raise NO_LINE_ITEM, Description:="No Line Item on Order"
            
            With LineItem
                
                .Asset.DBGet (Assets.FindAssetNo(TxtSearch.Value, "", ""))
                If .Asset Is Nothing Then Err.Raise NO_ASSET_ON_ORDER, Description:="No Asset on current Order"
            
            End With
                          
            'next page
            Hide
            If Not FrmCatSearch.ShowForm(LineItem) Then Err.Raise HANDLED_ERROR
            Unload Me
    End Select

    
    Set Assets = Nothing

Exit Sub
    
ErrorExit:
    
    Set Assets = Nothing
    
    FormTerminate
    Terminate
    
Exit Sub

ErrorHandler:
    
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        CustomErrorHandler Err.Number
        Resume Next
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' TxtSearch_Change
' detects when the search box changes and then re-runs the search
' ---------------------------------------------------------------
Private Sub TxtSearch_Change()
    Const StrPROCEDURE As String = "TxtSearch_Change()"
    
    Dim ListResults As String

    On Error GoTo ErrorHandler
        
    TxtSearch.BackColor = COLOUR_3
    LstResults.BackColor = COLOUR_3
    
    With LstResults
        If LstResults.ListIndex <> -1 Then ListResults = .List(.ListIndex)
    End With
    
    'if the search box has been changed since being updated by the results box, clear the result box
    If ListResults <> TxtSearch.Value Then LstResults.ListIndex = -1
    
    'if the results box has been clicked, add the selected result to the search box
    If LstResults.ListIndex = -1 Then
    
        'if no results selected, populate with new results
        If Len(TxtSearch.Value) > 2 Then
            If Not GetSearchItems(TxtSearch.Value) Then Err.Raise HANDLED_ERROR
        Else
            LstResults.Clear
        End If
    End If
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
    
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    'refresh asset list
    If Not ShtLists.RefreshAssetList Then Err.Raise HANDLED_ERROR
    
    
    FormInitialise = True
    
Exit Function

ErrorExit:
    
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
    
    Set LineItem = Nothing
    Unload Me
End Sub

' ===============================================================
' GetSearchItems
' Gets items from the asset list that match Txtsearch box
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
    
    LstResults.Clear
    'search item list and populate results.  Stop before looping back to start
    If Not RngResult Is Nothing Then
    
        Do
            Set RngResult = RngItems.FindNext(RngResult)
            LstResults.AddItem RngResult.Value
            
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
' LstResults_Click
' Gets items from the asset list that match Txtsearch box
' ---------------------------------------------------------------
Private Sub LstResults_Click()

    On Error Resume Next

    With LstResults
        Me.TxtSearch.Value = .List(.ListIndex)
    End With
End Sub

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
    
    With LstResults
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
    
Exit Function

ErrorExit:
    
    FormTerminate
    Terminate

    ValidateForm = FunctionalError

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

