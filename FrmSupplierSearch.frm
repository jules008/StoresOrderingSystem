VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmSupplierSearch 
   Caption         =   "Text Search"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7665
   OleObjectBlob   =   "FrmSupplierSearch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmSupplierSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 07 Jul 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "FrmSupplierSearch"

Private Suppliers As ClsSuppliers
Private Supplier As ClsSupplier

' ===============================================================
' ShowForm
' Initial entry point to form
' ---------------------------------------------------------------
Public Function ShowForm() As Boolean
    
    Const StrPROCEDURE As String = "ShowForm()"
    
    On Error GoTo ErrorHandler
    
    ResetForm
        
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
    
    Select Case ValidateForm
    
        Case Is = FunctionalError
            Err.Raise HANDLED_ERROR
        
        Case Is = FormOK
                                      
            'next page
            Hide
            
            With LstResults
                Supplier.DBGet .List(.ListIndex, 0)
            End With
            
            If Supplier Is Nothing Then Err.Raise HANDLED_ERROR, Description:="No Supplier Found"
            
            If Not FrmSupplier.ShowForm(Supplier) Then Err.Raise HANDLED_ERROR
            
            Unload Me
    End Select


Exit Sub
    
ErrorExit:
    
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
    
    Set Supplier = New ClsSupplier
    Set Suppliers = New ClsSuppliers
    
    Suppliers.GetCollection

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
    
    Set Supplier = Nothing
    Set Suppliers = Nothing
    
    Unload Me
End Sub

' ===============================================================
' GetSearchItems
' Gets items from the asset list that match Txtsearch box
' ---------------------------------------------------------------
Private Function GetSearchItems(StrSearch As String) As Boolean
    Dim LocSupplier As ClsSupplier
    Dim i As Integer
    
    Const StrPROCEDURE As String = "GetSearchItems()"

    On Error GoTo ErrorHandler
    
    LstResults.Clear
    'search item list and populate results.  Stop before looping back to start
    'get length of item list
    
    i = 0
    For Each LocSupplier In Suppliers
                    
        If InStr(UCase(LocSupplier.SupplierName), UCase(StrSearch)) Or InStr(UCase(LocSupplier.ItemsSupplied), UCase(StrSearch)) Then
            
            With LstResults
                .AddItem
                .List(i, 0) = LocSupplier.SupplierID
                .List(i, 1) = LocSupplier.SupplierName
                i = i + 1
            End With
        End If
    
    Next
    
    
    GetSearchItems = True
    
    
Exit Function

ErrorExit:
    
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

