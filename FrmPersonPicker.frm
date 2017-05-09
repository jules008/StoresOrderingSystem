VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmPersonPicker 
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6045
   OleObjectBlob   =   "FrmPersonPicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmPersonPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 09 May 17
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmPersonPicker"

Public SelectedUser As ClsPerson

' ===============================================================
' ShowForm
' Initial entry point to form
' ---------------------------------------------------------------
Public Function ShowForm() As ClsPerson
    
    Const StrPROCEDURE As String = "ShowForm()"
    
    On Error GoTo ErrorHandler
    
    Show
        
    Set ShowForm = SelectedUser
    
    Set SelectedUser = Nothing
    
    Unload Me
Exit Function

ErrorExit:
    
    Set SelectedUser = Nothing

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
    
    Set SelectedUser = Nothing
    
    Unload Me

End Function

' ===============================================================
' BtnClose_Click
' Event for page close button
' ---------------------------------------------------------------
Private Sub BtnClose_Click()

    On Error Resume Next
    
    FormTerminate
    
End Sub

' ===============================================================
' BtnSelect_Click
' Moves onto next form
' ---------------------------------------------------------------
Private Sub BtnSelect_Click()
    Dim StrUserName As String
    
    Const StrPROCEDURE As String = "BtnSelect_Click()"

    On Error GoTo ErrorHandler

    Select Case ValidateForm

        Case Is = FunctionalError
            Err.Raise HANDLED_ERROR
        
        Case Is = ValidationError
            
        Case Is = FormOK
                    
            Hide
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
' Entry for search string
' ---------------------------------------------------------------
Private Sub TxtSearch_Change()
    Const StrPROCEDURE As String = "TxtSearch_Change()"

    On Error GoTo ErrorHandler

    Dim ListResults As String

    On Error GoTo ErrorHandler
    
    TxtSearch.BackColor = COLOUR_3
    LstNames.BackColor = COLOUR_3

    With LstNames
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

    Set SelectedUser = New ClsPerson
    
    'refresh name list
    If Not ShtLists.RefreshNameList Then Err.Raise HANDLED_ERROR

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
    
    With LstNames
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
    ListLength = Application.WorksheetFunction.CountA(ShtLists.Range("C:C"))
    
    If IsNumeric(StrSearch) Then
    
        Set RngItems = ShtLists.Range("C1:C" & ListLength)
    Else
        Set RngItems = ShtLists.Range("D1:D" & ListLength)
    
    End If
        
    Set RngResult = RngItems.Find(StrSearch)
    Set RngFirstResult = RngResult
    
    LstNames.Clear
    'search item list and populate results.  Stop before looping back to start
    If Not RngResult Is Nothing Then
    
        i = 0
        Do
            Set RngResult = RngItems.FindNext(RngResult)
                With LstNames
                    .AddItem
                    If IsNumeric(StrSearch) Then
                        .List(i, 0) = RngResult.Value
                        .List(i, 1) = RngResult.Offset(0, 1)
                    Else
                        .List(i, 1) = RngResult.Value
                        .List(i, 0) = RngResult.Offset(0, -1)
                    End If
                    i = i + 1
            End With
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
' LstNames_Click
' Gets items from the name list that match Txtsearch box
' ---------------------------------------------------------------
Private Sub LstNames_Click()

    On Error Resume Next

    LstNames.BackColor = COLOUR_3
    
    With LstNames
        Me.TxtSearch.Value = .List(.ListIndex, 1)
        .ListIndex = 0
        
        SelectedUser.DBGet TxtSearch
    End With
End Sub



