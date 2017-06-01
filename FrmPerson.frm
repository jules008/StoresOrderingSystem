VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmPerson 
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6045
   OleObjectBlob   =   "FrmPerson.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'===============================================================
' v0,0 - Initial version
' v0,1 - changes for Remote Order functionality
' v0,2 - bug fix - change crewno to read string not integer
' v0,3 - Set Phone Order flag
' v0,4 - Phone Order Bug Fix
' v0,5 - Clean up if user cancels form
'---------------------------------------------------------------
' Date - 01 Jun 17
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmPerson"

Private Order As ClsOrder
Private Lineitem As ClsLineItem
Private RemoteOrder As Boolean

' ===============================================================
' ShowForm
' Initial entry point to form
' ---------------------------------------------------------------
Public Function ShowForm(LocRemoteOrder As Boolean, Optional LocLineItem As ClsLineItem) As Boolean
    
    Const StrPROCEDURE As String = "ShowForm()"
    
    On Error GoTo ErrorHandler
    
    RemoteOrder = LocRemoteOrder
    
    If RemoteOrder Then
        Set Order = New ClsOrder
        
        BtnPrev.Enabled = False
        LblAllocation = "Who are you raising the Order on behalf of?"
    Else
        If Not LocLineItem Is Nothing Then
            Set Lineitem = LocLineItem
            
            If Not PopulateForm Then Err.Raise HANDLED_ERROR
        End If
    End If
    
    
    Show
    ShowForm = True
    
Exit Function

ErrorExit:
    
    Set Lineitem = Nothing
    If Not Order Is Nothing Then Set Order = Nothing

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
' CancelOrder
' Cleans up after order is cancelled
' ---------------------------------------------------------------
Private Function CancelOrder() As Boolean
    Const StrPROCEDURE As String = "CancelOrder()"

    On Error GoTo ErrorHandler

    Lineitem.Parent.LineItems.RemoveItem (CStr(Lineitem.LineItemNo))

    CancelOrder = True

Exit Function

ErrorExit:

'    ***CleanUpCode***
    CancelOrder = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
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
    Dim CrewNo As String
    
    On Error GoTo ErrorHandler
    
    CrewNo = Lineitem.ForPerson.CrewNo
    
    If CrewNo = "" Or CrewNo = CurrentUser.CrewNo Then
        OptMe.Value = True
    Else
        OptElse.Value = True
    End If
    
    PopulateForm = True

Exit Function

ErrorExit:

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

    Set Lineitem = Nothing
    If Not Order Is Nothing Then Set Order = Nothing
    Unload Me

End Function

' ===============================================================
' BtnClose_Click
' Event for page close button
' ---------------------------------------------------------------
Private Sub BtnClose_Click()

    On Error Resume Next
    
    If Not CancelOrder Then Err.Raise HANDLED_ERROR
        
    FormTerminate
    
End Sub

' ===============================================================
' UserForm_QueryClose
' Tidies up if user cancels order
' ---------------------------------------------------------------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = 0 Then
        If Not CancelOrder Then Err.Raise HANDLED_ERROR
    End If
End Sub

' ===============================================================
' BtnNext_Click
' Moves onto next form
' ---------------------------------------------------------------
Private Sub BtnNext_Click()
    Dim StrUserName As String
    
    Const StrPROCEDURE As String = "BtnNext_Click()"

    On Error GoTo ErrorHandler

    Select Case ValidateForm

        Case Is = FunctionalError
            Err.Raise HANDLED_ERROR
        
        Case Is = FormOK
        
            If RemoteOrder Then
                If OptMe Then Order.Requestor = CurrentUser
                Order.PhoneOrder = True
            Else
                If OptMe Then Lineitem.ForPerson = CurrentUser
                
                If Lineitem.ForPerson.CrewNo = "" Then Err.Raise NO_NAMES_SELECTED
                If Not Order Is Nothing Then Order.PhoneOrder = False
            End If
            
            Hide
            
            If RemoteOrder Then
                If Not FrmOrder.ShowForm(Order) Then Err.Raise HANDLED_ERROR
            Else
                If Not FrmLossReport.ShowForm(Lineitem) Then Err.Raise HANDLED_ERROR
            End If
            
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
' BtnPrev_Click
' Back to previous screen event
' ---------------------------------------------------------------
Private Sub BtnPrev_Click()
    Const StrPROCEDURE As String = "BtnPrev_Click()"

    On Error GoTo ErrorHandler

    Unload Me
    If Not FrmCatSearch.ShowForm(Lineitem) Then Err.Raise HANDLED_ERROR
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
' OptElse_Click
' Processes who the item is for
' ---------------------------------------------------------------
Private Sub OptElse_Click()
    Const StrPROCEDURE As String = "OptElse_Click()"

    On Error GoTo ErrorHandler
    
    TxtSearch.Visible = True
    LstNames.Visible = True
    LblText1.Visible = True
    
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

    'refresh name list
    If Not ShtLists.RefreshNameList Then Err.Raise HANDLED_ERROR
    
    With LblAllocation
        .Visible = True
    End With
    
    With OptMe
        .Visible = True
    End With
    
    With OptElse
        .Visible = True
    End With
    
    With LblText1
        .Visible = False
    End With

    TxtSearch.Visible = False
    LstNames.Visible = False
    
    OptMe.Value = True
    
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
' OptMe_Click
' User has selected item is for them
' ---------------------------------------------------------------
Private Sub OptMe_Click()
    Const StrPROCEDURE As String = "OptMe_Click()"

    On Error GoTo ErrorHandler

    TxtSearch.Visible = False
    LstNames.Visible = False
    LblText1.Visible = False
    
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
' ValidateForm
' Ensures the form is filled out correctly before moving on
' ---------------------------------------------------------------
Private Function ValidateForm() As EnumFormValidation
    Const StrPROCEDURE As String = "ValidateForm()"

    On Error GoTo ErrorHandler

    If OptElse.Value = True Then
        
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
            
    End If
                    
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
        
        If RemoteOrder Then
            Order.Requestor.DBGet TxtSearch
        Else
            Lineitem.ForPerson.DBGet TxtSearch
        End If
    End With
End Sub



