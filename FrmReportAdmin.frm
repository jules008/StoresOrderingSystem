VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmReportAdmin 
   Caption         =   "Stores Ordering System"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10380
   OleObjectBlob   =   "FrmReportAdmin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmReportAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 30 Nov 17
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmReportAdmin"

Private SelectedUser As ClsPerson

' ===============================================================
' ShowForm
' Initial entry point to form
' ---------------------------------------------------------------
Public Function ShowForm() As Boolean
    
   Const StrPROCEDURE As String = "ShowForm()"
   
   On Error GoTo ErrorHandler
   
    If Not ResetForm Then Err.Raise HANDLED_ERROR
    Show

    ShowForm = True
Exit Function

ErrorExit:
    ShowForm = False
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
' BtnAddCC_Click
' Adds selected name to To list
' ---------------------------------------------------------------
Private Sub BtnAddCC_Click()
    On Error Resume Next
    AddNameToCC ("CC")
End Sub

' ===============================================================
' BtnAddTo_Click
' Adds selected name to To list
' ---------------------------------------------------------------
Private Sub BtnAddTo_Click()
    On Error Resume Next
    AddNameToCC ("To")
End Sub

' ===============================================================
' AddNameToCC
' Adds selected name to To and CC lists
' ---------------------------------------------------------------

Private Sub AddNameToCC(ToCC As String)
    Dim SelName As String
    Dim CrewNo As String
    Dim ReportNo As Integer
    Dim ActiveList As Control
    Dim i As Integer
    Dim NameFound As Boolean
    
    Const StrPROCEDURE As String = "AddNameToCC()"

    On Error GoTo ErrorHandler

    If LstUserList.ListIndex <> -1 Then
        
        If ToCC = "To" Then
            Set ActiveList = LstTo
        Else
            Set ActiveList = LstCC
        End If
        
        SelName = LstUserList.List(LstUserList.ListIndex, 1)
        CrewNo = LstUserList.List(LstUserList.ListIndex, 0)
        ReportNo = CmoSelectReport.List(CmoSelectReport.ListIndex, 0)
        
        'check if name already added
        With ActiveList
            For i = 0 To .ListCount - 1
                If .List(i, 1) = SelName Then NameFound = True
            Next
            
            'add name if not found and update database
            If Not NameFound Then
                
                With ActiveList
                    .AddItem
                    .List(.ListCount - 1, 0) = CrewNo
                    .List(.ListCount - 1, 1) = SelName
                End With
                
                ModReports.EmailReportsAddAddress CrewNo, ReportNo, ToCC
            End If
        End With
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
' BtnClose_Click
' Event for page close button
' ---------------------------------------------------------------
Private Sub BtnClose_Click()

    On Error Resume Next
    
    FormTerminate
    
End Sub

' ===============================================================
' BtnDelete_Click
' Removes name from selected list
' ---------------------------------------------------------------
Private Sub BtnDelete_Click()
    Dim SelName As String
    Dim CrewNo As String
    Dim ListActive As Boolean
    Dim ReportNo As Integer
    Dim ToCC As String
    Dim ActiveList As Control
    
    Const StrPROCEDURE As String = "BtnDelete_Click()"

    On Error GoTo ErrorHandler

    If LstTo.ListIndex <> -1 Then
        Set ActiveList = LstTo
        ToCC = "To"
        ListActive = True
    End If
    
    If LstCC.ListIndex <> -1 Then
        Set ActiveList = LstCC
        ToCC = "CC"
        ListActive = True
    End If
    
    If ListActive Then
    
        SelName = ActiveList.List(ActiveList.ListIndex, 1)
        CrewNo = ActiveList.List(ActiveList.ListIndex, 0)
        ReportNo = CmoSelectReport.List(CmoSelectReport.ListIndex, 0)
    
        With ActiveList
            If .ListIndex <> -1 Then
                .RemoveItem (.ListIndex)
                ModReports.EmailReportsRemoveAddress CrewNo, ReportNo, ToCC
            End If
            .ListIndex = -1
        End With
    End If
    
    Set ActiveList = Nothing
Exit Sub

ErrorExit:
    Set ActiveList = Nothing
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
' CmoSelectReport_Change
' Enable or disable search box dependent on whether report is selected
' ---------------------------------------------------------------
Private Sub CmoSelectReport_Change()
    Const StrPROCEDURE As String = "CmoSelectReport_Change()"

    On Error GoTo ErrorHandler
    
    If Not ResetForm Then Err.Raise HANDLED_ERROR
    
    With CmoSelectReport
        If .ListIndex = -1 Then
            TxtSearch.Enabled = False
            TxtSearch.Value = "Please select a report"
        Else
            TxtSearch.Enabled = True
            TxtSearch.Value = ""
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
' OptAlerts_Click
' Selects email alerts
' ---------------------------------------------------------------
Private Sub OptAlerts_Click()
    Const StrPROCEDURE As String = "OptAlerts_Click()"

    On Error GoTo ErrorHandler

    If Not EmailTypeChange(2) Then Err.Raise HANDLED_ERROR

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
' OptReports_Click
' Selects email reports
' ---------------------------------------------------------------
Private Sub OptReports_Click()
    Const StrPROCEDURE As String = "OptReports_Click()"

    On Error GoTo ErrorHandler

    If Not EmailTypeChange(1) Then Err.Raise HANDLED_ERROR

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
' TxtSearch_Change
' Entry for search string
' ---------------------------------------------------------------
Private Sub TxtSearch_Change()
    Const StrPROCEDURE As String = "TxtSearch_Change()"

    On Error GoTo ErrorHandler

    Dim ListResults As String

    On Error GoTo ErrorHandler
            
    With LstUserList
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
' ResetForm
' Resets form fields
' ---------------------------------------------------------------
Private Function ResetForm() As Boolean
    Const StrPROCEDURE As String = "ResetForm()"

    On Error GoTo ErrorHandler

    LstCC.Clear
    LstTo.Clear
    LstUserList.Clear
    TxtSearch.Value = ""
    
    ResetForm = True

Exit Function

ErrorExit:

    FormTerminate
    Terminate
    ResetForm = False

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
    ListLength = Application.WorksheetFunction.CountA(ShtLists.Range("C:C"))
    
    If IsNumeric(StrSearch) Then
    
        Set RngItems = ShtLists.Range("C1:C" & ListLength)
    Else
        Set RngItems = ShtLists.Range("D1:D" & ListLength)
    
    End If
        
    Set RngResult = RngItems.Find(StrSearch)
    Set RngFirstResult = RngResult
    
    LstUserList.Clear
    'search item list and populate results.  Stop before looping back to start
    If Not RngResult Is Nothing Then
    
        i = 0
        Do
            Set RngResult = RngItems.FindNext(RngResult)
                With LstUserList
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
    Dim i As Integer
    
    Const StrPROCEDURE As String = "FormInitialise()"

    On Error GoTo ErrorHandler

    
    With LstHeadings
        .AddItem
        .List(0, 0) = "No"
        .List(0, 1) = "Name"
    End With
    
    With CmoSelectReport
        .Clear
        .Value = ""
    End With
    
    With TxtSearch
        .Enabled = False
        .Value = "Please select a report"
    End With
    
    If Not ShtLists.RefreshNameList Then Err.Raise HANDLED_ERROR
    
    
    FormInitialise = True

Exit Function

ErrorExit:
    
    
    FormTerminate
    Terminate
    
    FormInitialise = False

Exit Function

ErrorHandler:
        
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        CustomErrorHandler Err.Number
    End If
    
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' UserForm_Initialize
' Automatic initialise event that triggers custom Initialise
' ---------------------------------------------------------------
Private Sub UserForm_Initialize()

    On Error Resume Next
    
    FormInitialise
    
End Sub

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
' PopulateForm
' Gets the To and CC data for the report addresses
' ---------------------------------------------------------------
Private Function PopulateForm() As Boolean
    Dim RstAddresses As Recordset
    Dim ReportNo As EnumReportNo
    
    Const StrPROCEDURE As String = "PopulateForm()"

    On Error GoTo ErrorHandler

    With CmoSelectReport
        ReportNo = .List(.ListIndex, 0)
    End With
    
    Set RstAddresses = ModReports.GetReportAddresses(ReportNo)

    With RstAddresses
        Do While Not .EOF
            If !ToCC = "To" Then
                LstTo.AddItem
                'debug.print !UserName
                'debug.print !CrewNo
                
                LstTo.List(LstTo.ListCount - 1, 0) = !CrewNo
                LstTo.List(LstTo.ListCount - 1, 1) = !UserName
            End If
            
            If !ToCC = "CC" Then
                LstCC.AddItem
                LstCC.List(LstCC.ListCount - 1, 0) = !CrewNo
                LstCC.List(LstCC.ListCount - 1, 1) = !UserName
            End If
            .MoveNext
        Loop
    End With
    Set RstAddresses = Nothing

    PopulateForm = True

Exit Function

ErrorExit:

    Set RstAddresses = Nothing
    
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
' EmailTypeChange
' Changes email type between alerts and reports
' ---------------------------------------------------------------
Public Function EmailTypeChange(EmailType As Integer) As Boolean
    Dim i As Integer
    Dim RstReports As Recordset
    
    Const StrPROCEDURE As String = "EmailTypeChange()"

    On Error GoTo ErrorHandler
    
    Set RstReports = ModReports.ReturnReportList(EmailType)
    
    If RstReports Is Nothing Then Err.Raise GENERIC_ERROR, , "No reports returned"

    i = 0
    With CmoSelectReport
        .Clear
        .Value = ""
        Do While Not RstReports.EOF
            .AddItem
            .List(i, 0) = RstReports!ReportNo
            .List(i, 1) = RstReports!ReportName
            i = i + 1
            RstReports.MoveNext
        Loop
    End With


    Set RstReports = Nothing


    EmailTypeChange = True

Exit Function

ErrorExit:
    Set RstReports = Nothing

'    ***CleanUpCode***
    EmailTypeChange = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

