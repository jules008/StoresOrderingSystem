VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmDataManagmt 
   Caption         =   "Category Search"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6735
   OleObjectBlob   =   "FrmDataManagmt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmDataManagmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 12 Jun 17
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmDataManagmt"

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
' BtnAssetExport_Click
' Exports all asset data into spreadsheet
' ---------------------------------------------------------------
Private Sub BtnAssetExport_Click()
    Dim ColWidths(0 To 23) As Integer
    Dim Headings(0 To 23) As String
    Dim RstAssets As Recordset
    
    Const StrPROCEDURE As String = "BtnAssetExport_Click()"

    On Error GoTo ErrorHandler

    'col widths
    ColWidths(0) = 10
    ColWidths(1) = 10
    ColWidths(2) = 30
    ColWidths(3) = 50
    ColWidths(4) = 10
    ColWidths(5) = 25
    ColWidths(6) = 25
    ColWidths(7) = 25
    ColWidths(8) = 20
    ColWidths(9) = 20
    ColWidths(10) = 20
    ColWidths(11) = 20
    ColWidths(12) = 20
    ColWidths(13) = 20
    ColWidths(14) = 20
    ColWidths(15) = 20
    ColWidths(16) = 20
    ColWidths(17) = 20
    ColWidths(18) = 50
    ColWidths(19) = 20
    ColWidths(20) = 20
    ColWidths(21) = 20
    ColWidths(22) = 20
    ColWidths(23) = 20
    
    'headings
    Headings(0) = "Asset No"
    Headings(1) = "Allocation Type"
    Headings(2) = "Brand"
    Headings(3) = "Description"
    Headings(4) = "Qty in Stock"
    Headings(5) = "Category 1"
    Headings(6) = "Category 2"
    Headings(7) = "Category 3"
    Headings(8) = "Size 1"
    Headings(9) = "Size 2"
    Headings(10) = "Purchase Unit"
    Headings(11) = "Min Amount"
    Headings(12) = "Max Amount"
    Headings(13) = "Order Lvl"
    Headings(14) = "Lead Time"
    Headings(15) = "Keywords"
    Headings(16) = "Allowed Order Reasons"
    Headings(17) = "Additional Info"
    Headings(18) = "No Order Message"
    Headings(19) = "Location"
    Headings(20) = "Status"
    Headings(21) = "Cost"
    Headings(22) = "Supplier 1"
    Headings(23) = "Supplier 2"
    
    Set RstAssets = ModDatabase.SQLQuery("TblAsset")

    If Not ModReports.CreateReport(RstAssets, ColWidths, Headings) Then Err.Raise HANDLED_ERROR
    
    Set RstAssets = Nothing
    
Exit Sub

ErrorExit:
    Set RstAssets = Nothing

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
' BtnAssetImport_Click
' Starts process for importing asset file
' ---------------------------------------------------------------
Private Sub BtnAssetImport_Click()
    Const StrPROCEDURE As String = "BtnAssetImport_Click()"

    On Error GoTo ErrorHandler

    If Not FrmDataImport.ShowForm Then Err.Raise HANDLED_ERROR

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
' BtnClosePage_Click
' Close page
' ---------------------------------------------------------------
Private Sub BtnClosePage_Click()
    FormTerminate
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
        
    On Error GoTo ErrorHandler
    
    
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
    
    Unload Me
    
End Sub

