VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmSupplier 
   Caption         =   "Supplier"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11280
   OleObjectBlob   =   "FrmSupplier.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmSupplier"
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

Private Const StrMODULE As String = "FrmSupplier"

Private Supplier As ClsSupplier

' ===============================================================
' ShowForm
' Initial entry point to form
' ---------------------------------------------------------------
Public Function ShowForm(Optional LocSupplier As ClsSupplier) As Boolean
    
    Const StrPROCEDURE As String = "ShowForm()"
    
    On Error GoTo ErrorHandler
    
    If LocSupplier Is Nothing Then
        Set Supplier = New ClsSupplier
    Else
        Set Supplier = LocSupplier
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
' Populates form controls
' ---------------------------------------------------------------
Private Function PopulateForm() As Boolean
    
    Const StrPROCEDURE As String = "PopulateForm()"
        
    On Error GoTo ErrorHandler
    
    With Supplier
        TxtAccountNo = .AccountNo
        TxtAddress1 = .Address1
        TxtAddress2 = .Address2
        TxtAggressNo = .AgressoNo
        TxtCategory = .Category
        TxtContactName = .ContactName
        TxtCounty = .County
        TxtEmail = .Email
        TxtItemsSupplied = .ItemsSupplied
        TxtName = .SupplierName
        ChkPCard = .PCard
        TxtPostcode = .Postcode
        TxtTelephone = .Telephone
        TxtTown = .TownCity
        TxtWebsite = .Website
        
    End With
    
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
' BtnCancel_Click
' Event for page close button
' ---------------------------------------------------------------
Private Sub BtnCancel_Click()

    On Error Resume Next
    
    FormTerminate
    
End Sub

' ===============================================================
' FormTerminate
' Terminates the form gracefully
' ---------------------------------------------------------------
Private Function FormTerminate() As Boolean
    Const StrPROCEDURE As String = "FormTerminate()"

    On Error GoTo ErrorHandler

    Set Supplier = Nothing
    
    Unload Me
    

    FormTerminate = True

Exit Function

ErrorExit:

    FormTerminate = False
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
' BtnOk_Click
' Closes and saves changes
' ---------------------------------------------------------------
Private Sub BtnOk_Click()
    Const StrPROCEDURE As String = "BtnOk_Click()"

    On Error GoTo ErrorHandler

    Unload Me

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
' Terminates the form gracefully
' ---------------------------------------------------------------
Private Sub UserForm_Terminate()
    Const StrPROCEDURE As String = "UserForm_Terminate()"

    On Error GoTo ErrorHandler

    If Not FormTerminate Then Err.Raise HANDLED_ERROR

Exit Sub

ErrorExit:

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
' FormInitialise
' initialises controls on form at start up
' ---------------------------------------------------------------
Private Function FormInitialise() As Boolean
    Const StrPROCEDURE As String = "FormInitialise()"
    
    On Error GoTo ErrorHandler
    

Exit Function

ErrorExit:
    
    FormInitialise = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

