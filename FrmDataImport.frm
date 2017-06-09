VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmDataImport 
   Caption         =   "Category Search"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6690
   OleObjectBlob   =   "FrmDataImport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmDataImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 09 Jun 17
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmDataImport"

Private Stage As Integer

' ===============================================================
' ShowForm
' Initial entry point to form
' ---------------------------------------------------------------
Public Function ShowForm() As Boolean
    
    Const StrPROCEDURE As String = "ShowForm()"
    
    On Error GoTo ErrorHandler
    
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
' BtnExit_Click
' Exit Form
' ---------------------------------------------------------------
Private Sub BtnExit_Click()
    FormTerminate
End Sub

' ===============================================================
' BtnStart_Click
' Starts Processing
' ---------------------------------------------------------------
Private Sub BtnStart_Click()
    Dim ErrorCnt As Integer
    Dim WarnCnt As Integer
    Dim Response As String
    
    Const StrPROCEDURE As String = "BtnStart_Click()"

    On Error GoTo ErrorHandler

    If MainScreen Is Nothing Then Err.Raise SYSTEM_RESTART
    
    Select Case Stage
        Case Is = 1
            If Not ModAssetImportExport.Stage1_LoadFile Then Err.Raise HANDLED_ERROR
            
            ErrorCnt = ModAssetImportExport.ErrorCount
            
            If ErrorCnt = 0 Then
                Stage = 2
                TxtS1Message = "Click Continue to start Stage 2"
                With Gge1Inner
                    .ForeColor = COLOUR_2
                    .BackColor = COLOUR_10
                    .BorderColor = COLOUR_1
                    .Caption = "Passed file check with no errors"
                End With
                BtnStart.Caption = "Continue"
                
                Frame2.Visible = True
                
            Else
                Stage = 1
                TxtS1Message = "Correct errors and Click Restart to re-run Stage 1"
                
                With Gge1Inner
                    If ErrorCnt = 1 Then .Caption = "File check failed with 1 Error"
                    If ErrorCnt > 1 Then .Caption = "File check failed with " & ErrorCnt & " errors"
                    .ForeColor = COLOUR_3
                    .BackColor = COLOUR_7
                    .BorderColor = COLOUR_1
                End With
                BtnStart.Caption = "Restart"
                DoEvents
                Sleep (1000)
                If Not FrmErrorLog.ShowForm(ModAssetImportExport.ErrorLog, ModAssetImportExport.WarningLog) Then Err.Raise HANDLED_ERROR
            End If
        
        Case Is = 2
            ModAssetImportExport.ClearLog
            
            If Not ModAssetImportExport.Stage2_PreBuild Then Err.Raise HANDLED_ERROR
            
                WarnCnt = ModAssetImportExport.WarningCount
            
            If WarnCnt = 0 Then
                Stage = 3
                TxtS2Message = "Click Continue to start Stage 3"
                With Gge2Inner
                    .ForeColor = COLOUR_2
                    .BackColor = COLOUR_10
                    .BorderColor = COLOUR_1
                    .Caption = "Passed Pre-build Check"
                End With
                BtnStart.Caption = "Continue"
                
                Frame3.Visible = True
                
            Else
                Stage = 3
                TxtS2Message = "Review warnings and click Continue to start Stage 3"
                
                With Gge2Inner
                    If WarnCnt = 1 Then .Caption = "Passed Pre-build Check with 1 Warning"
                    If WarnCnt > 1 Then .Caption = "Passed Pre-build Check with " & WarnCnt & " warnings"
                    .ForeColor = COLOUR_2
                    .BackColor = COLOUR_11
                    .BorderColor = COLOUR_1
                End With
                BtnStart.Caption = "Continue"
                Frame3.Visible = True
                DoEvents
                Sleep (1000)
                If Not FrmErrorLog.ShowForm(ModAssetImportExport.ErrorLog, ModAssetImportExport.WarningLog) Then Err.Raise HANDLED_ERROR
            End If

        Case Is = 3
            ModAssetImportExport.ClearLog
            
            Response = MsgBox("THE NEXT STAGE WILL PERMANENTLY WRITE DATE TO THE DATABASE, ARE YOU SURE YOU WISH TO CONTINUE?", vbCritical + vbYesNoCancel + vbDefaultButton2, APP_NAME)
                        
            If Response = 6 Then
            
                If Not ModAssetImportExport.Stage3_CopyData Then Err.Raise HANDLED_ERROR
                
                TxtS3Message = "Click Continue to start Stage 4"
                With Gge3Inner
                    .ForeColor = COLOUR_2
                    .BackColor = COLOUR_10
                    .BorderColor = COLOUR_1
                    .Caption = "Data copied to database"
                End With
            End If
                BtnStart.Caption = "Continue"
            Stage = 4
            Frame4.Visible = True
                    
            
        Case Is = 4
            ModAssetImportExport.ClearLog
            
            If Not ModAssetImportExport.Stage4_Validate Then Err.Raise HANDLED_ERROR
            
            ErrorCnt = ModAssetImportExport.ErrorCount
            
            If ErrorCnt = 0 Then
                Stage = 5
                TxtS4Message = "Data Validation passed with no errors"
                With Gge4Inner
                    .ForeColor = COLOUR_2
                    .BackColor = COLOUR_10
                    .BorderColor = COLOUR_1
                    .Caption = "Import successfully completed, now click Exit"
                End With
                BtnStart.Visible = False
                
                Frame3.Visible = True
                
            Else
                Stage = 4
                TxtS4Message = "Review errors and click Restart to re-run data copy"
                
                With Gge4Inner
                    If ErrorCnt = 1 Then .Caption = "Failed Validation Check with 1 Error"
                    If ErrorCnt > 1 Then .Caption = "Failed Validation Check with " & ErrorCnt & " errors"
                    .ForeColor = COLOUR_3
                    .BackColor = COLOUR_7
                    .BorderColor = COLOUR_1
                End With
                BtnStart.Caption = "Restart"
                DoEvents
                Sleep (1000)
                If Not FrmErrorLog.ShowForm(ModAssetImportExport.ErrorLog, ModAssetImportExport.WarningLog) Then Err.Raise HANDLED_ERROR
            End If

    End Select

Exit Sub

ErrorExit:
    FormTerminate
    Terminate
'    ***CleanUpCode***

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
    
    Frame2.Visible = False
    Frame3.Visible = False
    Frame4.Visible = False
    BtnStart.Caption = "Start"
    BtnRestart.Visible = False
    
    Stage = 1
    
    TxtS1Message = ""
    TxtS2Message = ""
    TxtS3Message = ""
    TxtS4Message = ""
    
    Gge1Inner.Width = 0
    Gge2Inner.Width = 0
    Gge3Inner.Width = 0
    Gge4Inner.Width = 0
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
    
    ModAssetImportExport.ImportTerminate
    Unload Me
    
End Sub

' ===============================================================
' UpdateProgrGges
' Updates the four progress gauges
' ---------------------------------------------------------------
Public Function UpdateProgrGges(Total As Integer, Progress As Integer, GaugeNo As Integer) As Integer
    Dim GaugeInner As Control
    Dim GaugeOuter As Control
    Dim ProgressPC As Single
    
    Const StrPROCEDURE As String = "UpdateProgrGges()"

    On Error GoTo ErrorHandler
    
    ProgressPC = Progress / Total
    
    Set GaugeInner = Me.Controls("gge" & GaugeNo & "Inner")
    Set GaugeOuter = Me.Controls("gge" & GaugeNo & "Outer")
    GaugeInner.Width = GaugeOuter.Width * ProgressPC

    UpdateProgrGges = Progress

    DoEvents
    
    Set GaugeInner = Nothing
    Set GaugeOuter = Nothing
Exit Function

ErrorExit:
    Set GaugeInner = Nothing
    Set GaugeOuter = Nothing

'    ***CleanUpCode***
    UpdateProgrGges = 0

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function
