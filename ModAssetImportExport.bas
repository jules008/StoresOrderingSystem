Attribute VB_Name = "ModAssetImportExport"
'===============================================================
' Module ModAssetImportExport
' v0,0 - Initial Version
' v0,1 - Improved version
' v0,2 - Test to ensure DBAsset is not nothing before copying qty
' v0,3 - Highlight if location changes
'---------------------------------------------------------------
' Date - 05 Jul 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModAssetImportExport"

Public ErrorLog(0 To 2000) As String
Public WarningLog(0 To 2000) As String
Public ErrorCount As Integer
Public WarningCount As Integer
Private DBAssets As ClsAssets
Private ShtAssets As ClsAssets
Private MaxAssetNo As Integer

' ===============================================================
' Stage1_LoadFile
' Loads and parses the asset file
' ---------------------------------------------------------------
Public Function Stage1_LoadFile() As Boolean
    Dim FSO As New FileSystemObject
    Dim Rw As Integer
    Dim LineInputString As String
    Dim AssetData() As String
    Dim FileName As String
    Dim FormValidation As Integer
    Dim AssetFileLoc As String
    Dim Asset As ClsAsset
    Dim AssetFile As Integer
    Dim x As Integer
    Dim RowNo As Integer
    Dim FuncPassFail As String
    
    Const StrPROCEDURE As String = "Stage1_LoadFile()"

    On Error GoTo ErrorHandler
    
    Set ShtAssets = New ClsAssets
    Set DBAssets = New ClsAssets
    Set Asset = New ClsAsset

    AssetFileLoc = OpenAssetFile
    DBConnect
    
    If AssetFileLoc = "Error" Then Err.Raise HANDLED_ERROR
    FileName = FSO.GetFileName(AssetFileLoc)
    RowNo = ModLibrary.GetTextLineNo(AssetFileLoc)
    
    'open workbook and sort by asset no
    Application.DisplayAlerts = False
    
    Workbooks.Open AssetFileLoc
    ActiveWorkbook.ActiveSheet.Range("A:Z").Sort key1:=Range("A2"), Header:=xlYes
    ActiveWorkbook.Close savechanges:=True
    
    Application.DisplayAlerts = True
    
    AssetFile = FreeFile()
        
    DBAssets.GetCollection
    
    Open AssetFileLoc For Input As AssetFile
    
    While Not EOF(AssetFile)
        Line Input #AssetFile, LineInputString
        AssetData = Split(LineInputString, ",")
        
        'remove any leading and trailing Quote marks
        For x = 1 To UBound(AssetData)
            
            If Left(AssetData(x), 1) = Chr(34) Then AssetData(x) = Right(AssetData(x), Len(AssetData(x)) - 1)
            If Right(AssetData(x), 1) = Chr(34) Then AssetData(x) = Left(AssetData(x), Len(AssetData(x)) - 1)
            AssetData(x) = Replace(AssetData(x), Chr(34) & Chr(34), Chr(34))
        Next
        
        Rw = Rw + 1
        
        MaxAssetNo = DBAssets.MaxAssetNo
        
        If Rw <> 1 Then
        
            If Not ParseAsset(AssetData, Rw) Then Err.Raise HANDLED_ERROR
                        
            Set Asset = BuildAsset(AssetData)
            
            If Asset Is Nothing Then Err.Raise HANDLED_ERROR
            
            ShtAssets.AddItem Asset
            
            'find maximum assetno
            If Asset.AssetNo > MaxAssetNo Then MaxAssetNo = Asset.AssetNo
                        
            'debug.print "Asset Added!"
        End If

        Rw = FrmDataImport.UpdateProgrGges(RowNo, Rw, 1)
            
        If Rw = 0 Then Err.Raise HANDLED_ERROR, Description:="Error updating gauges"
            
    Wend
        
    Close #AssetFile

    Set FSO = Nothing
    Stage1_LoadFile = True

Exit Function

ErrorExit:
    
    Set FSO = Nothing
    Stage1_LoadFile = False
    Application.DisplayAlerts = True

'    ***CleanUpCode***
    Set Asset = Nothing
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
' ParseAsset
' Checks asset data quality
' ---------------------------------------------------------------
Private Function ParseAsset(AssetData() As String, LineNo As Integer) As Boolean
    Dim i As Integer
    Dim TestValue As String
    Dim TestString() As String
    Dim AssetNo As String
    Static PrevAssetNo As Integer
    Dim PassGenericTests As Boolean
    
    Const StrPROCEDURE As String = "ParseAsset()"
    
    On Error Resume Next
    
    AssetNo = AssetData(0)
    
    PassGenericTests = True
    
    'generic tests first
    If UBound(AssetData) <> 23 Then
        AddToErrorLog AssetNo, "Incorrect use of commas"
        PassGenericTests = False
    End If
    
    If AssetNo = PrevAssetNo Then
        AddToErrorLog AssetNo, "Duplicate Asset No"
        PassGenericTests = False
    End If

    If PassGenericTests Then
        For i = 0 To 23
    
            TestValue = AssetData(i)
            
            If InStr(TestValue, "'") <> 0 Then
                AddToErrorLog AssetData(0), "Found apostrophe"
                PassGenericTests = False
            End If
            
            Select Case i
                Case Is = 0
                            
                Case Is = 1
                    If Not IsNumeric(TestValue) Then AddToErrorLog AssetData(0), "Allocation Type invalid"
                    If TestValue < 0 Or TestValue > 2 Then AddToErrorLog AssetData(0), "Allocation Type invalid"
                Case Is = 4
                    If IsNumeric(TestValue) Then
                        If TestValue < 0 Then AddToErrorLog AssetData(0), "Error in Quantity"
                    Else
                        If TestValue <> "" Then AddToErrorLog AssetData(0), "Error in Quantity"
                    End If
        
                Case Is = 5
                    If TestValue = "" Then AddToErrorLog AssetData(0), "Category 1 cannot be empty"
        
                Case Is = 11
                    If Not IsNumeric(TestValue) Then AddToErrorLog AssetData(0), "Number error in Min Amount"
                    If TestValue < 0 Then AddToErrorLog AssetData(0), "Number error in Min Amount"
                
                Case Is = 12
                    If Not IsNumeric(TestValue) Then AddToErrorLog AssetData(0), "Number error in Max Amount"
                    If TestValue < 0 Then AddToErrorLog AssetData(0), "Number error in Max Amount"
                
                Case Is = 13
                    If Not IsNumeric(TestValue) Then AddToErrorLog AssetData(0), "Number error in Order Levels"
                    If TestValue < 0 Then AddToErrorLog AssetData(0), "Number error in Order Levels"
                
                Case Is = 16
                    
                    If Len(TestValue) <> 13 Then AddToErrorLog AssetData(0), "Length of Allowed Reason string incorrect"
                    
                    TestString = Split(TestValue, ":")
                    
                    If TestString(0) <> "0" And TestString(0) <> "1" Then AddToErrorLog AssetData(0), "Error in Allowed Reason string"
                    If TestString(1) <> "0" And TestString(1) <> "1" Then AddToErrorLog AssetData(0), "Error in Allowed Reason string"
                    If TestString(2) <> "0" And TestString(2) <> "1" Then AddToErrorLog AssetData(0), "Error in Allowed Reason string"
                    If TestString(3) <> "0" And TestString(3) <> "1" Then AddToErrorLog AssetData(0), "Error in Allowed Reason string"
                    If TestString(4) <> "0" And TestString(4) <> "1" Then AddToErrorLog AssetData(0), "Error in Allowed Reason string"
                    If TestString(5) <> "0" And TestString(5) <> "1" Then AddToErrorLog AssetData(0), "Error in Allowed Reason string"
                    If TestString(6) <> "0" And TestString(6) <> "1" Then AddToErrorLog AssetData(0), "Error in Allowed Reason string"
        
                Case Is = 21
                    If TestValue <> "" Then
                    If Not IsNumeric(TestValue) Then AddToErrorLog AssetData(0), "Number error in Cost"
                    If TestValue < 0 Then AddToErrorLog AssetData(0), "Number error in Cost"
                    End If
                
            End Select
        Next
    End If
    
    PrevAssetNo = AssetNo
    
    ParseAsset = True

Exit Function
        
ValidationError:
    
Exit Function

ErrorExit:

'    ***CleanUpCode***
    ParseAsset = False

Exit Function

ErrorHandler:
        
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        If Err.Number = IMPORT_ERROR Then
            ParseAsset = i
            Stop
            Resume ValidationError
        End If
    End If

    
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildAsset
' Takes Asset array and build Asset Class
' ---------------------------------------------------------------
Private Function BuildAsset(AssetData() As String) As ClsAsset
    Dim Asset As ClsAsset
    
    Const StrPROCEDURE As String = "BuildAsset()"

    On Error Resume Next

    Set Asset = New ClsAsset

    With Asset
        .AssetNo = Trim(AssetData(0))
        .AllocationType = Trim(AssetData(1))
        .Brand = Trim(AssetData(2))
        .Description = Trim(AssetData(3))
        If AssetData(4) <> "" Then .QtyInStock = Trim(AssetData(4))
        .Category1 = Trim(AssetData(5))
        .Category2 = Trim(AssetData(6))
        .Category3 = Trim(AssetData(7))
        .Size1 = Trim(AssetData(8))
        .Size2 = Trim(AssetData(9))
        .PurchaseUnit = Trim(AssetData(10))
        .MinAmount = Trim(AssetData(11))
        .MaxAmount = Trim(AssetData(12))
        .OrderLevel = Trim(AssetData(13))
        If AssetData(14) <> "" Then .LeadTime = CInt(AssetData(14))
        .Keywords = Trim(AssetData(15))
        .AllowedOrderReasons = Trim(AssetData(16))
        .AdditInfo = Trim(AssetData(17))
        .NoOrderMessage = Trim(AssetData(18))
        .Location = Trim(AssetData(19))
        If AssetData(20) <> "" Then .Status = Trim(AssetData(20))
        If AssetData(21) <> "" Then .cost = AssetData(21)
        .Supplier1 = Trim(AssetData(22))
        .Supplier2 = Trim(AssetData(23))
    
    End With

    Set BuildAsset = Asset
    Set Asset = Nothing
Exit Function

ErrorExit:

'    ***CleanUpCode***
    BuildAsset = Nothing
    Set Asset = Nothing

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' OpenAssetFile
' Opens the Asset file for input
' ---------------------------------------------------------------
Private Function OpenAssetFile() As String
    Dim DlgOpen As FileDialog
    Dim ShtAssets As ClsAssets
    Dim NoFiles As Integer
    
    Const StrPROCEDURE As String = "OpenAssetFile()"

    On Error GoTo ErrorHandler

    Set DlgOpen = Application.FileDialog(msoFileDialogOpen)
    Set ShtAssets = New ClsAssets
    
     With DlgOpen
        .Filters.Clear
        .Filters.Add "CSV Files (*.csv)", "*.csv"
        .AllowMultiSelect = False
        .Title = "Select Spreadsheet of Doom"
        .Show
    End With
    
    'get no files selected
    NoFiles = DlgOpen.SelectedItems.Count
    
    'exit if no files selected
    If NoFiles = 0 Then Err.Raise NO_FILE_SELECTED

    OpenAssetFile = DlgOpen.SelectedItems(1)

    Set DlgOpen = Nothing
    Set ShtAssets = Nothing
Exit Function

ErrorExit:

'    ***CleanUpCode***
    Set DlgOpen = Nothing
    Set ShtAssets = Nothing
    OpenAssetFile = "Error"

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' Stage2_PreBuild
' Checks before writing to DB
' ---------------------------------------------------------------
Public Function Stage2_PreBuild() As Boolean
    Dim Rw As Integer
    Dim DBAssetNo As Integer
    Dim DBAsset As ClsAsset
    Dim ShtAsset As ClsAsset
    Dim DBAssetDescription As String
    Dim Response As Integer
    
    Const StrPROCEDURE As String = "Stage2_PreBuild()"

    On Error GoTo ErrorHandler

    For Rw = 1 To MaxAssetNo
    
        Set DBAsset = New ClsAsset
        Set ShtAsset = New ClsAsset
                
        Set DBAsset = DBAssets(CStr(Rw))
        Set ShtAsset = ShtAssets(CStr(Rw))
    
        If DBAsset Is Nothing Then
        
            If Not ShtAsset Is Nothing Then
        
                'Add
                AddToWarningLog Rw, ShtAsset.Description & " will be added to database"
        
            End If
        Else
        
            If ShtAsset Is Nothing Then
                
                'delete
                AddToWarningLog Rw, DBAsset.Description & " will be deleted from database"
            
        Else
                
                If ShtAsset.Description <> DBAsset.Description Then
                 
                    'changed description
                    AddToWarningLog Rw, "Asset will change from " & DBAsset.Description & " to " & ShtAsset.Description
                End If
                
                If ShtAsset.Location <> DBAsset.Location Then
                    
                    'changed location
                    AddToWarningLog Rw, "Location will change from " & DBAsset.Location & " to " & ShtAsset.Location
                End If
            End If
        End If
        
        Rw = FrmDataImport.UpdateProgrGges(MaxAssetNo, Rw, 2)
            
        If Rw = 0 Then Err.Raise HANDLED_ERROR, Description:="Error updating gauges"
        
    Next

        'debug.print DBAssetNo
    
    Set DBAsset = Nothing
    Set ShtAsset = Nothing
    
    Stage2_PreBuild = True

Exit Function

ErrorExit:

    Set DBAsset = Nothing
    Set ShtAsset = Nothing
'    ***CleanUpCode***
    Stage2_PreBuild = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' Stage3_CopyData
' Copies the asset file data to the DB
' ---------------------------------------------------------------
Public Function Stage3_CopyData() As Boolean
    Dim Rw As Integer
    Dim ShtAsset As ClsAsset
    Dim DBAsset As ClsAsset
    
    Const StrPROCEDURE As String = "Stage3_CopyData()"

    On Error GoTo ErrorHandler

    For Rw = 1 To MaxAssetNo
        
        'debug.print "Copying " & Rw & " of " & MaxAssetNo
        
        Set ShtAsset = ShtAssets(CStr(Rw))
        Set DBAsset = DBAssets(CStr(Rw))
        
        If ShtAsset Is Nothing Then
            If Not DBAsset Is Nothing Then DBAsset.DBDelete
        Else
        
            'don't overwrite quantity
            If Not DBAsset Is Nothing Then ShtAsset.QtyInStock = DBAsset.QtyInStock
            ShtAsset.DBSave Rw
        End If
        Rw = FrmDataImport.UpdateProgrGges(MaxAssetNo, Rw, 3)
        
        If Rw = 0 Then Err.Raise HANDLED_ERROR, Description:="Error updating gauges"
    Next

    Set ShtAsset = Nothing
    Set DBAsset = Nothing
    Stage3_CopyData = True

Exit Function

ErrorExit:

    Set ShtAsset = Nothing
    Set DBAsset = Nothing
'    ***CleanUpCode***
    Stage3_CopyData = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function


' ===============================================================
' Stage4_Validate
' Validates asset file with data in database
' ---------------------------------------------------------------
Public Function Stage4_Validate() As Boolean
    Dim Rw As Integer
    Dim DBAsset As ClsAsset
    Dim ShtAsset As ClsAsset
    
    Const StrPROCEDURE As String = "Stage4_Validate()"

    On Error GoTo ErrorHandler

    Set DBAssets = Nothing
    Set DBAssets = New ClsAssets
    DBAssets.GetCollection

    For Rw = 1 To MaxAssetNo
        
        Set ShtAsset = ShtAssets(CStr(Rw))
        Set DBAsset = DBAssets(CStr(Rw))
        
        'debug.print "Validating " & Rw & " of " & MaxAssetNo
        
        If ShtAsset Is Nothing Then
            If Not DBAsset Is Nothing Then AddToErrorLog Rw, "Failed Validation - Mismatch"
        Else
            If DBAsset Is Nothing Then
                AddToErrorLog Rw, "Failed Validation - Mismatch"
            Else
                With ShtAsset
                    If .AllocationType <> DBAsset.AllocationType Then AddToErrorLog Rw, "Failed Validation - Allocation Type"
                    If .Brand <> DBAsset.Brand Then AddToErrorLog Rw, "Failed Validation - Brand"
                    If .Description <> DBAsset.Description Then AddToErrorLog Rw, "Failed Validation - Description"
                    If .Category1 <> DBAsset.Category1 Then AddToErrorLog Rw, "Failed Validation - Category 1"
                    If .Category2 <> DBAsset.Category2 Then AddToErrorLog Rw, "Failed Validation - Category 2"
                    If .Category3 <> DBAsset.Category3 Then AddToErrorLog Rw, "Failed Validation - Category 3"
                    If .Size1 <> DBAsset.Size1 Then AddToErrorLog Rw, "Failed Validation - Size 1"
                    If .Size2 <> DBAsset.Size2 Then AddToErrorLog Rw, "Failed Validation - Size 2"
                    If .PurchaseUnit <> DBAsset.PurchaseUnit Then AddToErrorLog Rw, "Failed Validation - Purchase Unit"
                    If .MinAmount <> DBAsset.MinAmount Then AddToErrorLog Rw, "Failed Validation - Min Amount"
                    If .MaxAmount <> DBAsset.MaxAmount Then AddToErrorLog Rw, "Failed Validation - Max Amount"
                    If .OrderLevel <> DBAsset.OrderLevel Then AddToErrorLog Rw, "Failed Validation - Order Level"
                    If .LeadTime <> DBAsset.LeadTime Then AddToErrorLog Rw, "Failed Validation - Lead Time"
                    If .Keywords <> DBAsset.Keywords Then AddToErrorLog Rw, "Failed Validation - Keywords"
                    If .AllowedOrderReasons <> DBAsset.AllowedOrderReasons Then AddToErrorLog Rw, "Failed Validation - Order Reasons"
                    If .AdditInfo <> DBAsset.AdditInfo Then AddToErrorLog Rw, "Failed Validation - Addit Info"
                    If .NoOrderMessage <> DBAsset.NoOrderMessage Then AddToErrorLog Rw, "Failed Validation - No Order Message"
                    If .Location <> DBAsset.Location Then AddToErrorLog Rw, "Failed Validation - Location"
                    If .cost <> DBAsset.cost Then AddToErrorLog Rw, "Failed Validation - Cost"
                    If .Supplier1 <> DBAsset.Supplier1 Then AddToErrorLog Rw, "Failed Validation - Supplier 1"
                    If .Supplier2 <> DBAsset.Supplier2 Then AddToErrorLog Rw, "Failed Validation - Supplier 2"
                    If .QtyInStock <> DBAsset.QtyInStock And .QtyInStock > 0 Then AddToErrorLog Rw, "Failed Validation - Quantity"

            End With
        End If
        End If
        Rw = FrmDataImport.UpdateProgrGges(MaxAssetNo, Rw, 4)
            
        If Rw = 0 Then Err.Raise HANDLED_ERROR, Description:="Error updating gauges"

    Next

    Set ShtAsset = Nothing
    Set DBAsset = Nothing

    Stage4_Validate = True

Exit Function

ErrorExit:

    Set ShtAsset = Nothing
    Set DBAsset = Nothing
'    ***CleanUpCode***
    Stage4_Validate = False

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
' AddToErrorLog
' Adds import errors to error log
' ---------------------------------------------------------------
Private Sub AddToErrorLog(ByVal AssetNo As String, StrError As String)
    
    On Error Resume Next

    If ErrorCount < 2000 Then
    ErrorCount = ErrorCount + 1
    
    ErrorLog(ErrorCount) = "Asset No " & AssetNo & " - " & StrError
'    Debug.Print "Asset No " & AssetNo & " - " & StrError
    End If
End Sub

' ===============================================================
' AddToWarningLog
' Adds import warnings to warning log
' ---------------------------------------------------------------
Private Sub AddToWarningLog(ByVal AssetNo As String, StrWarning As String)
    
    On Error Resume Next

    If WarningCount < 2000 Then
    WarningCount = WarningCount + 1
    
    WarningLog(WarningCount) = "Asset No " & AssetNo & " - " & StrWarning
    Debug.Print "Asset No " & AssetNo & " - " & StrWarning
    End If
End Sub

' ===============================================================
' ImportTerminate
' Closes down asset collections
' ---------------------------------------------------------------
Public Function ImportTerminate() As Boolean
    Const StrPROCEDURE As String = "ImportTerminate()"

    On Error GoTo ErrorHandler

    ErrorCount = 0
    WarningCount = 0
    Set ShtAssets = Nothing
    Set DBAssets = Nothing

    ImportTerminate = True

Exit Function

ErrorExit:

'    ***CleanUpCode***
    ImportTerminate = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' ClearLog
' Clears all errors and warnings from log
' ---------------------------------------------------------------
Public Sub ClearLog()
    Dim i As Integer
    
    On Error Resume Next
    
    ErrorCount = 0
    WarningCount = 0
    
    For i = LBound(ErrorLog) To UBound(ErrorLog)
        ErrorLog(i) = ""
    Next
    
    For i = LBound(WarningLog) To UBound(WarningLog)
        WarningLog(i) = ""
    Next
End Sub
