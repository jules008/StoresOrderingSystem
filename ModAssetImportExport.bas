Attribute VB_Name = "ModAssetImportExport"
'===============================================================
' Module ModAssetImportExport
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 19 May 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModAssetImportExport"

Dim ErrorLog(1 To 2000) As String
Dim ErrorCount As Integer
' ===============================================================
' ImportAssetFile
' Imports the Asset file into the database
' ---------------------------------------------------------------
Private Sub ImportAssetFile()
    Dim LineInputString As String
    Dim AssetData() As String
    Dim FormValidation As Integer
    Dim AssetFileLoc As String
    Dim Asset As ClsAsset
    Dim ShtAssets As ClsAssets
    Dim AssetFile As Integer
    Dim DBAssets As ClsAssets
    Dim i As Integer
    Dim x As Integer
    Dim MaxAssetNo As Integer
    Dim FuncPassFail As String
    
    Const StrPROCEDURE As String = "ImportAssetFile()"

    On Error Resume Next
    
    Set ShtAssets = New ClsAssets
    Set DBAssets = New ClsAssets
    Set Asset = New ClsAsset

'    AssetFileLoc = OpenAssetFile
    
    DBConnect
    
    If AssetFileLoc = "Error" Then Err.Raise HANDLED_ERROR
    
    AssetFileLoc = "\\lincsfire.lincolnshire.gov.uk\folderredir$\Documents\julian.turner\Documents\RDS Project\Stores IT Project\Data\tblasset.csv"
    
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
        For x = 1 To 25
            If Left(AssetData(x), 1) = Chr(34) Then AssetData(x) = Right(AssetData(x), Len(AssetData(x)) - 1)
            If Right(AssetData(x), 1) = Chr(34) Then AssetData(x) = Left(AssetData(x), Len(AssetData(x)) - 1)
            AssetData(x) = Replace(AssetData(x), Chr(34) & Chr(34), Chr(34))
        Next
        
        i = i + 1
        
        MaxAssetNo = DBAssets.MaxAssetNo
        
        If i <> 1 Then
        
            If Not ParseAsset(AssetData, i) Then Err.Raise HANDLED_ERROR
                        
            Set Asset = BuildAsset(AssetData)
            
            If Asset Is Nothing Then Err.Raise HANDLED_ERROR
            
            ShtAssets.AddItem Asset
            
            'find maximum assetno
            If Asset.AssetNo > MaxAssetNo Then MaxAssetNo = Asset.AssetNo
                        
            Debug.Print "Asset Added!"
        End If
    Wend
    Close #AssetFile
    
    MsgBox ErrorCount & " errors have been found", vbCritical, APP_NAME
    Stop

    If Not PreBuildCheck(ShtAssets, DBAssets) Then Err.Raise HANDLED_ERROR
            
    MsgBox ErrorCount & " errors have been found", vbCritical, APP_NAME
    Stop
    
    If Not CopyAssetFile(ShtAssets, DBAssets, MaxAssetNo) Then Err.Raise HANDLED_ERROR
            
    MsgBox ErrorCount & " errors have been found", vbCritical, APP_NAME
    Stop
    
    If Not ValidateAssetFile(ShtAssets, MaxAssetNo) Then Err.Raise HANDLED_ERROR
        
    MsgBox ErrorCount & " errors have been found", vbCritical, APP_NAME
    Stop

    MsgBox "Complete"

    Set ShtAssets = Nothing
    Set Asset = Nothing
    Set DBAssets = Nothing

Exit Sub

ErrorExit:
    
    Application.DisplayAlerts = True

'    ***CleanUpCode***
    Set ShtAssets = Nothing
    Set Asset = Nothing
    Set DBAssets = Nothing
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
    If UBound(AssetData) < 25 Then
        AddToErrorLog AssetNo, "Incorrect use of commas"
        PassGenericTests = False
    End If
    
    If AssetNo = PrevAssetNo Then
        AddToErrorLog AssetNo, "Duplicate Asset No"
        PassGenericTests = False
    End If

    If AssetData(25) <> "!" Then AddToErrorLog AssetNo, "Number of columns incorrect, check use of commas"
    
    If PassGenericTests Then
        For i = 0 To 25
    
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
        
                Case Is = 22
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
        If AssetData(22) <> "" Then .cost = AssetData(22)
        .Supplier1 = Trim(AssetData(23))
        .Supplier2 = Trim(AssetData(24))
    
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

    On Error Resume Next

    Set DlgOpen = Application.FileDialog(msoFileDialogOpen)
    Set ShtAssets = New ClsAssets
    
     With DlgOpen
        .Filters.Clear
        .Filters.Add "CSV Files (*.csv)", "*.csv"
        .AllowMultiSelect = False
        .Title = "Select Spreadsheet of Doom"
'        .Show
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
' PreBuildCheck
' Checks before writing to DB
' ---------------------------------------------------------------
Private Function PreBuildCheck(ShtAssets As ClsAssets, DBAssets As ClsAssets) As Boolean
    Dim i As Integer
    Dim DBAssetNo As Integer
    Dim Asset As ClsAsset
    Dim DBAssetDescription As String
    Dim Response As Integer
    
    Const StrPROCEDURE As String = "PreBuildCheck()"

    On Error Resume Next

    Set Asset = New ClsAsset
    
    For Each Asset In DBAssets
        
        DBAssetNo = Asset.AssetNo
        
        Debug.Print DBAssetNo
        
        DBAssetDescription = Asset.Description
        
        If ShtAssets(CStr(DBAssetNo)) Is Nothing Then
            AddToErrorLog DBAssetNo, DBAssetDescription & " will be deleted from database"
            
        Else
            If ShtAssets(CStr(DBAssetNo)).Description <> DBAssetDescription Then
                AddToErrorLog DBAssetNo, "Asset will change from " & DBAssetDescription & " to " & ShtAssets(CStr(DBAssetNo)).Description
            End If
        End If
    Next

    Set Asset = Nothing
    
    PreBuildCheck = True

Exit Function

ErrorExit:

    Set Asset = Nothing
'    ***CleanUpCode***
    PreBuildCheck = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' CopyAssetFile
' Copies the asset file data to the DB
' ---------------------------------------------------------------
Private Function CopyAssetFile(ShtAssets As ClsAssets, DBAssets As ClsAssets, MaxAssetNo As Integer) As Boolean
    Dim i As Integer
    Dim ShtAsset As ClsAsset
    Dim DBAsset As ClsAsset
    
    Const StrPROCEDURE As String = "CopyAssetFile()"

    On Error Resume Next

    For i = 1 To MaxAssetNo
            
        Set ShtAsset = ShtAssets(CStr(i))
        Set DBAsset = DBAssets(CStr(i))
        
        If ShtAsset Is Nothing Then
            If Not DBAsset Is Nothing Then DBAsset.DBDelete
            Else
            
            'don't overwrite quantity
            If ShtAsset.QtyInStock = 0 Then ShtAsset.QtyInStock = DBAsset.QtyInStock
            ShtAsset.DBSave i
        End If
        
    Next

    Set ShtAsset = Nothing
    Set DBAsset = Nothing
    CopyAssetFile = True

Exit Function

ErrorExit:

    Set ShtAsset = Nothing
    Set DBAsset = Nothing
'    ***CleanUpCode***
    CopyAssetFile = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function


' ===============================================================
' ValidateAssetFile
' Validates asset file with data in database
' ---------------------------------------------------------------
Private Function ValidateAssetFile(ShtAssets As ClsAssets, MaxAssetNo As Integer) As Boolean
    Dim i As Integer
    Dim DBAssets As ClsAssets
    Dim DBAsset As ClsAsset
    Dim ShtAsset As ClsAsset
    
    Const StrPROCEDURE As String = "ValidateAssetFile()"

    On Error Resume Next

    Set DBAssets = New ClsAssets
    DBAssets.GetCollection

    For i = 1 To MaxAssetNo
        
        Set ShtAsset = ShtAssets(CStr(i))
        Set DBAsset = DBAssets(CStr(i))
        
        If ShtAsset Is Nothing Then
            If Not DBAsset Is Nothing Then AddToErrorLog i, "Failed Validation - Mismatch"
        Else
            If DBAsset Is Nothing Then
                AddToErrorLog i, "Failed Validation - Mismatch"
            Else
                With ShtAsset
                     If .AllocationType <> DBAsset.AllocationType Then AddToErrorLog i, "Failed Validation - Allocation Type"
                    If .Brand <> DBAsset.Brand Then AddToErrorLog i, "Failed Validation - Brand"
                    If .Description <> DBAsset.Description Then AddToErrorLog i, "Failed Validation - Description"
                    If .Category1 <> DBAsset.Category1 Then AddToErrorLog i, "Failed Validation - Category 1"
                    If .Category2 <> DBAsset.Category2 Then AddToErrorLog i, "Failed Validation - Category 2"
                    If .Category3 <> DBAsset.Category3 Then AddToErrorLog i, "Failed Validation - Category 3"
                    If .Size1 <> DBAsset.Size1 Then AddToErrorLog i, "Failed Validation - Size 1"
                    If .Size2 <> DBAsset.Size2 Then AddToErrorLog i, "Failed Validation - Size 2"
                    If .PurchaseUnit <> DBAsset.PurchaseUnit Then AddToErrorLog i, "Failed Validation - Purchase Unit"
                    If .MinAmount <> DBAsset.MinAmount Then AddToErrorLog i, "Failed Validation - Min Amount"
                    If .MaxAmount <> DBAsset.MaxAmount Then AddToErrorLog i, "Failed Validation - Max Amount"
                    If .OrderLevel <> DBAsset.OrderLevel Then AddToErrorLog i, "Failed Validation - Order Level"
                    If .LeadTime <> DBAsset.LeadTime Then AddToErrorLog i, "Failed Validation - Lead Time"
                    If .Keywords <> DBAsset.Keywords Then AddToErrorLog i, "Failed Validation - Keywords"
                    If .AllowedOrderReasons <> DBAsset.AllowedOrderReasons Then AddToErrorLog i, "Failed Validation - Order Reasons"
                    If .AdditInfo <> DBAsset.AdditInfo Then AddToErrorLog i, "Failed Validation - Addit Info"
                    If .NoOrderMessage <> DBAsset.NoOrderMessage Then AddToErrorLog i, "Failed Validation - No Order Message"
                    If .Location <> DBAsset.Location Then AddToErrorLog i, "Failed Validation - Location"
                    If .cost <> DBAsset.cost Then AddToErrorLog i, "Failed Validation - Cost"
                    If .Supplier1 <> DBAsset.Supplier1 Then AddToErrorLog i, "Failed Validation - Supplier 1"
                    If .Supplier2 <> DBAsset.Supplier2 Then AddToErrorLog i, "Failed Validation - Supplier 2"
                    If .QtyInStock <> DBAsset.QtyInStock And .QtyInStock > 0 Then AddToErrorLog i, "Failed Validation - Quantity"

            End With
        End If
        End If
    Next

    Set ShtAsset = Nothing
    Set DBAsset = Nothing

    ValidateAssetFile = True

Exit Function

ErrorExit:

    Set ShtAsset = Nothing
    Set DBAsset = Nothing
'    ***CleanUpCode***
    ValidateAssetFile = False

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
    Debug.Print "Asset No " & AssetNo & " - " & StrError
    End If
End Sub
