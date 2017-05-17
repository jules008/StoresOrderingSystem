Attribute VB_Name = "ModAssetImportExport"
'===============================================================
' Module ModAssetImportExport
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 17 May 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModAssetImportExport"

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
    Dim Assets As ClsAssets
    Dim AssetFile As Integer
    Dim RstAssets As Recordset
    Dim i As Integer
    Dim MaxAssetNo As Integer
    
    Const StrPROCEDURE As String = "ImportAssetFile()"

    On Error GoTo ErrorHandler
    
    Set Assets = New ClsAssets
    Set Asset = New ClsAsset

'    AssetFileLoc = OpenAssetFile
    
    If AssetFileLoc = "Error" Then Err.Raise HANDLED_ERROR
    
    AssetFileLoc = "\\lincsfire.lincolnshire.gov.uk\folderredir$\Documents\julian.turner\Documents\RDS Project\Stores IT Project\Data\tblasset.csv"
    
    AssetFile = FreeFile()
        
    Set RstAssets = Assets.GetAllAssets
    
    Open AssetFileLoc For Input As AssetFile
    
    While Not EOF(AssetFile)
        Line Input #AssetFile, LineInputString
        AssetData = Split(LineInputString, ",")
        i = i + 1
        
        Debug.Print "Starting Line: " & i
        
        If i <> 1 Then
        
            FormValidation = ParseAsset(AssetData, i)
            
            Select Case FormValidation
                Case 999
                    Err.Raise HANDLED_ERROR
                Case Is > 0
                    Err.Raise IMPORT_ERROR
            End Select
                        
            Set Asset = BuildAsset(AssetData)
            
            If Asset Is Nothing Then Err.Raise HANDLED_ERROR
            
            Assets.AddItem Asset
            
            'find maximum assetno
            If Asset.AssetNo > MaxAssetNo Then MaxAssetNo = Asset.AssetNo
                        
            Debug.Print "Asset Added!"
        End If
    Wend
    Close #AssetFile
    
    Stop
    
    If Not CopyAssetFile(Assets, RstAssets) Then Err.Raise HANDLED_ERROR
            
    Stop
    
    If Not ValidateAssetFile(Assets, RstAssets) Then Err.Raise HANDLED_ERROR

    MsgBox "Complete"

    Set Assets = Nothing
    Set Asset = Nothing
    Set RstAssets = Nothing

Exit Sub

ErrorExit:

'    ***CleanUpCode***
    Set Assets = Nothing
    Set Asset = Nothing
    Set RstAssets = Nothing
Exit Sub

ErrorHandler:

    If Err.Number >= 1000 And Err.Number <= 1500 Then
        
        If Err.Number = IMPORT_ERROR Then
            Select Case FormValidation
                Case Is < 25
                    MsgBox "There has been an error importing the data on line " & i & ", Field " & FormValidation + 1, vbExclamation, APP_NAME
                Case Is = 25
                    MsgBox "There is an error with the number of columns on line " _
                            & i & " This is commonly caused by the use of commas.  Please replace these by other puctuation marks", vbExclamation, APP_NAME
            End Select
            
            Stop
            Resume
        End If
    End If
      
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
Private Function ParseAsset(AssetData() As String, LineNo As Integer) As Integer
    Dim i As Integer
    Dim TestValue As String
    Dim TestString() As String
    
    Const StrPROCEDURE As String = "ParseAsset()"
    
    On Error GoTo ErrorHandler
    
    If UBound(AssetData) < 25 Then
        i = 25
        Err.Raise IMPORT_ERROR
    End If
    
    For i = 0 To 25
    
        TestValue = AssetData(i)
        
        'generic tests first
        If InStr(TestValue, "'") <> 0 Then Err.Raise IMPORT_ERROR

        Select Case i
            Case Is = 0
        
'** add check to ensure unique numeric key"
            
'** add check to ensure that asset description matches asset no

'** ensure category 1 is not NULL and other fields too
            
            Case Is = 1
                If Not IsNumeric(TestValue) Then Err.Raise IMPORT_ERROR
                If TestValue < 0 Or TestValue > 2 Then Err.Raise IMPORT_ERROR
        
            Case Is = 4
                If IsNumeric(TestValue) Then
                If TestValue < 0 Then Err.Raise IMPORT_ERROR
                Else
                    If TestValue <> "" Then Err.Raise IMPORT_ERROR
                End If
    
            Case Is = 11
                If Not IsNumeric(TestValue) Then Err.Raise IMPORT_ERROR
                If TestValue < 0 Then Err.Raise IMPORT_ERROR
            
            Case Is = 12
                If Not IsNumeric(TestValue) Then Err.Raise IMPORT_ERROR
                If TestValue < 0 Then Err.Raise IMPORT_ERROR
            
            Case Is = 13
                If Not IsNumeric(TestValue) Then Err.Raise IMPORT_ERROR
                If TestValue < 0 Then Err.Raise IMPORT_ERROR
            
            Case Is = 16
                
                If Len(TestValue) <> 13 Then Err.Raise IMPORT_ERROR
                
                On Error GoTo ValidationError
                
                TestString = Split(TestValue, ":")
                
                If TestString(0) <> "0" And TestString(0) <> "1" Then Err.Raise IMPORT_ERROR
                If TestString(1) <> "0" And TestString(1) <> "1" Then Err.Raise IMPORT_ERROR
                If TestString(2) <> "0" And TestString(2) <> "1" Then Err.Raise IMPORT_ERROR
                If TestString(3) <> "0" And TestString(3) <> "1" Then Err.Raise IMPORT_ERROR
                If TestString(4) <> "0" And TestString(4) <> "1" Then Err.Raise IMPORT_ERROR
                If TestString(5) <> "0" And TestString(5) <> "1" Then Err.Raise IMPORT_ERROR
                If TestString(6) <> "0" And TestString(6) <> "1" Then Err.Raise IMPORT_ERROR
    
                On Error GoTo ErrorHandler
    
            Case Is = 22
                If TestValue <> "" Then
                If Not IsNumeric(TestValue) Then Err.Raise IMPORT_ERROR
                If TestValue < 0 Then Err.Raise IMPORT_ERROR
                End If
            
            Case Is = 25
                If TestValue <> "!" Then Err.Raise IMPORT_ERROR
            
        End Select
        
        Next
        ParseAsset = 0
Exit Function
        
ValidationError:
    
Exit Function

ErrorExit:

'    ***CleanUpCode***
    ParseAsset = 999

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

    On Error GoTo ErrorHandler

    Set Asset = New ClsAsset

    With Asset
        .AssetNo = AssetData(0)
        .AllocationType = AssetData(1)
        .Brand = AssetData(2)
        .Description = AssetData(3)
        If AssetData(4) <> "" Then .QtyInStock = AssetData(4)
        .Category1 = AssetData(5)
        .Category2 = AssetData(6)
        .Category3 = AssetData(7)
        .Size1 = AssetData(8)
        .Size2 = AssetData(9)
        .PurchaseUnit = AssetData(10)
        .MinAmount = AssetData(11)
        .MaxAmount = AssetData(12)
        .OrderLevel = AssetData(13)
        If AssetData(14) <> "" Then .LeadTime = CInt(AssetData(14))
        .Keywords = AssetData(15)
        .AllowedOrderReasons = AssetData(16)
        .AdditInfo = AssetData(17)
        .NoOrderMessage = AssetData(18)
        .Location = AssetData(19)
        If AssetData(20) <> "" Then .Status = AssetData(20)
        If AssetData(21) <> "" Then .cost = CInt(AssetData(21))
'        .Supplier1 = AssetData(22)
'        .Supplier2 = AssetData(23)
    
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
    Dim Assets As ClsAssets
    Dim NoFiles As Integer
    
    Const StrPROCEDURE As String = "OpenAssetFile()"

    On Error GoTo ErrorHandler

    Set DlgOpen = Application.FileDialog(msoFileDialogOpen)
    Set Assets = New ClsAssets
    
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
    Set Assets = Nothing
Exit Function

ErrorExit:

'    ***CleanUpCode***
    Set DlgOpen = Nothing
    Set Assets = Nothing
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
' CopyAssetFile
' Copies the asset file data to the DB
' ---------------------------------------------------------------
Private Function CopyAssetFile(Assets As ClsAssets, RstAssets As Recordset) As Boolean
    Dim MaxAssetNo As Integer
    Dim i As Integer
    
    Const StrPROCEDURE As String = "CopyAssetFile()"

    On Error GoTo ErrorHandler

    RstAssets.MoveFirst
    For i = 1 To MaxAssetNo
        Debug.Print "Assets.AssetNo: " & i
        Debug.Print "RST.AssetNo: " & RstAssets!AssetNo
            
        If Assets(i).AssetNo = i Then
                Assets(i).DBSave
            RstAssets.MoveNext
            Debug.Print "Add to DB"
        Else
            If RstAssets!AssetNo = i Then
                RstAssets.Delete
                Debug.Print "Delete from DB"
            Else
                RstAssets.MoveNext
                Debug.Print "do nothing"
            End If
        End If
        Debug.Print
    Next




    CopyAssetFile = True

Exit Function

ErrorExit:

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
Private Function ValidateAssetFile(Assets As ClsAssets, RstAssets As Recordset) As Boolean
    Dim i As Integer
    Dim MaxAssetNo As Integer
    
    Const StrPROCEDURE As String = "ValidateAssetFile()"

    On Error GoTo ErrorHandler

    RstAssets.MoveFirst
    
    For i = 1 To MaxAssetNo
        Debug.Print "Assets.AssetNo: " & i
        Debug.Print "RST.AssetNo: " & RstAssets!AssetNo
        
        If Assets(i).AssetNo = i And RstAssets!AssetNo = i Then
            With Assets(i)
                If .AssetNo <> RstAssets!AssetNo Then Err.Raise IMPORT_ERROR
                If .AllocationType <> RstAssets!AllocationType Then Err.Raise IMPORT_ERROR
                If .Brand <> RstAssets!Brand Then Err.Raise IMPORT_ERROR
                If .Description <> RstAssets!Description Then Err.Raise IMPORT_ERROR
                If .QtyInStock <> RstAssets!QtyInStock And .QtyInStock <> "" Then Err.Raise IMPORT_ERROR
                If .Category1 <> RstAssets!Category1 Then Err.Raise IMPORT_ERROR
                If .Category2 <> RstAssets!Category2 Then Err.Raise IMPORT_ERROR
                If .Category3 <> RstAssets!Category3 Then Err.Raise IMPORT_ERROR
                If .Size1 <> RstAssets!Size1 Then Err.Raise IMPORT_ERROR
                If .Size2 <> RstAssets!Size2 Then Err.Raise IMPORT_ERROR
                If .PurchaseUnit <> RstAssets!PurchaseUnit Then Err.Raise IMPORT_ERROR
                If .MinAmount <> RstAssets!MinAmount Then Err.Raise IMPORT_ERROR
                If .MaxAmount <> RstAssets!MaxAmount Then Err.Raise IMPORT_ERROR
                If .OrderLevel <> RstAssets!OrderLevel Then Err.Raise IMPORT_ERROR
                If .LeadTime <> RstAssets!LeadTime Then Err.Raise IMPORT_ERROR
                If .Keywords <> RstAssets!Keywords Then Err.Raise IMPORT_ERROR
                If .AllowedOrderReasons <> RstAssets!AllowedOrderReasons Then Err.Raise IMPORT_ERROR
                If .AdditInfo <> RstAssets!AdditInfo Then Err.Raise IMPORT_ERROR
                If .NoOrderMessage <> RstAssets!NoOrderMessage Then Err.Raise IMPORT_ERROR
                If .Location <> RstAssets!Location Then Err.Raise IMPORT_ERROR
                If .Status <> RstAssets!Status Then Err.Raise IMPORT_ERROR
                If .cost <> RstAssets!cost Then Err.Raise IMPORT_ERROR
        '        .Supplier1 <> AssetData(22) Then Err.Raise IMPORT_ERROR
        '        .Supplier2 <> AssetData(23) Then Err.Raise IMPORT_ERROR

            End With
            RstAssets.MoveNext
        End If
        Debug.Print
    Next

    ValidateAssetFile = True

Exit Function

ErrorExit:

'    ***CleanUpCode***
    ValidateAssetFile = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function
