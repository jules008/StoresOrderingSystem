Attribute VB_Name = "ModPrint"
'===============================================================
' Module ModPrint
' v0,0 - Initial Version
' v0,1 - added PrintOrderList procedure
'---------------------------------------------------------------
' Date - 07 Apr 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModPrint"

' ===============================================================
' CreateTmpFile
' Creates and returns a new temp file
' ---------------------------------------------------------------
Public Function CreateTmpFile() As String
    Dim iFile As Integer
    Dim i As Integer
    Dim FileTxt As String
    Dim TmpFilePath As String
    
    Const StrPROCEDURE As String = "CreateTmpFile()"

    On Error GoTo ErrorHandler

    TmpFilePath = TMP_FILE_PATH

    If Right$(TmpFilePath, 1) <> "\" Then TmpFilePath = TmpFilePath & "\"

    iFile = FreeFile()
    
    Do
        i = i + 1
    Loop While Dir(TmpFilePath & "TmpFile" & i & ".txt") <> vbNullString
    
    Open TmpFilePath & "TmpFile" & i & ".txt" For Append As #iFile
    Close #iFile
    
    CreateTmpFile = TmpFilePath & "TmpFile" & i & ".txt"

Exit Function

ErrorExit:

'    ***CleanUpCode***
    CreateTmpFile = ""

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' PrintOrderReceipt
' Prints a text file
' ---------------------------------------------------------------
Public Function PrintOrderReceipt(Order As ClsOrder) As Boolean
    Dim PrintFilePath As String
    Dim iFile As Integer
    Dim LineItem As ClsLineItem
    
    Const StrPROCEDURE As String = "PrintOrderReceipt()"

    On Error GoTo ErrorHandler

    PrintFilePath = CreateTmpFile
    
    iFile = FreeFile()
    
    With Order
        Open PrintFilePath For Append As #iFile
            Print #iFile, "==================================================="
            Print #iFile,
            Print #iFile, "Order No: " & .OrderNo
            Print #iFile, "Order Date: " & .OrderDate
            Print #iFile, "Requested By: " & .Requestor.CrewNo & " " & .Requestor.UserName
            Print #iFile, "Station: " & .Requestor.Station.Name
            Print #iFile,
                        
            For Each LineItem In .LineItems
                With LineItem
                    Print #iFile,
                    Print #iFile, "---------------------------------------------------"
                    Print #iFile, "Desc: " & .Asset.Description
                    Print #iFile, "Qty: " & .Quantity
                    Print #iFile, "Size1: " & .Asset.Size1
                    Print #iFile, "Size2: " & .Asset.Size2
                End With
            Next
            Print #iFile, "==================================================="
            Print #iFile,
            Print #iFile,
            Print #iFile,
            Print #iFile,
        Close #iFile
        
        If ENABLE_PRINT Then Shell ("notepad.exe /p " & PrintFilePath)
        
        Kill PrintFilePath
        
        Set LineItem = Nothing
    End With
    
    PrintOrderReceipt = True

Exit Function

ErrorExit:

'    ***CleanUpCode***
    Set LineItem = Nothing
    PrintOrderReceipt = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' PrintOrderList
' Populates form with order items
' ---------------------------------------------------------------
Public Function PrintOrderList(Order As ClsOrder) As Boolean
    Dim RngOrderNo As Range
    Dim RngReqBy As Range
    Dim RngStation As Range
    Dim RngItemsRefPnt As Range
    Dim Lineitem As ClsLineItem
    Dim i As Integer
    
    Const StrPROCEDURE As String = "PrintOrderList()"

    On Error GoTo ErrorHandler

    ShtOrderList.ClearForm
    
    Set RngOrderNo = ShtOrderList.Range("C3")
    Set RngReqBy = ShtOrderList.Range("E3")
    Set RngStation = ShtOrderList.Range("G3")
    Set RngItemsRefPnt = ShtOrderList.Range("B6")

    With Order
        RngOrderNo = .OrderNo
        RngReqBy = .Requestor.UserName
        RngStation = .Requestor.Station.Name
        
        For i = 0 To .LineItems.Count - 1

            RngItemsRefPnt.Offset(i, 0) = .LineItems(i + 1).Asset.Description
            RngItemsRefPnt.Offset(i, 2) = .LineItems(i + 1).Quantity
            RngItemsRefPnt.Offset(i, 3) = .LineItems(i + 1).Asset.Size1
            RngItemsRefPnt.Offset(i, 4) = .LineItems(i + 1).Asset.Size2
            RngItemsRefPnt.Offset(i, 5) = .LineItems(i + 1).Asset.Location.Name
        Next
    End With
    
    If ENABLE_PRINT Then
        ShtOrderList.Visible = xlSheetVisible
        ShtOrderList.PrintOut
        ShtOrderList.Visible = xlSheetHidden
    End If
    
    PrintOrderList = True
    
    Set RngOrderNo = Nothing
    Set RngReqBy = Nothing
    Set RngStation = Nothing
    Set RngItemsRefPnt = Nothing
    Set Lineitem = Nothing

Exit Function

ErrorExit:
    Set RngOrderNo = Nothing
    Set RngReqBy = Nothing
    Set RngStation = Nothing
    Set RngItemsRefPnt = Nothing
    Set Lineitem = Nothing

'    ***CleanUpCode***
    PrintOrderList = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

