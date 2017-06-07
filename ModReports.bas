Attribute VB_Name = "ModReports"
'===============================================================
' Module ModReports
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 07 Jun 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModReports"

' ===============================================================
' CreateReport
' Creates Report from Recordset
' ---------------------------------------------------------------
Public Function CreateReport(RstData As Recordset, ColWidths() As Integer, Headings() As String) As Boolean
    Dim ReportBook As Workbook
    Dim RngQry As Range
    Dim RngHeader As Range
    Dim i As Integer
    
    Const StrPROCEDURE As String = "CreateReport()"

    On Error GoTo ErrorHandler

    Set ReportBook = Workbooks.Add
    
    With ReportBook.Worksheets(1)
        Set RngQry = .Range("A1")
        
        'headings and col widths
        For i = 0 To UBound(Headings)
            RngQry.Offset(0, i) = Headings(i)
            RngQry.Offset(0, i).ColumnWidth = ColWidths(i)
        Next
        
        'format heading
        Set RngHeader = .Range(Cells(1, 1), Cells(1, UBound(Headings) + 1))
    
        With RngHeader
            .Interior.Color = COLOUR_9
            .Borders.Color = COLOUR_2
            .Font.Bold = True
        
            'set filter
            .AutoFilter
        End With

        RngQry.Offset(1, 0).CopyFromRecordset RstData
    
    End With
    
    Set RngQry = Nothing
    Set RngHeader = Nothing
    Set ReportBook = Nothing
    CreateReport = True

Exit Function

ErrorExit:

    Set RngQry = Nothing
    Set RngHeader = Nothing
    Set ReportBook = Nothing
'    ***CleanUpCode***
    CreateReport = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' Report1Query
' SQL query for Report 1 returning results as recordset
' ---------------------------------------------------------------
Public Function Report1Query() As Recordset
    Dim RstQuery As Recordset
    
    Const StrPROCEDURE As String = "Report1Query()"

    On Error GoTo ErrorHandler

    Set RstQuery = ModDatabase.SQLQuery("SELECT " _
                                  & "TblAsset.AssetNo AS [Asset No], " _
                                  & "TblAsset.Description AS Description, " _
                                  & "TblAsset.Category1, " _
                                  & "TblAsset.Category2, " _
                                  & "TblAsset.Category3, " _
                                  & "TblAsset.Size1, " _
                                  & "TblAsset.Size2, " _
                                  & "TblLineItem.Quantity AS Quantity, " _
                                  & "TblPerson.Username AS [For Person], " _
                                  & "TblStation.Name AS [For Station], " _
                                  & "TblVehicle.VehReg AS [For Vehicle], " _
                                  & "TblStation1.Name AS [Vehicle Station], " _
                                  & "TblReqReason.ReqReason AS [Request Reason] " _
                                & "From " _
                                  & "(((((TblLineItem " _
                                  & "LEFT JOIN TblAsset ON TblLineItem.AssetID = TblAsset.AssetNo) " _
                                  & "LEFT JOIN TblPerson ON TblLineItem.ForPersonID = TblPerson.CrewNo) " _
                                  & "LEFT JOIN TblStation ON TblLineItem.ForStationID = TblStation.StationID) " _
                                  & "LEFT JOIN TblVehicle ON TblLineItem.ForVehicleID = TblVehicle.VehNo) " _
                                  & "LEFT JOIN TblReqReason ON TblLineItem.ReqReason = TblReqReason.ReqReasonNo) " _
                                  & "LEFT JOIN TblStation TblStation1 ON TblVehicle.StationID = TblStation1.StationID " _
                                & "WHERE " _
                                  & "TblAsset.AssetNo IS NOT NULL " _
                                  & "ORDER BY TblLineItem.OrderNo")



    Set Report1Query = RstQuery

    Set RstQuery = Nothing
Exit Function

ErrorExit:

'    ***CleanUpCode***
    Set Report1Query = Nothing
    Set RstQuery = Nothing
Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function
