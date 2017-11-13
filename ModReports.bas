Attribute VB_Name = "ModReports"
'===============================================================
' Module ModReports
' v0,0 - Initial Version
' v0,1 - Updated Query 1
' v0,2 - Added query for Report 2
' v0,3 - Prevent Deleted orders being included in Order Report
' v0,4 - Prevent deleted line items being included in Order Report
' v0,5 - Exclude Orders with Null or 0 Order No in Report 1
' v0,6 - Add cost to Order Report
' v0,7 - Added query for Report 3
'---------------------------------------------------------------
' Date - 13 Nov 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModReports"

' ===============================================================
' CreateReport
' Creates Report from Recordset
' ---------------------------------------------------------------
Public Function CreateReport(RstData As Recordset, ColWidths() As Integer, Headings() As String, ColFormats() As String) As Boolean
    Dim ReportBook As Workbook
    Dim RngQry As Range
    Dim RngHeader As Range
    Dim ShtReport As Worksheet
    
    Dim i As Integer
    
    Const StrPROCEDURE As String = "CreateReport()"

    On Error GoTo ErrorHandler

    Set ReportBook = Workbooks.Add
    Set ShtReport = ReportBook.Worksheets(1)
    
    With ShtReport
        Set RngQry = .Range("A1")
        
        'headings and col widths
        For i = 0 To UBound(Headings)
            RngQry.Offset(0, i) = Headings(i)
            RngQry.Offset(0, i).ColumnWidth = ColWidths(i)
        Next
        
        'formats
        For i = 0 To UBound(ColFormats)
            .Columns(i + 1).NumberFormat = ColFormats(i)
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
    Set ShtReport = Nothing
    CreateReport = True

Exit Function

ErrorExit:

    Set RngQry = Nothing
    Set RngHeader = Nothing
    Set ReportBook = Nothing
    Set ShtReport = Nothing
    
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
    Dim StrSelect As String
    Dim StrFrom As String
    Dim StrWhere As String
    Dim StrOrderBy As String
    
    Const StrPROCEDURE As String = "Report1Query()"

    On Error GoTo ErrorHandler

    StrSelect = "SELECT " _
                    & "TblOrder.OrderNo AS [Order No], " _
                    & "TblOrder.OrderDate AS [Order Date], " _
                    & "TblPerson1.Username AS [Ordered By], " _
                    & "TblAsset.Description AS Description, " _
                    & "TblAsset.Category1 AS [Category 1], " _
                    & "TblAsset.Category2 AS [Category 2], " _
                    & "TblAsset.Category3 AS [Category 3], " _
                    & "TblAsset.Size1 AS [Size 1], " _
                    & "TblAsset.Size2 AS [Size 2], " _
                    & "TblLineItem.Quantity AS Quantity, " _
                    & "TblPerson.Username AS [For Person], " _
                    & "TblStation.Name AS [For Station], " _
                    & "TblVehicle.VehReg AS [For Vehicle], " _
                    & "TblStation1.Name AS [Vehicle Station], " _
                    & "TblReqReason.ReqReason AS [Request Reason], " _
                    & "TblAsset.Cost * TblLineItem.Quantity AS [Total Cost] "
                    
    StrFrom = "FROM " _
                    & "(((((((TblLineItem " _
                    & "LEFT JOIN TblAsset ON TblLineItem.AssetID = TblAsset.AssetNo) " _
                    & "LEFT JOIN TblPerson ON TblLineItem.ForPersonID = TblPerson.CrewNo) " _
                    & "LEFT JOIN TblStation ON TblLineItem.ForStationID = TblStation.StationID) " _
                    & "LEFT JOIN TblVehicle ON TblLineItem.ForVehicleID = TblVehicle.VehNo) " _
                    & "LEFT JOIN TblReqReason ON TblLineItem.ReqReason = TblReqReason.ReqReasonNo) " _
                    & "LEFT JOIN TblStation TblStation1 ON TblVehicle.StationID = TblStation1.StationID) " _
                    & "LEFT JOIN TblOrder ON TblLineItem.OrderNo = TblOrder.OrderNo) " _
                    & "LEFT JOIN TblPerson TblPerson1 ON TblOrder.RequestorID = TblPerson1.CrewNo "
                    
    StrWhere = "WHERE " _
                    & "TblAsset.AssetNo IS NOT NULL " _
                    & "AND TblOrder.Deleted IS NULL " _
                    & "AND TblLineItem.Deleted IS NULL " _
                    & "AND TblOrder.OrderNo IS NOT NULL " _
                    & "AND TblOrder.OrderNo <> 0 "
                    
    StrOrderBy = "ORDER BY " _
                    & "Tblorder.OrderNo"
                    
    Set RstQuery = ModDatabase.SQLQuery(StrSelect & StrFrom & StrWhere & StrOrderBy)

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

' ===============================================================
' Report2Query
' SQL query for Report 2 returning results as recordset
' ---------------------------------------------------------------
Public Function Report2Query() As Recordset
    Dim RstQuery As Recordset
    Dim StrSelect As String
    Dim StrFrom As String
    Dim StrWhere As String
    
    Const StrPROCEDURE As String = "Report2Query()"

    On Error GoTo ErrorHandler

    StrSelect = "SELECT " _
                    & "TblAsset.AssetNo, " _
                    & "TblAsset.Description, " _
                    & "TblAsset.QtyInStock, " _
                    & "TblAsset.Category1, " _
                    & "TblAsset.Category2, " _
                    & "TblAsset.Category3, " _
                    & "TblAsset.Size1, " _
                    & "TblAsset.Size2, " _
                    & "TblAsset.Cost AS [Item Cost], " _
                    & "TblAsset.QtyInStock * TblAsset.Cost AS [Total Cost] "
                    
    StrFrom = "FROM " _
                    & "TblAsset "
    StrWhere = ""

    Set RstQuery = ModDatabase.SQLQuery(StrSelect & StrFrom & StrWhere)

    Set Report2Query = RstQuery

    Set RstQuery = Nothing
Exit Function

ErrorExit:

'    ***CleanUpCode***
    Set Report2Query = Nothing
    Set RstQuery = Nothing
Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' Report3Query
' SQL query for Report 3 returning results as recordset
' ---------------------------------------------------------------
Public Function Report3Query() As Recordset
    Dim StrSQL1 As String
    Dim StrSQL2 As String
    Dim StrSQL3 As String
    Dim RstQueryAll As Recordset
    Dim StrSelect As String
    Dim StrFrom As String
    Dim StrWhere As String
    Dim StrOrderBy As String
    
    Const StrPROCEDURE As String = "Report3Query()"

    On Error GoTo ErrorHandler

    'Create Query 1 for Station allocation type
    '-------------------------------------------
    StrSelect = "SELECT " _
                    & "TblOrder.OrderNo AS [Order No], " _
                    & "TblOrder.OrderDate AS [Date], " _
                    & "TblAsset.Description, " _
                    & "TblLineItem.Quantity, " _
                    & "TblLineItem.Quantity * TblAsset.Cost AS [Total Cost], " _
                    & "TblStation.StationNo AS [Station No], " _
                    & "TblStation.Name AS [Station Name], " _
                    & "TblStation.Division "
                    
    StrFrom = "FROM " _
                    & "((TblLineItem " _
                    & "LEFT JOIN TblAsset ON TblLineItem.AssetID = TblAsset.AssetNo) " _
                    & "LEFT JOIN TblOrder ON TblOrder.OrderNo = TblLineItem.OrderNo) " _
                    & "LEFT JOIN TblStation ON TblLineItem.ForStationID = TblStation.StationID "
    
    StrWhere = "WHERE " _
                    & "TblLineItem.ReturnReqd = YES AND " _
                    & "TblLineItem.itemsReturned = NO AND " _
                    & "TblAsset.AllocationType = 2 "
                    
    StrSQL1 = StrSelect & StrFrom & StrWhere & StrOrderBy

    'Create Query 2 for Vehicle allocation type
    '-------------------------------------------
    StrSelect = "SELECT " _
                    & "TblOrder.OrderNo AS [Order No], " _
                    & "TblOrder.OrderDate AS [Date], " _
                    & "TblAsset.Description, " _
                    & "TblLineItem.Quantity, " _
                    & "TblLineItem.Quantity * TblAsset.Cost AS [Total Cost], " _
                    & "TblStation.StationNo AS [Station No], " _
                    & "TblStation.Name AS [Station Name], " _
                    & "TblStation.Division "
                    
    StrFrom = "FROM " _
                    & "(((TblLineItem " _
                    & "LEFT JOIN TblAsset ON TblLineItem.AssetID = TblAsset.AssetNo) " _
                    & "LEFT JOIN TblOrder ON TblOrder.OrderNo = TblLineItem.OrderNo) " _
                    & "LEFT JOIN TblVehicle ON TblLineItem.ForVehicleID = TblVehicle.VehNo) " _
                    & "LEFT JOIN TblStation ON TblVehicle.StationID = TblStation.StationID "
                     
    StrWhere = "WHERE " _
                    & "TblLineItem.ReturnReqd = YES AND " _
                    & "TblLineItem.itemsReturned = NO AND " _
                    & "TblAsset.AllocationType = 1 AND " _
                    & "TblOrder.OrderDate IS NOT NULL AND " _
                    & "TblStation.StationNo IS NOT NULL "
                                    
    StrSQL2 = StrSelect & StrFrom & StrWhere & StrOrderBy
    
    'Create Query 3 for Person allocation type
    '-------------------------------------------
    StrSelect = "SELECT " _
                    & "TblOrder.OrderNo AS [Order No], " _
                    & "TblOrder.OrderDate AS [Date], " _
                    & "TblAsset.Description, " _
                    & "TblLineItem.Quantity, " _
                    & "TblLineItem.Quantity * TblAsset.Cost AS [Total Cost], " _
                    & "TblStation.StationNo AS [Station No], " _
                    & "TblStation.Name AS [Station Name], " _
                    & "TblStation.Division "
                    
    StrFrom = "FROM " _
                    & "(((TblLineItem " _
                    & "LEFT JOIN TblAsset ON TblLineItem.AssetID = TblAsset.AssetNo) " _
                    & "LEFT JOIN TblOrder ON TblOrder.OrderNo = TblLineItem.OrderNo) " _
                    & "LEFT JOIN TblPerson ON TblLineItem.ForPersonID = TblPerson.CrewNo) " _
                    & "LEFT JOIN TblStation ON TblPerson.StationID = TblStation.StationID "
                    
    StrWhere = "WHERE " _
                    & "TblLineItem.ReturnReqd = YES AND " _
                    & "TblLineItem.itemsReturned = NO AND " _
                    & "TblAsset.AllocationType = 0 AND " _
                    & "TblOrder.OrderDate IS NOT NULL "
    
    StrOrderBy = "ORDER BY " _
                    & "[Order No]"
                                            
    StrSQL3 = StrSelect & StrFrom & StrWhere & StrOrderBy
    
    ' Run Query
    '----------
    Set RstQueryAll = SQLQuery(StrSQL1 & " UNION ALL " & StrSQL2 & " UNION ALL " & StrSQL3)
 
    Set Report3Query = RstQueryAll

    Set RstQueryAll = Nothing
Exit Function

ErrorExit:

'    ***CleanUpCode***
    Set Report3Query = Nothing
    Set RstQueryAll = Nothing
Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function



