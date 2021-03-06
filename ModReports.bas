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
' v0,81 - Schedule email reports
' v0,93 - Added Email Reports
'---------------------------------------------------------------
' Date - 30 Nov 17
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
                    & "TblLineitem.Quantity AS Quantity, " _
                    & "TblPerson.Username AS [For Person], " _
                    & "TblStation.Name AS [For Station], " _
                    & "TblVehicle.VehReg AS [For Vehicle], " _
                    & "TblStation1.Name AS [Vehicle Station], " _
                    & "TblReqReason.ReqReason AS [Request Reason], " _
                    & "TblAsset.Cost * TblLineitem.Quantity AS [Total Cost] "
                    
    StrFrom = "FROM " _
                    & "(((((((TblLineitem " _
                    & "LEFT JOIN TblAsset ON TblLineitem.AssetID = TblAsset.AssetNo) " _
                    & "LEFT JOIN TblPerson ON TblLineitem.ForPersonID = TblPerson.CrewNo) " _
                    & "LEFT JOIN TblStation ON TblLineitem.ForStationID = TblStation.StationID) " _
                    & "LEFT JOIN TblVehicle ON TblLineitem.ForVehicleID = TblVehicle.VehNo) " _
                    & "LEFT JOIN TblReqReason ON TblLineitem.ReqReason = TblReqReason.ReqReasonNo) " _
                    & "LEFT JOIN TblStation TblStation1 ON TblVehicle.StationID = TblStation1.StationID) " _
                    & "LEFT JOIN TblOrder ON TblLineitem.OrderNo = TblOrder.OrderNo) " _
                    & "LEFT JOIN TblPerson TblPerson1 ON TblOrder.RequestorID = TblPerson1.CrewNo "
                    
    StrWhere = "WHERE " _
                    & "TblAsset.AssetNo IS NOT NULL " _
                    & "AND TblOrder.Deleted IS NULL " _
                    & "AND TblLineitem.Deleted IS NULL " _
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
                    & "TblLineitem.Quantity, " _
                    & "TblLineitem.Quantity * TblAsset.Cost AS [Total Cost], " _
                    & "TblStation.StationNo AS [Station No], " _
                    & "TblStation.Name AS [Station Name], " _
                    & "TblStation.Division "
                    
    StrFrom = "FROM " _
                    & "((TblLineitem " _
                    & "LEFT JOIN TblAsset ON TblLineitem.AssetID = TblAsset.AssetNo) " _
                    & "LEFT JOIN TblOrder ON TblOrder.OrderNo = TblLineitem.OrderNo) " _
                    & "LEFT JOIN TblStation ON TblLineitem.ForStationID = TblStation.StationID "
    
    StrWhere = "WHERE " _
                    & "TblLineitem.ReturnReqd = YES AND " _
                    & "TblLineitem.itemsReturned = NO AND " _
                    & "TblAsset.AllocationType = 2 "
                    
    StrSQL1 = StrSelect & StrFrom & StrWhere & StrOrderBy

    'Create Query 2 for Vehicle allocation type
    '-------------------------------------------
    StrSelect = "SELECT " _
                    & "TblOrder.OrderNo AS [Order No], " _
                    & "TblOrder.OrderDate AS [Date], " _
                    & "TblAsset.Description, " _
                    & "TblLineitem.Quantity, " _
                    & "TblLineitem.Quantity * TblAsset.Cost AS [Total Cost], " _
                    & "TblStation.StationNo AS [Station No], " _
                    & "TblStation.Name AS [Station Name], " _
                    & "TblStation.Division "
                    
    StrFrom = "FROM " _
                    & "(((TblLineitem " _
                    & "LEFT JOIN TblAsset ON TblLineitem.AssetID = TblAsset.AssetNo) " _
                    & "LEFT JOIN TblOrder ON TblOrder.OrderNo = TblLineitem.OrderNo) " _
                    & "LEFT JOIN TblVehicle ON TblLineitem.ForVehicleID = TblVehicle.VehNo) " _
                    & "LEFT JOIN TblStation ON TblVehicle.StationID = TblStation.StationID "
                     
    StrWhere = "WHERE " _
                    & "TblLineitem.ReturnReqd = YES AND " _
                    & "TblLineitem.itemsReturned = NO AND " _
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
                    & "TblLineitem.Quantity, " _
                    & "TblLineitem.Quantity * TblAsset.Cost AS [Total Cost], " _
                    & "TblStation.StationNo AS [Station No], " _
                    & "TblStation.Name AS [Station Name], " _
                    & "TblStation.Division "
                    
    StrFrom = "FROM " _
                    & "(((TblLineitem " _
                    & "LEFT JOIN TblAsset ON TblLineitem.AssetID = TblAsset.AssetNo) " _
                    & "LEFT JOIN TblOrder ON TblOrder.OrderNo = TblLineitem.OrderNo) " _
                    & "LEFT JOIN TblPerson ON TblLineitem.ForPersonID = TblPerson.CrewNo) " _
                    & "LEFT JOIN TblStation ON TblPerson.StationID = TblStation.StationID "
                    
    StrWhere = "WHERE " _
                    & "TblLineitem.ReturnReqd = YES AND " _
                    & "TblLineitem.itemsReturned = NO AND " _
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

' ===============================================================
' ScheduleReports
' Schedules automatic email reports
' ---------------------------------------------------------------
Public Function ScheduleReports() As Boolean
    Dim ReportNo As EnumReportNo
    Dim RstDueReports As Recordset
    Dim StrReportHTML As String
    Dim RstReportData As Recordset
    
    Const StrPROCEDURE As String = "ScheduleReports()"

    On Error GoTo ErrorHandler

    Set RstDueReports = SQLQuery("SELECT * FROM TblReports WHERE DueDate <= NOW()")

    With RstDueReports
        If .RecordCount > 0 Then
            
            ReportNo = !ReportNo
            
            Select Case ReportNo
                Case Is = EnumCFSStockCountReport
                    
                    Debug.Print "Run Report " & ReportNo
                    
                    ' Send CFS Report
                    Set RstReportData = CFS_emailQuery
                    If RstReportData Is Nothing Then Err.Raise HANDLED_ERROR
                    
                    StrReportHTML = CFSEmailReportGen(RstReportData)
                    If StrReportHTML = "Error" Then Err.Raise HANDLED_ERROR, , "No Report Text"

                    If Not SendEmailReports("CFS Stock Report", StrReportHTML, EnumCFSStockCountReport) Then Err.Raise HANDLED_ERROR

                    'Reset due date
                    .Edit
                    !DueDate = !DueDate + !Frequency
                    .Update
            End Select
            
        End If
    End With
    
    Set RstDueReports = Nothing
    Set RstReportData = Nothing
    
    ScheduleReports = True

Exit Function

ErrorExit:

'    ***CleanUpCode***
    Set RstDueReports = Nothing
    Set RstReportData = Nothing
    
    ScheduleReports = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' CFS_emailQuery
' Gets data for CFS Stock List
' ---------------------------------------------------------------
Private Function CFS_emailQuery() As Recordset
    Dim StrSelect As String
    Dim StrFrom As String
    Dim StrWhere As String
    Dim StrOrderBy As String
    Dim RstCFSStock As Recordset
    
    Const StrPROCEDURE As String = "CFS_emailQuery()"

    On Error GoTo ErrorHandler
    
    StrSelect = "SELECT " _
                    & "TblAsset.Size1 & ' ' & TblAsset.Description AS [CFS Item], " _
                    & "TblAsset.QtyInStock, " _
                    & "TblAsset.MinAmount, " _
                    & "TblAsset.MaxAmount, " _
                    & "TblAsset.OrderLevel "
                                      
    StrFrom = "FROM " _
                    & "TblAsset "
                    
    StrWhere = "WHERE " _
                    & "TblAsset.Category3 = 'CFS Consumables'"
                    
    Set RstCFSStock = ModDatabase.SQLQuery(StrSelect & StrFrom & StrWhere & StrOrderBy)
  
    Set CFS_emailQuery = RstCFSStock
    
    Set RstCFSStock = Nothing
    
Exit Function

ErrorExit:

    Set RstCFSStock = Nothing
    Set CFS_emailQuery = Nothing

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' CFSEmailReportGen
' Generates the automatic email reports
' ---------------------------------------------------------------
Private Function CFSEmailReportGen(RstReportData As Recordset) As String
    Dim StrHead As String
    Dim StrBody As String
    Dim StrReport As String
    
    Const StrPROCEDURE As String = "CFSEmailReportGen()"

    On Error GoTo ErrorHandler

    StrHead = "<!DOCTYPE html><html><head><style>" _
                & "body    {" _
                    & "background-color:RGB(248,255,244);} " _
                & "#Line       {" _
                    & "color:RGB(28,71,84);" _
                    & "font-family:'Tahoma';" _
                    & "font-size:12px;}" _
                & "#Stock th{" _
                    & "padding-top: 4px;" _
                    & "padding-bottom: 4px;" _
                    & "text-align: center;" _
                    & "background-color: RGB(32,163,158);" _
                    & "color: RGB(248,255,244);" _
                    & "font-size:14px;}" _
                & "#Stock   {" _
                    & "width: 90%;" _
                    & "font-family: Tahoma; " _
                    & "border-collapse: collapse; }" _
                & "#Stock td      { " _
                    & "border: 1px solid #ddd;" _
                    & "padding: 8px;" _
                    & "color:RGB(28,71,84);" _
                    & "font-size:14px;} "

    StrHead = StrHead _
                & "#Stock tr:nth -Child(Even){" _
                    & "background-color: #f2f2f2;}" _
                & "#Stock tr:hover{" _
                    & "background-color: #ddd};" _
                & "#Header th{" _
                    & "font-size:22px;" _
                    & "padding: 10px;" _
                    & "text-align:center;" _
                    & "color:RGB(28,71,84);" _
                    & "font-family:'Calibri';" _
                    & "background-color:RGB(255,166,52);}" _
                & "#Header{" _
                    & "width: 50%;}" _
          & "</style></head>"

  StrBody = "<body>" _
                & "<table id='Header'>" _
                    & "<tr>" _
                        & "<th>Stores Ordering System</th>" _
                    & "</tr>" _
                & "</table>" _
                & "<p id = 'Line'>The current CFS stock levels as of  " & Format(Now, "dd mmm yy") & " are:</p>" _
                & "<table id = 'Stock'>" _
                    & "<tr>" _
                        & "<th>CFS Item</th>" _
                        & "<th>Quantity</th>" _
                    & "</tr>"
                    
                    Do While Not RstReportData.EOF
                        StrBody = StrBody _
                            & "<tr>" _
                                & "<td>" & Trim(RstReportData![CFS Item]) & "</td>" _
                                & "<td style='text-align:Center'>" & Trim(RstReportData!QtyInStock) & "</td>" _
                            & "</tr>"
                        RstReportData.MoveNext
                        Loop
                    StrBody = StrBody _
              & "</table>" _
              & "<p id = 'Line'>This is an automated system message from Ops Support </p>" _
        & "</body></html>"
    
    StrReport = StrHead & StrBody
        
    CFSEmailReportGen = StrReport

Exit Function

ErrorExit:

'    ***CleanUpCode***
    CFSEmailReportGen = "Error"

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' ReturnReportList
' Returns a list of reports or alerts in the system
' ---------------------------------------------------------------
Public Function ReturnReportList(EmailType As Integer) As Recordset
    Dim RstReports As Recordset
    
    Const StrPROCEDURE As String = "ReturnReportList()"

    On Error GoTo ErrorHandler

    Set RstReports = SQLQuery("SELECT * FROM TblReports WHERE ReportType = " & EmailType)

    Set ReturnReportList = RstReports

Exit Function

ErrorExit:

'    ***CleanUpCode***
    Set ReturnReportList = Nothing

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' GetReportAddresses
' Returns the To and CC addresses for a given report
' ---------------------------------------------------------------
Public Function GetReportAddresses(ReportNo As EnumReportNo) As Recordset
    Dim RstAddresses As Recordset
    
    Const StrPROCEDURE As String = "GetReportAddresses()"

    On Error GoTo ErrorHandler

    Set RstAddresses = SQLQuery("SELECT " _
                                  & "TblPerson.UserName, " _
                                  & "TblPerson.CrewNo, " _
                                  & "TblRptsAlerts.ToCC " _
                                & "From " _
                                  & "(TblReports " _
                                  & "RIGHT JOIN TblRptsAlerts ON TblRptsAlerts.ReportNo = TblReports.ReportNo) " _
                                  & "LEFT JOIN TblPerson ON TblRptsAlerts.CrewNo = TblPerson.CrewNo " _
                                & "WHERE " _
                                  & "TblRptsAlerts.ReportNo = " & ReportNo)

    Set GetReportAddresses = RstAddresses

Exit Function

ErrorExit:

'    ***CleanUpCode***
    Set GetReportAddresses = Nothing

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' EmailReportsAddAddress
' Adds a CC or To addressee to an email report
' ---------------------------------------------------------------
Public Sub EmailReportsAddAddress(CrewNo As String, ReportNo As Integer, ToCC As String)
    Dim SQL As String
    
    On Error Resume Next
    
    SQL = "INSERT INTO TblRptsAlerts VALUES ('" & CrewNo & "'," & ReportNo & ",'" & ToCC & "')"
    DB.Execute (SQL)
    
    Debug.Print SQL
End Sub

' ===============================================================
' EmailReportsRemoveAddress
' Adds a CC or To addressee to an email report
' ---------------------------------------------------------------
Public Sub EmailReportsRemoveAddress(CrewNo As String, ReportNo As Integer, ToCC As String)
    Dim SQL As String
    
'    On Error Resume Next
    
    SQL = "DELETE FROM TblRptsAlerts WHERE CrewNo = '" & CrewNo _
                & "' AND ReportNo = " & ReportNo _
                & " AND ToCC = '" & ToCC & "'"
    
    DB.Execute (SQL)

    Debug.Print SQL
End Sub

' ===============================================================
' SendEmailReports
' Sends the Email Reports to the addresses that have been set up
' ---------------------------------------------------------------
Public Function SendEmailReports(ReportSubject As String, ReportBody As String, ReportNo As EnumReportNo) As Boolean
    Dim TestFlag As String
    Dim RstToCC As Recordset
    Dim MailRecipients As Outlook.Recipients
    Dim MailRecipient As Outlook.Recipient
    Dim i As Integer
    
    Const StrPROCEDURE As String = "SendEmailReports()"

    On Error GoTo ErrorHandler

    Set MailSystem = New ClsMailSystem
    Set RstToCC = GetReportAddresses(ReportNo)
    
    If Not ModLibrary.OutlookRunning Then
        Shell "Outlook.exe"
    End If
    
    If TEST_MODE Then
        TestFlag = TEST_PREFIX
    Else
        TestFlag = ""
    End If

    With RstToCC
        
        Set MailRecipients = MailSystem.MailItem.Recipients
        
        Do While Not .EOF
            Set MailRecipient = MailRecipients.Add(!UserName)
            MailRecipient.Resolve
            If MailRecipient.Resolved Then
                If !ToCC = "To" Then MailRecipient.Type = olTo
                If !ToCC = "CC" Then MailRecipient.Type = olCC
            Else
                For i = 1 To MailRecipients.Count
                    If MailRecipients(i).Name = !UserName Then MailRecipients.Remove i
                Next
            End If
            RstToCC.MoveNext
        Loop
    End With
    With MailSystem.MailItem
        .Subject = TestFlag & ReportSubject
        .HTMLBody = TestFlag & ReportBody
        If SEND_EMAILS Then .Send Else .Display
    End With
    
    Set MailSystem = Nothing
    Set MailRecipient = Nothing
    Set MailRecipients = Nothing

    SendEmailReports = True

Exit Function

ErrorExit:

'    ***CleanUpCode***
    SendEmailReports = False
    Set MailRecipient = Nothing
    Set MailRecipients = Nothing

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function
