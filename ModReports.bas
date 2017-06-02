Attribute VB_Name = "ModReports"
'===============================================================
' Module ModReports
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 02 Jun 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModReports"

' ===============================================================
' CreateReport1
' Creates Report 1
' ---------------------------------------------------------------
Private Function CreateReport1() As Boolean
    Dim ReportBook As Workbook
    Dim RstQuery As Recordset
    Dim RngQry As Range
    
    Const StrPROCEDURE As String = "CreateReport1()"

    On Error GoTo ErrorHandler

    Set RstQuery = ModReports.Report1Query
    
    If RstQuery Is Nothing Then Err.Raise HANDLED_ERROR
    
    Set ReportBook = Workbooks.Add

    With ReportBook
        Set RngQry = .Worksheets(1).Range("A1")
        RngQry.CopyFromRecordset RstQuery
    
    End With
    
    Set RngQry = Nothing
    Set RstQuery = Nothing
    CreateReport1 = True

Exit Function

ErrorExit:

    Set RngQry = Nothing
    Set RstQuery = Nothing
'    ***CleanUpCode***
    CreateReport1 = False

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
                                  & "LEFT JOIN TblPerson ON TblLineItem.ForPersonID = TblPerson.ID) " _
                                  & "LEFT JOIN TblStation ON TblLineItem.ForStationID = TblStation.StationID) " _
                                  & "LEFT JOIN TblVehicle ON TblLineItem.ForVehicleID = TblVehicle.VehNo) " _
                                  & "LEFT JOIN TblReqReason ON TblLineItem.ReqReason = TblReqReason.ID) " _
                                  & "LEFT JOIN TblStation TblStation1 ON TblVehicle.StationID = TblStation1.StationID " _
                                & "WHERE " _
                                  & "TblAsset.AssetNo IS NOT NULL " _
                                & "ORDER BY " _
                                  & "TblLineItem.OrderNo")



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