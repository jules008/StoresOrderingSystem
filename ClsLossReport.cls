VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsLossReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsLossReport
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 22 Feb 17
'===============================================================
' Methods
'---------------------------------------------------------------
' DBGet - Gets class from Database
' DBSave - Saves class to Database
' DBDelete(Optional FullDelete As Boolean) - Marks record as deleted or fully deletes
'===============================================================

Option Explicit
Private pDeleted As Date
Private pLossReportNo As Integer
Private pReportDate As Date
Private pTheft As Boolean
Private pActionsTaken As String
Private pAdditInfo As String
Private pReportingOfficer As ClsPerson
Private pHQAction As String
Private pOpsSupportAction As String
Private pIncNo As String
Private pStatus As EnumLossRepStatus
Private pAuthorised As Boolean
Private pRejected As Boolean
Private pParent As Long

'---------------------------------------------------------------
Public Property Get Deleted() As Date
    Deleted = pDeleted
End Property

Public Property Let Deleted(ByVal VnewValue As Date)
    pDeleted = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get LossReportNo() As Integer
    LossReportNo = pLossReportNo
End Property

Public Property Let LossReportNo(ByVal VnewValue As Integer)
    pLossReportNo = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get ReportDate() As Date
    ReportDate = pReportDate
End Property

Public Property Let ReportDate(ByVal VnewValue As Date)
    pReportDate = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get Theft() As Boolean
    Theft = pTheft
End Property

Public Property Let Theft(ByVal VnewValue As Boolean)
    pTheft = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get ActionsTaken() As String
    ActionsTaken = pActionsTaken
End Property

Public Property Let ActionsTaken(ByVal VnewValue As String)
    pActionsTaken = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get AdditInfo() As String
    AdditInfo = pAdditInfo
End Property

Public Property Let AdditInfo(ByVal VnewValue As String)
    pAdditInfo = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get ReportingOfficer() As ClsPerson
    Set ReportingOfficer = pReportingOfficer
End Property

Public Property Let ReportingOfficer(ByVal VnewValue As ClsPerson)
    Set pReportingOfficer = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get HQAction() As String
    HQAction = pHQAction
End Property

Public Property Let HQAction(ByVal VnewValue As String)
    pHQAction = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get OpsSupportAction() As String
    OpsSupportAction = pOpsSupportAction
End Property

Public Property Let OpsSupportAction(ByVal VnewValue As String)
    pOpsSupportAction = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get IncNo() As String
    IncNo = pIncNo
End Property

Public Property Let IncNo(ByVal VnewValue As String)
    pIncNo = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get Status() As EnumLossRepStatus
    Status = pStatus
End Property

Public Property Let Status(ByVal VnewValue As EnumLossRepStatus)
    pStatus = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get Authorised() As Boolean
    Authorised = pAuthorised
End Property

Public Property Let Authorised(ByVal VnewValue As Boolean)
    pAuthorised = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get Rejected() As Boolean
    Rejected = pRejected
End Property

Public Property Let Rejected(ByVal VnewValue As Boolean)
    pRejected = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get Parent() As ClsLineItem
    If pParent <> 0 Then
        Set Parent = GetParentFromPtr(pParent)
    End If
End Property

Friend Function SetParent(ByVal Ptr As Long) As Boolean
    pParent = Ptr
End Function

' ===============================================================
' Method DBGet
' Gets class from Database
'---------------------------------------------------------------
Public Sub DBGet(LossReportNo As Integer)
    Dim RstLossReport As Recordset

    Set RstLossReport = ModDatabase.SQLQuery("SELECT * FROM TblLossReport WHERE LossReportNo = " & LossReportNo & " AND Deleted IS NULL")
    With RstLossReport
        If .RecordCount > 0 Then
            If Not IsNull(!Deleted) Then pDeleted = !Deleted
            If Not IsNull(!LossReportNo) Then pLossReportNo = !LossReportNo
            If Not IsNull(!ReportDate) Then pReportDate = !ReportDate
            If Not IsNull(!Theft) Then pTheft = !Theft
            If Not IsNull(!ActionsTaken) Then pActionsTaken = !ActionsTaken
            If Not IsNull(!AdditInfo) Then pAdditInfo = !AdditInfo
            If Not IsNull(!HQAction) Then pHQAction = !HQAction
            If Not IsNull(!OpsSupportAction) Then pOpsSupportAction = !OpsSupportAction
            If Not IsNull(!IncNo) Then pIncNo = !IncNo
            If Not IsNull(!Status) Then pStatus = !Status
            If Not IsNull(!Authorised) Then pAuthorised = !Authorised
            If Not IsNull(!Rejected) Then pRejected = !Rejected

            If Not IsNull(!ReportingOfficerID) Then pReportingOfficer.DBGet !ReportingOfficerID
        End If
    End With
Set RstLossReport = Nothing
End Sub

' ===============================================================
' Method DBSave
' Saves class to Database
'---------------------------------------------------------------
Public Sub DBSave()
    Dim RstLossReport As Recordset
    Dim RstMaxNo As Recordset
    Dim LastNo As Integer

    Set RstLossReport = ModDatabase.SQLQuery("SELECT * FROM TblLossReport WHERE LossReportNo = " & pLossReportNo & " AND Deleted IS NULL")
    Set RstMaxNo = ModDatabase.SQLQuery("SELECT MAX(LossReportNo) FROM TblLossReport ")

    If RstMaxNo.Fields(0).Value <> 0 Then
        LastNo = RstMaxNo.Fields(0).Value
    Else
        LastNo = 0
    End If

    With RstLossReport
        If .RecordCount = 0 Then
            .AddNew
            pLossReportNo = LastNo + 1
        Else
            .Edit
        End If
        !LossReportNo = pLossReportNo
        !ReportDate = pReportDate
        !Theft = pTheft
        !ActionsTaken = pActionsTaken
        !AdditInfo = pAdditInfo
        !ReportingOfficerID = pReportingOfficer.CrewNo
        !HQAction = pHQAction
        !OpsSupportAction = pOpsSupportAction
        !IncNo = pIncNo
        !Status = pStatus
        !Authorised = pAuthorised
        !Rejected = pRejected
        .Update

    End With
    Set RstLossReport = Nothing
    Set RstMaxNo = Nothing
End Sub

' ===============================================================
' Method DBDelete(Optional FullDelete As Boolean)
' Marks record as deleted or fully deletes
'---------------------------------------------------------------
Public Sub DBDelete(Optional FullDelete As Boolean)
    Dim RstLossReport As Recordset
    Dim i As Integer

    Set RstLossReport = ModDatabase.SQLQuery("SELECT * FROM TblLossReport WHERE LossReportNo = " & pLossReportNo & " AND Deleted IS NULL")
    With RstLossReport
        For i = .RecordCount To 1 Step -1
            If FullDelete Then
                .Delete
                .MoveNext
            Else
                .Edit
                !Deleted = Now
                .Update
            End If
        Next
    End With

    Set RstLossReport = Nothing
End Sub

' ===============================================================
' Method GetParentFromPtr
' Private routine to copy memory address of parent class
' ---------------------------------------------------------------
Private Function GetParentFromPtr(ByVal Ptr As Long) As ClsLineItem
    Dim tmp As ClsLineItem

    CopyMemory tmp, Ptr, 4
    Set GetParentFromPtr = tmp
    CopyMemory tmp, 0&, 4
End Function

' ===============================================================
Private Sub Class_Initialize()
    Set pReportingOfficer = New ClsPerson

End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pReportingOfficer = Nothing

End Sub

'---------------------------------------------------------------
