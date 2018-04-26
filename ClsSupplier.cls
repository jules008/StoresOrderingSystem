VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsSupplier
' v0,0 - Initial Version
' v0,1 - DBGet from ID or Name
' v0,2 - check for "" before writing to DB
'---------------------------------------------------------------
' Date - 19 Jul 17
'===============================================================
' Methods
'---------------------------------------------------------------
' DBGet - Gets class from Database
' DBSave - Saves class to Database
' DBDelete(Optional FullDelete As Boolean) - Marks record as deleted or fully deletes
'===============================================================

Option Explicit
Private pDeleted As Date
Private pSupplierID As Integer
Private pSupplierName As String
Private pAgressoNo As String
Private pPCard As Boolean
Private pAddress1 As String
Private pAddress2 As String
Private pTownCity As String
Private pCounty As String
Private pPostcode As String
Private pTelephone As String
Private pWebsite As String
Private pEmail As String
Private pContactName As String
Private pCategory As String
Private pItemsSupplied As String
Private pAccountNo As String
Private pParent As Long

'---------------------------------------------------------------
Public Property Get Deleted() As Date
    Deleted = pDeleted
End Property

Public Property Let Deleted(ByVal vNewValue As Date)
    pDeleted = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get SupplierID() As Integer
    SupplierID = pSupplierID
End Property

Public Property Let SupplierID(ByVal vNewValue As Integer)
    pSupplierID = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get SupplierName() As String
    SupplierName = pSupplierName
End Property

Public Property Let SupplierName(ByVal vNewValue As String)
    pSupplierName = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get AgressoNo() As String
    AgressoNo = pAgressoNo
End Property

Public Property Let AgressoNo(ByVal vNewValue As String)
    pAgressoNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get PCard() As Boolean
    PCard = pPCard
End Property

Public Property Let PCard(ByVal vNewValue As Boolean)
    pPCard = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Address1() As String
    Address1 = pAddress1
End Property

Public Property Let Address1(ByVal vNewValue As String)
    pAddress1 = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Address2() As String
    Address2 = pAddress2
End Property

Public Property Let Address2(ByVal vNewValue As String)
    pAddress2 = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get TownCity() As String
    TownCity = pTownCity
End Property

Public Property Let TownCity(ByVal vNewValue As String)
    pTownCity = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get County() As String
    County = pCounty
End Property

Public Property Let County(ByVal vNewValue As String)
    pCounty = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Postcode() As String
    Postcode = pPostcode
End Property

Public Property Let Postcode(ByVal vNewValue As String)
    pPostcode = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Telephone() As String
    Telephone = pTelephone
End Property

Public Property Let Telephone(ByVal vNewValue As String)
    pTelephone = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Website() As String
    Website = pWebsite
End Property

Public Property Let Website(ByVal vNewValue As String)
    pWebsite = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Email() As String
    Email = pEmail
End Property

Public Property Let Email(ByVal vNewValue As String)
    pEmail = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get ContactName() As String
    ContactName = pContactName
End Property

Public Property Let ContactName(ByVal vNewValue As String)
    pContactName = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Category() As String
    Category = pCategory
End Property

Public Property Let Category(ByVal vNewValue As String)
    pCategory = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get AccountNo() As String
    AccountNo = pAccountNo
End Property

Public Property Let AccountNo(ByVal vNewValue As String)
    pAccountNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get ItemsSupplied() As String
    ItemsSupplied = pItemsSupplied
End Property

Public Property Let ItemsSupplied(ByVal vNewValue As String)
    pItemsSupplied = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Parent() As ClsAsset
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
Public Sub DBGet(IDorName As Variant)
    Dim RstSupplier As Recordset
    Dim SQL As String
    
    If IsNumeric(IDorName) Then
        SQL = "SELECT * FROM TblSupplier WHERE SupplierID = " & IDorName
    Else
        SQL = "SELECT * FROM TblSupplier WHERE SupplierName = " & SupplierName
    End If
    
    Set RstSupplier = ModDatabase.SQLQuery(SQL)
    
    With RstSupplier
        If .RecordCount > 0 Then
            If Not IsNull(!SupplierID) Then pSupplierID = !SupplierID
            If Not IsNull(!SupplierName) Then pSupplierName = !SupplierName
            If Not IsNull(!AgressoNo) Then pAgressoNo = !AgressoNo
            If Not IsNull(!Address1) Then pAddress1 = !Address1
            If Not IsNull(!Address2) Then pAddress2 = !Address2
            If Not IsNull(!TownCity) Then pTownCity = !TownCity
            If Not IsNull(!County) Then pCounty = !County
            If Not IsNull(!Postcode) Then pPostcode = !Postcode
            If Not IsNull(!Telephone) Then pTelephone = !Telephone
            If Not IsNull(!Website) Then pWebsite = !Website
            If Not IsNull(!Email) Then pEmail = !Email
            If Not IsNull(!ContactName) Then pContactName = !ContactName
            If Not IsNull(!Category) Then pCategory = !Category
            If Not IsNull(!ItemsSupplied) Then pItemsSupplied = !ItemsSupplied
            If Not IsNull(!AccountNo) Then pAccountNo = !AccountNo
            pPCard = !PCard
        End If
    End With
Set RstSupplier = Nothing
End Sub

' ===============================================================
' Method DBSave
' Saves class to Database
'---------------------------------------------------------------
Public Sub DBSave()
    Dim RstSupplier As Recordset
    Dim RstMaxNo As Recordset
    Dim LastNo As Integer

    Set RstSupplier = ModDatabase.SQLQuery("SELECT * FROM TblSupplier WHERE SupplierID = " & pSupplierID)
    Set RstMaxNo = ModDatabase.SQLQuery("SELECT MAX(SupplierID) FROM TblSupplier ")

    If RstMaxNo.Fields(0).Value <> 0 Then
        LastNo = RstMaxNo.Fields(0).Value
    Else
        LastNo = 0
    End If

    With RstSupplier
        If .RecordCount = 0 Then
            .AddNew
            pSupplierID = LastNo + 1
        Else
            .Edit
        End If
        !SupplierID = pSupplierID
        If pAccountNo <> "" Then !AccountNo = pAccountNo
        If pAddress1 <> "" Then !Address1 = pAddress1
        If pAddress2 <> "" Then !Address2 = pAddress2
        If pAgressoNo <> "" Then !AgressoNo = pAgressoNo
        If pCategory <> "" Then !Category = pCategory
        If pContactName <> "" Then !ContactName = pContactName
        If pCounty <> "" Then !County = pCounty
        If pEmail <> "" Then !Email = pEmail
        If pItemsSupplied <> "" Then !ItemsSupplied = pItemsSupplied
        !PCard = pPCard
        If pPostcode <> "" Then !Postcode = pPostcode
        If pSupplierName <> "" Then !SupplierName = pSupplierName
        If pTelephone <> "" Then !Telephone = pTelephone
        If pTownCity <> "" Then !TownCity = pTownCity
        If pWebsite <> "" Then !Website = pWebsite

        .Update

    End With
    Set RstSupplier = Nothing
    Set RstMaxNo = Nothing
End Sub

' ===============================================================
' Method DBDelete(Optional FullDelete As Boolean)
' Marks record as deleted or fully deletes
'---------------------------------------------------------------
Public Sub DBDelete(Optional FullDelete As Boolean)
    Dim RstSupplier As Recordset
    Dim i As Integer

    Set RstSupplier = ModDatabase.SQLQuery("SELECT * FROM TblSupplier WHERE SupplierID = " & pSupplierID & " AND Deleted IS NULL")
    With RstSupplier
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

    Set RstSupplier = Nothing
End Sub

' ===============================================================
' Method GetParentFromPtr
' Private routine to copy memory address of parent class
' ---------------------------------------------------------------
Private Function GetParentFromPtr(ByVal Ptr As Long) As ClsAsset
    Dim tmp As ClsAsset

    CopyMemory tmp, Ptr, 4
    Set GetParentFromPtr = tmp
    CopyMemory tmp, 0&, 4
End Function
