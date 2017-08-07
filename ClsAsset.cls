VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsAsset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsAsset
' v0,0 - Initial Version
' v0,1 - Change location to string
' v0,2 - Trimmed values from Database
' v0,3 - Temporary Supplier strings instead of classes
' v0,4 - Removed deleted property and changes for Asset Import
' v0,5 - Temporary disable Asset DBsave 
'---------------------------------------------------------------
' Date - 07 Aug 17
'===============================================================
' Methods
'---------------------------------------------------------------
' DBGet - Gets class from Database
' DBSave - Saves class to Database
' DBDelete(Optional FullDelete As Boolean) - Marks record as deleted or fully deletes
' ReturnAssetStatus - Returns string value from enumAssetStatus
'===============================================================

Option Explicit
Private pAssetNo As Integer
Private pBrand As String
Private pDescription As String
Private pCost As Single
Private pQtyInStock As Integer
Private pCategory1 As String
Private pCategory2 As String
Private pCategory3 As String
Private pSizing As String
Private pSize1 As String
Private pSize2 As String
Private pPurchaseUnit As String
Private pMinAmount As Integer
Private pMaxAmount As Integer
Private pOrderLevel As Integer
Private pLeadTime As Integer
Private pStatus As EnumAssetStatus
Private pSupplier1 As String
Private pSupplier2 As String
Private pKeywords As String
Private pAllowedOrderReasons As String
Private pAllocationType As EnumAllocationType
Private pLocation As String
Private pAdditInfo As String
Private pNoOrderMessage As String

'---------------------------------------------------------------
Public Property Get AssetNo() As Integer
    AssetNo = pAssetNo
End Property

Public Property Let AssetNo(ByVal VnewValue As Integer)
    pAssetNo = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get Brand() As String
    Brand = pBrand
End Property

Public Property Let Brand(ByVal VnewValue As String)
    pBrand = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get Description() As String
    Description = pDescription
End Property

Public Property Let Description(ByVal VnewValue As String)
    pDescription = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get cost() As Single
    cost = pCost
End Property

Public Property Let cost(ByVal VnewValue As Single)
    pCost = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get QtyInStock() As Integer
    QtyInStock = pQtyInStock
End Property

Public Property Let QtyInStock(ByVal VnewValue As Integer)
    pQtyInStock = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get Category1() As String
    Category1 = pCategory1
End Property

Public Property Let Category1(ByVal VnewValue As String)
    pCategory1 = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get Category2() As String
    Category2 = pCategory2
End Property

Public Property Let Category2(ByVal VnewValue As String)
    pCategory2 = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get Category3() As String
    Category3 = pCategory3
End Property

Public Property Let Category3(ByVal VnewValue As String)
    pCategory3 = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get Sizing() As String
    Sizing = pSizing
End Property

Public Property Let Sizing(ByVal VnewValue As String)
    pSizing = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get Size1() As String
    Size1 = pSize1
End Property

Public Property Let Size1(ByVal VnewValue As String)
    pSize1 = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get Size2() As String
    Size2 = pSize2
End Property

Public Property Let Size2(ByVal VnewValue As String)
    pSize2 = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get PurchaseUnit() As String
    PurchaseUnit = pPurchaseUnit
End Property

Public Property Let PurchaseUnit(ByVal VnewValue As String)
    pPurchaseUnit = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get MinAmount() As Integer
    MinAmount = pMinAmount
End Property

Public Property Let MinAmount(ByVal VnewValue As Integer)
    pMinAmount = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get MaxAmount() As Integer
    MaxAmount = pMaxAmount
End Property

Public Property Let MaxAmount(ByVal VnewValue As Integer)
    pMaxAmount = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get OrderLevel() As Integer
    OrderLevel = pOrderLevel
End Property

Public Property Let OrderLevel(ByVal VnewValue As Integer)
    pOrderLevel = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get LeadTime() As Integer
    LeadTime = pLeadTime
End Property

Public Property Let LeadTime(ByVal VnewValue As Integer)
    pLeadTime = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get Status() As EnumAssetStatus
    
    Select Case pQtyInStock
        Case Is > pOrderLevel
            pStatus = Ok
        Case Is > pMinAmount
            pStatus = ReOrder
        Case Is > 0
            pStatus = LowLevel
        Case Else
            pStatus = NoStock
    End Select
    
    Status = pStatus
End Property

Public Property Let Status(ByVal VnewValue As EnumAssetStatus)
    pStatus = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get Supplier1() As String
    Supplier1 = pSupplier1
End Property

Public Property Let Supplier1(ByVal VnewValue As String)
    pSupplier1 = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get Supplier2() As String
    Supplier2 = pSupplier2
End Property

Public Property Let Supplier2(ByVal VnewValue As String)
    pSupplier2 = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get Keywords() As String
    Keywords = pKeywords
End Property

Public Property Let Keywords(ByVal VnewValue As String)
    pKeywords = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get AllowedOrderReasons() As String
    AllowedOrderReasons = pAllowedOrderReasons
End Property

Public Property Let AllowedOrderReasons(ByVal VnewValue As String)
    pAllowedOrderReasons = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get AllocationType() As EnumAllocationType
    AllocationType = pAllocationType
End Property

Public Property Let AllocationType(ByVal VnewValue As EnumAllocationType)
    pAllocationType = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get Location() As String
    Location = pLocation
End Property

Public Property Let Location(ByVal VnewValue As String)
    pLocation = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get AdditInfo() As String
    AdditInfo = pAdditInfo
End Property

Public Property Let AdditInfo(ByVal VnewValue As String)
    pAdditInfo = VnewValue
End Property

'---------------------------------------------------------------
Public Property Get NoOrderMessage() As String
    NoOrderMessage = pNoOrderMessage
End Property

Public Property Let NoOrderMessage(ByVal VnewValue As String)
    pNoOrderMessage = VnewValue
End Property

'---------------------------------------------------------------
' ===============================================================
' Method DBGet
' Gets class from Database
'---------------------------------------------------------------
Public Sub DBGet(AssetNo As Integer)
    Dim RstAsset As Recordset

    Set RstAsset = ModDatabase.SQLQuery("SELECT * FROM TblAsset WHERE AssetNo = " & AssetNo)
    With RstAsset
        If .RecordCount > 0 Then
            If Not IsNull(!AssetNo) Then pAssetNo = Trim(!AssetNo)
            If Not IsNull(!Brand) Then pBrand = Trim(!Brand)
            If Not IsNull(!Description) Then pDescription = Trim(!Description)
            If Not IsNull(!cost) Then pCost = Trim(!cost)
            If Not IsNull(!QtyInStock) Then QtyInStock = Trim(!QtyInStock)
            If Not IsNull(!Category1) Then pCategory1 = Trim(!Category1)
            If Not IsNull(!Category2) Then pCategory2 = Trim(!Category2)
            If Not IsNull(!Category3) Then pCategory3 = Trim(!Category3)
            If Not IsNull(!Size1) Then pSize1 = Trim(!Size1)
            If Not IsNull(!Size2) Then pSize2 = Trim(!Size2)
            If Not IsNull(!PurchaseUnit) Then pPurchaseUnit = Trim(!PurchaseUnit)
            If Not IsNull(!MinAmount) Then pMinAmount = Trim(!MinAmount)
            If Not IsNull(!MaxAmount) Then pMaxAmount = Trim(!MaxAmount)
            If Not IsNull(!OrderLevel) Then pOrderLevel = Trim(!OrderLevel)
            If Not IsNull(!LeadTime) Then pLeadTime = Trim(!LeadTime)
            If Not IsNull(!Status) Then pStatus = Trim(!Status)
            If Not IsNull(!Supplier1) Then pSupplier1 = Trim(!Supplier1)
            If Not IsNull(!Supplier2) Then pSupplier2 = Trim(!Supplier2)
            If Not IsNull(!Keywords) Then pKeywords = Trim(!Keywords)
            If Not IsNull(!AllowedOrderReasons) Then pAllowedOrderReasons = Trim(!AllowedOrderReasons)
            If Not IsNull(!AllocationType) Then pAllocationType = Trim(!AllocationType)
            If Not IsNull(!Location) Then pLocation = Trim(!Location)
            If Not IsNull(!AdditInfo) Then pAdditInfo = Trim(!AdditInfo)
            If Not IsNull(!NoOrderMessage) Then pNoOrderMessage = Trim(!NoOrderMessage) Else pNoOrderMessage = ""

'            pSupplier1.DBGet
'            pSupplier2.DBGet
        End If
    End With
    Set RstAsset = Nothing
End Sub

' ===============================================================
' Method DBSave
' Saves class to Database
'---------------------------------------------------------------
Public Sub DBSave(Optional OWAssetNo As Integer)
    Dim RstAsset As Recordset
    Dim RstMaxNo As Recordset
    Dim LastNo As Integer

    Set RstAsset = ModDatabase.SQLQuery("SELECT * FROM TblAsset WHERE AssetNo = " & pAssetNo)
    Set RstMaxNo = ModDatabase.SQLQuery("SELECT MAX(AssetNo) FROM TblAsset ")

    If RstMaxNo.Fields(0).Value <> 0 Then
        LastNo = RstMaxNo.Fields(0).Value
    Else
        LastNo = 0
    End If

    With RstAsset
        If .RecordCount = 0 Then
            .addnew
            
            'if an asset no has been passed, then overwrite that asset in the database, otherwise
            'create new record
            If OWAssetNo <> 0 Then
                pAssetNo = OWAssetNo
            Else
            pAssetNo = LastNo + 1
            End If
            
            !AssetNo = pAssetNo
        Else
            .Edit
        End If
        !Brand = pBrand
        !Description = pDescription
        !cost = pCost
        !QtyInStock = pQtyInStock
        !Category1 = pCategory1
        !Category2 = pCategory2
        !Category3 = pCategory3
        !Size1 = pSize1
        !Size2 = pSize2
        !PurchaseUnit = pPurchaseUnit
        !MinAmount = pMinAmount
        !MaxAmount = pMaxAmount
        !OrderLevel = pOrderLevel
        !LeadTime = pLeadTime
        !Status = pStatus
        !Supplier1 = pSupplier1
        !Supplier2 = pSupplier2
        !Keywords = pKeywords
        !AllowedOrderReasons = pAllowedOrderReasons
        !AllocationType = pAllocationType
        !Location = pLocation
        !AdditInfo = pAdditInfo
        !NoOrderMessage = pNoOrderMessage
'        .Update

    End With
    Set RstAsset = Nothing
    Set RstMaxNo = Nothing
End Sub

' ===============================================================
' Method DBDelete(Optional FullDelete As Boolean)
' Fully deletes record from DB
'---------------------------------------------------------------
Public Sub DBDelete()
    Dim RstAsset As Recordset
    Dim i As Integer

    Set RstAsset = ModDatabase.SQLQuery("SELECT * FROM TblAsset WHERE AssetNo = " & pAssetNo)
    With RstAsset
        For i = .RecordCount To 1 Step -1
                .Delete
                .MoveNext
        Next
    End With

    Set RstAsset = Nothing
End Sub

' ===============================================================
' ReturnAssetStatus
' Returns string value from enumAssetStatus
' ---------------------------------------------------------------
Public Function ReturnAssetStatus() As String
    Const StrPROCEDURE As String = "ReturnAssetStatus()"
    
    Dim RstAssetStatus As Recordset
    
    Set RstAssetStatus = ModDatabase.SQLQuery("SELECT AssetStatus FROM TblAssetStatus WHERE AssetStatusNo = " & Status)
    
    With RstAssetStatus
        If .RecordCount > 0 Then ReturnAssetStatus = !AssetStatus
    
    End With
    
    Set RstAssetStatus = Nothing

End Function


