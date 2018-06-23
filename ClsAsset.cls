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
' v0,6 - Re-enabled Asset DBSave
' v0,7 - Added HideFromView property
' v0,8 - Added Field Report
'---------------------------------------------------------------
' Date - 23 Jun 18
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
Private pHideFromView As Boolean
Private pReport As Boolean

'---------------------------------------------------------------
Public Property Get AssetNo() As Integer
    AssetNo = pAssetNo
End Property

Public Property Let AssetNo(ByVal vNewValue As Integer)
    pAssetNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Brand() As String
    Brand = pBrand
End Property

Public Property Let Brand(ByVal vNewValue As String)
    pBrand = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Description() As String
    Description = pDescription
End Property

Public Property Let Description(ByVal vNewValue As String)
    pDescription = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get cost() As Single
    cost = pCost
End Property

Public Property Let cost(ByVal vNewValue As Single)
    pCost = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get QtyInStock() As Integer
    QtyInStock = pQtyInStock
End Property

Public Property Let QtyInStock(ByVal vNewValue As Integer)
    pQtyInStock = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Category1() As String
    Category1 = pCategory1
End Property

Public Property Let Category1(ByVal vNewValue As String)
    pCategory1 = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Category2() As String
    Category2 = pCategory2
End Property

Public Property Let Category2(ByVal vNewValue As String)
    pCategory2 = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Category3() As String
    Category3 = pCategory3
End Property

Public Property Let Category3(ByVal vNewValue As String)
    pCategory3 = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Sizing() As String
    Sizing = pSizing
End Property

Public Property Let Sizing(ByVal vNewValue As String)
    pSizing = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Size1() As String
    Size1 = pSize1
End Property

Public Property Let Size1(ByVal vNewValue As String)
    pSize1 = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Size2() As String
    Size2 = pSize2
End Property

Public Property Let Size2(ByVal vNewValue As String)
    pSize2 = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get PurchaseUnit() As String
    PurchaseUnit = pPurchaseUnit
End Property

Public Property Let PurchaseUnit(ByVal vNewValue As String)
    pPurchaseUnit = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get MinAmount() As Integer
    MinAmount = pMinAmount
End Property

Public Property Let MinAmount(ByVal vNewValue As Integer)
    pMinAmount = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get MaxAmount() As Integer
    MaxAmount = pMaxAmount
End Property

Public Property Let MaxAmount(ByVal vNewValue As Integer)
    pMaxAmount = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get OrderLevel() As Integer
    OrderLevel = pOrderLevel
End Property

Public Property Let OrderLevel(ByVal vNewValue As Integer)
    pOrderLevel = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get LeadTime() As Integer
    LeadTime = pLeadTime
End Property

Public Property Let LeadTime(ByVal vNewValue As Integer)
    pLeadTime = vNewValue
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

Public Property Let Status(ByVal vNewValue As EnumAssetStatus)
    pStatus = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Supplier1() As String
    Supplier1 = pSupplier1
End Property

Public Property Let Supplier1(ByVal vNewValue As String)
    pSupplier1 = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Supplier2() As String
    Supplier2 = pSupplier2
End Property

Public Property Let Supplier2(ByVal vNewValue As String)
    pSupplier2 = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Keywords() As String
    Keywords = pKeywords
End Property

Public Property Let Keywords(ByVal vNewValue As String)
    pKeywords = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get AllowedOrderReasons() As String
    AllowedOrderReasons = pAllowedOrderReasons
End Property

Public Property Let AllowedOrderReasons(ByVal vNewValue As String)
    pAllowedOrderReasons = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get AllocationType() As EnumAllocationType
    AllocationType = pAllocationType
End Property

Public Property Let AllocationType(ByVal vNewValue As EnumAllocationType)
    pAllocationType = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Location() As String
    Location = pLocation
End Property

Public Property Let Location(ByVal vNewValue As String)
    pLocation = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get AdditInfo() As String
    AdditInfo = pAdditInfo
End Property

Public Property Let AdditInfo(ByVal vNewValue As String)
    pAdditInfo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get NoOrderMessage() As String
    NoOrderMessage = pNoOrderMessage
End Property

Public Property Let NoOrderMessage(ByVal vNewValue As String)
    pNoOrderMessage = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get HideFromView() As Boolean
    HideFromView = pHideFromView
End Property

Public Property Let HideFromView(ByVal vNewValue As Boolean)
    pHideFromView = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Report() As Boolean
    Report = pReport
End Property

Public Property Let Report(ByVal vNewValue As Boolean)
    pReport = vNewValue
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
            If Not IsNull(!HideFromView) Then pHideFromView = !HideFromView
            If Not IsNull(!Report) Then pReport = !Report
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
            .AddNew
            
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
        !HideFromView = pHideFromView
        !Report = pReport
        .Update

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


