VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsAssets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsAssets
' v0,0 - Initial Version
' v0,1 - Improved size list counting to prevent empty size lists
' v0,2 - Added GetAllAssets function
' v0,3 - Removed GetAllAssets and added MaxAssetNo property
' v0,4 - Delete collection on terminate
' v0,5 - Changed get size lists
'---------------------------------------------------------------
' Date - 23 May 17
'===============================================================
' Methods
'---------------------------------------------------------------
' FindItem - Finds Item from index
' AddItem - Add item to Collection
' RemoveItem - Deletes item from collection
' ItemCount - Returns number of items in collection
' GetCollection - Gets all items in collection
' DeleteCollection - Deletes all items in collection
' GetCategoryLists - Gets items in category lists 1,2 or 3, returning them in an recordset
' GetItemList - Gets list of items from category lists, returning them in a recordset
' GetSizeLists - Returns size lists for a given asset no
' FindAssetNo - Finds Asset No from Asset Description
'===============================================================

Option Explicit
Private pAssets As Collection
Private pMaxAssetNo As Integer

'---------------------------------------------------------------

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pAssets.[_NewEnum]
End Function

'---------------------------------------------------------------
Public Property Get MaxAssetNo() As Integer
    MaxAssetNo = pMaxAssetNo
End Property

Public Property Let MaxAssetNo(ByVal vNewValue As Integer)
    pMaxAssetNo = vNewValue
End Property

' ===============================================================
' Method FindItem
' Finds Item from index
'---------------------------------------------------------------

Public Function FindItem(AssetNo As Variant) As ClsAsset
Attribute FindItem.VB_UserMemId = 0
    On Error Resume Next
    Set FindItem = pAssets.Item(AssetNo)
End Function

' ===============================================================
' Method AddItem
' Add item to Collection
'---------------------------------------------------------------

Public Sub AddItem(Asset As ClsAsset)
    pAssets.Add Asset, Key:=CStr(Asset.AssetNo)
    If Asset.AssetNo > MaxAssetNo Then MaxAssetNo = Asset.AssetNo
End Sub

' ===============================================================
' Method RemoveItem
' Deletes item from collection
'---------------------------------------------------------------

Public Sub RemoveItem(AssetNo As Variant)
    pAssets.Remove AssetNo
End Sub

' ===============================================================
' Method ItemCount
' Returns number of items in collection
'---------------------------------------------------------------

Public Function Count() As Integer
    Count = pAssets.Count
End Function

' ===============================================================
' Method GetCollection
' Gets all items in collection
'---------------------------------------------------------------

Public Sub GetCollection()
    Dim RstAsset As Recordset
    Dim Asset As ClsAsset
    Dim i As Integer

    Set RstAsset = ModDatabase.SQLQuery("TblAsset")
    With RstAsset
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                Set Asset = New ClsAsset
                Asset.DBGet !AssetNo
                Me.AddItem Asset
                .MoveNext
            Loop
        End If
    End With
    Set RstAsset = Nothing
End Sub

' ===============================================================
' Method DeleteCollection
' Deletes all items in collection
'---------------------------------------------------------------

Public Sub DeleteCollection()
    Dim Asset As ClsAsset
    For Each Asset In pAssets
        pAssets.Remove CStr(Asset.AssetNo)
        Asset.DBDelete
    Next
End Sub

' ===============================================================
' Method GetCategoryLists
' Gets items in category lists 1,2 or 3, returning them in an recordset
'---------------------------------------------------------------
Public Function GetCategoryLists(Optional Category1 As String, Optional Category2 As String, Optional Category3 As String) As Recordset
    Dim RstCategory As Recordset
    Dim SearchLvl As Integer
    Dim RecordFound As Boolean
    Dim FieldVal As String
    
    If Category1 = "" Then SearchLvl = 1
    If Category1 <> "" Then SearchLvl = 2
    If Category2 <> "" Then SearchLvl = 3
    
    Select Case SearchLvl
        Case 1
            Set RstCategory = ModDatabase.SQLQuery("SELECT DISTINCT Category1 FROM TblAsset")
        Case 2
            Set RstCategory = ModDatabase.SQLQuery("SELECT DISTINCT Category2 FROM TblAsset" & _
                                " WHERE Category1 = '" & Category1 & "'")
        Case 3
            Set RstCategory = ModDatabase.SQLQuery("SELECT DISTINCT Category3 FROM TblAsset" & _
                                " WHERE Category1 = '" & Category1 & "'" & _
                                " AND Category2 = '" & Category2 & "'")
    End Select
        
    'if only a null record is received, return Recordset is nothing
    With RstCategory
        RecordFound = False
        .MoveFirst
        Do While Not .EOF
            If Not IsNull(.Fields(0)) Then
                FieldVal = Trim(.Fields(0))
                If FieldVal <> "" Then RecordFound = True
            End If
            .MoveNext
        Loop
    End With
    
    If RecordFound Then Set GetCategoryLists = RstCategory Else Set GetCategoryLists = Nothing

    Set RstCategory = Nothing
End Function

' ===============================================================
' Method GetItemList
' Gets list of items from category lists, returning them in a recordset
'---------------------------------------------------------------
Public Function GetItemList(Optional Category1 As String, Optional Category2 As String, Optional Category3 As String) As Recordset
    Dim RstItemList As Recordset
    Dim StrSearch As String
    
    If Category1 <> "" Then StrSearch = " WHERE Category1 = '" & Category1 & "'"
    If Category2 <> "" Then StrSearch = StrSearch & " AND Category2 = '" & Category2 & "'"
    If Category3 <> "" Then StrSearch = StrSearch & " AND Category3 = '" & Category3 & "'"

    Set RstItemList = ModDatabase.SQLQuery("SELECT DISTINCT Description FROM TblAsset" & StrSearch)
    
    Set GetItemList = RstItemList
    Set RstItemList = Nothing
End Function

' ===============================================================
' Method GetSizeLists
' Returns size lists for a given asset no
'---------------------------------------------------------------
Public Function GetSizeLists(Description As String, ListNo As Integer) As String()
    Dim RstSizeList As Recordset
    Dim StrSizeArry() As String
    Dim RecordCnt As Integer
    Dim i As Integer
    
    Set RstSizeList = ModDatabase.SQLQuery("SELECT DISTINCT Size" & ListNo & " FROM TblAsset" & _
                                    " WHERE Description = '" & Description & "'")


    With RstSizeList
        If .RecordCount = 0 Then
            ReDim StrSizeArry(0 To 0)
        Else
        ReDim StrSizeArry(1 To .RecordCount)
        
        RecordCnt = .RecordCount
        i = 0
        .MoveFirst
        Do While Not .EOF
            If IsNull(.Fields(0)) Then
                RecordCnt = RecordCnt - 1
            Else
                If Trim(.Fields(0)) = "" Then
                    RecordCnt = RecordCnt - 1
                Else
                    i = i + 1
                    StrSizeArry(i) = .Fields(0)
                End If
            End If
            .MoveNext
        Loop
        End If
        
    End With
    
    GetSizeLists = StrSizeArry
    Set RstSizeList = Nothing
End Function

' ===============================================================
' FindAssetNo
' Finds Asset No from Asset Description
' ---------------------------------------------------------------
Public Function FindAssetNo(Description As String, Size1 As String, Size2 As String) As Integer
    Dim RstAssetNo As Recordset
    Dim StrSearch As String
    
    If Description = "" Then
        FindAssetNo = 0
        Exit Function
    End If
    
    If Size1 <> "" Then StrSearch = " AND Size1 = '" & Size1 & "'"
    If Size2 <> "" Then StrSearch = StrSearch & " AND Size2 = '" & Size2 & "'"
    
    Set RstAssetNo = ModDatabase.SQLQuery("SELECT AssetNo FROM TblAsset WHERE Description = '" & Description & "'" & StrSearch)
    
    If RstAssetNo.RecordCount > 0 Then
    FindAssetNo = RstAssetNo!AssetNo
    Else
        FindAssetNo = 0
    End If
    
    Set RstAssetNo = Nothing
End Function

' ===============================================================
Private Sub Class_Initialize()
    Set pAssets = New Collection

End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
    Dim Asset As ClsAsset

    For Each Asset In pAssets
        pAssets.Remove CStr(Asset.AssetNo)
        Set Asset = Nothing
    Next
    
    Set pAssets = Nothing
End Sub

'---------------------------------------------------------------




