VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================
' Sheet ShtLists
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 22 Feb 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ShtLists"

' ===============================================================
' RefreshAssetList
' Updates the list of assets from the database
' ---------------------------------------------------------------
Public Function RefreshAssetList() As Boolean
    Const StrPROCEDURE As String = "RefreshAssetList()"

    On Error GoTo ErrorHandler

    Dim RstAssets As Recordset
    
    Set RstAssets = ModDatabase.SQLQuery("SELECT DISTINCT Description FROM tblasset")

    Me.Range("A1").CopyFromRecordset RstAssets

    RefreshAssetList = True
    
    Set RstAssets = Nothing

Exit Function

ErrorExit:

    Terminate

    Set RstAssets = Nothing
    
    RefreshAssetList = False

Exit Function

ErrorHandler:

    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' RefreshNameList
' Updates the list of assets from the database
' ---------------------------------------------------------------
Public Function RefreshNameList() As Boolean
    Dim RstNames As Recordset
    
    Const StrPROCEDURE As String = "RefreshNameList()"

    On Error GoTo ErrorHandler

    Range("C2:D3000").ClearContents
    
    Set RstNames = ModDatabase.SQLQuery("SELECT CrewNo, UserName FROM TblPerson ORDER BY CrewNo")

    Me.Range("C1").CopyFromRecordset RstNames

    RefreshNameList = True
    
    Set RstNames = Nothing

Exit Function

ErrorExit:
    
    Terminate

    Set RstNames = Nothing
    
    RefreshNameList = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

