Attribute VB_Name = "ModCloseDown"
'===============================================================
' Module ModCloseDown
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 17 Jan 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModCloseDown"

' ===============================================================
' Terminate
' Functions for graceful close down of system
' ---------------------------------------------------------------
Public Sub Terminate()
    Dim Frame As ClsUIFrame
    Dim DashObj As ClsUIDashObj
    Dim MenuItem As ClsUIMenuItem
    
    Const StrPROCEDURE As String = "Terminate()"

    On Error Resume Next
        
    ShtMain.Unprotect
    
    For Each Frame In MainScreen.Frames
        'debug.print Frame.Name
        For Each DashObj In Frame.DashObs
            'debug.print DashObj.Name
            DashObj.ShpDashObj.Delete
            Set DashObj = Nothing
        Next
        
        For Each MenuItem In Frame.Menu
            'debug.print MenuItem.Name
            MenuItem.ShpMenuItem.Delete
            MenuItem.Icon.Delete
            Set MenuItem = Nothing
        Next
        
        Frame.Header.Icon.Delete
        Frame.Header.ShpHeader.Delete
        Set Frame.Header = Nothing
        
        Frame.ShpFrame.Delete
        Set Frame = Nothing
    Next
    
    Set MainScreen = Nothing
    
    If Not Stations Is Nothing Then Set Stations = Nothing
    If Not CurrentUser Is Nothing Then Set CurrentUser = Nothing
    If Not Vehicles Is Nothing Then Set Vehicles = Nothing

    ModDatabase.DBTerminate
    DeleteAllShapes
    
End Sub


' ===============================================================
' DeleteAllShapes
' Deletes all shapes on screen except templates
' ---------------------------------------------------------------
Private Sub DeleteAllShapes()
    Dim i As Integer
    
    Const StrPROCEDURE As String = "DeleteAllShapes()"

    On Error Resume Next

    Dim Shp As Shape
    
    For i = ShtMain.Shapes.Count To 1 Step -1
    
        Set Shp = ShtMain.Shapes(i)
        'debug.print i & "/" & ShtMain.Shapes.Count & " " & Shp.Name
        
        If Left(Shp.Name, 8) <> "TEMPLATE" Then Shp.Delete
    Next

End Sub
