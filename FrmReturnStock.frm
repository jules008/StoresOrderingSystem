VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmReturnStock 
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8505
   OleObjectBlob   =   "FrmReturnStock.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmReturnStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
' v0,0 - Initial version
' v0,1 - Mark Return Order as closed
'---------------------------------------------------------------
' Date - 07 Jan 19
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmReturnStock"

Dim RetOrder As ClsOrder

' ===============================================================
' ShowForm
' Initial entry point to form
' ---------------------------------------------------------------
Public Function ShowForm() As Boolean
    
    Const StrPROCEDURE As String = "ShowForm()"
    
    On Error GoTo ErrorHandler
    
    Show
    ShowForm = True

Exit Function

ErrorExit:

    FormTerminate
    Terminate
    ShowForm = False

Exit Function

ErrorHandler:
    
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        CustomErrorHandler Err.Number
        Resume Next
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' FormTerminate
' Terminates the form gracefully
' ---------------------------------------------------------------
Private Function FormTerminate() As Boolean

    On Error Resume Next
    
    Set RetOrder = Nothing
    Unload Me

End Function

' ===============================================================
' BtnReturn_Click
' Moves onto next form
' ---------------------------------------------------------------
Private Sub BtnReturn_Click()
    Dim StrUserName As String
    
    Const StrPROCEDURE As String = "BtnReturn_Click()"

    On Error GoTo ErrorHandler

        Select Case ValidateForm
    
            Case Is = FunctionalError
                Err.Raise HANDLED_ERROR
            
            Case Is = FormOK
                
                With RetOrder
                    .OrderDate = Now
                    .Requestor = CurrentUser
                    .Lineitems(1).Quantity = 0 - TxtQty
                    .Lineitems(1).ReqReason = ItemReturn
                    .Status = OrderClosed
                    .DBSave
                End With
                
                MsgBox "Return has been successfully processed", vbOKCancel + vbInformation, APP_NAME
                Hide
                Unload Me
                 
        End Select
        
GracefulExit:

Exit Sub

ErrorExit:

    FormTerminate
    Terminate

Exit Sub

ErrorHandler:
    
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        CustomErrorHandler Err.Number
        Resume GracefulExit
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' BtnClose_Click
' Closes form
' ---------------------------------------------------------------
Private Sub BtnClose_Click()
    Dim StrUserName As String
    
    Const StrPROCEDURE As String = "BtnClose_Click()"

    On Error GoTo ErrorHandler

    Hide
    Unload Me
                       
GracefulExit:

Exit Sub

ErrorExit:

    FormTerminate
    Terminate

Exit Sub

ErrorHandler:
    
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        CustomErrorHandler Err.Number
        Resume GracefulExit
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' LstFrom_Click
' Adds either a person or vehicle to the order dependant on the allocation type
' ---------------------------------------------------------------
Private Sub LstFrom_click()
    Dim AllocType As EnumAllocationType
    Dim Person As ClsPerson
    Dim Persons As ClsPersons
    Dim Vehicle As ClsVehicle
    Dim SelVehicleID As Integer
    Dim SelPersonID As Double
    
    Set Person = New ClsPerson
    Set Persons = New ClsPersons
    Set Vehicle = New ClsVehicle
    
    AllocType = RetOrder.Lineitems(1).Asset.AllocationType
    
    Select Case AllocType
        Case Is = 0
        
            With LstFrom
                SelPersonID = .List(.ListIndex, 1)
                RetOrder.Lineitems(1).ForPerson.DBGet CStr(SelPersonID)
            End With
            
        Case Is = 1
        
             With LstFrom
                SelVehicleID = .List(.ListIndex, 0)
                RetOrder.Lineitems(1).ForVehicle.DBGet CStr(SelVehicleID)
           End With
       
    End Select

    Set Persons = Nothing
    Set Person = Nothing
    Set Vehicle = Nothing
End Sub

' ===============================================================
' LstStations_Click
' When Station selected, show either vehicles or people depending on asset type
' ---------------------------------------------------------------
Private Sub LstStations_Click()
    Dim ErrNo As Integer
    Dim AssetType As EnumAllocationType
    Dim Vehicle As ClsVehicle
    Dim Persons As ClsPersons
    Dim SelStation As String
    Dim RstCrewMembers As Recordset
    Dim i As Integer
    
    Const StrPROCEDURE As String = "LstStations_Click()"

    On Error GoTo ErrorHandler

Restart:
    
    If MainScreen Is Nothing Then Err.Raise SYSTEM_RESTART
    
    AssetType = RetOrder.Lineitems(1).Asset.AllocationType
    
    With LstStations
        .BackColor = COLOUR_4
        SelStation = .List(.ListIndex, 0)
    End With
    
    Select Case AssetType
        Case Is = 0
            Set Persons = New ClsPersons
            
            Set RstCrewMembers = Persons.ReturnStationCrew(CInt(SelStation))
            
            LblFrom.Visible = True
            LblFrom.Caption = "Select the Person that the item is being returned from"
            
             With LstFrom
                .BackColor = COLOUR_3
                .Enabled = True
                .Clear
                i = 0
                
                Do While Not RstCrewMembers.EOF
                    .AddItem
                    If Not IsNull(RstCrewMembers!CrewNo) Then .List(i, 1) = RstCrewMembers!CrewNo
                    .List(i, 2) = RstCrewMembers!UserName
                    RstCrewMembers.MoveNext
                    i = i + 1
                    
                Loop
                Set Persons = Nothing
            End With
       
        Case Is = 1
        
            LblFrom.Visible = True
            LblFrom.Caption = "Select the vehicle that the item is being returned from"
            
             With LstFrom
                .BackColor = COLOUR_3
                .Enabled = True
                .Clear
                i = 0
                
                For Each Vehicle In Vehicles
                    If Vehicle.StationID = SelStation Then
                        .AddItem
                        .List(i, 0) = Vehicle.VehNo
                        .List(i, 1) = Vehicle.CallSign
                        .List(i, 2) = Vehicles.GetVehicleType(Vehicle.VehType)
                        i = i + 1
                    End If
                Next
            
            End With
        
        Case Is = 2
            
            RetOrder.Lineitems(1).ForStation = Stations(SelStation)
           
    End Select

GracefulExit:


Exit Sub

ErrorExit:
    '***CleanUpCode***
    Set Persons = Nothing
Exit Sub

ErrorHandler:
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        ErrNo = Err.Number
        CustomErrorHandler (Err.Number)
        If ErrNo = SYSTEM_RESTART Then Resume Restart Else Resume GracefulExit
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub
' ===============================================================
' UserForm_Initialize
' Automatic initialise event that triggers custom Initialise
' ---------------------------------------------------------------
Private Sub UserForm_Initialize()

    On Error Resume Next
    
    FormInitialise
    
End Sub

' ===============================================================
' UserForm_Terminate
' Automatic Terminate event that triggers custom Terminate
' ---------------------------------------------------------------
Private Sub UserForm_Terminate()

    On Error Resume Next
    
    FormTerminate
    
End Sub

' ===============================================================
' FormInitialise
' initialises controls on form at start up
' ---------------------------------------------------------------
Private Function FormInitialise() As Boolean
    Dim i As Integer
    Dim Lineitem As ClsLineItem
    
    Const StrPROCEDURE As String = "FormInitialise()"
    
    On Error GoTo ErrorHandler
    
    Set RetOrder = New ClsOrder
    Set Lineitem = New ClsLineItem
    
    RetOrder.Lineitems.AddItem Lineitem
    
    LstFrom.Enabled = False
    LstFrom.BackColor = COLOUR_9
    LstStations.Enabled = False
    LstStations.BackColor = COLOUR_9
    LblFrom.Visible = False
    LblStations.Visible = False
    
    If Not ClearSearch Then Err.Raise HANDLED_ERROR

    If Not ShtLists.RefreshAssetList Then Err.Raise HANDLED_ERROR
    
    Set Lineitem = Nothing
    
    FormInitialise = True

Exit Function

ErrorExit:
    
    Set Lineitem = Nothing

    FormTerminate
    Terminate
    
    FormInitialise = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' ValidateForm
' Ensures the form is filled out correctly before moving on
' ---------------------------------------------------------------
Private Function ValidateForm() As EnumFormValidation
    Const StrPROCEDURE As String = "ValidateForm()"

    On Error GoTo ErrorHandler
    
    With TxtSearch
        If .Value = "" Then
            .BackColor = COLOUR_6
            ValidateForm = ValidationError
        Else
            .BackColor = COLOUR_3
        End If
    End With
            
    With LstAssets
        If .ListIndex = -1 Then
            .BackColor = COLOUR_6
            ValidateForm = ValidationError
        Else
            .BackColor = COLOUR_3
        End If
    End With
    
    With TxtQty
        If .Value = "" Then
            .BackColor = COLOUR_6
            ValidateForm = ValidationError
        Else
            .BackColor = COLOUR_3
        End If
    End With
    
    With CmoSize1
        If .Visible = True And .Value = "" Then
            .BackColor = COLOUR_6
            ValidateForm = ValidationError
        Else
            .BackColor = COLOUR_3
        End If
    End With
    
    With CmoSize2
        If .Visible = True And .Value = "" Then
            .BackColor = COLOUR_6
            ValidateForm = ValidationError
        Else
            .BackColor = COLOUR_3
        End If
    End With
    
    With LstStations
        If .ListIndex = -1 And .Enabled Then
            .BackColor = COLOUR_6
            ValidateForm = ValidationError
        Else
            .BackColor = COLOUR_3
        End If
    End With
    
     With LstFrom
        If .ListIndex = -1 And .Enabled Then
            .BackColor = COLOUR_6
            ValidateForm = ValidationError
        Else
            .BackColor = COLOUR_3
        End If
    End With
                       
    If ValidateForm = ValidationError Then
        Err.Raise FORM_INPUT_EMPTY
    Else
        ValidateForm = FormOK
    End If
    
    ValidateForm = FormOK

GracefulExit:


Exit Function

ErrorExit:

    ValidateForm = FunctionalError
    FormTerminate
    Terminate

Exit Function

ErrorHandler:
    
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        CustomErrorHandler Err.Number
        Resume GracefulExit
    End If

If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' TxtQty_Change
' Change event for quantity txt box
' ---------------------------------------------------------------
Private Sub TxtQty_Change()
    TxtQty.BackColor = COLOUR_3

End Sub

' ===============================================================
' TxtSearch_Change
' Entry for search string
' ---------------------------------------------------------------
Private Sub TxtSearch_Change()
    Const StrPROCEDURE As String = "TxtSearch_Change()"

    On Error GoTo ErrorHandler

    Dim ListResults As String

    On Error GoTo ErrorHandler
    
    TxtSearch.BackColor = COLOUR_3
    LstAssets.BackColor = COLOUR_3

    With LstAssets
        If .ListIndex <> -1 Then ListResults = .List(.ListIndex)
    
        'if the search box has been changed since being updated by the results box, clear the result box
        If ListResults <> TxtSearch.Value Then .ListIndex = -1
        
        'if the results box has been clicked, add the selected result to the search box
        If .ListIndex = -1 Then
        
            'if no results selected, populate with new results
            If Len(TxtSearch.Value) > 1 Then
                If Not GetSearchItems(TxtSearch.Value) Then Err.Raise HANDLED_ERROR
            Else
                .Clear
            End If
        End If
    End With

Exit Sub

ErrorExit:

    FormTerminate
    Terminate

Exit Sub

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub


' ===============================================================
' ClearSearch
' Clears form ready for new item search
' ---------------------------------------------------------------
Private Function ClearSearch() As Boolean
    Const StrPROCEDURE As String = "ClearSearch()"

    On Error GoTo ErrorHandler

    TxtSearch = ""
    TxtQty = ""
    CmoSize1 = ""
    CmoSize2 = ""
    CmoSize1.Visible = False
    CmoSize2.Visible = False
    LblSize1.Visible = False
    LblSize2.Visible = False

    ClearSearch = True

Exit Function

ErrorExit:

    ClearSearch = False

    FormTerminate
    Terminate

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function
' ===============================================================
' GetSearchItems
' Gets items from the name list that match Txtsearch box
' ---------------------------------------------------------------
Private Function GetSearchItems(StrSearch As String) As Boolean
    Const StrPROCEDURE As String = "GetSearchItems()"

    On Error GoTo ErrorHandler

    Dim ListLength As Integer
    Dim RngResult As Range
    Dim RngItems As Range
    Dim RngFirstResult As Range
    Dim i As Integer
    
    'get length of item list
    ListLength = Application.WorksheetFunction.CountA(ShtLists.Range("A:A"))
    
    Set RngItems = ShtLists.Range("A1:A" & ListLength)
            
    Set RngResult = RngItems.Find(StrSearch)
    Set RngFirstResult = RngResult
    
    LstAssets.Clear
    'search item list and populate results.  Stop before looping back to start
    If Not RngResult Is Nothing Then
        Do
            Set RngResult = RngItems.FindNext(RngResult)
            LstAssets.AddItem RngResult
        Loop While RngResult <> 0 And RngResult.Address <> RngFirstResult.Address
    End If

    GetSearchItems = True
    
    Set RngItems = Nothing
    Set RngResult = Nothing
    Set RngFirstResult = Nothing
    
Exit Function

ErrorExit:

    Set RngItems = Nothing
    Set RngResult = Nothing
    Set RngFirstResult = Nothing
    
    FormTerminate
    Terminate

    GetSearchItems = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' LstAssets_Click
' Gets items from the name list that match Txtsearch box
' ---------------------------------------------------------------
Private Sub LstAssets_Click()

    On Error Resume Next

    LstAssets.BackColor = COLOUR_3
    
    LblStations.Visible = True
    LblFrom.Visible = False
    LstFrom.Enabled = False
    
    With LstFrom
        .Clear
        .BackColor = COLOUR_9
    End With
    
    With LstAssets
        Me.TxtSearch.Value = .List(.ListIndex)
        
    End With
                
    If Not UpdateSizeLists Then Err.Raise HANDLED_ERROR
    If Not GetAsset Then Err.Raise HANDLED_ERROR
    
    TxtQty.SetFocus
End Sub

' ===============================================================
' GetAsset
' if all information completed, retrieves Asset and adds to order
' ---------------------------------------------------------------
Private Function GetAsset() As Boolean
    Dim LocAsset As ClsAsset
    Dim Assets As ClsAssets
    Dim AssetAvail As Boolean
    Dim AssetType As EnumAllocationType
    
    Const StrPROCEDURE As String = "GetAsset()"

    On Error GoTo ErrorHandler
    
    Set Assets = New ClsAssets
    Set LocAsset = New ClsAsset
    
    If Not CmoSize1.Visible And Not CmoSize2.Visible Then AssetAvail = True
    If CmoSize2.Visible And CmoSize2.ListIndex <> -1 Then AssetAvail = True
    If CmoSize1.Visible And CmoSize1.ListIndex <> -1 And CmoSize2.Visible And CmoSize2.ListIndex <> -1 Then AssetAvail = True
    If CmoSize1.Visible And Not CmoSize2.Visible Then AssetAvail = True
    
    Debug.Print "Asset Avail = " & AssetAvail
    
    If AssetAvail Then
        LocAsset.DBGet (Assets.FindAssetNo(TxtSearch, CmoSize1, CmoSize2))
        RetOrder.Lineitems(1).Asset = LocAsset
        AssetType = LocAsset.AllocationType
        
        If Not ShowStations Then Err.Raise HANDLED_ERROR


    End If

    GetAsset = True
    Set LocAsset = Nothing
    Set Assets = Nothing

Exit Function

ErrorExit:

    '***CleanUpCode***
    Set LocAsset = Nothing
    GetAsset = False
    Set Assets = Nothing

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
' CmoSize1_Change
' Event when entry changes
' ---------------------------------------------------------------
Private Sub CmoSize1_Change()
    Dim Assets As ClsAssets
    Dim StrSize2Arry() As String
    Dim i As Integer
    
    Const StrPROCEDURE As String = "CmoSize1_Change()"

    On Error GoTo ErrorHandler
   
    Set Assets = New ClsAssets
    
    CmoSize1.BackColor = COLOUR_3
    
    If CmoSize1 = "" Then CmoSize2 = ""

    StrSize2Arry() = Assets.GetSizeLists(TxtSearch, 2, CmoSize1)

    If UBound(StrSize2Arry) <> LBound(StrSize2Arry) Then
        LblSize2.Visible = True
        CmoSize2.Visible = True
        
        CmoSize2.Clear
        For i = LBound(StrSize2Arry) To UBound(StrSize2Arry)
            CmoSize2.AddItem StrSize2Arry(i)
        Next
    End If
    
    If Not GetAsset Then Err.Raise HANDLED_ERROR

    Set Assets = Nothing
Exit Sub

ErrorExit:

    Set Assets = Nothing
    FormTerminate
    Terminate

Exit Sub

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE, True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' CmoSize2_Change
' Event when entry changes
' ---------------------------------------------------------------
Private Sub CmoSize2_Change()
    Dim Assets As ClsAssets
    
    Const StrPROCEDURE As String = "CmoSize1_Change()"

    On Error GoTo ErrorHandler
   
    Set Assets = New ClsAssets
    
    CmoSize2.BackColor = COLOUR_3
    
    If Not GetAsset Then Err.Raise HANDLED_ERROR

    Set Assets = Nothing

Exit Sub

ErrorExit:

    Set Assets = Nothing
    FormTerminate
    Terminate

Exit Sub

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE, True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' UpdateSizeLists
' Event when item changes
' ---------------------------------------------------------------
Private Function UpdateSizeLists() As Boolean
    Dim AssetNo As Integer
    Dim Assets As ClsAssets
    Dim StrSize1Arry() As String
    Dim StrSize2Arry() As String
    Dim i As Integer
    
    Const StrPROCEDURE As String = "UpdateSizeLists()"

    On Error GoTo ErrorHandler
        
    TxtSearch.BackColor = COLOUR_3
    
    If TxtSearch <> "" Then
        
        Set Assets = New ClsAssets
        
        StrSize1Arry() = Assets.GetSizeLists(TxtSearch, 1)
        StrSize2Arry() = Assets.GetSizeLists(TxtSearch, 2, CmoSize1)
        
        CmoSize1.Clear
        CmoSize2.Clear
        CmoSize1.Visible = False
        CmoSize2.Visible = False
        LblSize1.Visible = False
        LblSize2.Visible = False
        TxtQty = ""
        
        If UBound(StrSize1Arry) <> LBound(StrSize1Arry) Then
            LblSize1.Visible = True
            CmoSize1.Visible = True
            
            For i = LBound(StrSize1Arry) To UBound(StrSize1Arry)
                CmoSize1.AddItem StrSize1Arry(i)
            Next
        End If
        
        If UBound(StrSize2Arry) <> LBound(StrSize2Arry) Then
            LblSize2.Visible = True
            CmoSize2.Visible = True
            
            For i = LBound(StrSize2Arry) To UBound(StrSize2Arry)
                CmoSize2.AddItem StrSize2Arry(i)
            Next
        End If
        
    End If
    
    UpdateSizeLists = True
    
GracefulExit:

Exit Function

ErrorExit:

    UpdateSizeLists = False
    
    FormTerminate
    Terminate

Exit Function

ErrorHandler:
    
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        CustomErrorHandler Err.Number
        Resume GracefulExit
    End If


    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' ProcessReturn
' Creates a negative order to return stock to Stores
' ---------------------------------------------------------------
Private Function ProcessReturn() As Boolean
    Dim RetOrder As ClsOrder
    Dim RetLineitem As ClsLineItem
    Dim Assets As ClsAssets
    Dim AssetNo As Integer
    
    Const StrPROCEDURE As String = "ProcessReturn()"

    On Error GoTo ErrorHandler

    With RetLineitem
        .Asset = Assets.FindItem(AssetNo)
    End With
    
    With RetOrder
        
    
    
    End With



    Set RetOrder = Nothing
    Set RetLineitem = Nothing
    Set Assets = Nothing
    
    ProcessReturn = True


Exit Function

ErrorExit:
    
    Set RetOrder = Nothing
    Set RetLineitem = Nothing
    Set Assets = Nothing
    
    '***CleanUpCode***
    ProcessReturn = False

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
' ShowStations
' Fills out list of Stations, Vehicles or People
' ---------------------------------------------------------------
Private Function ShowStations() As Boolean
    Dim Station As ClsStation
    Dim Vehicle As ClsVehicle
    Dim i As Integer
    
    Const StrPROCEDURE As String = "ShowStations()"

    On Error GoTo ErrorHandler

     With LstStations
        .BackColor = COLOUR_3
        .Enabled = True
        .Clear
        i = 0
        For Each Station In Stations
            If Station.StnActive Then
                .AddItem
                .List(i, 0) = Station.StationID
                .List(i, 1) = Station.StationNo
                .List(i, 2) = Station.Name
                i = i + 1
            End If
        Next
    End With
    
    ShowStations = True

Exit Function

ErrorExit:

    '***CleanUpCode***
    ShowStations = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function
