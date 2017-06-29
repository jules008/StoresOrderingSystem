Attribute VB_Name = "ModGlobals"
'===============================================================
' Module ModGlobals
' v0,0 - Initial Version
' v0,1 - Added no ini file error
' v0,2 - Added Order Switch Button
' v0,3 - Add no stock avail error code
' v0,4 - added Remote Order button
' v0,5 - Added Deleted Order Status
' v0,6 - Added No File Selected Error
' v0,61 - Added DB Version constant
' v0,7 - Added Import Error
' v0,8 - DB Change
' v0,9 - Delivery Button
' v0,10 - Reports1 Button
' v0,11 - Data Management Button
' v0,12 - Added Report2 button
' v0,13 - Removed hard numbering on enumbuttons
' v0,14 - Added FindOrder Button
'---------------------------------------------------------------
' Date - 29 Jun 17
'===============================================================

Option Explicit
' ===============================================================
' Global Constants
' ---------------------------------------------------------------
Private Const StrMODULE As String = "ModGlobals"
Public Const INI_FILE_PATH As String = "\System Files\"
Public Const INI_FILE As String = "System.ini"
Public Const APP_NAME As String = "Stores IT System"
Public Const TEST_PREFIX As String = "TEST - "
Public Const FILE_ERROR_LOG As String = "Error.log"
Public Const VERSION = "0.0"
Public Const DB_VER = "v0,34"
Public Const VER_DATE = "15/01/17"

' ===============================================================
' Error Constants
' ---------------------------------------------------------------
Public Const UNKNOWN_USER As Long = 1000
Public Const NO_ITEM_SELECTED As Long = 1001
Public Const SYSTEM_RESTART As Long = 1002
Public Const NO_QUANTITY_ENTERED As Long = 1003
Public Const NO_SIZE_ENTERED As Long = 1004
Public Const NO_CREW_NO_ENTERED As Long = 1005
Public Const NUMBERS_ONLY As Long = 1006
Public Const CREWNO_UNRECOGNISED As Long = 1007
Public Const NO_DATABASE_FOUND As Long = 1008
Public Const NO_VEHICLE_SELECTED As Long = 1009
Public Const NO_STATION_SELECTED As Long = 1010
Public Const FIELDS_INCOMPLETE As Long = 1011
Public Const NO_NAMES_SELECTED As Long = 1012
Public Const FORM_INPUT_EMPTY As Long = 1013
Public Const ACCESS_DENIED As Long = 1014
Public Const NO_ORDER_MESSAGE As Long = 1015
Public Const NO_INI_FILE As Long = 1016
Public Const NO_STOCK_AVAIL As Long = 1017
Public Const DB_WRONG_VER As Long = 1018
Public Const NO_FILE_SELECTED As Long = 1018
Public Const IMPORT_ERROR As Long = 1019

Public Const NO_ASSET_ON_ORDER As Long = 1501
Public Const NO_LINE_ITEM As Long = 1502
Public Const NO_ORDER As Long = 1503
Public Const NO_RECORDSET_RETURNED As Long = 1504
Public Const SYSTEM_FAILURE As Long = 1505
Public Const NO_LOSS_REPORT As Long = 1506
Public Const NO_USER_AVAILABLE As Long = 1507
Public Const NO_REQUESTOR As Long = 1508
Public Const NO_ASSET_FOUND As Long = 1509
Public Const HANDLED_ERROR As Long = 9999
Public Const USER_CANCEL As Long = 18

' ===============================================================
' Global Variables
' ---------------------------------------------------------------
Public DB_PATH As String
Public ALLOW_EMAILS As Boolean
Public DEBUG_MODE As Boolean
Public TEST_MODE As Boolean
Public OUTPUT_MODE As String
Public ENABLE_PRINT As Boolean
Public SEND_EMAILS As Boolean
Public DEV_MODE As Boolean
Public TMP_FILE_PATH As String
Public MENU_ITEM_SEL As Integer

' ===============================================================
' Global Class Declarations
' ---------------------------------------------------------------
' Main Screen
' ---------------------------------------------------------------
Public MainScreen As ClsUIScreen
Public MenuBar As ClsUIFrame
Public Logo As ClsUIDashObj
Public Menu As ClsUIMenu
Public MenuItem As ClsUIMenuItem
Public MainFrame As ClsUIFrame
Public LeftFrame As ClsUIFrame
Public RightFrame As ClsUIFrame
Public Header As ClsUIHeader
Public BtnNewOrder As ClsUIMenuItem

' ---------------------------------------------------------------
' Support Screen
' ---------------------------------------------------------------
Public SupportFrame1 As ClsUIFrame

' ---------------------------------------------------------------
' Stores Screen
' ---------------------------------------------------------------
Public StoresFrame1 As ClsUIFrame
Public BtnUserMangt As ClsUIMenuItem
Public BtnOrderSwitch As ClsUIMenuItem
Public BtnRemoteOrder As ClsUIMenuItem
Public BtnDelivery As ClsUIMenuItem
Public BtnManageData As ClsUIMenuItem
Public BtnFindOrder As ClsUIMenuItem

' ---------------------------------------------------------------
' Reports Screen
' ---------------------------------------------------------------
Public BtnReport1 As ClsUIMenuItem
Public BtnReport2 As ClsUIMenuItem

' ---------------------------------------------------------------
' Others
' ---------------------------------------------------------------
Public MailSystem As ClsMailSystem
Public CurrentUser As ClsPerson
Public Vehicles As ClsVehicles
Public Stations As ClsStations

' ===============================================================
' Colours
' ---------------------------------------------------------------
Public Const COLOUR_1 As Long = 5525013
Public Const COLOUR_2 As Long = 2369842
Public Const COLOUR_3 As Long = 16777215
Public Const COLOUR_4 As Long = 10396448
Public Const COLOUR_5 As Long = 5266544
Public Const COLOUR_6 As Long = 3450623
Public Const COLOUR_7 As Long = 6893787
Public Const COLOUR_8 As Long = 16056312
Public Const COLOUR_9 As Long = 12439241
Public Const COLOUR_10 As Long = 7864234
Public Const COLOUR_11 As Long = 52479

' ===============================================================
' Enum Declarations
' ---------------------------------------------------------------
Enum EnumOrderStatus
    OrderOpen = 0
    OrderAssigned = 1
    OrderOnHold = 2
    OrderIssued = 3
    OrderClosed = 4
    OrderDeleted = 5
End Enum

Enum EnumPersonRole
    Requestor = 0
    OpsSupport = 1
    Stores = 2
    Supervisor = 3
End Enum

Enum EnumAccessLvl
    BasicLvl_1 = 0
    StoresLvl_2 = 1
    SupervisorLvl_3 = 2
    ManagerLvl_4 = 3
    AdminLvl_5 = 4
End Enum

Enum EnumReqReason
    UsedConsumed = 0
    lost = 1
    Stolen = 2
    DamagedOpTraining = 3
    DamagedOther = 4
    Malfunction = 5
    NewIssue = 6
End Enum

Enum EnumStationID
    Alford = 0
    Bardney = 1
    Billingborough = 2
    Billinghay = 3
    Binbrook = 4
    Boston = 5
    BostonAccomPods = 6
    Bourne = 7
    BrantBroughton = 8
    Caistor = 9
    CorbyGlen = 10
    Crowland = 11
    Donington = 12
    Gainsborough = 13
    GainsbouroughAccomPods = 14
    Grantham = 15
    GranthamAccomPods = 16
    Holbeach = 17
    Horncastle = 18
    Kirton = 19
    Leverton = 20
    LincolnNorth = 21
    LincolnNorthAccomPods = 22
    LincolnSouth = 23
    LongSutton = 24
    Louth = 25
    LouthAccomPods = 26
    Mablethorpe = 27
    MarketDeeping = 28
    MarketRasen = 29
    Metheringham = 30
    NorthHykeham = 31
    NorthSomercotes = 32
    Saxilby = 33
    Skegness = 34
    Sleaford = 35
    Spalding = 36
    SpaldingAccomPods = 37
    Spilsby = 38
    Stamford = 39
    Waddington = 40
    Wainfleet = 41
    WoodhallSpa = 42
    Wragby = 43
    SleafordWT = 44
    SleafordAccomPods = 45
    HQ = 46
    WTF = 47
    Control = 48
End Enum

Enum EnumStnType
    WholeTime = 1
    Retained = 2
    HQ = 3
    WTF = 4
End Enum

Enum EnumAssetStatus
    Ok = 0
    ReOrder = 1
    LowLevel = 2
    NoStock = 3
End Enum

Enum EnumLossRepStatus
    RepOpen = 0
    RepAssigned = 1
    RepOnHold = 2
    RepApproved = 3
    RepRejected = 4
End Enum

Enum EnumVehType
    Pump = 0
    ALP = 1
    RSU = 2
    WelfareUnit
    CommandUnit
    WaterCarrier
    Wayfarer
    Autoroller500
    TransitAWD
    Vivaro
    Movano
    IRU
    ForkLiftTruck
    Primemover
    Plant
    Hatchback
    EstateCar
    Ambulance
    Estate
    Panelvan
    TransitSWB
    Minibus
    FordTransitCustom
    FDSCar
End Enum

Enum EnumFormValidation
    FormOK = 2
    ValidationError = 1
    FunctionalError = 0
End Enum

Enum EnumAllocationType
    Person = 0
    Vehicle = 1
    Station = 2
End Enum

Enum EnumSrchType
    TextSearch = 1
    CategorySearch = 2
End Enum

Enum EnumLineItemStatus
    LineOpen = 0
    LineOnHold = 1
    LineIssued = 2
    LineDelivered = 3
    LineComplete = 4
End Enum

Enum EnumObjType
    ObjImage = 1
    ObjChart = 2
End Enum

Enum EnumBtnNo
    EnumMyStation = 1
    EnumStores
    EnumReports
    EnumMyProfile
    EnumSupport
    EnumExit
    EnumNewOrder
    EnumSupportMsg
    EnumUserMngt
    EnumOrderSwitch
    EnumRemoteOrder
    EnumDeliveryBtn
    EnumReport1Btn
    EnumManageDataBtn
    EnumReport2Btn
    EnumFindOrderBtn
End Enum

' ===============================================================
' Type Declarations
' ---------------------------------------------------------------
Type TypeStyle
    ForeColour As Long
    BorderColour As Long
    BorderWidth As Long
    FontStyle As String
    FontBold As Boolean
    FontSize As Integer
    FontColour As Long
    FontXJust As XlHAlign
    FontYJust As XlVAlign
    Fill1 As Long
    Fill2 As Long
    Shadow As MsoShadowType
End Type

