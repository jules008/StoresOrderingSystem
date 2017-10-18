Attribute VB_Name = "ModUISettings"
'===============================================================
' Module ModUISettings
' v0,0 - Initial Version
' v0,1 - Added Order Switch Button
' v0,2 - Added Remote Order Button
' v0,3 - Right Frame Order List
' v0,4 - Left Frame Order List
' v0,5 - Delivery button and tidy up
' v0,6 - Report1 Button
' v0,7 - Data Management Button
' v0,8 - Added Report2 Button
' v0,9 - Added Exit Button
' v0,10 - Added Order Age Column
' v0,11 - Added FindOrder Button
' v0,12 - Change Delivery button to Supplier
' v0,13 - Added My Profile screen
'---------------------------------------------------------------
' Date - 18 Oct 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModUISettings"

' ===============================================================
' Global Constants
' ---------------------------------------------------------------
' Main Screen
' ---------------------------------------------------------------
Public Const SCREEN_HEIGHT As Integer = 800
Public Const SCREEN_WIDTH As Integer = 1025
' ---------------------------------------------------------------
' Main Frame
' ---------------------------------------------------------------
Public Const MAIN_FRAME_TOP As Integer = 10
Public Const MAIN_FRAME_LEFT As Integer = 175
Public Const MAIN_FRAME_WIDTH As Integer = 764
Public Const MAIN_FRAME_HEIGHT As Integer = 215
' ---------------------------------------------------------------
' Left Frame
' ---------------------------------------------------------------
Public Const LEFT_FRAME_TOP As Integer = 240
Public Const LEFT_FRAME_LEFT As Integer = 175
Public Const LEFT_FRAME_WIDTH As Integer = 373
Public Const LEFT_FRAME_HEIGHT As Integer = 215

Public Const RCT_ORDER_LINEITEM_HEIGHT As Integer = 15
Public Const RCT_ORDER_LINEITEM_WIDTH As Integer = 550
Public Const RCT_ORDER_LINEITEM_LEFT As Integer = 0
Public Const RCT_ORDER_LINEITEM_TOP As Integer = 25
Public Const RCT_ORDER_LINEITEM_NOCOLS As Integer = 4
Public Const RCT_ORDER_LINEITEM_COL_WIDTHS As String = "70:100:100:100"
Public Const RCT_ORDER_LINEITEM_ROWOFFSET As Integer = 15
Public Const RCT_ORDER_LINEITEM_TITLES As String = "Order No:Order Date:Ordered By:Order Status"
Public Const RCT_ORDER_MAX_LINES As Integer = 10

' ---------------------------------------------------------------
' Right Frame
' ---------------------------------------------------------------
Public Const RIGHT_FRAME_TOP As Integer = 240
Public Const RIGHT_FRAME_LEFT As Integer = 566
Public Const RIGHT_FRAME_WIDTH As Integer = 373
Public Const RIGHT_FRAME_HEIGHT As Integer = 125

Public Const MY_ORDER_LINEITEM_HEIGHT As Integer = 15
Public Const MY_ORDER_LINEITEM_WIDTH As Integer = 550
Public Const MY_ORDER_LINEITEM_LEFT As Integer = 0
Public Const MY_ORDER_LINEITEM_TOP As Integer = 25
Public Const MY_ORDER_LINEITEM_NOCOLS As Integer = 4
Public Const MY_ORDER_LINEITEM_COL_WIDTHS As String = "70:100:100:100"
Public Const MY_ORDER_LINEITEM_ROWOFFSET As Integer = 15
Public Const MY_ORDER_LINEITEM_TITLES As String = "Order No:Order Date:Assigned To:Order Status"
Public Const MY_ORDER_MAX_LINES As Integer = 5

' ---------------------------------------------------------------
' Menu Bar
' ---------------------------------------------------------------
Public Const MENUBAR_HEIGHT As Integer = 800
Public Const MENUBAR_WIDTH As Integer = 150
Public Const MENUBAR_TOP As Integer = 0
Public Const MENUBAR_LEFT As Integer = 0
Public Const MENU_TOP As Integer = 180
Public Const MENU_LEFT As Integer = 10
Public Const MENUITEM_HEIGHT As Integer = 31
Public Const MENUITEM_WIDTH As Integer = 150
Public Const MENUITEM_COUNT As Integer = 6
Public Const MENUITEM_TEXT = "My Station:Stores:Reports:My Profile:Support:Exit"
Public Const MENUITEM_ICONS = "TEMPLATE - Icon_Station:TEMPLATE - Icon_Stores:TEMPLATE - Icon_Document:TEMPLATE - Icon_Head:TEMPLATE - Icon_Support:TEMPLATE - Exit"
Public Const MENUITEM_ICON_TOP As Integer = 5
Public Const MENUITEM_ICON_LEFT As Integer = 5
Public Const LOGO_TOP As Integer = 15
Public Const LOGO_LEFT As Integer = 20
Public Const LOGO_WIDTH As Integer = 126
Public Const LOGO_HEIGHT As Integer = 60

Public Const HEADER_HEIGHT As Integer = 25
Public Const HEADER_ICON_TOP As Integer = 5
Public Const HEADER_ICON_RIGHT As Integer = 5

Public Const BTN_NEWORDER_TOP As Integer = 390
Public Const BTN_NEWORDER_LEFT As Integer = 566
Public Const BTN_NEWORDER_WIDTH As Integer = 373
Public Const BTN_NEWORDER_HEIGHT As Integer = 75

' ---------------------------------------------------------------
' Support Screen
' ---------------------------------------------------------------
Public Const SUPPORT_FRAME_1_HEIGHT As Integer = 200
Public Const SUPPORT_FRAME_1_WIDTH As Integer = 370
Public Const SUPPORT_FRAME_1_LEFT As Integer = 175
Public Const SUPPORT_FRAME_1_TOP As Integer = 10

Public Const COMMENT_BOX_HEIGHT As Integer = 100
Public Const COMMENT_BOX_WIDTH As Integer = 175
Public Const COMMENT_BOX_LEFT As Integer = 10
Public Const COMMENT_BOX_TOP As Integer = 35

Public Const COMMENT_BTN_HEIGHT As Integer = 30
Public Const COMMENT_BTN_WIDTH As Integer = 145
Public Const COMMENT_BTN_LEFT As Integer = 25
Public Const COMMENT_BTN_TOP As Integer = 150

' ---------------------------------------------------------------
' My Profile Screen
' ---------------------------------------------------------------
Public Const MY_PROFILE_1_HEIGHT As Integer = 245
Public Const MY_PROFILE_1_WIDTH As Integer = 440
Public Const MY_PROFILE_1_LEFT As Integer = 175
Public Const MY_PROFILE_1_TOP As Integer = 10

Public Const MY_PROFILE_TEXTBOX_HEIGHT As Integer = 20
Public Const MY_PROFILE_TEXTBOX_WIDTH As Integer = 120

Public Const MY_PROFILE_LABEL_HEIGHT As Integer = 20
Public Const MY_PROFILE_LABEL_WIDTH As Integer = 75

Public Const MY_PROFILE_BUTTON_HEIGHT As Integer = 25
Public Const MY_PROFILE_BUTTON_WIDTH As Integer = 100

Public Const MY_PROFILE_LBLCREWNO_LEFT As Integer = 10
Public Const MY_PROFILE_LBLCREWNO_TOP As Integer = 40

Public Const MY_PROFILE_TXTCREWNO_LEFT As Integer = 85
Public Const MY_PROFILE_TXTCREWNO_TOP As Integer = 40

Public Const MY_PROFILE_LBLRANKGRADE_LEFT As Integer = 225
Public Const MY_PROFILE_LBLRANKGRADE_TOP As Integer = 40

Public Const MY_PROFILE_TXTRANKGRADE_LEFT As Integer = 300
Public Const MY_PROFILE_TXTRANKGRADE_TOP As Integer = 40

Public Const MY_PROFILE_LBLFORENAME_LEFT As Integer = 10
Public Const MY_PROFILE_LBLFORENAME_TOP As Integer = 80

Public Const MY_PROFILE_TXTFORENAME_LEFT As Integer = 85
Public Const MY_PROFILE_TXTFORENAME_TOP As Integer = 80

Public Const MY_PROFILE_LBLSURNAME_LEFT As Integer = 225
Public Const MY_PROFILE_LBLSURNAME_TOP As Integer = 80

Public Const MY_PROFILE_TXTSURNAME_LEFT As Integer = 300
Public Const MY_PROFILE_TXTSURNAME_TOP As Integer = 80

Public Const MY_PROFILE_LBLROLE_LEFT As Integer = 10
Public Const MY_PROFILE_LBLROLE_TOP As Integer = 120

Public Const MY_PROFILE_TXTROLE_LEFT As Integer = 85
Public Const MY_PROFILE_TXTROLE_TOP As Integer = 120

Public Const MY_PROFILE_LBLACCESSLVL_LEFT As Integer = 225
Public Const MY_PROFILE_LBLACCESSLVL_TOP As Integer = 120

Public Const MY_PROFILE_TXTACCESSLVL_LEFT As Integer = 300
Public Const MY_PROFILE_TXTACCESSLVL_TOP As Integer = 120

Public Const MY_PROFILE_LBLLOCATION_LEFT As Integer = 10
Public Const MY_PROFILE_LBLLOCATION_TOP As Integer = 160

Public Const MY_PROFILE_CMOLOCATION_LEFT As Integer = 85
Public Const MY_PROFILE_CMOLOCATION_TOP As Integer = 160

Public Const MY_PROFILE_LBLWATCH_LEFT As Integer = 225
Public Const MY_PROFILE_LBLWATCH_TOP As Integer = 160

Public Const MY_PROFILE_TXTWATCH_LEFT As Integer = 300
Public Const MY_PROFILE_TXTWATCH_TOP As Integer = 160

Public Const MY_PROFILE_BTNUPDATE_LEFT As Integer = 320
Public Const MY_PROFILE_BTNUPDATE_TOP As Integer = 200


' ---------------------------------------------------------------
' Stores Screen
' ---------------------------------------------------------------
Public Const STORES_FRAME_1_HEIGHT As Integer = 300
Public Const STORES_FRAME_1_WIDTH As Integer = 650
Public Const STORES_FRAME_1_LEFT As Integer = 175
Public Const STORES_FRAME_1_TOP As Integer = 10

Public Const BTN_USER_MANGT_HEIGHT As Integer = 30
Public Const BTN_USER_MANGT_WIDTH As Integer = 175
Public Const BTN_USER_MANGT_LEFT As Integer = 850
Public Const BTN_USER_MANGT_TOP As Integer = 20

Public Const BTN_ORDER_SWITCH_HEIGHT As Integer = 30
Public Const BTN_ORDER_SWITCH_WIDTH As Integer = 175
Public Const BTN_ORDER_SWITCH_LEFT As Integer = 850
Public Const BTN_ORDER_SWITCH_TOP As Integer = 60

Public Const BTN_REMOTE_ORDER_HEIGHT As Integer = 30
Public Const BTN_REMOTE_ORDER_WIDTH As Integer = 175
Public Const BTN_REMOTE_ORDER_LEFT As Integer = 850
Public Const BTN_REMOTE_ORDER_TOP As Integer = 100

Public Const BTN_SUPPLIER_HEIGHT As Integer = 30
Public Const BTN_SUPPLIER_WIDTH As Integer = 175
Public Const BTN_SUPPLIER_LEFT As Integer = 850
Public Const BTN_SUPPLIER_TOP As Integer = 140

Public Const BTN_MANAGE_DATA_HEIGHT As Integer = 30
Public Const BTN_MANAGE_DATA_WIDTH As Integer = 175
Public Const BTN_MANAGE_DATA_LEFT As Integer = 850
Public Const BTN_MANAGE_DATA_TOP As Integer = 180

Public Const BTN_FIND_ORDER_HEIGHT As Integer = 30
Public Const BTN_FIND_ORDER_WIDTH As Integer = 175
Public Const BTN_FIND_ORDER_LEFT As Integer = 850
Public Const BTN_FIND_ORDER_TOP As Integer = 220

Public Const ORDER_LINEITEM_HEIGHT As Integer = 15
Public Const ORDER_LINEITEM_WIDTH As Integer = 550
Public Const ORDER_LINEITEM_LEFT As Integer = 20
Public Const ORDER_LINEITEM_TOP As Integer = 30
Public Const ORDER_LINEITEM_NOCOLS As Integer = 7
Public Const ORDER_LINEITEM_COL_WIDTHS As String = "70:70:70:120:100:100:70"
Public Const ORDER_LINEITEM_ROWOFFSET As Integer = 20
Public Const ORDER_LINEITEM_TITLES As String = "Order No:Days Old:Items:Requested By:Station:Assigned To:Order Status"

' ---------------------------------------------------------------
' Report Screen
' ---------------------------------------------------------------
Public Const BTN_REPORT_1_HEIGHT As Integer = 30
Public Const BTN_REPORT_1_WIDTH As Integer = 175
Public Const BTN_REPORT_1_LEFT As Integer = 180
Public Const BTN_REPORT_1_TOP As Integer = 30

Public Const BTN_REPORT_2_HEIGHT As Integer = 30
Public Const BTN_REPORT_2_WIDTH As Integer = 175
Public Const BTN_REPORT_2_LEFT As Integer = 365
Public Const BTN_REPORT_2_TOP As Integer = 30

' ===============================================================
' Style Declarations
' ---------------------------------------------------------------
' Main Screen
' ---------------------------------------------------------------
Public SCREEN_STYLE As TypeStyle
Public MENUBAR_STYLE As TypeStyle
Public MENUITEM_SET_STYLE As TypeStyle
Public MENUITEM_UNSET_STYLE As TypeStyle
Public MAIN_FRAME_STYLE As TypeStyle
Public HEADER_STYLE As TypeStyle
Public BTN_NEWORDER_STYLE As TypeStyle
Public GENERIC_BUTTON As TypeStyle
Public GENERIC_LINEITEM As TypeStyle
Public GENERIC_LINEITEM_HEADER As TypeStyle
Public GENERIC_LABEL As TypeStyle

' ---------------------------------------------------------------
' New Order Workflow
' ---------------------------------------------------------------
Public WF_MAINSCREEN_STYLE As TypeStyle

' ===============================================================
' Style Definitions
' ---------------------------------------------------------------
' Generic Styles
' ---------------------------------------------------------------
Public Const GENERIC_BUTTON_BORDER_WIDTH As Long = 0
Public Const GENERIC_BUTTON_FILL_1 As Long = COLOUR_11
Public Const GENERIC_BUTTON_FILL_2 As Long = COLOUR_6
Public Const GENERIC_BUTTON_SHADOW As Long = msoShadow21
Public Const GENERIC_BUTTON_FONT_STYLE As String = "Eras Medium ITC"
Public Const GENERIC_BUTTON_FONT_SIZE As Integer = 12
Public Const GENERIC_BUTTON_FONT_COLOUR As Long = COLOUR_2
Public Const GENERIC_BUTTON_FONT_BOLD As Boolean = False
Public Const GENERIC_BUTTON_FONT_X_JUST As Integer = xlHAlignCenter
Public Const GENERIC_BUTTON_FONT_Y_JUST As Integer = xlVAlignCenter

Public Const GENERIC_LINEITEM_BORDER_WIDTH As Long = 0
Public Const GENERIC_LINEITEM_FILL_1 As Long = COLOUR_3
Public Const GENERIC_LINEITEM_FILL_2 As Long = COLOUR_3
Public Const GENERIC_LINEITEM_SHADOW As Long = 0
Public Const GENERIC_LINEITEM_FONT_STYLE As String = "Eras Medium ITC"
Public Const GENERIC_LINEITEM_FONT_SIZE As Integer = 10
Public Const GENERIC_LINEITEM_FONT_COLOUR As Long = COLOUR_1
Public Const GENERIC_LINEITEM_FONT_BOLD As Boolean = False
Public Const GENERIC_LINEITEM_FONT_X_JUST As Integer = xlHAlignCenter
Public Const GENERIC_LINEITEM_FONT_Y_JUST As Integer = xlVAlignCenter

Public Const GENERIC_LINEITEM_HEADER_BORDER_WIDTH As Long = 0
Public Const GENERIC_LINEITEM_HEADER_FILL_1 As Long = COLOUR_3
Public Const GENERIC_LINEITEM_HEADER_FILL_2 As Long = COLOUR_3
Public Const GENERIC_LINEITEM_HEADER_SHADOW As Long = 0
Public Const GENERIC_LINEITEM_HEADER_FONT_STYLE As String = "Calibri"
Public Const GENERIC_LINEITEM_HEADER_FONT_SIZE As Integer = 10
Public Const GENERIC_LINEITEM_HEADER_FONT_COLOUR As Long = COLOUR_2
Public Const GENERIC_LINEITEM_HEADER_FONT_BOLD As Boolean = True
Public Const GENERIC_LINEITEM_HEADER_FONT_X_JUST As Integer = xlHAlignCenter
Public Const GENERIC_LINEITEM_HEADER_FONT_Y_JUST As Integer = xlVAlignCenter

Public Const GENERIC_LABEL_BORDER_WIDTH As Long = 0
Public Const GENERIC_LABEL_FILL_1 As Long = COLOUR_3
Public Const GENERIC_LABEL_FILL_2 As Long = COLOUR_3
Public Const GENERIC_LABEL_SHADOW As Long = 0
Public Const GENERIC_LABEL_FONT_STYLE As String = "Eras Medium ITC"
Public Const GENERIC_LABEL_FONT_SIZE As Integer = 10
Public Const GENERIC_LABEL_FONT_COLOUR As Long = COLOUR_1
Public Const GENERIC_LABEL_FONT_BOLD As Boolean = False
Public Const GENERIC_LABEL_FONT_X_JUST As Integer = xlHAlignLeft
Public Const GENERIC_LABEL_FONT_Y_JUST As Integer = xlVAlignCenter

' ---------------------------------------------------------------
' Main Screen
' ---------------------------------------------------------------
Public Const SCREEN_BORDER_WIDTH As Long = 0
Public Const SCREEN_FILL_1 As Long = COLOUR_1
Public Const SCREEN_FILL_2 As Long = COLOUR_1
Public Const SCREEN_SHADOW As Long = msoShadow21

Public Const MENUBAR_BORDER_WIDTH As Long = 0
Public Const MENUBAR_FILL_1 As Long = COLOUR_2
Public Const MENUBAR_FILL_2 As Long = COLOUR_2
Public Const MENUBAR_SHADOW As Long = msoShadow21

Public Const MENUITEM_UNSET_BORDER_WIDTH As Long = 0
Public Const MENUITEM_UNSET_FILL_1 As Long = COLOUR_5
Public Const MENUITEM_UNSET_FILL_2 As Long = COLOUR_2
Public Const MENUITEM_UNSET_SHADOW As Long = 0
Public Const MENUITEM_UNSET_FONT_STYLE As String = "Eras Medium ITC"
Public Const MENUITEM_UNSET_FONT_SIZE As Integer = 12
Public Const MENUITEM_UNSET_FONT_COLOUR As Long = COLOUR_3
Public Const MENUITEM_UNSET_FONT_X_JUST As Integer = xlHAlignCenter
Public Const MENUITEM_UNSET_FONT_Y_JUST As Integer = xlVAlignCenter

Public Const MENUITEM_SET_BORDER_WIDTH As Long = 0
Public Const MENUITEM_SET_FILL_1 As Long = COLOUR_4
Public Const MENUITEM_SET_FILL_2 As Long = COLOUR_4
Public Const MENUITEM_SET_SHADOW As Long = 0
Public Const MENUITEM_SET_FONT_STYLE As String = "Eras Medium ITC"
Public Const MENUITEM_SET_FONT_SIZE As Integer = 12
Public Const MENUITEM_SET_FONT_COLOUR As Long = COLOUR_3
Public Const MENUITEM_SET_FONT_X_JUST As Integer = xlHAlignCenter
Public Const MENUITEM_SET_FONT_Y_JUST As Integer = xlVAlignCenter

Public Const MAIN_FRAME_BORDER_WIDTH As Long = 0
Public Const MAIN_FRAME_FILL_1 As Long = COLOUR_3
Public Const MAIN_FRAME_FILL_2 As Long = COLOUR_3
Public Const MAIN_FRAME_SHADOW As Long = msoShadow21

Public Const HEADER_BORDER_WIDTH As Long = 0
Public Const HEADER_FILL_1 As Long = COLOUR_4
Public Const HEADER_FILL_2 As Long = COLOUR_4
Public Const HEADER_SHADOW As Long = 0
Public Const HEADER_FONT_STYLE As String = "Calibri"
Public Const HEADER_FONT_SIZE As Integer = 12
Public Const HEADER_FONT_COLOUR As Long = COLOUR_3
Public Const HEADER_FONT_BOLD As Boolean = True
Public Const HEADER_FONT_X_JUST As Integer = xlHAlignCenter
Public Const HEADER_FONT_Y_JUST As Integer = xlVAlignCenter

Public Const BTN_NEWORDER_BORDER_WIDTH As Long = 0
Public Const BTN_NEWORDER_FILL_1 As Long = COLOUR_11
Public Const BTN_NEWORDER_FILL_2 As Long = COLOUR_6
Public Const BTN_NEWORDER_SHADOW As Long = msoShadow21
Public Const BTN_NEWORDER_FONT_STYLE As String = "Calibri"
Public Const BTN_NEWORDER_FONT_SIZE As Integer = 32
Public Const BTN_NEWORDER_FONT_COLOUR As Long = COLOUR_2
Public Const BTN_NEWORDER_FONT_BOLD As Boolean = True
Public Const BTN_NEWORDER_FONT_X_JUST As Integer = xlHAlignCenter
Public Const BTN_NEWORDER_FONT_Y_JUST As Integer = xlVAlignCenter

