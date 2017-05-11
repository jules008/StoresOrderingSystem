Attribute VB_Name = "ModUISettings"
'===============================================================
' Module ModUISettings
' v0,0 - Initial Version
' v0,1 - Added Order Switch Button
' v0,2 - Added Remote Order Button
' v0,31 - Right Frame Order List
'---------------------------------------------------------------
' Date - 11 May 17
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
Public Const MENUITEM_COUNT As Integer = 5
Public Const MENUITEM_TEXT = "My Station:Stores:Reports:My Profile:Support"
Public Const MENUITEM_ICONS = "TEMPLATE - Icon_Station:TEMPLATE - Icon_Stores:TEMPLATE - Icon_Document:TEMPLATE - Icon_Head:TEMPLATE - Icon_Support"
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

Public Const ORDER_LINEITEM_HEIGHT As Integer = 15
Public Const ORDER_LINEITEM_WIDTH As Integer = 550
Public Const ORDER_LINEITEM_LEFT As Integer = 25
Public Const ORDER_LINEITEM_TOP As Integer = 30
Public Const ORDER_LINEITEM_NOCOLS As Integer = 6
Public Const ORDER_LINEITEM_COL_WIDTHS As String = "70:100:100:100:100:100"
Public Const ORDER_LINEITEM_ROWOFFSET As Integer = 20
Public Const ORDER_LINEITEM_TITLES As String = "Order No:No of Items:Requested By:Station:Assigned To:Order Status"


' ---------------------------------------------------------------
' New Order Workflow
' ---------------------------------------------------------------
Public Const WF_MAINSCREEN_TOP As Integer = 70
Public Const WF_MAINSCREEN_LEFT As Integer = 100
Public Const WF_MAINSCREEN_HEIGHT As Integer = 500
Public Const WF_MAINSCREEN_WIDTH As Integer = 800
Public Const WF_CLOSE_ICON_TOP As Integer = 5
Public Const WF_CLOSE_ICON_LEFT As Integer = 5
Public Const WF_CLOSE_ICON_HEIGHT As Integer = 10
Public Const WF_CLOSE_ICON_WIDTH As Integer = 10

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

' ---------------------------------------------------------------
' Main Screen
' ---------------------------------------------------------------

' ---------------------------------------------------------------
' New Order Workflow
' ---------------------------------------------------------------
Public Const WF_MAINSCREEN_BORDER_WIDTH As Long = 2
Public Const WF_MAINSCREEN_BORDER_COLOUR As Long = COLOUR_1
Public Const WF_MAINSCREEN_FILL_1 As Long = COLOUR_8
Public Const WF_MAINSCREEN_FILL_2 As Long = COLOUR_8
Public Const WF_MAINSCREEN_SHADOW As Long = msoShadow21
Public Const WF_MAINSCREEN_FONT_STYLE As String = "Calibri"
Public Const WF_MAINSCREEN_FONT_SIZE As Integer = 36
Public Const WF_MAINSCREEN_FONT_COLOUR As Long = COLOUR_2
Public Const WF_MAINSCREEN_FONT_BOLD As Boolean = True
Public Const WF_MAINSCREEN_FONT_X_JUST As Integer = xlHAlignCenter
Public Const WF_MAINSCREEN_FONT_Y_JUST As Integer = xlVAlignCenter

