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
'---------------------------------------------------------------
' Date - 22 Feb 17
'===============================================================
' Methods
'---------------------------------------------------------------
'===============================================================

Option Explicit
Private pDeleted As Date

'---------------------------------------------------------------
Public Property Get Deleted() As Date
    Deleted = pDeleted
End Property

Public Property Let Deleted(ByVal VnewValue As Date)
    pDeleted = VnewValue
End Property

'---------------------------------------------------------------
