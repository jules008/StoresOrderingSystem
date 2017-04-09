VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub BtnDBConnect_Click()
    ModDatabase.SelectDB
End Sub

Private Sub CommandButton1_Click()
    ModCloseDown.Terminate
End Sub



