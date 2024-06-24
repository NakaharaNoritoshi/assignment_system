VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'--------------------------------------------------------------------------------
'   メニューをダブルクリックで選択
'--------------------------------------------------------------------------------
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Application.Run "Eiyo_Macro.xla!Eiyo01_420Menu", Target.Address
    Cancel = True
End Sub
