VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'--------------------------------------------------------------------------------
'   摂食のワークシートが選択された
'--------------------------------------------------------------------------------
Private Sub Worksheet_Activate()
    Application.Run macro:="Eiyo_Macro.xla!Eiyo01_400MealDisp"
End Sub
'--------------------------------------------------------------------------------
'   ワークシート値が変更された
'--------------------------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
    Application.EnableEvents = False            'イベント発生抑止
    Application.Run "Eiyo_Macro.xla!Eiyo01_410MealChange", Target.Address
    Application.EnableEvents = True             'イベント発生再開
End Sub
'--------------------------------------------------------------------------------
'   ダブルクリックされた
'--------------------------------------------------------------------------------
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Application.EnableEvents = False            'イベント発生抑止
    Cancel = True
    Range("b1") = Target.Address
    Application.EnableEvents = True             'イベント発生再開
    Sheets("ﾒﾆｭｰ").Select
'    Application.Run "Eiyo_Macro.xla!Eiyo01_420Menu", Target.Address
End Sub
'--------------------------------------------------------------------------------
'   検証_Click
'--------------------------------------------------------------------------------
Private Sub 検証_Click()
    Application.Run "Eiyo_Macro.xla!Eiyo01_500MealCalc", 2
End Sub
'--------------------------------------------------------------------------------
'   登録_Click
'--------------------------------------------------------------------------------
Private Sub 登録_Click()
    Application.Run "Eiyo_Macro.xla!Eiyo01_500MealCalc", 1
End Sub
