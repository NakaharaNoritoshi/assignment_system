VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'--------------------------------------------------------------------------------
'   職業と労作をダブルクリックで選択
'--------------------------------------------------------------------------------
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim Wline As Long '行
    Dim Wcolu As Long '欄
    Dim 基礎シート As Worksheet

    ' 基礎シートをオブジェクトとして取得
    Set 基礎シート = ThisWorkbook.Sheets("基礎")

    ' ダブルクリックのイベントをキャンセル
    Cancel = True

    ' ダブルクリックされたセルの行と列を取得
    Wline = Target.Row
    Wcolu = Target.Column

    ' 列の値を調整
    Select Case Wcolu
        Case 2: Wcolu = 3
        Case 6: Wcolu = 7
        Case 10: Wcolu = 11
    End Select

    ' 基礎シートのセルに値を設定
    基礎シート.Range("Qjob1").Value = Target.Parent.Cells(Wline, Wcolu).Value

    ' 基礎シートを選択
    基礎シート.Select
End Sub
