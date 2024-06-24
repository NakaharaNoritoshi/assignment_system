VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub クリア_Click()
    Application.Run macro:="Eiyo_Macro.xla!Eiyo01_030クリアClick"
End Sub

Private Sub 検索_Click()
    Application.Run macro:="Eiyo_Macro.xla!Eiyo01_100検索Click"
End Sub

Private Sub 更新_Click()
    Application.Run macro:="Eiyo_Macro.xla!Eiyo01_200更新Click"
End Sub

Private Sub 取消_Click()
    If MsgBox("取消を実行しますか", vbOKCancel) = vbOK Then
        Application.Run macro:="Eiyo_Macro.xla!Eiyo01_300取消Click"
    End If
End Sub

Private Sub 作表_Click()
    Application.Run macro:="Eiyo_Macro.xla!Eiyo01_700作表Click"
End Sub

Private Sub 終了_Click()
    ActiveWorkbook.Close savechanges:=False     '保存せずに終了
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Left(Target.Address, 4) = "$AH$" Then        'ファミリーコード一覧からの選択
        Range("Fcode") = Target.Value
        Cancel = True
        検索_Click
    ElseIf Left(Target.Address, 4) = "$Q$8" Then    '職業と労作
        Cancel = True
        Sheets("職業と労作").Select
    End If
End Sub
'--------------------------------------------------------------------------------
'   ワークシートの値が変更された
'--------------------------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address = Range("Adrno").Address Then         '郵便番号
        If Len(Target.Text) > 3 And Range("adrs1") = Empty Then
            Application.EnableEvents = False    'イベント発生抑止
            ' 郵便番号を全角変換し住所に転記
            Range("adrs1").Value = StrConv(Range("Adrno").Value, vbWide)
            ' 住所のｾﾙを選択
            Range("adrs1").Select
            ' F2 → Shift+Home → F13 を擬装入力する
            SendKeys "{F2}", True       ' 編集ﾓｰﾄﾞ
            SendKeys "+{HOME}", True    ' 文字列全体を選択
            SendKeys "{F13}", True      ' 再変換(MS-IME)
            Application.EnableEvents = True     'イベント発生再開
        End If
    End If
End Sub
'--------------------------------------------------------------------------------
'   アクティブになったセルを強調表示するためセルが選択させたとき画面更新する
'--------------------------------------------------------------------------------
Private Sub Worksheet_SelectionChange(ByVal Traget As Range)
    Application.ScreenUpdating = True
End Sub
