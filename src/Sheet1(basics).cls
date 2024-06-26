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

' クリアボタンがクリックされたときの処理
Private Sub クリア_Click()
    Application.Run macro:="Eiyo_Macro.xla!Eiyo01_030クリアClick"
End Sub

' 検索ボタンがクリックされたときの処理
Private Sub 検索_Click()
    Application.Run macro:="Eiyo_Macro.xla!Eiyo01_100検索Click"
End Sub

' 更新ボタンがクリックされたときの処理
Private Sub 更新_Click()
    Application.Run macro:="Eiyo_Macro.xla!Eiyo01_200更新Click"
End Sub

' 取消ボタンがクリックされたときの処理
Private Sub 取消_Click()
    If MsgBox("取消を実行しますか", vbOKCancel) = vbOK Then
        Application.Run macro:="Eiyo_Macro.xla!Eiyo01_300取消Click"
    End If
End Sub

' 作表ボタンがクリックされたときの処理
Private Sub 作表_Click()
    ' Eiyo_Macro.xla にある Eiyo01_700作表Click マクロを実行
    Application.Run macro:="Eiyo_Macro.xla!Eiyo01_700作表Click"
End Sub

' 終了ボタンがクリックされたときの処理
Private Sub 終了_Click()
    ActiveWorkbook.Close savechanges:=False ' 保存せずにアクティブなブックを閉じる
End Sub

' シート内でダブルクリックが発生したときの処理
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Left(Target.Address, 4) = "$AH$" Then    ' セルが$AH$列にある場合の処理
        Range("Fcode") = Target.Value   ' "Fcode" 範囲にターゲットの値を設定し、検索ボタンクリック処理を呼び出す
        Cancel = True
        検索_Click
    ElseIf Left(Target.Address, 4) = "$Q$8" Then    ' セルが$Q$8にある場合の処理
        Cancel = True
        Sheets("職業と労作").Select                   ' "職業と労作" シートを選択
    End If
End Sub
'--------------------------------------------------------------------------------
'   ワークシートの値が変更された
'--------------------------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
    ' 郵便番号が入力されたセルが変更された場合の処理
    If Target.Address = Range("Adrno").Address Then ' 郵便番号セルのアドレスが変更された場合
        If Len(Target.Text) > 3 And Range("adrs1") = Empty Then
            Application.EnableEvents = False    ' イベント発生を一時停止
            ' 郵便番号を全角変換し住所に転記
            Range("adrs1").Value = StrConv(Range("Adrno").Value, vbWide)
            ' 住所セルを選択
            Range("adrs1").Select
            ' キー操作をシミュレート F2 → Shift+Home → F13 を擬装入力する
            SendKeys "{F2}", True       ' 編集モードにする
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
    ' SelectionChangeイベントが発生すると、このプロシージャが実行
    ' Target は変更されたセル範囲を表す
    Application.ScreenUpdating = True   ' 画面の更新を有効にする
End Sub
