Attribute VB_Name = "Module1"
Option Explicit
Private Declare PtrSafe Function SetCurrentDirectory Lib "kernel32" Alias _
        "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Const myFileName = "Eiyo.mdb"""
Const Tbl_Kiso = "F_Kiso"
Const Tbl_Syoyo = "F_Syoyo"
Const Tbl_Energ = "F_Energ"
Const Tbl_Meal = "F_Meal"
Const Tbl_Food = "F_Food"
Const Tbl_Field = "F_Field"
Const Tbl_Need = "F_Need"
Const Tbl_Advic = "F_Advic"
Dim myCon       As New ADODB.Connection
Dim Rst_Kiso    As New ADODB.Recordset
Dim Rst_Syoyo   As New ADODB.Recordset
Dim Rst_Energ   As New ADODB.Recordset
Dim Rst_Meal    As New ADODB.Recordset
Dim Rst_Food    As New ADODB.Recordset
Dim Rst_Field   As New ADODB.Recordset
Dim Rst_Need    As New ADODB.Recordset
Dim Rst_Advic   As New ADODB.Recordset
Dim Fld_Adrs1   As Variant
'Dim Fld_Adrs2   As Variant
Dim Fld_Area    As Variant
Dim Fld_Field   As Variant
'--------------------------------------------------------------------------------
'   Eiyo01_000 画面項目定義
'   0:Field-Name
'   1:セル範囲
'   2:i/o 入力可否を明確化
'   3:Field (D:DB G:Guid)
'   4:Type
'   5:Sample
'--------------------------------------------------------------------------------
Function Eiyo01_000init()
Dim Wtext   As String
    Wtext = "Gmesg,a025:a025,o,00,X ,Message"
    Wtext = Wtext & vbLf & "Fcode,g003:k003,i,D,X,1234567890"    'Fcode
    Wtext = Wtext & vbLf & "Fsave,l003:p003,o,G,90,1234567890"   'Fcode save
    Wtext = Wtext & vbLf & "Date1,g004:k004,i,D,Ds,2008/10/10"   '調査期間自
    Wtext = Wtext & vbLf & "Nissu,l004:l004,i,D,90,1"            '期間
    Wtext = Wtext & vbLf & "Namej,g005:o005,i,D,J ,氏名１氏名２氏名３氏"
    Wtext = Wtext & vbLf & "Sex  ,g006:g006,i,D,X ,1"            '性別"
    Wtext = Wtext & vbLf & "Birth,g007:k007,i,D,Ds,2001/1/1"     '生年月日"
    Wtext = Wtext & vbLf & "Gyyyy,l007:m007,o,G,X ,"             '和暦生年
    Wtext = Wtext & vbLf & "Age  ,n007:o007,o,D,90,"             '年齢"
    Wtext = Wtext & vbLf & "Hight,g008:i008,i,D,91,123.4"        '身長
    Wtext = Wtext & vbLf & "Weght,g009:i009,i,D,91,123.4"        '体重
    Wtext = Wtext & vbLf & "Sibou,g010:i010,i,D,91,123.4"        '皮下脂肪
    Wtext = Wtext & vbLf & "Adrno,g011:j011,i,D,X ,123-4567"     '郵便番号
    Wtext = Wtext & vbLf & "Adrs1,g012:v012,i,D,J ,住所ー１＃住所ー１＃住所ー１＃住所ー"
    Wtext = Wtext & vbLf & "Adrs2,g013:v013,i,D,J ,住所ー２＃住所ー２＃住所ー２＃住所ー"
    Wtext = Wtext & vbLf & "Area1,g014:h014,o,D,X ,12"           '地区
    Wtext = Wtext & vbLf & "Gare1,i014:i014,o,G,X ,地域名"       '地域
    Wtext = Wtext & vbLf & "Area2,g015:h015,o,D,X ,12"           '地区
    Wtext = Wtext & vbLf & "Gare2,i015:i015,o,G,X ,都府県"       '地域
    Wtext = Wtext & vbLf & "Q3rec,g016:k016,i,D,X ,1234567890"   'Q3.食習
    Wtext = Wtext & vbLf & "Q4rec,g017:i017,i,D,X ,12345"        'Q4.休養
    Wtext = Wtext & vbLf & "Q5rec,g018:h018,i,D,X ,123"          'Q5.運動
    Wtext = Wtext & vbLf & "Q6r_a,g019:k019,i,D,X ,1234567890"   'Q6.健康-1
    Wtext = Wtext & vbLf & "Q6r_b,g020:k020,i,D,X ,1234567890"   'Q6.健康-2
    Wtext = Wtext & vbLf & "Q6r_c,g021:k021,i,D,X ,1234567890"   'Q6.健康-3
    Wtext = Wtext & vbLf & "Q6r_d,g022:k022,i,D,X ,1234567890"   'Q6.健康-4
    Wtext = Wtext & vbLf & "Q6r_e,g023:k023,i,D,X ,1234567890"   'Q6.健康-5
    Wtext = Wtext & vbLf & "Qjob1,q008:r008,i,D,X ,1234"         'Q7.職業-1
    Wtext = Wtext & vbLf & "Qjob5,s008:s008,i,D,X ,1"            'Q7.職業-2
    Wtext = Wtext & vbLf & "Qsyuf,q009:q009,i,D,X ,1"            'Qa.主婦
    Wtext = Wtext & vbLf & "Qcnd1,q010:q010,i,D,X ,1"            'Qb.妊娠
    Wtext = Wtext & vbLf & "Qtony,q011:q011,i,G,X ,1"            'Qc.頻尿
    Wtext = Wtext & vbLf & "Qill1,r011:r011,o,D,X ,123456"       'Qc.頻尿
    Wtext = Wtext & vbLf & "Qkoke,q014:q014,i,G,X ,1"            'Qd.高血圧
    Wtext = Wtext & vbLf & "Qill2,r014:r014,o,D,X ,123456"       'Qd.高血圧
    Wtext = Wtext & vbLf & "Qsrmr,q015:r015,i,D,90,123"          'Qe.Spot-1
    Wtext = Wtext & vbLf & "Qsmin,s015:t015,i,D,90,123"          'Qe.Spot-2
    Wtext = Wtext & vbLf & "Qclab,q016:q016,i,D,90,1"            'Qf.運動部
    Wtext = Wtext & vbLf & "Qtobc,q017:q017,i,D,90,1"            'Qt.喫煙
    Wtext = Wtext & vbLf & "Qsyog,q018:r018,i,D,90,12"           'Qg.身体障害
    Wtext = Wtext & vbLf & "Qwcnt,q019:q019,i,D,90,1"            'Qh.ｳｴｲﾄCT
    Wtext = Wtext & vbLf & "Tenes,q020:q020,i,D,90,1"            'Qi.ｴﾈﾙｷﾞｰ指定-1
    Wtext = Wtext & vbLf & "Tenee,r020:u020,i,D,92,12345.67"     'Qi.ｴﾈﾙｷﾞｰ指定-2
    Wtext = Wtext & vbLf & "Tanps,q021:q021,i,D,90,1"            'Qj.ﾀﾝﾊﾟｸ指定-1
    Wtext = Wtext & vbLf & "Tanpe,r021:u021,i,D,92,12345.67"     'Qj.ﾀﾝﾊﾟｸ指定-2
    Wtext = Wtext & vbLf & "ｶｳﾝｾﾗ1,q023:af23,i,D,J ,ｶｳﾝｾﾗ1"      '
    Wtext = Wtext & vbLf & "ｶｳﾝｾﾗ2,q024:af24,i,D,J ,ｶｳﾝｾﾗ2"      '
    Wtext = Wtext & vbLf & "ｶｳﾝｾﾗ3,q025:af25,i,D,J ,ｶｳﾝｾﾗ3"      '
    Wtext = Wtext & vbLf & "Blood,ab03:ac03,i,D,X ,12"           'B1.血液型
    Wtext = Wtext & vbLf & "Bscd1,ab04:ac04,i,D,X ,123"          'B2.支社部-1
    Wtext = Wtext & vbLf & "Bscd2,ad04:ae04,i,D,X ,12"           'B2.支社部-2
    Wtext = Wtext & vbLf & "Bhok1,ab05:ae05,i,D,X ,12345678"     'B3.保健記号
    Wtext = Wtext & vbLf & "Bhok2,ab06:ae06,i,D,X ,12345678"     'B4.保健記号
    Wtext = Wtext & vbLf & "Bhant,ab07:ac07,i,D,X ,12"           'B5.定期検診判定
    Wtext = Wtext & vbLf & "Barm ,ab08:ab08,i,D,X ,1"            'B6.検査腕
    Wtext = Wtext & vbLf & "Bdate,ab09:af09,i,D,Ds,2008/10/10"   'B7.検査日
    Wtext = Wtext & vbLf & "Bbl01,ab10:ad10,i,D,91,123.41"       'B8.赤血球数
    Wtext = Wtext & vbLf & "Bbl02,ab11:ad11,i,D,91,123.41"       'B8.血色素量
    Wtext = Wtext & vbLf & "Bbl03,ab12:ad12,i,D,91,123.41"       'B8.ﾍﾏﾄｸﾘｯﾄ
    Wtext = Wtext & vbLf & "Bbl04,ab13:ad13,i,D,91,123.41"       'B8.ｺﾚｽﾃﾛｰﾙ
    Wtext = Wtext & vbLf & "Bbl05,ab14:ad14,i,D,91,123.41"       'B8.HDL
    Wtext = Wtext & vbLf & "Bbl06,ab15:ad15,i,D,91,123.41"       'B8.中性脂肪
    Wtext = Wtext & vbLf & "Bbl07,ab16:ad16,i,D,91,123.41"       'B8.G.O.T.
    Wtext = Wtext & vbLf & "Bbl08,ab17:ad17,i,D,91,123.41"       'B8.G.P.T.
    Wtext = Wtext & vbLf & "Bbl09,ab18:ad18,i,D,91,123.41"       'B8.尿酸
    Wtext = Wtext & vbLf & "Bbl10,ab19:ad19,i,D,91,123.41"       'B8.血糖
    Wtext = Wtext & vbLf & "Bbl11,ab20:ad20,i,D,91,123.41"       'B8.血圧最高
    Wtext = Wtext & vbLf & "Bbl12,ab21:ad21,i,D,91,123.41"       'B8.血圧最低
    Fld_Adrs1 = Split(Wtext, vbLf)

    Wtext = "100,関東Ⅰ,100" & vbLf & "101,関東Ⅱ,101" & vbLf & "102,北　陸,102"
    Wtext = Wtext & vbLf & "103,東　海,103" & vbLf & "104,近畿Ⅰ,104" & vbLf & "105,近畿Ⅱ,105"
    Wtext = Wtext & vbLf & "106,中　国,106" & vbLf & "107,四　国,107" & vbLf & "108,北九州,108"
    Wtext = Wtext & vbLf & "109,南九州,109" & vbLf & "110,北海道,110" & vbLf & "111,東　北,111"
    Wtext = Wtext & vbLf & "201,北海道,110" & vbLf & "202,青森　,111" & vbLf & "203,岩手　,111"
    Wtext = Wtext & vbLf & "204,宮城　,111" & vbLf & "205,秋田　,111" & vbLf & "206,山形　,111"
    Wtext = Wtext & vbLf & "207,福島　,111" & vbLf & "208,茨城　,101" & vbLf & "209,栃木　,101"
    Wtext = Wtext & vbLf & "210,群馬　,101" & vbLf & "211,埼玉　,100" & vbLf & "212,千葉　,100"
    Wtext = Wtext & vbLf & "213,東京　,100" & vbLf & "214,神奈川,100" & vbLf & "215,新潟　,102"
    Wtext = Wtext & vbLf & "216,富山　,102" & vbLf & "217,石川　,102" & vbLf & "218,福井　,102"
    Wtext = Wtext & vbLf & "219,山梨　,101" & vbLf & "220,長野　,101" & vbLf & "221,岐阜　,103"
    Wtext = Wtext & vbLf & "222,静岡　,103" & vbLf & "223,愛知　,103" & vbLf & "224,三重　,103"
    Wtext = Wtext & vbLf & "225,滋賀　,105" & vbLf & "226,京都　,104" & vbLf & "227,大阪　,104"
    Wtext = Wtext & vbLf & "228,兵庫　,104" & vbLf & "229,奈良　,105" & vbLf & "230,和歌山,105"
    Wtext = Wtext & vbLf & "231,鳥取　,106" & vbLf & "232,島根　,106" & vbLf & "233,岡山　,106"
    Wtext = Wtext & vbLf & "234,広島　,106" & vbLf & "235,山口　,106" & vbLf & "236,徳島　,107"
    Wtext = Wtext & vbLf & "237,香川　,107" & vbLf & "238,愛媛　,107" & vbLf & "239,高知　,107"
    Wtext = Wtext & vbLf & "240,福岡　,108" & vbLf & "241,佐賀　,108" & vbLf & "242,長崎　,108"
    Wtext = Wtext & vbLf & "243,熊本　,109" & vbLf & "244,大分　,108" & vbLf & "245,宮崎　,109"
    Wtext = Wtext & vbLf & "246,鹿児島,109" & vbLf & "247,沖縄　,109"
    Fld_Area = Split(Wtext, vbLf)
End Function
'--------------------------------------------------------------------------------
'   01_010 摂食画面のワークシートがアクティブになった
'--------------------------------------------------------------------------------
Function Eiyo01_010�ېH_Activate()
    ActiveSheet.Unprotect                           'シートの保護を解除
'    ActiveSheet.Protect UserInterfaceOnly:=True     '保護を有効にする
End Function
'--------------------------------------------------------------------------------
'   01_020 基礎画面のダブルクリック
'   AA列（検索該当複数時）のダブルクリックは該当番号をセル[G3]に設定
'--------------------------------------------------------------------------------
Function Eiyo01_020��b_BeforedoubleClick()
Dim Wadrs   As String
Dim Wcoul   As String
Dim Wtext   As String
Dim i1      As Long     'Fld_Adrs Index
Dim i3      As Long     'ダブルクリックの行番号

End Function
'--------------------------------------------------------------------------------
'   01_030 クリア_Click
'   入力項目の消去、帳票・検証シートの削除
'--------------------------------------------------------------------------------
Function Eiyo01_030クリアClick()
Dim i1      As Long
Dim FldItem As Variant
Dim Lmax    As Long

    Call Eiyo01_000init
    Call Eiyo930Screen_Hold     '画面抑止ほか
    
    For i1 = 0 To UBound(Fld_Adrs1)
        FldItem = Split(Fld_Adrs1(i1), ",")
        If FldItem(0) = "Gyyyy" Or _
           FldItem(0) = "Age  " Or _
           IsEmpty(Range(Trim(FldItem(0)))) Then
        Else
           Range(Trim(FldItem(0))) = Empty
        End If
    Next i1

    Call Eiyo01_820操作ガイド
    Lmax = Sheets("摂食").UsedRange.Rows.Count
    If Lmax > 4 Then: Sheets("摂食").Rows("5:" & Lmax).Delete Shift:=xlUp
    Call Eiyo99_指定シート削除("検証")
    Call Eiyo99_指定シート削除("検証2")
    Call Eiyo99_指定シート削除("DBmirror")
    Call Eiyo99_指定シート削除("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ")
    Range("Fcode").Select
    Call Eiyo940Screen_Start    '画面描画ほか
End Function
'--------------------------------------------------------------------------------
'   01_100 検索_Click
'       基礎情報の検索、特定された場合に摂食情報も取得する
'--------------------------------------------------------------------------------
Function Eiyo01_100検索Click()
Dim FldItem     As Variant
Dim i1          As Long

    Call Eiyo930Screen_Hold     '画面抑止ほか
    Call Eiyo01_000init
    Range("Gmesg") = Empty
    For i1 = 1 To UBound(Fld_Adrs1)
        FldItem = Split(Fld_Adrs1(i1), ",")
        If FldItem(2) = "i" And Range(Trim(FldItem(0))) <> Empty Then: Exit For
    Next i1
    
    If FldItem(0) = "Qtony" Or _
       FldItem(0) = "Qkoke" Then
        i1 = i1 + 1
        FldItem = Split(Fld_Adrs1(i1), ",")
        If FldItem(0) = "Qill1" Then
            Range(Trim(FldItem(0))) = "000321"
        Else
            Range(Trim(FldItem(0))) = "000313"
        End If
    End If
    
    If i1 > UBound(Fld_Adrs1) Then
        Range("Gmesg") = "検索キーがありません"
    Else
        Call Eiyo01_110検索(i1)
    End If
    
    If IsEmpty(Range("Fcode")) = False And _
       Range("Fcode") = Range("Fsave") Then         '特定された場合は摂食情報
        Application.ScreenUpdating = False          '画面描画抑止
        Call Eiyo01_130MealGet
        Sheets("基礎").Select
    End If
    Range("Fcode").Select
    Call Eiyo940Screen_Start                        '画面描画ほか
End Function
'--------------------------------------------------------------------------------
'   01_110 ＤＢ検索処理     F-024
'--------------------------------------------------------------------------------
Function Eiyo01_110検索(i1 As Long)
Dim mySqlStr    As String
Dim i2          As Long
Dim in_key      As String
Dim Wtbl        As String
Dim FldItem     As Variant
Dim Wtext       As String
Dim FldName     As String

    Columns("ah:hz").Delete Shift:=xlToLeft
    FldItem = Split(Fld_Adrs1(i1), ",")
    Range("Fsave") = Empty
        
    'SQLで読み込むデータを指定する
    in_key = Range(Trim(FldItem(0))).Text
    If Left(in_key, 1) = "%" Then: in_key = "%" & Right(in_key, Len(in_key) - 1)
    Call Eiyo91DB_Open      'DB Open
    If FldItem(0) = "Fcode" Then
        mySqlStr = "SELECT * FROM " & Tbl_Kiso & " Where Fcode = """ & in_key & """"
    Else
        mySqlStr = "SELECT * FROM " & Tbl_Kiso & " Where " & _
                   Trim(FldItem(0)) & " like """ & in_key & "%"""
    End If
    Set Rst_Kiso = myCon.Execute(mySqlStr)
    If Rst_Kiso.EOF Then
        Range("Gmesg") = "該当データはありません"
        Range("Fcode").Select
    Else
        With Rst_Kiso
            Range("Ah2").CopyFromRecordset Rst_Kiso           'レコード
            If Range("Ah3") = Empty Then                        '該当が１件のとき
                For i1 = 1 To UBound(Fld_Adrs1)                 '画面項目の順次処理
                    FldItem = Split(Fld_Adrs1(i1), ",")
                    If FldItem(3) = "D" Then
                        For i2 = 0 To .Fields.Count - 1             'フィールド名
                            If .Fields(i2).Name = Trim(FldItem(0)) And _
                               .Fields(i2).Name <> "Age" Then
                                Range(Trim(FldItem(0))) = Range("ah2").Offset(0, i2)
                                Exit For
                            End If
                        Next i2
                    End If
                Next i1
                If Range("Qill1") = "000000" Then
                    Range("Qtony") = "0"
                Else
                    Range("Qtony") = "1"
                End If
                If Range("Qill2") = "000000" Then
                    Range("Qkoke") = "0"
                Else
                    Range("Qkoke") = "1"
                End If
                If Len(Range("Q6r_a")) = 50 Then
                    Wtext = Range("Q6r_a")
                    Range("Q6r_a") = Left(Wtext, 10)
                    Range("Q6r_b") = Mid(Wtext, 11, 10)
                    Range("Q6r_c") = Mid(Wtext, 21, 10)
                    Range("Q6r_d") = Mid(Wtext, 31, 10)
                    Range("Q6r_e") = Mid(Wtext, 41, 10)
                End If
                Range("Gare1") = Eiyo01_120�n��("1" & Range("area1"))
                Range("Gare2") = Eiyo01_120�n��("2" & Range("area2"))
                Range("Fsave") = Range("Fcode")
                Call Eiyo01_820操作ガイド
            Else
                For i1 = 1 To .Fields.Count                     'フィールド名
                    Cells(1, i1 + 33).Value = .Fields(i1 - 1).Name
                Next
                Columns("ah:hz").EntireColumn.AutoFit           '幅
                i1 = Range("ah1").End(xlDown).Row
                Range("Ah2:ah" & i1).Locked = False             '入力可
                Range("Ah2:ah" & i1).Interior.ColorIndex = 34
            End If
            .Close
        End With
    End If
    Set Rst_Kiso = Nothing                        'オブジェクトの解放
    Call Eiyo920DB_Close    'DB Close
End Function
'--------------------------------------------------------------------------------
'   01_120 地域・都道府県表示
'--------------------------------------------------------------------------------
Function Eiyo01_120地域(in_code As String) As String
Dim i1      As Long
Dim Witem   As Variant

    Eiyo01_120地域 = Empty
    For i1 = 0 To UBound(Fld_Area)
        Witem = Split(Fld_Area(i1), ",")
        If Witem(0) = in_code Then
            Eiyo01_120地域 = Witem(1)
            Exit For
        End If
    Next i1
End Function
'--------------------------------------------------------------------------------
'   01_130　摂食取得
'--------------------------------------------------------------------------------
Function Eiyo01_130MealGet()
Dim mySqlStr    As String
Dim Lmax        As Long
Dim i1          As Long

    Sheets("摂食").Select
    Application.EnableEvents = False                'イベント発生抑止
'    ActiveSheet.Unprotect                           'シートの保護を解除
    Range("b1") = Empty
    Range("a2") = Range("Fcode") & ":" & Range("Namej")
    Lmax = ActiveSheet.UsedRange.Rows.Count
    If Lmax > 4 Then: Rows("5:" & Lmax).Delete Shift:=xlUp
        
    'SQLで読み込むデータを指定する
    Call Eiyo91DB_Open      'DB Open
    mySqlStr = "SELECT Sdate,Ekubn,Foodc,Suryo FROM " & Tbl_Meal & " Where Fcode = """ & Range("Fcode") & """"
    Set Rst_Meal = myCon.Execute(mySqlStr)
    If Rst_Meal.EOF Then
        Lmax = 0
    Else
        Range("A5").CopyFromRecordset Rst_Meal           'レコード
    End If
    
    Lmax = ActiveSheet.UsedRange.Rows.Count
    For i1 = 5 To Lmax
        Cells(i1, 6) = Cells(i1, 4)
        Cells(i1, 4) = Cells(i1, 3)
        Cells(i1, 3) = Eiyo01_401食事区分(Cells(i1, 2))
        Call Eiyo01_402食品マスタ(i1)
    Next i1
    Set Rst_Meal = Nothing                    'オブジェクトの解放
    Call Eiyo920DB_Close    'DB Close
End Function
'--------------------------------------------------------------------------------
'   01_200 更新_Click
'--------------------------------------------------------------------------------
Function Eiyo01_200更新Click()
Dim Rtn As Long
    Call Eiyo930Screen_Hold                     '画面抑止ほか
    Call Eiyo01_000init
    Rtn = Eiyo01_210KeyCheck                    'キーチェック
    If Rtn = 0 Then: Rtn = Eiyo01_220項目Check  '項目チェック
    If Rtn = 0 Then: Rtn = Eiyo01_230DB更新     'DB更新
    Call Eiyo940Screen_Start                    '画面描画ほか
End Function
'--------------------------------------------------------------------------------
'   01_210 キーチェック
'--------------------------------------------------------------------------------
Function Eiyo01_210KeyCheck() As Long
Dim mySqlStr    As String

    Call Eiyo91DB_Open      'DB Open
    mySqlStr = "SELECT * FROM " & Tbl_Kiso & " Where Fcode = """ & Range("Fcode") & """"
    Set Rst_Kiso = myCon.Execute(mySqlStr)
    If Rst_Kiso.EOF Then
        If Range("Fcode") = Range("Fsave") Then
            Range("Gmesg") = "Program Error Non Key & Save Key Same"    '×：キーなし、Save同じ
            Eiyo01_210KeyCheck = 1
        Else
            Eiyo01_210KeyCheck = 0                                      '○：キーなし、Save異なる(新規)
        End If
    Else
        If Range("Fcode") = Range("Fsave") Then
            Eiyo01_210KeyCheck = 0                                      '○：キーあり、Save同じ(更新)
        Else
            Range("Gmesg") = "コードが重複しています"                   '×：キーあり、Save異なる
            Eiyo01_210KeyCheck = 1
        End If
    End If
    Set Rst_Kiso = Nothing                        'オブジェクトの解放
    Call Eiyo920DB_Close    'DB Close
End Function
'--------------------------------------------------------------------------------
'   01_220 項目チェック
'--------------------------------------------------------------------------------
Function Eiyo01_220項目Check() As Long
Dim Witem   As Variant
Dim Wlen    As Long
Dim i1      As Long
Dim Wtemp   As String

    Eiyo01_220項目Check = 1
    Range("Gmesg") = Empty
'   コード
'    Witem = Range("Fcode")
'    If IsNumeric(Witem) = True And Len(Witem) <= 10 Then
'    Else
'        Range("Gmesg") = "コードは１０桁以内の数字にしてください" & Len(Witem)
'        Range("Fcode").Activate
'        Exit Function
'    End If
'   調査期間開始日
    If IsDate(Range("Date1")) Then
    Else
        Range("Gmesg") = "調査期間開始日を実在日にしてください"
        Range("Date1").Activate
        Exit Function
    End If
'   調査期間日数
    Witem = Range("Nissu")
    If IsNumeric(Witem) = True And Len(Witem) = 1 Then
    Else
        Range("Gmesg") = "調査期間日数は１桁の数字にしてください"
        Range("Nissu").Activate
        Exit Function
    End If
'   氏名
    If Eiyo01_221����check("Namej", "����", 10) = 1 Then: Exit Function
'   性別
    Witem = Range("sex")
    If Witem = "" Or Witem = "0" Or Witem = "1" Then
    Else
        Range("Gmesg") = "性別は１桁の数字にしてください"
        Range("sex").Activate
        Exit Function
    End If
'   生年月日
    If IsDate(Range("Birth")) Then
    Else
        Range("Gmesg") = "生年月日を実在日にしてください"
        Range("Birth").Activate
        Exit Function
    End If
    
    If Eiyo01_223数値lcheck("Hight", "身長", 3, 1, 300) = 1 Then: Exit Function
    If Eiyo01_223数値lcheck("Weght", "体重", 3, 1, 300) = 1 Then: Exit Function
    If Eiyo01_223数値lcheck("Sibou", "皮下脂肪", 2, 1, 50) = 1 Then: Exit Function
    
    If Eiyo01_221桁数check("Adrno", "郵便番号", 18) = 1 Then: Exit Function
    If Eiyo01_221桁数check("Adrs1", "住所ー１", 18) = 1 Then: Exit Function
    If Eiyo01_221桁数check("Adrs2", "住所ー２", 18) = 1 Then: Exit Function
'   地区・地域
    Wtemp = Left(Range("adrs1"), 2)
    For i1 = 0 To UBound(Fld_Area)
        Witem = Split(Fld_Area(i1), ",")
        If Left(Witem(0), 1) = "2" And _
           Left(Witem(1), 2) = Wtemp Then
            Range("Area1") = Right(Witem(2), 2)
            Range("Gare1") = Eiyo01_120地域("1" & Range("Area1"))
            Range("Area2") = Right(Witem(0), 2)
            Range("Gare2") = Witem(1)
            Exit For
        End If
    Next i1
    If Eiyo01_222数字check("Q3rec", "Q3.食習慣", 10) = 1 Then: Exit Function
    If Eiyo01_222数字check("Q4rec", "Q4.休養", 5) = 1 Then: Exit Function
    If Eiyo01_222数字check("Q5rec", "Q5.運動", 3) = 1 Then: Exit Function
    If Eiyo01_222数字check("Q6r_a", "Q6.健康１", 10) = 1 Then: Exit Function
    If Eiyo01_222数字check("Q6r_b", "Q6.健康２", 10) = 1 Then: Exit Function
    If Eiyo01_222数字check("Q6r_c", "Q6.健康３", 10) = 1 Then: Exit Function
    If Eiyo01_222数字check("Q6r_d", "Q6.健康４", 10) = 1 Then: Exit Function
    If Eiyo01_222数字check("Q6r_e", "Q6.健康５", 10) = 1 Then: Exit Function
'   職業
    Range("Qjob1") = UCase(Range("Qjob1"))
    If Len(Range("Qjob1")) = 4 Then
    Else
        Range("Gmesg") = "職業は４桁としてください　" & Len(Range("Qjob1"))
        Range("Qjob1").Activate
    End If

    If Eiyo01_222数字check("Qsyuf", "QA.主婦", 1) = 1 Then: Exit Function
    If Eiyo01_222数字check("Qcnd1", "QB.妊娠", 1) = 1 Then: Exit Function
    If Eiyo01_222数字check("Qtony", "QC.糖尿", 1) = 1 Then: Exit Function
    If Range("Qtony") = "0" Then
        Range("Qill1") = "000000"
    Else
        Range("Qill1") = "000321"
    End If
    If Eiyo01_222数字check("Qkoke", "QC.糖尿", 1) = 1 Then: Exit Function
    If Range("Qkoke") = "0" Then
        Range("Qill2") = "000000"
    Else
        Range("Qill2") = "000313"
    End If
    If Eiyo01_223数値lcheck("Qsrmr", "QE.ｽﾎﾟｰﾂ1", 3, 0, 1000) = 1 Then: Exit Function
    If Eiyo01_223数値lcheck("Qsmin", "QE.ｽﾎﾟｰﾂ2", 3, 0, 1000) = 1 Then: Exit Function
    If Eiyo01_222数字check("Qclab", "QF.運動部", 1) = 1 Then: Exit Function
    If Eiyo01_222数字check("Qclab", "Q .喫煙", 1) = 1 Then: Exit Function
    If Eiyo01_223数値lcheck("Qsyog", "QG.身障害", 2, 0, 100) = 1 Then: Exit Function
    If Eiyo01_222数字check("Qwcnt", "QG.ｳｴｲﾄCT", 1) = 1 Then: Exit Function
    If Eiyo01_222数字check("Tenes", "ｴﾈﾙｷﾞ指定", 1) = 1 Then: Exit Function
    If Eiyo01_222数字check("Tanps", "ﾀﾝﾊﾟｸ指定", 1) = 1 Then: Exit Function
    If Eiyo01_223数値lcheck("Tenee", "ｴﾈﾙｷﾞ指定", 5, 2, 100000) = 1 Then: Exit Function
    If Eiyo01_223数値lcheck("Tanpe", "ﾀﾝﾊﾟｸ指定", 5, 2, 100000) = 1 Then: Exit Function
'
    Range("Blood") = UCase(Range("Blood"))
    Wtemp = Range("Blood")
    If Wtemp = "" Or Wtemp = "A" Or Wtemp = "B" Or Wtemp = "O" Or Wtemp = "AB" Then
    Else
        Range("Gmesg") = "血液型が不正です"
        Range("Blood").Activate
        Exit Function
    End If

    If Eiyo01_221桁数check("Bscd1", "支店", 3) = 1 Then: Exit Function
    If Eiyo01_221桁数check("Bscd2", "支部", 2) = 1 Then: Exit Function
    If Eiyo01_221桁数check("Bhok1", "保険証記号", 8) = 1 Then: Exit Function
    If Eiyo01_221桁数check("Bhok2", "保険証No", 8) = 1 Then: Exit Function
    If Eiyo01_221桁数check("Bhant", "定期健診", 2) = 1 Then: Exit Function
'
    Range("Barm") = UCase(Range("Barm"))
    Wtemp = Range("Barm")
    If Wtemp = "" Or Wtemp = "L" Or Wtemp = "R" Then
    Else
        Range("Gmesg") = "検査腕が不正です"
        Range("Barm").Activate
        Exit Function
    End If
'   血液検査日
    If IsEmpty(Range("Bdate")) Or IsDate(Range("Bdate")) Then
    Else
        Range("Gmesg") = "血液検査日を実在日にしてください"
        Range("Bdate").Activate
        Exit Function
    End If
    If Eiyo01_223数値check("Bbl01", "赤血球数", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223数値check("Bbl02", "血色素量", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223数値check("Bbl03", "ﾍﾏﾄｸﾘｯﾄ", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223数値check("Bbl04", "ｺﾚｽﾃﾛｰﾙ", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223数値check("Bbl05", "HDL", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223数値check("Bbl06", "中性脂肪", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223数値check("Bbl07", "G.O.T.", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223数値check("Bbl08", "G.P.T.", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223数値check("Bbl09", "尿酸", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223数値check("Bbl10", "血糖", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223数値check("Bbl11", "血圧最高", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223数値check("Bbl12", "血圧最低", 3, 1, 10000) = 1 Then: Exit Function
    Eiyo01_220項目Check = 0
End Function
'--------------------------------------------------------------------------------
'   01_221 桁数チェック
'--------------------------------------------------------------------------------
Function Eiyo01_221桁数check(Ifld As String, Iname As String, Ilen As Long) As Long
    If Len(Range(Ifld)) > Ilen Then
        Range("Gmesg") = Iname & "は" & Ilen & "桁以内にしてください" & Len(Range(Ifld))
        Range(Ifld).Activate
        Eiyo01_221桁数check = 1
    Else
        Eiyo01_221桁数check = 0
    End If
End Function
'--------------------------------------------------------------------------------
'   01_222 固定桁数字項目チェック
'--------------------------------------------------------------------------------
Function Eiyo01_222数字check(Ifld As String, Iname As String, Ilen As Long) As Long
Dim Witem   As Variant
Dim Wlen    As Long

    If Range(Ifld) = Empty Then: Range(Ifld) = String(Ilen, "0")
    Witem = Range(Ifld)
    Wlen = Len(Witem)
    If IsNumeric(Witem) And Wlen = Ilen Then
        Eiyo01_222数字check = 0
    Else
        Range("Gmesg") = Iname & "は" & Ilen & "桁の数字にしてください　" & Wlen
        Range(Ifld).Activate
        Eiyo01_222数字check = 1
    End If
End Function
'--------------------------------------------------------------------------------
'   01_223 数値項目チェック
'--------------------------------------------------------------------------------
Function Eiyo01_223数値check(Ifld As String, Iname As String, _
                              Ilen1 As Long, Ilen2 As Long, Imax As Long) As Long
Dim Witem   As Variant
    
    Witem = Range(Ifld)
    If IsNumeric(Witem) And Witem < Imax Then
        Eiyo01_223数値lcheck = 0
    Else
        Range("Gmesg") = Iname & "は上" & Ilen1 & "桁下" & Ilen2 & "桁以内の数値にしてください"
        Range(Ifld).Activate
        Eiyo01_223数値check = 1
    End If
End Function
'--------------------------------------------------------------------------------
'   01_230 ＤＢ更新                                     F-026
'   Microsoft ActiveX Data Objects 2.X Library 参照設定
'--------------------------------------------------------------------------------
Function Eiyo01_230DB更新() As Long
Dim FldItem     As Variant
Dim FldName     As String
Dim i1          As Long

    Call Eiyo91DB_Open      'DB Open
    '準備ここまで
    With Rst_Kiso
        'インデックスの設定
        .Index = "PrimaryKey"
        'レコードセットを開く
        Rst_Kiso.Open Source:=Tbl_Kiso, ActiveConnection:=myCon, _
            CursorType:=adOpenKeyset, LockType:=adLockOptimistic, _
            Options:=adCmdTableDirect
        '番号が登録されているか検索する
        If Not .EOF Then .Seek Range("Fcode")
        If .EOF Then
            .AddNew
            Range("Gmesg") = "追加登録されました。"
            Range("Fsave") = Range("Fcode")
        Else
            Range("Gmesg") = "更新されました。"
        End If
        For i1 = 1 To UBound(Fld_Adrs1)                 '画面項目の順次処理
            FldItem = Split(Fld_Adrs1(i1), ",")         '
            If FldItem(3) = "D" Then
                FldName = Trim(FldItem(0))
                .Fields(FldName).Value = Range(FldName).Value
            End If
        Next i1
        .Update
        .Close
    End With
    Set Rst_Kiso = Nothing      'オブジェクトの解放
    Call Eiyo920DB_Close        'DB Close
    Eiyo01_230DB更新 = 0
End Function
'--------------------------------------------------------------------------------
'   01_300　取消_Click
'--------------------------------------------------------------------------------
Function Eiyo01_300取消Click()
    If Range("Fcode") = Range("Fsave") And _
        IsEmpty(Range("Fcode")) = False Then
        Call Eiyo91DB_Open      'DB Open
        myCon.Execute "DELETE FROM " & Tbl_Kiso & " Where Fcode = """ & Range("Fcode") & """"
        myCon.Execute "DELETE FROM " & Tbl_Meal & " Where Fcode = """ & Range("Fcode") & """"
        Range("Gmesg") = "取消削除されました。"
        Range("Fsave") = Empty
        Call Eiyo920DB_Close    'DB Close
    Else
        Range("Gmesg") = "検索されていません。"
    End If
End Function
'--------------------------------------------------------------------------------
'   01_400　摂食表示
'--------------------------------------------------------------------------------
Function Eiyo01_400MealDisp()
Dim Rtn     As Long
Dim Wmsg    As String
    
    Range("a2") = Range("Fcode") & ":" & Range("Namej")
    Wmsg = "基礎情報の検索が行われていません"
    If IsEmpty(Range("Fcode")) Or Range("Fcode") <> Range("Fsave") Then
        Rtn = CreateObject("WScript.Shell").Popup(Wmsg, 3, "Microsoft Excel", 0)
        Sheets("基礎").Select
    End If
End Function
'--------------------------------------------------------------------------------
'   01_401　食事区分
'--------------------------------------------------------------------------------
Function Eiyo01_401食事区分(kbn As Long) As String
    Select Case kbn
        Case 1: Eiyo01_401食事区分 = "朝"
        Case 2: Eiyo01_401食事区分 = "昼"
        Case 3: Eiyo01_401食事区分 = "夕"
        Case 4: Eiyo01_401食事区分 = "夜"
        Case 5: Eiyo01_401食事区分 = "間"
        Case Else: Eiyo01_401食事区分 = Empty
    End Select
End Function
'--------------------------------------------------------------------------------
'   01_402　食品マスタ取得
'--------------------------------------------------------------------------------
Function Eiyo01_402食品マスタ(in_line As Long)
Dim mySqlStr    As String
    If IsEmpty(Cells(in_line, 4)) Then
        Cells(in_line, 5) = Empty
        Range("g" & in_line & ":z" & in_line) = Empty
        Exit Function
    End If
        
    mySqlStr = "SELECT * FROM " & Tbl_Food & " Where Foodc = " & Cells(in_line, 4)
    Set Rst_Food = myCon.Execute(mySqlStr)
    If Rst_Food.EOF Then
        Cells(in_line, 5) = "キーなし"
        Range("g" & in_line & ":z" & in_line) = Empty
    Else
        Cells(in_line, 7).CopyFromRecordset Rst_Food
        Cells(in_line, 5) = Cells(in_line, 8)
    End If
    Rst_Food.Close
    Set Rst_Food = Nothing
End Function
'--------------------------------------------------------------------------------
'   01_410　摂食画面が変更された
'--------------------------------------------------------------------------------
Function Eiyo01_410MealChange(ChangeCell As String)
Dim Wl      As Long
Dim Wc      As Long

    Wl = Range(ChangeCell).Row
    Wc = Range(ChangeCell).Column
'    ActiveSheet.Unprotect                           'シートの保護を解除
    Select Case Wc
        Case 1: Cells(Wl, 2).Select
        Case 2
            Cells(Wl, 3) = Eiyo01_401食事区分(Cells(Wl, 2))
            Cells(Wl, 4).Select
        Case 4
            'SQLで読み込むデータを指定する
            Call Eiyo91DB_Open      'DB Open
            Call Eiyo01_402食品マスタ(Wl)
            Call Eiyo920DB_Close    'DB Close
            If Wl > 5 Then
                If IsEmpty(Cells(Wl, 1)) Then: Cells(Wl, 1) = Cells(Wl - 1, 1)
                If IsEmpty(Cells(Wl, 2)) Then
                    Cells(Wl, 2) = Cells(Wl - 1, 2)
                    Cells(Wl, 3) = Cells(Wl - 1, 3)
                End If
            End If
            Cells(Wl, 6).Select
        Case 6: Cells(Wl + 1, 4).Select
    End Select
'    ActiveSheet.Protect UserInterfaceOnly:=True     '保護を有効にする
End Function
'--------------------------------------------------------------------------------
'   01_420　メニューが選択された
'--------------------------------------------------------------------------------
Function Eiyo01_420Menu(in_cell As String)
Dim Wcode   As String
Dim Wname   As String
Dim Rtn     As Long
Dim Wmsg    As String
Dim Wcell   As String

    Wcode = Range(in_cell).Offset(0, 10)
    If Wcode = "" Then: Exit Function
    Wname = Range(in_cell)
    If IsEmpty(Sheets("摂食").Range("b1")) Then
        Wmsg = Wcode & ":" & Wname & "が選択されました。"
        Rtn = CreateObject("WScript.Shell").Popup(Wmsg, 3, "Microsoft Excel", 0)
    Else
        Application.EnableEvents = False            'イベント発生抑止
        Sheets("摂食").Select
        Wcell = Range("b1")
        Range("b1") = Empty
        Application.EnableEvents = True             'イベント発生再開
        Range(Wcell) = Wcode
    End If
End Function
'--------------------------------------------------------------------------------
'   01_500　登録 Click
'--------------------------------------------------------------------------------
Function Eiyo01_500MealCalc(in_Func As Long)
    Call Eiyo930Screen_Hold                                 '画面抑止ほか
    Call Eiyo91DB_Open                                      'DB Open
    If Eiyo01_501MealEntry = 1 Then: GoTo Eiyo01_503Exit    '未入力チェック
    If Eiyo01_502Mealscope = 1 Then: GoTo Eiyo01_503Exit    '摂食量の範囲チェック
    If Eiyo01_503Mealzerod = 1 Then: GoTo Eiyo01_503Exit    '摂食量ゼロの削除
    If Eiyo01_504MealDoubl = 1 Then: GoTo Eiyo01_503Exit    '摂食の重複入力
    If Eiyo01_510MealUdate = 1 Then: GoTo Eiyo01_503Exit    'ＤＢ更新
    If Eiyo01_511MealFldgt = 1 Then: GoTo Eiyo01_501Exit    '項目要素取得
    If Eiyo01_512MealSheet = 1 Then: GoTo Eiyo01_501Exit    '摂食計算シート
    If Eiyo01_513kenso2sht = 1 Then: GoTo Eiyo01_501Exit    '検証２シート作成
    If Eiyo01_514MealCalc1 = 1 Then: GoTo Eiyo01_501Exit    '摂食計算
    If Eiyo01_515MealTotal = 1 Then: GoTo Eiyo01_501Exit    '摂食量合計
    If Eiyo01_521CalcDbGet(1) = 1 Then: GoTo Eiyo01_501Exit '摂食量合計
    If Eiyo01_522Mealcalc2 = 1 Then: GoTo Eiyo01_501Exit    '標準体重ほか
    If Eiyo01_525MealDiffe = 1 Then: GoTo Eiyo01_501Exit    '過不足アドバイス
    If Eiyo01_528Eiyohirit = 1 Then: GoTo Eiyo01_501Exit    '栄養比率
    If in_Func = 2 Then
        If Eiyo01_540Old_Check = 1 Then: GoTo Eiyo01_501Exit    '旧計算値
'    Else
'        Call Eiyo99_指定シート削除("DBmirror")
    End If
Eiyo01_501Exit:
    Call Eiyo01_550RstClose
Eiyo01_503Exit:
    Call Eiyo920DB_Close                'DB Close
    Sheets("基礎").Select
    Call Eiyo940Screen_Start            '画面描画ほか
End Function
'--------------------------------------------------------------------------------
'   01_501　摂食情報未入力チェック
'--------------------------------------------------------------------------------
Function Eiyo01_501MealEntry() As Long
Dim Lmax    As Long
Dim i1      As Long
Dim Wnon    As Long

    Eiyo01_501MealEntry = 1
    Lmax = ActiveSheet.UsedRange.Rows.Count
    If Lmax < 5 Then
        MsgBox "データがありません"
        Exit Function
    End If
        
    Wnon = 0
    Range("a5:b" & Lmax).Interior.ColorIndex = xlNone
    Range("d5:d" & Lmax).Interior.ColorIndex = xlNone
    Range("f5:f" & Lmax).Interior.ColorIndex = xlNone
    For i1 = 5 To Lmax
        If Cells(i1, 6) <> 0 Then
            If IsDate(Cells(i1, 1)) = False Or _
               Cells(i1, 1) < Range("Date1") Or _
               Cells(i1, 1) > Range("Date1") + Range("Nissu") - 1 Then
                Cells(i1, 1).Interior.ColorIndex = 6
                Exit For
            ElseIf IsEmpty(Cells(i1, 2)) Then
                Cells(i1, 2).Interior.ColorIndex = 6
                Exit For
            ElseIf IsEmpty(Cells(i1, 4)) Then
                Cells(i1, 4).Interior.ColorIndex = 6
                Exit For
            End If
        End If
    Next i1
    If i1 <= Lmax Then
        MsgBox ("誤りの項目を修正してください。")
    Else
        Eiyo01_501MealEntry = 0
    End If
End Function
'--------------------------------------------------------------------------------
'   01_502　摂食量の範囲チェック
'--------------------------------------------------------------------------------
Function Eiyo01_502Mealscope() As Long
Dim Lmax    As Long
Dim i1      As Long
Dim Wover   As Long
Dim Wmsg    As String

    Eiyo01_502Mealscope = 1
    Lmax = ActiveSheet.UsedRange.Rows.Count
    Wover = 0
    For i1 = 5 To Lmax
        If Cells(i1, 6) <> 0 Then
            If Cells(i1, 6) < Cells(i1, 17) Or _
               Cells(i1, 6) > Cells(i1, 18) Then
                Wover = Wover + 1
                Cells(i1, 6).Interior.ColorIndex = 6
            Else
                Cells(i1, 6).Interior.ColorIndex = xlNone
            End If
        End If
    Next i1
    If Wover > 0 Then
        Wmsg = "摂食量の異常値が" & Wover & "ヵ所あります"
        i1 = CreateObject("WScript.Shell").Popup(Wmsg, 1, "Microsoft Excel", 0)
        Eiyo01_502Mealscope = 0
    End If
    Eiyo01_502Mealscope = 0
End Function
'--------------------------------------------------------------------------------
'   01_503　摂食量ゼロの削除
'--------------------------------------------------------------------------------
Function Eiyo01_503Mealzerod() As Long
Dim Lmax    As Long
Dim i1      As Long
Dim Wzero   As Long
Dim Wmsg    As String

    Lmax = ActiveSheet.UsedRange.Rows.Count
    Wzero = 0
    For i1 = 5 To Lmax
        If Cells(i1, 6) = 0 Then
            Wzero = Wzero + 1
            Rows(i1).Delete Shift:=xlUp
            Lmax = Lmax - 1
        End If
    Next i1
    If Wzero > 0 Then
        Wmsg = "摂食量ゼロの" & Wzero / 2 & "行を削除しました。"
        i1 = CreateObject("WScript.Shell").Popup(Wmsg, 1, "Microsoft Excel", 0)
    End If
    Eiyo01_503Mealzerod = 0
End Function
'--------------------------------------------------------------------------------
'   01_504　摂食の重複チェック
'--------------------------------------------------------------------------------
Function Eiyo01_504MealDoubl() As Long
Dim Lmax    As Long
Dim i1      As Long

    Eiyo01_504MealDoubl = 1
    Lmax = ActiveSheet.UsedRange.Rows.Count
    For i1 = 5 To Lmax
        Cells(i1, 2) = Val(Cells(i1, 2))
    Next i1
    Rows("5:" & Lmax).Sort key1:=Range("A5"), order1:=xlAscending, _
                           key2:=Range("B5"), order2:=xlAscending, _
                           key3:=Range("D5"), order3:=xlAscending, Header:=xlNo
    i1 = 6
    Do Until IsEmpty(Cells(i1, 4))
        If Cells(i1 - 1, 1) & Cells(i1 - 1, 2) & Cells(i1 - 1, 4) = _
            Cells(i1, 1) & Cells(i1, 2) & Cells(i1, 4) Then
            Cells(i1 - 1, 6) = Cells(i1 - 1, 6) + Cells(i1, 6)
            Rows(i1).Delete Shift:=xlUp
        Else
            i1 = i1 + 1
        End If
    Loop
    Eiyo01_504MealDoubl = 0
End Function
'--------------------------------------------------------------------------------
'   01_510　摂食DB登録
'--------------------------------------------------------------------------------
Function Eiyo01_510MealUdate() As Long
Dim Lmax    As Long
Dim i1      As Long
Dim Wkey    As Variant

    Lmax = Range("a4").End(xlDown).Row
    myCon.Execute "DELETE FROM " & Tbl_Meal & " Where Fcode = """ & Range("Fcode") & """"
    '準備ここまで
    With Rst_Meal
        'インデックスの設定
        .Index = "PrimaryKey"
        'レコードセットを開く
        Rst_Meal.Open Source:=Tbl_Meal, ActiveConnection:=myCon, _
            CursorType:=adOpenKeyset, LockType:=adLockOptimistic, _
            Options:=adCmdTableDirect
        For i1 = 5 To Lmax
        '番号が登録されているか検索する
            Wkey = Array(Range("Fcode"), Cells(i1, 1), Cells(i1, 2), Cells(i1, 4))
            If Not .EOF Then .Seek Wkey
            If .EOF Then: .AddNew
            .Fields(0).Value = Range("Fcode").Value
            .Fields(1).Value = Cells(i1, 1).Value
            .Fields(2).Value = Cells(i1, 2).Value
            .Fields(3).Value = Cells(i1, 4).Value
            .Fields(4).Value = Cells(i1, 6).Value
            .Fields(5).Value = Cells(i1, 16).Value
            .Update
        Next i1
        .Close
    End With
    Set Rst_Meal = Nothing                    'オブジェクトの解放
    Eiyo01_510MealUdate = 0
End Function
'--------------------------------------------------------------------------------
'   01_511　栄養素項目の各種情報取得   F-018
'--------------------------------------------------------------------------------
Function Eiyo01_511MealFldgt() As Long
    
    Sheets.Add After:=Sheets(Sheets.Count)      'シート追加
    'レコードセットを開く
    Rst_Field.Open Source:=Tbl_Field, _
                ActiveConnection:=myCon, _
                CursorType:=adOpenForwardOnly, _
                LockType:=adLockReadOnly, _
                Options:=adCmdTableDirect
    'レコード
    Range("a1").CopyFromRecordset Rst_Field
    Fld_Field = ActiveSheet.UsedRange
    Rst_Field.Close
    Set Rst_Field = Nothing    'オブジェクトの解放
    Application.DisplayAlerts = False               '確認抑止
    ActiveSheet.Delete
    Application.DisplayAlerts = True                '確認復活
    Eiyo01_511MealFldgt = 0
End Function
'--------------------------------------------------------------------------------
'   01_512　摂食計算シート作成
'--------------------------------------------------------------------------------
Function Eiyo01_512MealSheet() As Long
Dim i1      As Long     '行Index
Dim i2      As Long     '欄Index
Dim Wno     As String
Dim Wtext   As String

    Call Eiyo99_指定シート削除("検証")
    Sheets.Add After:=Sheets(Sheets.Count)      'シート追加
    ActiveSheet.Name = "検証"
    Range("d1") = "栄養計算　摂食検証"
    Range("a2") = Sheets("摂食").Range("a2")
    Wtext = Empty
    For i1 = 1 To 27                            '栄養素
        Wno = Format(i1, "00")
        Wtext = Wtext & "摂取量" & Wno & vbTab
        Wtext = Wtext & "熱損後" & Wno & vbTab
    Next i1
    For i1 = 1 To 15
        Wtext = Wtext & "ｴﾈﾙｷﾞC" & Format(i1, "00") & vbTab
    Next i1
    For i1 = 1 To 15
        Wtext = Wtext & "ｴﾈﾙｷﾞW" & Format(i1, "00") & vbTab
    Next i1
    For i1 = 1 To 15
        Wtext = Wtext & "ｶﾙｼｳﾑ1" & Format(i1, "00") & vbTab
    Next i1
    For i1 = 1 To 15
        Wtext = Wtext & "ｶﾙｼｳﾑ2" & Format(i1, "00") & vbTab
    Next i1
    Wtext = Wtext & "脂質動物" & vbTab
    Wtext = Wtext & "脂質魚介" & vbTab
    Wtext = Wtext & "脂質植物" & vbTab
    Wtext = Wtext & "熱損動物" & vbTab
    Wtext = Wtext & "熱損魚介" & vbTab
    Wtext = Wtext & "熱損植物"
    Range("a4:dp4") = Split(Wtext, vbTab)
    ActiveWindow.FreezePanes = False        'ウインド枠固定の解除
    Range("a5").Select
    ActiveWindow.FreezePanes = True         'ウインド枠固定の設定
    Cells.NumberFormatLocal = "#,##0.00;[赤]-#,##0.00"
    Eiyo01_512MealSheet = 0
End Function
'--------------------------------------------------------------------------------
'   01_513　検証２シート作成
'--------------------------------------------------------------------------------
Function Eiyo01_513kenso2sht() As Long
Dim i1      As Long
    Call Eiyo99_指定シート削除("検証2")
    Sheets.Add After:=Sheets(Sheets.Count)      'シート追加
    ActiveSheet.Name = "検証2"
    Cells.Interior.ColorIndex = 36              '全画面背景色
'   表題
    Range("C1:F1").Select
    Selection.MergeCells = True                 '表題セル連結
    Selection.HorizontalAlignment = xlCenter    '表題センタリング
    Selection.Interior.ColorIndex = 37          '表題色（ペールブルー）
    With Selection.Font                         'フォント
        .FontStyle = "太字"
        .Size = 16
    End With
    Range("C1") = "栄養計算　検証資料２"
    
'   栄養素名
    Range("a4") = "No.栄養素名"
    Range("b4") = "単位"
    For i1 = 1 To 27
        Cells(4 + i1, 1) = Format(i1, "00") & "." & Fld_Field(i1, 4)
        Cells(4 + i1, 2) = Fld_Field(i1, 5)
    Next i1
'   摂取量
    Range("c3") = "<========= 熱損後摂取量 ==========>"
    Range("c4") = "総量"
    Range("d4") = "／日"
    Range("e4") = "式"
    Range("f4") = "補正後"
    Range("c4:p4").HorizontalAlignment = xlCenter
    Range("c5:d31,f5:f31").Interior.ColorIndex = xlNone      '白抜き化
    With Range("c5:d31,f5:f31").Borders                      '枠罫線
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .Weight = xlThin
    End With
    Range("c5:d31,f5:f31").NumberFormatLocal = "#,##0.00;[赤]-#,##0.00"
    Range("c5").Name = "ks2_eiyoso"
'   摂取量の補正条件
    Range("e:e,o:o").HorizontalAlignment = xlCenter '横中央
    Range("e15,e20,e24,e27").Interior.ColorIndex = xlNone   '白抜き化
    With Range("e15,e20,e24,e27").Borders                   '枠罫線
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .Weight = xlThin
    End With
    Range("e15").Name = "ks2_hosei11"
    Range("e20").Name = "ks2_hosei16"
    Range("e24").Name = "ks2_hosei20"
    Range("e27").Name = "ks2_hosei23"
'   基礎情報
    Range("i4") = "実体重"
    Range("j4") = "標準体重"
    Range("h5") = "a.体重"
    Range("H6") = "b.体表面積"
    Range("H8") = "基礎代謝"
    Range("h9") = "c.生活指数"
    Range("H10") = "d.コード"
    Range("H11") = "e.面積当り"
    Range("H12") = "f.／日"
    Range("H13") = "g.／分"
    Range("H15") = "エネルギー"
    Range("H16") = "h.標準量"
    Range("H17") = "i.適用条件"
    Range("H18") = "j.ｴﾈﾙｷﾞｰ1"
    Range("H19") = "k.ｴﾈﾙｷﾞｰ2"
    Range("i21") = "f = b * e * 24"
    Range("i22") = "h = f * (1+c) * 1.1"
    Range("F05").Copy Range("I5:j6,i9:i11,i12:j13,i16:j16,i18:i19")
    Range("E15").Copy Range("I10,i17")
    Range("i05").Name = "ks2_weght"     '体重
    Range("i06").Name = "ks2_Aansa"     '体表面積
    Range("i09").Name = "ks2_Aansx"     '生活指数
    Range("i10").Name = "ks2_kisocd"    '基礎コード
    Range("i11").Name = "ks2_kisot"     '面積当り基礎代謝
    Range("i12").Name = "ks2_Aansb"     '基礎代謝／日
    Range("i13").Name = "ks2_Aansc"     '基礎代謝／日
    Range("i16").Name = "ks2_Aansd"     'ｴﾈﾙｷﾞｰ標準量
    Range("i17").Name = "ks2_energ"     '所要量ｴﾈﾙｷﾞｰ条件
'   所要量(ｴﾈﾙｷﾞｰ以外)
    Range("l3") = "所要TBL"
    Range("m3") = "生活強度"
    Range("n3") = "妊娠補正"
    Range("o3") = "式"
    Range("p3") = "所要量"
    Range("q3") = "摘要"
    Range("E15").Copy Range("l4:n4")
    Range("F05").Copy Range("l6:n31")
    Range("l4").Name = "ks2_syoyo"
    Range("F05").Copy Range("p5:p31")

    With ActiveSheet.PageSetup
        .Orientation = xlLandscape                      '横長
'        .PrintHeadings = True                           '行列番号
        .LeftMargin = Application.InchesToPoints(0.4)   '左余白
        .RightMargin = Application.InchesToPoints(0.2)  '右余白
        .Zoom = False
        .FitToPagesWide = 1                             '横１頁
        .FitToPagesTall = 1                             '縦１頁
    End With
    Cells.EntireColumn.AutoFit                          '列幅
    Range("C:D,F:F,I:J,L:N,P:P").ColumnWidth = 10
    Range("G:G,K:K").ColumnWidth = 4
    Range("a01") = Range("Namej")
    Range("a02") = "[" & Range("Fcode") & "]"
    Range("q05") = "ｴﾈﾙｷﾞｰ1"
    Range("q06") = "年令性別ほか"
    Range("q07") = "たんぱく質 1/2"
    Range("q08") = "たんぱく質 1/2"
    Range("q09") = "ｴﾈﾙｷﾞｰ1 & 生活強度"
    Range("q10") = "ｴﾈﾙｷﾞｰ1から計算"
    Range("q11") = "ｴﾈﾙｷﾞｰ1 * 0.0099"
    Range("q12") = "実体重と係数"
    Range("q13") = "ｶﾙｼｳﾑ同値"
    Range("q14") = "年令"
    Range("q15") = "TBL値(高血圧は指定値)"
    Range("q16") = "TBL値"
    Range("q17") = "ｴﾈﾙｷﾞｰ2 * 0.0004"
    Range("q18") = "ｴﾈﾙｷﾞｰ2 * 0.00055"
    Range("q19") = "ｴﾈﾙｷﾞｰ2 * 0.0066"
    Range("q20") = "TBL値"
    Range("q21") = "(TBL値)"
    Range("q22") = "(TBL値)"
    Range("q23") = "(TBL値)"
    Range("q24") = "不飽和脂肪酸 * 0.6"
    Range("q25") = "ﾅﾄﾘｳﾑ同値"
    Range("q26") = "ｶﾙｼｳﾑ 1/2"
    Range("q27") = "高血圧は指定値"
    Range("q28") = "一律"
    Range("q29") = "脂質 66%"
    Range("q30") = "脂質 34%"
    Range("q31") = "糖尿/ｳｴｲﾄ･ｺﾝﾄﾛｰﾙ/一般"
    Sheets("検証").Selec
    Eiyo01_513kenso2sht = 0
End Function
'--------------------------------------------------------------------------------
'   01_514　摂食計算
'--------------------------------------------------------------------------------
Function Eiyo01_514MealCalc1() As Long
Dim aa      As Variant
Dim bb      As Worksheet
Dim i1      As Long     '行Index
Dim i2      As Long     '欄Index
Dim Lmax    As Long     '行Max
Dim Wtemp1  As Double
Dim wtemp2  As Double
    
    aa = Sheets("摂食").UsedRange
    Set bb = Sheets("検証")
    Lmax = UBound(aa, 1)
    For i1 = 5 To Lmax
        For i2 = 1 To 27
            If i1 = 10 And i2 = 2 Then
                Wtemp1 = 1
            End If
            
'           栄養素計算 =   摂取量(F) * 栄養素(S)       * 換算値(M)
            Wtemp1 = WorksheetFunction.Round(aa(i1, 6) * aa(i1, i2 + 18) * aa(i1, 13), 2)
            If aa(i1, 16) = 2 Then
                wtemp2 = Wtemp1
            Else
                wtemp2 = WorksheetFunction.Round(Wtemp1 * (100 - Fld_Field(i2, 20)) / 100, 2)
            End If
            bb.Cells(i1, i2 * 2 - 1) = Wtemp1
            bb.Cells(i1, i2 * 2) = wtemp2
        Next i2
        For i2 = 1 To 15
'           ｴﾈﾙｷﾞｰC              =      摂取量(F) * ｴﾈﾙｷﾞｰC(AT)     * 換算値(M)
            Cells(i1, i2 + 54) = WorksheetFunction.Round(aa(i1, 6) * aa(i1, i2 + 45) * aa(i1, 13), 2)
'           ｴﾈﾙｷﾞｰW              =      摂取量(F) * ｴﾈﾙｷﾞｰW(BI)     * 換算値(M)
            Cells(i1, i2 + 69) = WorksheetFunction.Round(aa(i1, 6) * aa(i1, i2 + 60) * aa(i1, 13), 2)
'           ｶﾙｼｳﾑ  =       摂取量(F) * ｶﾙｼｳﾑ           * 換算値
            Wtemp1 = WorksheetFunction.Round(aa(i1, 6) * aa(i1, i2 + 75) * aa(i1, 13), 2)
            If aa(i1, 16) = 2 Then
                wtemp2 = Wtemp1
            Else
                wtemp2 = WorksheetFunction.Round(Wtemp1 * (100 - Fld_Field(8, 20)) / 100, 2)  '�M����(8)
            End If
            Cells(i1, i2 + 84) = Wtemp1
            Cells(i1, i2 + 99) = wtemp2
        Next i2
        For i2 = 1 To 3
'           脂質    =      摂取量(F) * 脂質(CM)        * 換算値(M)
            Wtemp1 = WorksheetFunction.Round(aa(i1, 6) * aa(i1, i2 + 90) * aa(i1, 13), 2)
            If aa(i1, 16) = 2 Then
                wtemp2 = Wtemp1
            Else
                wtemp2 = WorksheetFunction.Round(Wtemp1 * (100 - Fld_Field(i2, 20)) / 100, 2) '�M��?
            End If
            bb.Cells(i1, i2 + 114) = Wtemp1
            bb.Cells(i1, i2 + 117) = wtemp2
        Next i2
    Next i1
    Set aa = Nothing
    Set bb = Nothing
    Eiyo01_514MealCalc1 = 0
End Function
'--------------------------------------------------------------------------------
'   01_515　摂食量合計
'--------------------------------------------------------------------------------
Function Eiyo01_515MealTotal() As Long
Dim Lmax    As Long     '行Max
Dim i1      As Long     '行Index
Dim i2      As Long     '欄Index
Dim Wnissu  As Long

    Wnissu = Range("Nissu")
    Lmax = Sheets("摂食").UsedRange.Rows.Count
    i1 = Lmax + 2
    For i2 = 1 To 120
        Cells(i1, i2) = "=SUM(R5C:R[-2]C)"
        Cells(i1 + 1, i2) = "=round(R[-1]C/" & Wnissu & ",2)"
    Next i2
    For i2 = 2 To 54 Step 2
        Range("ks2_eiyoso").Offset(i2 / 2 - 1, 0) = Cells(i1, i2)
        Range("ks2_eiyoso").Offset(i2 / 2 - 1, 1) = Cells(i1 + 1, i2)
    Next i2
    
    i1 = i1 + 1                             '一日当たり合計行Index
    If Mid(Range("Q3rec"), 7, 1) = "3" Then '薄味による補正
        If Cells(i1, 45) > 17 Then                           'ｼｵ16G
            Cells(i1, 45) = Cells(i1, 45) - 8                 ' -8G
            Cells(i1, 21) = Cells(i1, 21) - 3149              'ﾅﾄﾘｳﾑ
        ElseIf Cells(i1, 45) > 9 Then                        ' 9Gｲｶ
            Cells(i1, 21) = Cells(i1, 21) - _
                          WorksheetFunction.RoundDown( _
                          (Cells(i1, 45) - 9) / 0.00254, 2)
            Cells(i1, 45) = 9                                 'ｲﾁﾘﾂ
        End If
        If Cells(i1, 46) > 17 Then                           'ｼｵ16G
            Cells(i1, 46) = Cells(i1, 46) - 8                ' -8G
            Cells(i1, 22) = Cells(i1, 22) - 3149             'ﾅﾄﾘｳﾑ
            Range("ks2_hosei23") = 1
        ElseIf Cells(i1, 46) > 9 Then                        ' 9Gｲｶ
            Cells(i1, 22) = Cells(i1, 22) - _
                          WorksheetFunction.RoundDown( _
                          (Cells(i1, 46) - 9) / 0.00254, 2)
            Cells(i1, 46) = 9                                 'ｲﾁﾘﾂ
            Range("ks2_hosei23") = 2
        Else
            Range("ks2_hosei23") = 3
        End If
    Else
        Range("ks2_hosei23") = 4
    End If
    Range("ks2_hosei11") = Range("ks2_hosei23")
    
    
    If Cells(i1, 32) < 40 Then              'VC補正(16)
        If Cells(i1, 31) > 40 Then
            Cells(i1, 32) = 40
            Range("ks2_hosei16") = 1
        Else
            Cells(i1, 32) = Cells(i1, 31)
            Range("ks2_hosei16") = 2
        End If
    Else
        Range("ks2_hosei16") = 3
    End If
    If Cells(i1, 40) < 3 Then               'VE補正(20)
        If Cells(i1, 39) > 3 Then
            Cells(i1, 40) = 3
            Range("ks2_hosei20") = 1
        Else
            Cells(i1, 40) = Cells(i1, 39)
            Range("ks2_hosei20") = 2
        End If
    Else
        Range("ks2_hosei20") = 3
    End If

    For i2 = 2 To 54 Step 2
        Range("ks2_eiyoso").Offset(i2 / 2 - 1, 3) = Cells(i1, i2)
    Next i2
End Function
'--------------------------------------------------------------------------------
'   01_521　基礎情報ほか取得
'           引数 Func 1:更新あり 2:更新なし
'--------------------------------------------------------------------------------
Function Eiyo01_521CalcDbGet(Func As Long) As Long
Dim mySqlStr    As String
Dim i1          As Long
Dim i2          As Long

    Call Eiyo99_指定シート削除("DBmirror")
    Sheets.Add After:=Sheets(Sheets.Count)      'シート追加
    ActiveSheet.Name = "DBmirror"
    Sheets("DBmirror").Range("m2:m3").NumberFormatLocal = "@"     '住所ー２
'   基礎情報取得
    With Rst_Kiso
        .Index = "PrimaryKey"
        Rst_Kiso.Open Source:=Tbl_Kiso, ActiveConnection:=myCon, _
            CursorType:=adOpenKeyset, LockType:=adLockOptimistic, _
            Options:=adCmdTableDirect
'        If Int(Range("Fcode") / 10000) = 33 Then
'            .Seek 110000 + Range("Fcode") Mod 10000
'            For i1 = 1 To .Fields.Count
'                Cells(4, i1).Value = .Fields(i1 - 1).Value
'            Next
'        End If
        .Seek Range("Fcode")
        If .EOF Then
            .AddNew
            .Fields(0).Value = Range("Fcode")
'            .Fields(1).Value = Range("Namej")
        Else
            For i1 = 1 To .Fields.Count
                Cells(1, i1).Value = .Fields(i1 - 1).Name
                Cells(2, i1).Value = .Fields(i1 - 1).Value
            Next
        End If
    End With
'   所要量取得
    With Rst_Syoyo
        .Index = "PrimaryKey"
        Rst_Syoyo.Open Source:=Tbl_Syoyo, ActiveConnection:=myCon, _
            CursorType:=adOpenKeyset, LockType:=adLockOptimistic, _
            Options:=adCmdTableDirect
        .Seek Range("Fcode")
        If .EOF Then
            .AddNew
            .Fields(0).Value = Range("Fcode")
        Else
            For i1 = 1 To .Fields.Count
                Cells(6, i1).Value = .Fields(i1 - 1).Name
                Cells(7, i1).Value = .Fields(i1 - 1).Value
            Next
        End If
    End With
'   エネルギー／カロリー
    With Rst_Energ
        .Index = "PrimaryKey"
        Rst_Energ.Open Source:=Tbl_Energ, ActiveConnection:=myCon, _
            CursorType:=adOpenKeyset, LockType:=adLockOptimistic, _
            Options:=adCmdTableDirect
        .Seek Range("Fcode")
        If .EOF Then
            .AddNew
            .Fields(0).Value = Range("Fcode")
        Else
            For i1 = 1 To .Fields.Count
                Cells(11, i1).Value = .Fields(i1 - 1).Name
                Cells(12, i1).Value = .Fields(i1 - 1).Value
            Next
        End If
    End With
    
    If Func = 1 Then
        i2 = Sheets("検証").UsedRange.Rows.Count
'       脂質の戻し
        Rst_Kiso.Fields("Sfods1").Value = Sheets("検証").Cells(i2, 115)
        Rst_Kiso.Fields("Sfods2").Value = Sheets("検証").Cells(i2, 116)
        Rst_Kiso.Fields("Sfods3").Value = Sheets("検証").Cells(i2, 117)
        Rst_Kiso.Fields("Sfodh1").Value = Sheets("検証").Cells(i2, 118)
        Rst_Kiso.Fields("Sfodh2").Value = Sheets("検証").Cells(i2, 119)
        Rst_Kiso.Fields("Sfodh3").Value = Sheets("検証").Cells(i2, 120)
'       所要量の戻し
        For i1 = 1 To 27
            Rst_Syoyo.Fields(i1 * 5 - 3).Value = Sheets("検証").Cells(i2, i1 * 2 - 1)
            Rst_Syoyo.Fields(i1 * 5 - 2).Value = Sheets("検証").Cells(i2, i1 * 2)
        Next i1
'       エネルギー／カロリーの戻し
        For i1 = 1 To 15
            Rst_Energ.Fields(i1 + 1).Value = Sheets("検証").Cells(i2, i1 + 54)
            Rst_Energ.Fields(i1 + 16).Value = Sheets("検証").Cells(i2, i1 + 69)
            Rst_Energ.Fields(i1 + 31).Value = WorksheetFunction.Round(Sheets("検証").Cells(i2, i1 + 54) / 80, 2)
            Rst_Energ.Fields(i1 + 58).Value = Sheets("検証").Cells(i2, i1 + 84)
            Rst_Energ.Fields(i1 + 73).Value = Sheets("検証").Cells(i2, i1 + 99)
        Next i1
    End If
    
    Eiyo01_521CalcDbGet = 0
End Function
'--------------------------------------------------------------------------------
'   01_522　標準体重ほか
'--------------------------------------------------------------------------------
Function Eiyo01_522Mealcalc2() As Long
Dim mySqlStr    As String
Dim 実体重      As Double
Dim 標準体重    As Double
Dim 労作        As String   '労作強度　Q7.職業の４桁目
Dim 妊娠        As Long
Dim Wtemp       As Double
Dim i1          As Long
Dim Wcondition  As Long     '所要量エネルギー適用条件
Dim Wenerg1     As Double   '指定ありの調整ｴﾈﾙｷﾞ
Dim Wenerg2     As Double   '指定除外の調整ｴﾈﾙｷﾞ
Dim KisoCd1     As Long     '必要量マスタの基礎コード  基礎代謝、活動代謝
Dim KisoCd2     As Long     '必要量マスタの基礎コード　所要量
Dim KisoCd3     As Long     '必要量マスタの基礎コード　妊娠授乳
Dim 年齢        As Long
Dim Warray      As Variant
Dim Wtext       As String

    実体重 = Range("Weght")
    年齢 = Range("Age")
    労作 = Mid(Range("Qjob1"), 4, 1)
    妊娠 = Range("Qcnd1")
'   標準体重
    If 年齢 <= 12 Then
        標準体重 = 実体重
    ElseIf Range("Hight") <= 150 Then
        標準体重 = Range("Hight") - 100
    ElseIf Range("Hight") <= 165 Then
        標準体重 = WorksheetFunction.Round((Range("Hight") - 100) * 0.9, 1)
    Else
        標準体重 = Range("Hight") - 110
    End If
    Rst_Kiso.Fields("Aans1").Value = 標準体重
    Range("ks2_weght") = 実体重
    Range("ks2_weght").Offset(0, 1) = 標準体重
'   肥満度
    Rst_Kiso.Fields("Himanp").Value = WorksheetFunction.Round( _
                                 (実体重 - 標準体重) / 実体重 * 100, 0)
'   体格指数
    If 年齢 <= 2 Then
        Wtemp = 実体重 / (Range("Hight") ^ 2) * 10 ^ 4              'ｶｳﾌﾟ指数
    ElseIf 年齢 <= 12 Then
        Wtemp = 実体重 / (Range("Hight") ^ 3) * 10 ^ 7              'ﾛｰﾚﾙ指数
    Else
        Wtemp = WorksheetFunction.Round(実体重 / 標準体重 * 100, 0) 'ﾌﾞﾛｰｶｰ指数
    End If
    Rst_Kiso.Fields("Taiis").Value = Wtemp
'   体表面積
    Rst_Kiso.Fields("Aansa").Value = Eiyo01_523_taihyou(実体重)
    Rst_Kiso.Fields("Bansa").Value = Eiyo01_523_taihyou(標準体重)
    Range("ks2_Aansa") = Rst_Kiso.Fields("Aansa").Value
    Range("ks2_Aansa").Offset(0, 1) = Rst_Kiso.Fields("Bansa").Value
'   生活指数
    Select Case 労作
        Case "A":  Wtemp = 0.35
        Case "B":  Wtemp = 0.5
        Case "C":  Wtemp = 0.75
        Case Else: Wtemp = 1#
    End Select
    Rst_Kiso.Fields("Aansx").Value = Wtemp
    Range("ks2_Aansx") = Wtemp
'   単位表面積あたりの基礎代謝
    If 年齢 < 20 Then
        KisoCd1 = 年齢
    ElseIf 年齢 < 80 Then
        KisoCd1 = Int(年齢 / 10) * 10
    Else
        KisoCd1 = 80
    End If
    If Range("Sex") = 1 Then: KisoCd1 = KisoCd1 + 100
    mySqlStr = "SELECT 基礎代謝,活動代謝 FROM " & Tbl_Need & " Where Ncode = " & KisoCd1
    Set Rst_Need = myCon.Execute(mySqlStr)
    If Rst_Need.EOF Then
        MsgBox "必要量マスタのキーなし:" & KisoCd1
    End If
    Rst_Kiso.Fields("Aans3").Value = Rst_Need.Fields("基礎代謝").Value
    Range("ks2_kisocd") = KisoCd1
    Range("ks2_kisot") = Rst_Kiso.Fields("Aans3").Value
'   基礎代謝
    Rst_Kiso.Fields("Aansb").Value = WorksheetFunction.Round(Rst_Kiso.Fields("Aans3").Value _
                                                           * Rst_Kiso.Fields("Aansa").Value * 24, 2)    '実体重の基礎代謝／日
    Rst_Kiso.Fields("Aansc").Value = WorksheetFunction.Round(Rst_Kiso.Fields("Aansb").Value / 1440, 2)  '実体重の基礎代謝／分
    Rst_Kiso.Fields("Aansd").Value = Eiyo01_524ansd(Rst_Kiso.Fields("Aansb").Value)                     '実体重のE標準量
    
    Rst_Kiso.Fields("Bansb").Value = WorksheetFunction.Round(Rst_Kiso.Fields("Aans3").Value _
                                                           * Rst_Kiso.Fields("Bansa").Value * 24, 2)    '標準体重の基礎代謝／日
    Rst_Kiso.Fields("Bansc").Value = WorksheetFunction.Round(Rst_Kiso.Fields("Bansb").Value / 1440, 2)  '標準体重の基礎代謝／分
    Rst_Kiso.Fields("Bansd").Value = Eiyo01_524ansd(Rst_Kiso.Fields("Bansb").Value)                     '標準体重のE標準量
    Range("ks2_Aansb") = Rst_Kiso.Fields("Aansb").Value                 '実体重の基礎代謝／日
    Range("ks2_Aansc") = Rst_Kiso.Fields("Aansc").Value                 '実体重の基礎代謝／分
    Range("ks2_Aansd") = Rst_Kiso.Fields("Aansd").Value                 '実体重のE標準量
    Range("ks2_Aansb").Offset(0, 1) = Rst_Kiso.Fields("Bansb").Value    '標準体重の基礎代謝／日
    Range("ks2_Aansc").Offset(0, 1) = Rst_Kiso.Fields("Bansc").Value    '標準体重の基礎代謝／分
    Range("ks2_Aansd").Offset(0, 1) = Rst_Kiso.Fields("Bansd").Value    '標準体重のE標準量

'   所要量　エネルギー  -------------------------------------------------------------------------------------
    If Rst_Kiso.Fields("Tenes").Value = 1 Then              'エネルギー指定・値
        Wenerg1 = Rst_Kiso.Fields("Tenee").Value
        Wenerg2 = Rst_Kiso.Fields("Aansd").Value
        Wcondition = 1
    ElseIf Rst_Kiso.Fields("Tenes").Value = 2 Then          'エネルギー指定・実体重
        Wenerg1 = WorksheetFunction.RoundDown(Rst_Kiso.Fields("Tenee").Value * ���̏d, 2)
        Wenerg2 = Rst_Kiso.Fields("Aansd").Value
        Wcondition = 2
    ElseIf Rst_Kiso.Fields("Tenes").Value = 3 Then          'エネルギー指定・標準体重
        Wenerg1 = WorksheetFunction.RoundDown(Rst_Kiso.Fields("Tenee").Value * �W���̏d, 2)
        Wenerg2 = Rst_Kiso.Fields("Aansd").Value
        Wcondition = 3
    ElseIf �D�P = 1 Then                                    '妊娠前期
        Wenerg1 = Rst_Kiso.Fields("Bansd").Value + 150
        Wenerg2 = Rst_Kiso.Fields("Aansd").Value + 150
        Wcondition = 4
    ElseIf �D�P = 2 Then                                    '妊娠後期
        Wenerg1 = Rst_Kiso.Fields("Bansd").Value + 350
        Wenerg2 = Rst_Kiso.Fields("Aansd").Value + 350
        Wcondition = 5
    ElseIf �D�P = 3 Then                                    '授乳期
        Wenerg1 = Rst_Kiso.Fields("Bansd").Value + 700
        Wenerg2 = Rst_Kiso.Fields("Aansd").Value + 700
        Wcondition = 6
    ElseIf Rst_Kiso.Fields("Qill1").Value <> 0 Then         '糖尿病
        Wenerg1 = Rst_Kiso.Fields("Bansd").Value - 200
        Wenerg2 = Rst_Kiso.Fields("Aansd").Value
        Wcondition = 7
    ElseIf Rst_Kiso.Fields("Qsrmr").Value <> 0 Then         'スポーツ
        Wenerg1 = Rst_Kiso.Fields("Aansb").Value _
                + WorksheetFunction.RoundDown((Rst_Kiso.Fields("Qsrmr").Value + 1.2) _
                                             * Rst_Kiso.Fields("Qsmin").Value _
                                             * Rst_Need.Fields("活動代謝").Value _
                                             * Rst_Kiso.Fields("Aansc").Value, 2)
        Wenerg2 = Wenerg1
        Wcondition = 8
    ElseIf Rst_Kiso.Fields("Taiis").Value <= 90 Or _
           Rst_Kiso.Fields("Taiis").Value >= 120 Then      '肥満
        Wenerg1 = Rst_Kiso.Fields("Bansd").Value
        Wenerg2 = Rst_Kiso.Fields("Aansd").Value
        Wcondition = 9
    Else                                                    'その他一般
        Wenerg1 = Rst_Kiso.Fields("Aansd").Value
        Wenerg2 = Rst_Kiso.Fields("Aansd").Value
        Wcondition = 10
    End If
    Rst_Syoyo.Fields("Syoyo01").Value = Wenerg1
    Range("ks2_energ") = Wcondition
    Range("ks2_energ").Offset(1, 0) = Wenerg1
    Range("ks2_energ").Offset(2, 0) = Wenerg2
    
'   所要量　その他  ------------------------------------------------------------------------------------------
    KisoCd2 = KisoCd1 + 1000
    mySqlStr = "SELECT * FROM " & Tbl_Need & " Where Ncode = " & KisoCd2
    Set Rst_Need = myCon.Execute(mySqlStr)
    If Rst_Need.EOF Then
        MsgBox "必要量マスタのキーなし:" & KisoCd2
    End If
    Range("ks2_syoyo") = KisoCd2
    For i1 = 2 To 27
        Rst_Syoyo.Fields(i1 * 5 - 1).Value = Rst_Need.Fields(i1 + 1).Value
        Range("ks2_syoyo").Offset(i1, 0) = Rst_Need.Fields(i1 + 1).Value
    Next i1
    If 労作 = "B" Or 年齢 < 15 Then
    Else
        Select Case 労作
            Case "A":  KisoCd2 = 1200
            Case "C":  KisoCd2 = 1220
            Case Else: KisoCd2 = 1230
        End Select
        If Range("Sex") = 1 Then: KisoCd2 = KisoCd2 + 100
        Select Case Range("age")
            Case 15 To 19: KisoCd2 = KisoCd2 + 1
            Case 20 To 39: KisoCd2 = KisoCd2 + 2
            Case 40 To 59: KisoCd2 = KisoCd2 + 3
            Case Else:     KisoCd2 = KisoCd2 + 4
        End Select
        mySqlStr = "SELECT * FROM " & Tbl_Need & " Where Ncode = " & KisoCd2
        Set Rst_Need = myCon.Execute(mySqlStr)
        If Rst_Need.EOF Then
            MsgBox "必要量マスタのキーなし:" & KisoCd2
        End If
        Range("ks2_syoyo").Offset(0, 1) = KisoCd2
        For i1 = 2 To 27
            Rst_Syoyo.Fields(i1 * 5 - 1).Value = _
            Rst_Syoyo.Fields(i1 * 5 - 1).Value + Rst_Need.Fields(i1 + 1).Value
            Range("ks2_syoyo").Offset(i1, 1) = Rst_Need.Fields(i1 + 1).Value
        Next i1
    End If
    Select Case 妊娠
        Case 1:    KisoCd3 = 1401   '妊娠前期
        Case 2:    KisoCd3 = 1402   '妊娠後期
        Case 3:    KisoCd3 = 1403   '授乳期
        Case Else: KisoCd3 = 0
    End Select
    If KisoCd3 > 0 Then
        mySqlStr = "SELECT * FROM " & Tbl_Need & " Where Ncode = " & KisoCd3
        Set Rst_Need = myCon.Execute(mySqlStr)
        If Rst_Need.EOF Then
            MsgBox "必要量マスタのキーなし:" & KisoCd3
        End If
        Range("ks2_syoyo").Offset(0, 2) = KisoCd3
        For i1 = 2 To 27
            Rst_Syoyo.Fields(i1 * 5 - 1).Value = _
            Rst_Syoyo.Fields(i1 * 5 - 1).Value + Rst_Need.Fields(i1 + 1).Value
            Range("ks2_syoyo").Offset(i1, 2) = Rst_Need.Fields(i1 + 1).Value
        Next i1
    End If
'   所要量　たんぱく質  --------------------------------------------------------------------------------------
    Wcondition = 0
    If Rst_Kiso.Fields("Tanps").Value = 1 Then              'たんぱく指定・値
        Wtemp = Rst_Kiso.Fields("Tanpe").Value
        Wcondition = 1
    ElseIf Rst_Kiso.Fields("Tanps").Value = 2 Then          'たんぱく指定・実体重
        Wtemp = WorksheetFunction.RoundDown(Rst_Kiso.Fields("Tanpe").Value * 実体重, 2)
        Wcondition = 2
    ElseIf Rst_Kiso.Fields("Tanps").Value = 3 Then          'たんぱく指定・標準体重
        Wtemp = WorksheetFunction.RoundDown(Rst_Kiso.Fields("Tanpe").Value * 標準体重, 2)
        Wcondition = 3
    ElseIf Rst_Kiso.Fields("Tanps").Value = 4 Then          'たんぱく指定・エネルギー比
        Wtemp = WorksheetFunction.RoundDown(Rst_Kiso.Fields("Tanpe").Value _
                                          * Wenerg1 / 400, 2)
        Wcondition = 4
    ElseIf Rst_Kiso.Fields("Qsrmr").Value <> 0 Then         'スポーツ
        Wtemp = 実体重 * 1.4
        Wcondition = 5
    ElseIf Rst_Kiso.Fields("Qwcnt").Value <> 0 Then         'ｳｴｲﾄ･ｺﾝﾄﾛｰﾙ
        If Wenerg1 < 1412 Then
            Wtemp = 60
            Wcondition = 6
        Else
            Wtemp = Wenerg1 * 17 / 400
            Wcondition = 7
        End If
    Else
        If 年齢 < 21 Then
            If Range("Sex") = 0 Then
'                          3   4   5   6   7   8   9  10  11  12  13  14  15  16  17  18  19  20��
                Wtext = "117,116,122,124,128,135,138,144,140,138,136,129,125,121,117,117,113,113"
            Else
                Wtext = "117,120,126,132,137,144,149,148,144,146,139,133,131,129,129,124,121,121"
            End If
            Warray = Split(Wtext, ",")
            Wtemp = Warray(年齢 - 3)
            If Rst_Kiso.Fields("Qsyog").Value = 1 Then: Wtemp = Wtemp + 20  '障害者
            Wcondition = 8
        Else
            i1 = Int(年齢 / 10) - 2
            If i1 > 6 Then: i1 = 6
            Select Case 労作       ' 20  30  40  50  60  70  80      才代
                Case "A":  Wtext = "130,140,150,150,160,160,160"    'X=0.35  (A)
                Case "B":  Wtext = "120,130,130,135,140,145,150"    'X=0.5   (B)
                Case "C":  Wtext = "120,130,130,130,140,140,140"    'X=0.75  (C)
                Case Else: Wtext = "120,130,130,130,135,135,135"    'X=1     (D)
            End Select
            Warray = Split(Wtext, ",")
            Wtemp = Warray(i1)
            Wcondition = 9
        End If
        Wtemp = Wenerg1 * Wtemp / 4000
        Select Case 妊娠
            Case 1: Wtemp = Wtemp + 10  '妊娠前期
            Case 2: Wtemp = Wtemp + 20  '妊娠後期
            Case 3: Wtemp = Wtemp + 20  '授乳期
        End Select
    End If
    Wtemp = WorksheetFunction.RoundDown(Wtemp, 2)
    Rst_Syoyo.Fields("Syoyo02").Value = WorksheetFunction.RoundDown(Wtemp, 2)       'たんぱく質  (02)
    Rst_Syoyo.Fields("Syoyo03").Value = WorksheetFunction.RoundDown(Wtemp / 2, 2)   '動物たんぱく(03)
    Rst_Syoyo.Fields("Syoyo04").Value = WorksheetFunction.RoundDown(Wtemp / 2, 2)   '植物たんぱく(04)
    Range("ks2_syoyo").Offset(2, 3) = Wcondition
'   所要量　脂質  --------------------------------------------------------------------------
    If 妊娠 > 0 Then
        Wtemp = 275
    ElseIf 年齢 < 21 Then
        If 労作 = "A" Then
            Wtemp = 225
        Else
            Wtemp = 275
        End If
    Else
        Select Case 労作
            Case "A", "B": Wtemp = 225
            Case Else:     Wtemp = 275
        End Select
    End If
    Wtemp = WorksheetFunction.RoundDown(Wenerg1 * Wtemp / 9000, 2)
    Rst_Syoyo.Fields("Syoyo05").Value = Wtemp                                       '脂質  (05)
    Rst_Syoyo.Fields("Syoyo26").Value = WorksheetFunction.Round(Wtemp * 0.34, 2)    'S      (24)
    Rst_Syoyo.Fields("Syoyo25").Value = WorksheetFunction.Round(Wtemp * 0.66, 2)    'P      (25)
    Rst_Syoyo.Fields("Syoyo24").Value = 300                                         'ｺﾚｽﾃﾛｰﾙ(24)
    
    Rst_Syoyo.Fields("Syoyo06").Value = WorksheetFunction.RoundDown((Wenerg1 _
                                      - Rst_Syoyo.Fields("Syoyo02").Value * 4 _
                                      - Rst_Syoyo.Fields("Syoyo05").Value * 9) / 4, 2)  '糖質(06)
    If Rst_Kiso.Fields("Qill1").Value <> 0 Then                                         '砂糖(27)
        Rst_Syoyo.Fields("Syoyo27").Value = 10          '糖尿病
    ElseIf Rst_Kiso.Fields("Qwcnt").Value <> 0 Then
        Rst_Syoyo.Fields("Syoyo27").Value = 10          'ｳｴｲﾄ･ｺﾝﾄﾛｰﾙ
    Else
        Rst_Syoyo.Fields("Syoyo27").Value = 30          'その他(一般)
    End If
    Rst_Syoyo.Fields("Syoyo07").Value = WorksheetFunction.RoundDown(Wenerg1 * 0.0099, 2)    '食物せんい(07)
'   ｶﾙｼｳﾑ(08) ------------------------------------------------------------------------------------------------
    Wcondition = 0
    If 年齢 < 21 Then
        If Range("Sex") = 0 Then
'                      3   4   5   6   7   8   9  10  11  12  13  14  15  16  17  18  19  20��
            Wtext = "171,168,169,173,176,165,169,177,187,188,177,156,134,124,115,109,103,103"
        Else
            Wtext = "173,169,169,174,182,177,184,184,175,158,142,133,119,108,100,100,100,100"
        End If
        Warray = Split(Wtext, ",")
        Wtemp = WorksheetFunction.RoundDown(実体重 * Warray(年齢 - 3) / 10, 2)
        Wcondition = 1
    ElseIf 年齢 < 60 Then
        Wtemp = 実体重 * 10
        Wcondition = 2
    Else
        Wtemp = 600
        Wcondition = 3
    End If
    Select Case 妊娠
        Case 1, 2
            Wtemp = Wtemp + 400 '妊娠前後期
            Wcondition = 4
        Case 3
            Wtemp = Wtemp + 500 '授乳期
            Wcondition = 5
    End Select
    Rst_Syoyo.Fields("Syoyo08").Value = Wtemp   'ｶﾙｼｳﾑ(08)
    Rst_Syoyo.Fields("Syoyo09").Value = Wtemp   'リン (09)
    Range("ks2_syoyo").Offset(8, 3) = Wcondition
'   鉄  ------------------------------------------------------------------------------------------------------
    Select Case 妊娠
        Case 1:    Wtemp = 15               '妊娠前期
        Case 2, 3: Wtemp = 20               '妊娠後期・授乳期
        Case Else
            Select Case 年齢
                Case 1 To 5:   Wtemp = 8    '    ５才以下
                Case 6 To 8:   Wtemp = 9    ' 6〜 8才
                Case 9 To 11:  Wtemp = 10   ' 9〜11才
                Case 12 To 19: Wtemp = 12   '12〜19才
                Case 20 To 49
                    Select Case Range("Sex")
                        Case 0:    Wtemp = 10   '20〜49才の男
                        Case Else: Wtemp = 12   '20〜49才の女
                    End Select
                Case Else: Wtemp = 10           '50才以上
            End Select
    End Select
    Rst_Syoyo.Fields("Syoyo10").Value = Wtemp
'   VB1/VB2/ﾅｲｱｼﾝ  -------------------------------------------------------------------------------------------
    Rst_Syoyo.Fields("Syoyo13").Value = WorksheetFunction.RoundDown(Wenerg2 * 0.0004, 2)
    Rst_Syoyo.Fields("Syoyo14").Value = WorksheetFunction.RoundDown(Wenerg2 * 0.00055, 2)
    Rst_Syoyo.Fields("Syoyo15").Value = WorksheetFunction.RoundDown(Wenerg2 * 0.0066, 2)
    Select Case 妊娠
        Case 1
            Rst_Syoyo.Fields("Syoyo13").Value = Rst_Syoyo.Fields("Syoyo13").Value + 0.1
            Rst_Syoyo.Fields("Syoyo14").Value = Rst_Syoyo.Fields("Syoyo14").Value + 0.1
            Rst_Syoyo.Fields("Syoyo15").Value = Rst_Syoyo.Fields("Syoyo15").Value + 1
        Case 2
            Rst_Syoyo.Fields("Syoyo13").Value = Rst_Syoyo.Fields("Syoyo13").Value + 0.2
            Rst_Syoyo.Fields("Syoyo14").Value = Rst_Syoyo.Fields("Syoyo14").Value + 0.2
            Rst_Syoyo.Fields("Syoyo15").Value = Rst_Syoyo.Fields("Syoyo15").Value + 2
        Case 3
            Rst_Syoyo.Fields("Syoyo13").Value = Rst_Syoyo.Fields("Syoyo13").Value + 0.3
            Rst_Syoyo.Fields("Syoyo14").Value = Rst_Syoyo.Fields("Syoyo14").Value + 0.4
            Rst_Syoyo.Fields("Syoyo15").Value = Rst_Syoyo.Fields("Syoyo15").Value + 5
    End Select
    
    If Rst_Kiso.Fields("Qill2").Value = 313 Then
        Rst_Syoyo.Fields("Syoyo23").Value = 6       'ｼｵ
        Rst_Syoyo.Fields("Syoyo11").Value = 2800    'ﾅﾄﾘｳﾑ
        Range("ks2_syoyo").Offset(11, 3) = 1
        Range("ks2_syoyo").Offset(23, 3) = 1
    Else
        Rst_Syoyo.Fields("Syoyo23").Value = 10      'ｼｵ
        Range("ks2_syoyo").Offset(23, 3) = 2
    End If
    Rst_Syoyo.Fields("Syoyo20").Value = WorksheetFunction.Round(Rst_Syoyo.Fields("Syoyo25").Value * 0.6, 2)     'VE
    Rst_Syoyo.Fields("Syoyo21").Value = Rst_Syoyo.Fields("Syoyo11").Value                                       'ｶﾘｳﾑ <= ﾅﾄﾘｳﾑ
    Rst_Syoyo.Fields("Syoyo22").Value = WorksheetFunction.RoundDown(Rst_Syoyo.Fields("Syoyo08").Value / 2, 2)   'Mg=Ca/2
'   栄養素比率
    If Rst_Syoyo.Fields("Foodh01").Value = 0 Then
        Rst_Kiso.Fields("Per01").Value = 0
        Rst_Kiso.Fields("Per02").Value = 0
        Rst_Kiso.Fields("Per03").Value = 0
        Rst_Kiso.Fields("Per05").Value = 0
        Rst_Kiso.Fields("Per06").Value = 0
    Else
        Rst_Kiso.Fields("Per01").Value = WorksheetFunction.RoundDown( _
                                     Rst_Syoyo.Fields("Foodh06").Value * 400 / Rst_Syoyo.Fields("Foodh01").Value, 1)
        Rst_Kiso.Fields("Per02").Value = WorksheetFunction.RoundDown( _
                                     Rst_Energ.Fields("Enec11").Value * 100 / Rst_Syoyo.Fields("Foodh01").Value, 1)
        Rst_Kiso.Fields("Per03").Value = WorksheetFunction.RoundDown( _
                                     Rst_Syoyo.Fields("Foodh02").Value * 400 / Rst_Syoyo.Fields("Foodh01").Value, 1)
        Rst_Kiso.Fields("Per05").Value = WorksheetFunction.RoundDown( _
                                     Rst_Syoyo.Fields("Foodh05").Value * 900 / Rst_Syoyo.Fields("Foodh01").Value, 1)
        Rst_Kiso.Fields("Per06").Value = WorksheetFunction.RoundDown( _
                                     Rst_Kiso.Fields("Sfodh1").Value * 900 / Rst_Syoyo.Fields("Foodh01").Value, 1)
    End If
    If Rst_Syoyo.Fields("Foodh02").Value = 0 Then
        Rst_Kiso.Fields("Per04").Value = 0
    Else
        Rst_Kiso.Fields("Per04").Value = WorksheetFunction.RoundDown( _
                                     Rst_Syoyo.Fields("Foodh03").Value * 100 / Rst_Syoyo.Fields("Foodh02").Value, 1)
    End If
    If Rst_Syoyo.Fields("Foodh09").Value = 0 Then
        Rst_Kiso.Fields("Per07").Value = 0
    Else
        Rst_Kiso.Fields("Per07").Value = WorksheetFunction.RoundDown( _
                                     Rst_Syoyo.Fields("Foodh08").Value * 100 _
                                     / Rst_Syoyo.Fields("Foodh09").Value, 1)
    End If
    If Rst_Syoyo.Fields("Foodh26").Value = 0 Then
        Rst_Kiso.Fields("Per08").Value = 0
    Else
        Rst_Kiso.Fields("Per08").Value = WorksheetFunction.RoundDown( _
                                     Rst_Syoyo.Fields("Foodh25").Value * 100 _
                                     / Rst_Syoyo.Fields("Foodh26").Value, 1)
    End If
    Rst_Need.Close
    Set Rst_Need = Nothing

    '   更新結果表示
    For i1 = 1 To Rst_Kiso.Fields.Count
        Cells(3, i1).Value = Rst_Kiso.Fields(i1 - 1).Value
    Next
    For i1 = 1 To 27
        Range("ks2_syoyo").Offset(i1, 4) = Rst_Syoyo.Fields(i1 * 5 - 1).Value
    Next i1
End Function
'--------------------------------------------------------------------------------
'   01_523 体表面積
'       ５歳以下    体重^0.423 * 身長^0.362 * 382.89 / 10000
'       ６歳以上    体重^0.444 * 身長^0.663 *  88.83 / 10000
'--------------------------------------------------------------------------------
Function Eiyo01_523_taihyou(体重 As Double) As Double
Dim Wtemp   As Double

    If Range("Age") < 6 Then
        Wtemp = WorksheetFunction.Round(�̏d ^ 0.423 * Range("hight") ^ 0.362 * 382.89 / 10000, 2)
    Else
        Wtemp = WorksheetFunction.Round(�̏d ^ 0.444 * Range("hight") ^ 0.663 * 88.83 / 10000, 2)
    End If
    Eiyo01_523_taihyou = Wtemp
End Function
'--------------------------------------------------------------------------------
'   01_524 エネルギー標準量　生活活動強度補正
'       障害者      50%
'       ６０歳代    90%
'       ７０歳代    80%
'       ８０歳以上  70%
'       基礎代謝／日 * (補正生活活動強度+1) * 1.1
'--------------------------------------------------------------------------------
Function Eiyo01_524ansd(基礎代謝 As Double) As Double
Dim Wtemp   As Double
    
    If Rst_Kiso.Fields("Qsyog").Value = 1 Then                      '障害者
        Wtemp = WorksheetFunction.Round(Rst_Kiso.Fields("Aansx").Value * 0.5, 2)
    ElseIf Range("Age") < 60 Then                                   '６０歳未満
        Wtemp = Rst_Kiso.Fields("Aansx").Value
    ElseIf Range("Age") >= 60 And Range("Age") <= 69 Then           '６０歳代
        Wtemp = WorksheetFunction.Round(Rst_Kiso.Fields("Aansx").Value * 0.9, 2)
    ElseIf Range("Age") >= 70 And Range("Age") <= 79 Then           '７０歳代
        Wtemp = WorksheetFunction.Round(Rst_Kiso.Fields("Aansx").Value * 0.8, 2)
    Else                                                            '８０歳以上
        Wtemp = WorksheetFunction.Round(Rst_Kiso.Fields("Aansx").Value * 0.7, 2)
    End If
    Wtemp = WorksheetFunction.Round(基礎代謝 * (1 + Wtemp) * 1.1, 2)
    Eiyo01_524ansd = Wtemp
End Function
'--------------------------------------------------------------------------------
'   01_525 過不足計算、アドバイス
'--------------------------------------------------------------------------------
Function Eiyo01_525MealDiffe() As Long
Dim i1      As Long
Dim Wans1   As Long
Dim Wans2   As Long
Dim Wtext   As String
Dim Wtext2  As String

    For i1 = 1 To 27
        If Rst_Syoyo.Fields(i1 * 5 - 1).Value = 0 Then
            Wans1 = 0
            Wans2 = 5
        Else
            Wans1 = WorksheetFunction.Round(Rst_Syoyo.Fields(i1 * 5 - 2).Value _
                                          / Rst_Syoyo.Fields(i1 * 5 - 1).Value * 100, 0) - 100   '過不足率
            If Wans1 <= Fld_Field(i1, 16) Then
                Wans2 = 1
            ElseIf Wans1 <= Fld_Field(i1, 17) Then
                Wans2 = 2
            ElseIf Wans1 < Fld_Field(i1, 18) Then
                Wans2 = 3
            ElseIf Wans1 < Fld_Field(i1, 19) Then
                Wans2 = 4
            Else
                Wans2 = 5
            End If
        End If
        Rst_Syoyo.Fields(i1 * 5 + 0).Value = Wans1
        Rst_Syoyo.Fields(i1 * 5 + 1).Value = Wans2
    Next i1
'   アドバイス
    Wtext = Rst_Kiso.Fields("Q3rec").Value      '食習慣
    Wans2 = 0
    For i1 = 1 To 10
        Wans1 = Val(Mid(Wtext, i1, 1))
        If Wans1 <= 0 Then: Exit For
        If i1 = 4 And Wans1 = 4 Then
            Wans2 = Wans2 + 1
        Else
            Wans2 = Wans2 + Wans1
        End If
    Next i1
    If i1 > 10 Then
        Select Case Wans2
            Case 0 To 15:  Wans1 = 3040
            Case 16 To 20: Wans1 = 3030
            Case 21 To 25: Wans1 = 3020
            Case Else:     Wans1 = 3010
        End Select
    Else
        Wans1 = 98  '2008/4/25 3050を0098に変更
    End If
    Rst_Kiso.Fields("Badv1").Value = Wans1
    
    Wtext = Rst_Kiso.Fields("Q4rec").Value      '休養
    Wans2 = 0
    For i1 = 1 To 5
        Wans1 = Val(Mid(Wtext, i1, 1))
        If Wans1 <= 0 Then: Exit For
        Wans2 = Wans2 + Wans1
    Next i1
    If i1 > 5 Then
        Select Case Wans2
            Case 0 To 6:   Wans1 = 3140
            Case 7 To 9:   Wans1 = 3130
            Case 10 To 12: Wans1 = 3120
            Case Else:     Wans1 = 3110
        End Select
    Else
        Wans1 = 98  '2008/4/25 3150を0098に変更
    End If
    Rst_Kiso.Fields("Badv2").Value = Wans1
    
    Wtext = Rst_Kiso.Fields("Q5rec").Value      '運動
    Wans2 = 0
    For i1 = 1 To 3
        Wans1 = Val(Mid(Wtext, i1, 1))
        If Wans1 <= 0 Then: Exit For
        Wans2 = Wans2 + Wans1
    Next i1
    If i1 > 3 Then
        Select Case Wans2
            Case 0 To 5:  Wans1 = 3240
            Case 6 To 8:  Wans1 = 3230
            Case 9 To 11: Wans1 = 3220
            Case Else:    Wans1 = 3210
        End Select
    Else
        Wans1 = 98  '2008/4/25 3250を0098に変更
    End If
    Rst_Kiso.Fields("Badv3").Value = Wans1
    
    Wtext = Rst_Kiso.Fields("Q6r_a").Value & Rst_Kiso.Fields("Q6r_b").Value _
          & Rst_Kiso.Fields("Q6r_c").Value & Rst_Kiso.Fields("Q6r_d").Value _
          & Rst_Kiso.Fields("Q6r_e").Value                              '健康調査
    Wans2 = 0
    For i1 = 1 To 35
        If Mid(Wtext, i1, 1) = "9" Then: Exit For
        If Mid(Wtext, i1, 1) = "0" Then
            Wans2 = Wans2 + 1
        End If
    Next i1
    If i1 > 35 Then
        Select Case Wans2
            Case 0 To 7:  Wans1 = 3310
            Case 8 To 9:  Wans1 = 3320
            Case Else:    Wans1 = 3340
        End Select
    Else
        Wans1 = 98
    End If
    Rst_Kiso.Fields("Badv4").Value = Wans1
    
'             ....+....1....+....2....+....3....+....4....+....5
    Wtext2 = "00101100011010000110101101000101100111111111111111"
    Wans2 = 0
    For i1 = 1 To 50
        If Mid(Wtext, i1, 1) = "9" Then: Exit For
        If Mid(Wtext, i1, 1) = "0" And _
           Mid(Wtext2, i1, 1) = "1" Then
            Wans2 = Wans2 + 1
        End If
    Next i1
    If i1 > 50 Then
        Select Case Wans2
            Case 0 To 3:   Wans1 = 3510
            Case 4 To 7:   Wans1 = 3520
            Case 8 To 11:  Wans1 = 3530
            Case Else:     Wans1 = 3540
        End Select
    Else
        Wans1 = 98
    End If
    Rst_Kiso.Fields("Badv5").Value = Wans1
    
    If Rst_Kiso.Fields("Qsrmr").Value <> 0 Then            '  ｳｴｲﾄ ｱﾄﾞﾊﾞｲｽ
       Wans1 = 2801
    ElseIf Rst_Kiso.Fields("age").Value <= 12 Then
       Wans1 = 2806
    ElseIf Rst_Kiso.Fields("Taiis").Value < 120 Or _
           Rst_Kiso.Fields("Qcnd1").Value = 1 Or _
           Rst_Kiso.Fields("Qcnd1").Value = 2 Then      ' 120%ﾐﾏﾝ OR ﾆﾝｼﾝ
       Wans1 = 0
    Else
       Wans1 = 2803
    End If
    Rst_Kiso.Fields("Wadvs").Value = Wans1
    
    Rst_Kiso.Fields("Cadv1").Value = 98
    Rst_Kiso.Fields("Cadv2").Value = 98
    Rst_Kiso.Fields("Cadv3").Value = 98
    Rst_Kiso.Fields("Cadv4").Value = 98
    i1 = 1
    If Rst_Syoyo.Fields("Syort20").Value < -37 Then     'VE ﾌｿｸ
       If Rst_Kiso.Fields("age").Value < 40 And _
          Rst_Kiso.Fields("Sex").Value = 1 Then
           Call Eiyo01_526Cadvs(3630)                   '39ｲｶ  ｵﾝﾅ
       Else
           Call Eiyo01_526Cadvs(3610)                   'ｵﾄｺ & 40ｲｼﾞｮｳ ｵﾝﾅ
       End If
    End If
    If Rst_Syoyo.Fields("Syort12") < -37 Then: Call Eiyo01_526Cadvs(3640)   'VA
    If Rst_Syoyo.Fields("Syort13") < -37 Then: Call Eiyo01_526Cadvs(3620)   'VB1
    If Rst_Syoyo.Fields("Syort08") < -37 Then: Call Eiyo01_526Cadvs(3650)   'CA
    If Rst_Syoyo.Fields("Syort07") < -37 Then: Call Eiyo01_526Cadvs(3660)   'ｾﾝｲ

    If Rst_Syoyo.Fields("Syort23") > 12 Then: Call Eiyo01_527Cadvs(3730)    'ｼｵ
    If Rst_Energ.Fields("Enet08") < 85 Then: Call Eiyo01_527Cadvs(3720)     '3ｸﾞﾝ ｴｲﾖｳ ｾｯｼｭ
    If Rst_Syoyo.Fields("Syort08") < -37 Then: Call Eiyo01_527Cadvs(3760)   'CA
    If Rst_Energ.Fields("Enet08") < 85 Or _
       Rst_Energ.Fields("Enet09") < 85 Then: Call Eiyo01_527Cadvs(3740)     '4ｸﾞﾝ
    If Rst_Kiso.Fields("PER04") > 50 Then: Call Eiyo01_527Cadvs(3710)      'ﾄﾞｳﾌﾞﾂ ﾀﾝﾊﾟｸｼﾂ ﾋ
    
'                     ---- 0.35 ----  ---- 0.5 -----
    Wtext = Empty   '<=-110=><=111-=><=-110=><=111-=>
    Wtext = Wtext & "00012023000120270001021500010211"  '   -20 ｵﾄｺ
    Wtext = Wtext & "00012028000120210001021600010214"  ' 21-30
    Wtext = Wtext & "00013335000133340001202800012029"  ' 31-40
    Wtext = Wtext & "00013337000133360001202100012025"  ' 41-50
    Wtext = Wtext & "00012025000120290001021100010214"  ' 51-60
    Wtext = Wtext & "00012030000120260001021700010218"  ' 61-70
    Wtext = Wtext & "00012032000120310001020900010219"  ' 71-
    Wtext = Wtext & "00010204000102030001021000010212"  '   -20 ｵﾝﾅ
    Wtext = Wtext & "00012021000120220001020300010211"  ' 21-30
    Wtext = Wtext & "00012024000120230001021300010212"  ' 31-40
    Wtext = Wtext & "00012026000120250001020700010214"  ' 41-50
    Wtext = Wtext & "00010205000120260001020500010208"  ' 51-60
    Wtext = Wtext & "00010206000102050001020600010209"  ' 61-70
    Wtext = Wtext & "00010209000102080001020900010208"  ' 71-
    If Rst_Kiso.Fields("Aansx") > 50 Then
        Wtext2 = "38394041"
    Else
        If Rst_Kiso.Fields("Taiis").Value < 111 Then   'ﾀｲｶｸ ｼｽｳ
            Wans1 = 0
        Else
            Wans1 = 1
        End If
        If Rst_Kiso.Fields("Aansx") = 0.5 Then: Wans1 = Wans1 + 2        'ｾｲｶﾂ ｼｽｳ
        Select Case Rst_Kiso.Fields("Age")
            Case 0 To 20:
            Case 21 To 30: Wans1 = Wans1 + 4
            Case 31 To 40: Wans1 = Wans1 + 8
            Case 41 To 50: Wans1 = Wans1 + 12
            Case 51 To 60: Wans1 = Wans1 + 16
            Case 61 To 70: Wans1 = Wans1 + 20
            Case Else:     Wans1 = Wans1 + 24
        End Select
        If Rst_Kiso.Fields("Sex") = 1 Then: Wans1 = Wans1 + 28        'ｵﾝﾅ
        Wtext2 = Mid(Wtext, Wans1 * 8 + 1, 8)
    End If
    Rst_Kiso.Fields("Dadv1").Value = Left(Wtext2, 2)
    Rst_Kiso.Fields("Dadv2").Value = Mid(Wtext2, 3, 2)
    Rst_Kiso.Fields("Dadv3").Value = Mid(Wtext2, 5, 2)
    Rst_Kiso.Fields("Dadv4").Value = Mid(Wtext2, 7, 2)
    
    Eiyo01_525MealDiffe = 0
End Function
'--------------------------------------------------------------------------------
'   01_526 Cアドバイス１
'--------------------------------------------------------------------------------
Function Eiyo01_526Cadvs(advc As Long)
    If Rst_Kiso.Fields("Cadv1").Value = 98 Then
        Rst_Kiso.Fields("Cadv1").Value = advc
    ElseIf Rst_Kiso.Fields("Cadv2").Value = 98 Then
        Rst_Kiso.Fields("Cadv2").Value = advc
    End If
End Function
'--------------------------------------------------------------------------------
'   01_527 Cアドバイス２
'--------------------------------------------------------------------------------
Function Eiyo01_527Cadvs(advc As Long)
    If Rst_Kiso.Fields("Cadv3").Value = 98 Then
        Rst_Kiso.Fields("Cadv3").Value = advc
    ElseIf Rst_Kiso.Fields("Cadv4").Value = 98 Then
        Rst_Kiso.Fields("Cadv4").Value = advc
    End If
End Function
'--------------------------------------------------------------------------------
'   01_528 栄養比率
'--------------------------------------------------------------------------------
Function Eiyo01_528Eiyohirit() As Long
Dim i1      As Long
Dim Wtext2  As String
Dim Wtemp   As Double

    i1 = WorksheetFunction.Round(Rst_Syoyo.Fields("Syoyo01").Value / 80, 0)
    Select Case i1
        Case 11: Wtext2 = "030101010401"
        Case 12: Wtext2 = "030101010501"
        Case 13: Wtext2 = "030101010601"
        Case 14: Wtext2 = "030101010701"
        Case 15: Wtext2 = "030201010701"
        Case 16: Wtext2 = "030201010801"
        Case 17: Wtext2 = "040201010801"
        Case 18: Wtext2 = "040201010802"
        Case 19: Wtext2 = "040201010902"
        Case 20: Wtext2 = "040201011002"
        Case 21: Wtext2 = "040201011102"
        Case 22: Wtext2 = "040201011202"
        Case 23: Wtext2 = "040201011302"
        Case 24: Wtext2 = "040201011402"
        Case 25: Wtext2 = "040201011502"
        Case 26: Wtext2 = "040201011602"
        Case 27: Wtext2 = "040201011702"
        Case 28: Wtext2 = "040201011802"
        Case 29: Wtext2 = "040201011902"
        Case 30: Wtext2 = "040201012002"
        Case 31: Wtext2 = "050201012002"
        Case 32: Wtext2 = "050201012102"
        Case 33: Wtext2 = "050201022102"
        Case 34: Wtext2 = "050201022103"
        Case 35: Wtext2 = "050201022203"
        Case 36: Wtext2 = "050301022203"
        Case 37: Wtext2 = "050301022303"
        Case 38: Wtext2 = "050301022403"
        Case 39: Wtext2 = "050301022404"
        Case 40: Wtext2 = "050301022504"
        Case 41: Wtext2 = "060301022504"
        Case 42: Wtext2 = "060301022604"
        Case 43: Wtext2 = "060301022704"
        Case 44: Wtext2 = "060301022804"
        Case 45: Wtext2 = "060301022904"
        Case 46: Wtext2 = "060301023004"
        Case 47: Wtext2 = "060301023104"
        Case 48: Wtext2 = "060301023204"
        Case 49: Wtext2 = "060301023304"
        Case 50: Wtext2 = "060301023404"
        Case 63: Wtext2 = "150801032907"
        Case Else: Wtext2 = "000000000000"
    End Select
    Rst_Energ.Fields("Enes01").Value = Val(Mid(Wtext2, 1, 2))
    Rst_Energ.Fields("Enes02").Value = Val(Mid(Wtext2, 3, 2))
    Rst_Energ.Fields("Enes03").Value = Val(Mid(Wtext2, 5, 2))
    Rst_Energ.Fields("Enes04").Value = Val(Mid(Wtext2, 7, 2))
    Rst_Energ.Fields("Enes05").Value = Val(Mid(Wtext2, 9, 2))
    Rst_Energ.Fields("Enes06").Value = Val(Mid(Wtext2, 11, 2))
    
    Wtemp = Rst_Energ.Fields("Enec01").Value _
          + Rst_Energ.Fields("Enec02").Value _
          + Rst_Energ.Fields("Enec03").Value _
          + Rst_Energ.Fields("Enec04").Value
    Wtemp = WorksheetFunction.Round(Wtemp / 80, 1)
    Rst_Energ.Fields("Enek01").Value = Wtemp - Rst_Energ.Fields("Enes01").Value
    
    Wtemp = Rst_Energ.Fields("Enec05").Value _
          + Rst_Energ.Fields("Enec06").Value _
          + Rst_Energ.Fields("Enec07").Value
    Wtemp = WorksheetFunction.Round(Wtemp / 80, 1)
    Rst_Energ.Fields("Enek02").Value = Wtemp - Rst_Energ.Fields("Enes02").Value

    Wtemp = WorksheetFunction.Round(Rst_Energ.Fields("Enec08").Value / 80, 1)
    Rst_Energ.Fields("Enek03").Value = Wtemp - Rst_Energ.Fields("Enes03").Value
    
    Wtemp = Rst_Energ.Fields("Enec09").Value _
          + Rst_Energ.Fields("Enec10").Value
    Wtemp = WorksheetFunction.Round(Wtemp / 80, 1)
    Rst_Energ.Fields("Enek04").Value = Wtemp - Rst_Energ.Fields("Enes04").Value
    
    Wtemp = Rst_Energ.Fields("Enec11").Value _
          + Rst_Energ.Fields("Enec12").Value _
          + Rst_Energ.Fields("Enec13").Value
    Wtemp = WorksheetFunction.Round(Wtemp / 80, 1)
    Rst_Energ.Fields("Enek05").Value = Wtemp - Rst_Energ.Fields("Enes05").Value

    Wtemp = Rst_Energ.Fields("Enec14").Value _
          + Rst_Energ.Fields("Enec15").Value
    Wtemp = WorksheetFunction.Round(Wtemp / 80, 1)
    Rst_Energ.Fields("Enek06").Value = Wtemp - Rst_Energ.Fields("Enes06").Value

    Eiyo01_528Eiyohirit = 0
End Function
'--------------------------------------------------------------------------------
'   01_540 旧摂食計算値の比較用
'--------------------------------------------------------------------------------
Function Eiyo01_540Old_Check() As Long
Dim mySqlStr    As String
Dim i1          As Long
Dim i2          As Long
Dim Lmax1       As Long
Dim Lmax2       As Long
Dim Lmax3       As Long
Dim Errcnt      As Long

'   更新結果表示
    For i1 = 1 To Rst_Kiso.Fields.Count
        Cells(3, i1).Value = Rst_Kiso.Fields(i1 - 1).Value
    Next
    For i1 = 1 To Rst_Syoyo.Fields.Count
        Cells(8, i1).Value = Rst_Syoyo.Fields(i1 - 1).Value
    Next
    For i1 = 1 To Rst_Energ.Fields.Count
        Cells(13, i1).Value = Rst_Energ.Fields(i1 - 1).Value
    Next
    
    Lmax1 = Range("c1").End(xlToRight).Column
    Lmax2 = Range("c6").End(xlToRight).Column
    Lmax3 = Range("c11").End(xlToRight).Column
    If IsEmpty(Cells(4, 1)) Then
        Errcnt = Eiyo01_541diff(2, 3, Lmax1, 0)
        Errcnt = Eiyo01_541diff(7, 8, Lmax2, Errcnt)
        Errcnt = Eiyo01_541diff(12, 13, Lmax3, Errcnt)
    Else
        Errcnt = Eiyo01_541diff(3, 4, Lmax1, 0)
        Errcnt = Eiyo01_541diff(8, 9, Lmax2, Errcnt)
        Errcnt = Eiyo01_541diff(13, 14, Lmax3, Errcnt)
    End If
    If Errcnt > 0 Then: MsgBox "不一致 " & Errcnt
    
End Function
'--------------------------------------------------------------------------------
'   01_541 比較
'--------------------------------------------------------------------------------
Function Eiyo01_541diff(i1 As Long, i2 As Long, Max As Long, Errcnt As Long) As Long
Dim ii  As Long
    For ii = 3 To Max
        If Cells(i1, ii) <> Cells(i2, ii) Then
            Cells(i2, ii).Interior.ColorIndex = 4
            Errcnt = Errcnt + 1
        End If
    Next ii
    Eiyo01_541diff = Errcnt
End Function
'--------------------------------------------------------------------------------
'   01_550 基礎情報ほかClose
'--------------------------------------------------------------------------------
Function Eiyo01_550RstClose()
    Rst_Kiso.Update
    Rst_Syoyo.Update
    Rst_Energ.Update
    Rst_Kiso.Close
    Rst_Syoyo.Close
    Rst_Energ.Close
    Set Rst_Kiso = Nothing
    Set Rst_Syoyo = Nothing
    Set Rst_Energ = Nothing
End Function
'--------------------------------------------------------------------------------
'   01_700 カウンセリングシート作表
'--------------------------------------------------------------------------------
Function Eiyo01_700作表Click()
    
    If IsEmpty(Range("Fcode")) Or _
       Range("Fcode") <> Range("Fsave") Then
        MsgBox "基礎情報の検索が行われていません"
        Exit Function
    End If
    Application.ScreenUpdating = False  '画面描画抑止
    Call Eiyo91DB_Open                  'DB Open
    Call Eiyo01_511MealFldgt            '項目要素取得
    Call Eiyo01_701Sheet                'ｶｳﾝｾﾘﾝｸﾞｼｰﾄ追加
    Call Eiyo01_702DbGet                'DB Get(521)
    Call Eiyo01_703Pset                 '印刷項目の設定
    Call Eiyo01_704Advic                'アドバイス
    Call Eiyo01_705Footer               'コード＆日付、カンウセラー
    Call Eiyo920DB_Close                'DB Close
'    Call Eiyo99_指定シート削除("DBmirror")
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Select
End Function
'--------------------------------------------------------------------------------
'   01_701 シート追加
'--------------------------------------------------------------------------------
Function Eiyo01_701Sheet()
Const ShtName = "ｶｳﾝｾﾘﾝｸﾞｼｰﾄ"
Const Eiyo01Bk = "Eiyo01_基礎摂食入力.xls"
Const Eiyo02Bk = "Eiyo02_ｶｳﾝｾﾘﾝｸﾞｼｰﾄ.xls"
    
    Call Eiyo99_指定シート削除(ShtName)
    Workbooks.Open Filename:=ThisWorkbook.Path & "" & Eiyo02Bk, ReadOnly:=False    'openn
    Windows(Eiyo02Bk).Activate
    Sheets(ShtName).Copy After:=Workbooks(Eiyo01Bk).Sheets(3)                       'copy
    Windows(Eiyo02Bk).Close savechanges:=False                                      'close
    Windows(Eiyo01Bk).Activate
End Function
'--------------------------------------------------------------------------------
'   01_702 DB Get
'--------------------------------------------------------------------------------
Function Eiyo01_702DbGet()
Dim i1  As Long
    If IsError(Evaluate("DBmirror!a1")) Then
        Call Eiyo01_521CalcDbGet(2)
        Rst_Kiso.Close
        Rst_Syoyo.Close
        Rst_Energ.Close
        Set Rst_Kiso = Nothing
        Set Rst_Syoyo = Nothing
        Set Rst_Energ = Nothing
    End If
End Function
'--------------------------------------------------------------------------------
'   01_703 印刷項目の設定
'--------------------------------------------------------------------------------
Function Eiyo01_703Pset()
Dim aa  As Worksheet
Dim bb  As Worksheet
Dim i1  As Long
Dim i2  As Long
Dim i3  As Long
Dim i4  As Long

    Set aa = Sheets("DBmirror")
    Set bb = Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ")
'   調査日・期間
    bb.Range("p_date1") = Format(aa.Range("b2"), " yyyy""年"" mm""月"" dd""日から") & _
                     Format(aa.Range("b2") + aa.Range("c2") - 1, " mm""月"" dd""日まで(") & _
                     aa.Range("c2") & "日間)"
'   性別
    If aa.Range("e2") = 0 Then
        bb.Range("P_sex") = "男"
    Else
        bb.Range("P_sex") = "女"
    End If
'
    bb.Range("P_age") = aa.Range("g2")              '年齢
    bb.Range("P_adrno") = aa.Range("k2")            '郵便番号
    bb.Range("P_adrs1") = aa.Range("l2")            '住所ー１
    bb.Range("P_adrs2") = "'" & aa.Range("m2")      '住所ー２
    bb.Range("P_namej") = aa.Range("d2") & "　様"   '氏名
    bb.Range("P_fcode") = aa.Range("a2")            'Fcode
    bb.Range("P_hok1") = aa.Range("at2")            '保険証記号
    bb.Range("P_hok2") = aa.Range("au2")            '保険証ＮＯ
'   体位
    bb.Range("P_hight") = aa.Range("h2")            '身長
    bb.Range("P_weght") = aa.Range("i2")            '体重
    If aa.Range("g2") > 12 Or _
       aa.Range("ad2") = 0 Or _
       aa.Range("ag2") = 0 Then
        bb.Range("P_aans1") = aa.Range("bl2")       '標準体重
    Else
        bb.Range("P_aans1") = Empty
    End If
    If aa.Range("j2") = 0 Then                      '皮下脂肪
        bb.Range("P_sibou") = Empty
    Else
        bb.Range("P_sibou") = aa.Range("j2")
    End If
    If aa.Range("ad2") = 0 And aa.Range("ag2") = 0 Then '妊娠/ｽﾎﾟｰﾂ
        bb.Range("P_taii") = aa.Range("bx2")        '体位指数
        If aa.Range("g2") < 3 Then
            bb.Range("P_tsisu") = "(カウプ指数)"    '２才以下
        ElseIf aa.Range("g2") < 13 Then
            bb.Range("P_tsisu") = "(ローレル指数)"  '３〜１２才
        ElseIf aa.Range("i2") < 150 Then
            bb.Range("P_tsisu") = "(ブローカー指数変法)"  '身長150cm未満
        Else
            bb.Range("P_tsisu") = "(ブローカー指数変法)"  '身長150cm以上
        End If
    Else
        bb.Range("P_taii") = Empty
        bb.Range("P_tsisu") = Empty
    End If
'   食品摂取バランス
    bb.Range("P_enec01") = aa.Cells(12, 3)
    bb.Range("P_enec02") = aa.Cells(12, 4)
    bb.Range("P_enec03") = aa.Cells(12, 5)
    bb.Range("P_enec04") = aa.Cells(12, 6)
    bb.Range("P_enec05") = aa.Cells(12, 7)
    bb.Range("P_enec06") = aa.Cells(12, 8)
    bb.Range("P_enec07") = aa.Cells(12, 9)
    bb.Range("P_enec08") = aa.Cells(12, 10)
    bb.Range("P_enec09") = aa.Cells(12, 11)
    bb.Range("P_enec10") = aa.Cells(12, 12)
    bb.Range("P_enec11") = aa.Cells(12, 13)
    bb.Range("P_enec12") = aa.Cells(12, 14)
    bb.Range("P_enec13") = aa.Cells(12, 15)
    bb.Range("P_enec14") = aa.Cells(12, 16)
    bb.Range("P_enec15") = aa.Cells(12, 17)
    bb.Range("P_enew01") = aa.Cells(12, 18)
    bb.Range("P_enew02") = aa.Cells(12, 19)
    bb.Range("P_enew03") = aa.Cells(12, 20)
    bb.Range("P_enew04") = aa.Cells(12, 21)
    bb.Range("P_enew05") = aa.Cells(12, 22)
    bb.Range("P_enew06") = aa.Cells(12, 23)
    bb.Range("P_enew07") = aa.Cells(12, 24)
    bb.Range("P_enew08") = aa.Cells(12, 25)
    bb.Range("P_enew09") = aa.Cells(12, 26)
    bb.Range("P_enew10") = aa.Cells(12, 27)
    bb.Range("P_enew11") = aa.Cells(12, 28)
    bb.Range("P_enew12") = aa.Cells(12, 29)
    bb.Range("P_enew13") = aa.Cells(12, 30)
    bb.Range("P_enew14") = aa.Cells(12, 31)
    bb.Range("P_enew15") = aa.Cells(12, 32)
    bb.Range("P_enet01") = aa.Cells(12, 33)
    bb.Range("P_enet02") = aa.Cells(12, 34)
    bb.Range("P_enet03") = aa.Cells(12, 35)
    bb.Range("P_enet04") = aa.Cells(12, 36)
    bb.Range("P_enet05") = aa.Cells(12, 37)
    bb.Range("P_enet06") = aa.Cells(12, 38)
    bb.Range("P_enet07") = aa.Cells(12, 39)
    bb.Range("P_enet08") = aa.Cells(12, 40)
    bb.Range("P_enet09") = aa.Cells(12, 41)
    bb.Range("P_enet10") = aa.Cells(12, 42)
    bb.Range("P_enet11") = aa.Cells(12, 43)
    bb.Range("P_enet12") = aa.Cells(12, 44)
    bb.Range("P_enet13") = aa.Cells(12, 45)
    bb.Range("P_enet14") = aa.Cells(12, 46)
    bb.Range("P_enet15") = aa.Cells(12, 47)
    bb.Range("P_enes01") = aa.Cells(12, 48)
    bb.Range("P_enes02") = aa.Cells(12, 49)
    bb.Range("P_enes03") = aa.Cells(12, 50)
    bb.Range("P_enes04") = aa.Cells(12, 51)
    bb.Range("P_enes05") = aa.Cells(12, 52)
    bb.Range("P_enes06") = aa.Cells(12, 53)
    bb.Range("P_enek01") = aa.Cells(12, 54)
    bb.Range("P_enek02") = aa.Cells(12, 55)
    bb.Range("P_enek03") = aa.Cells(12, 56)
    bb.Range("P_enek04") = aa.Cells(12, 57)
    bb.Range("P_enek05") = aa.Cells(12, 58)
    bb.Range("P_enek06") = aa.Cells(12, 59)
    bb.Range("P_enec99") = "=sum(r17:r36)"
    bb.Range("P_enet99") = "=sum(V17:v36)"
    bb.Range("P_enes99") = "=sum(z20:z36)"
    bb.Range("P_enes98") = "=sum(z20:z36)*80"
'   血液検査
    bb.Range("P_bdate") = aa.Cells(2, 50)
    For i1 = 1 To 12
        bb.Range("P_bbl01").Offset(i1 - 1, 0) = aa.Cells(2, i1 + 50)
    Next i1
'   栄養素摂取バランス  i1:栄養素Index 1〜27  i2:行Index 1〜24
    For i1 = 1 To 27
        i2 = Fld_Field(i1, 24)
        If i2 > 0 Then
            bb.Range("i44").Offset(i2, 0) = aa.Range("e7").Offset(0, (i1 * 5 - 5))
            bb.Range("l44").Offset(i2, 0) = aa.Range("d7").Offset(0, (i1 * 5 - 5))
            bb.Range("o44").Offset(i2, 0) = bb.Range("l44").Offset(i2, 0) - bb.Range("i44").Offset(i2, 0)
            i3 = aa.Range("f7").Offset(0, (i1 * 5 - 5))
            If i3 > -62.5 Then
'                bb.Range("r44").Offset(i2, 0) = String((i3 + 62.5) * 52 / 125, "*")
                bb.Range("r44").Offset(i2, 0) = String((i3 + 62.5) * 52 / 250, "■")
            End If
            bb.Range("al44").Offset(i2, 0) = i3 + 100
            If i3 < -37 And Fld_Field(i1, 22) = 1 Then
                For i4 = 0 To 9
                    If IsEmpty(bb.Range("ap51").Offset(i4, 0)) Then
                        bb.Range("ap51").Offset(i4, 0) = Fld_Field(i1, 4)
                        Exit For
                    ElseIf IsEmpty(bb.Range("av51").Offset(i4, 0)) Then
                        bb.Range("av51").Offset(i4, 0) = Fld_Field(i1, 4)
                        Exit For
                    End If
                Next i4
            ElseIf i3 > 37 And Fld_Field(i1, 23) = 1 Then
                For i4 = 0 To 9
                    If IsEmpty(bb.Range("bc51").Offset(i4, 0)) Then
                        bb.Range("bc51").Offset(i4, 0) = Fld_Field(i1, 4)
                        Exit For
                    ElseIf IsEmpty(bb.Range("bi51").Offset(i4, 0)) Then
                        bb.Range("bi51").Offset(i4, 0) = Fld_Field(i1, 4)
                        Exit For
                    End If
                Next i4
            End If
            
        End If
    Next i1
'   栄養素比率
    bb.Range("P_Per01") = aa.Cells(2, 78)
    bb.Range("P_Per02") = aa.Cells(2, 79)
    bb.Range("P_Per03") = aa.Cells(2, 80)
    bb.Range("P_Per04") = aa.Cells(2, 81)
    bb.Range("P_Per05") = aa.Cells(2, 82)
    bb.Range("P_Per06") = aa.Cells(2, 83)
    bb.Range("P_Per07") = aa.Cells(2, 84) / 100
    bb.Range("P_Per08") = aa.Cells(2, 85) / 100
    
    Set aa = Nothing
    Set bb = Nothing
End Function
'--------------------------------------------------------------------------------
'   01_704 アドバイス項目の設定
'--------------------------------------------------------------------------------
Function Eiyo01_704Advic()
Dim Wkey        As Variant
Dim Wtext       As String
Dim i1          As Long
Dim Wadvic1(13) As String
Dim Wadvic2(13) As String
Dim Wadvic3(5)  As String

    Wtext = Empty
    For i1 = 1 To 9
        Wtext = Wtext & vbTab & Sheets("DBmirror").Range("ch2").Offset(0, i1 - 1)
    Next i1
    For i1 = 10 To 13
        Wtext = Wtext & vbTab & "38" & Format(Sheets("DBmirror").Range("ch2").Offset(0, i1 - 1), "00")
    Next i1
    Wkey = Split(Wtext, vbTab)
    
    With Rst_Advic
        .Index = "PrimaryKey"
        Rst_Advic.Open Source:=Tbl_Advic, ActiveConnection:=myCon, _
            CursorType:=adOpenKeyset, LockType:=adLockOptimistic, _
            Options:=adCmdTableDirect
        For i1 = 1 To 13
            .Seek Wkey(i1)
            If .EOF Then
                Wadvic1(i1) = Empty
                Wadvic2(i1) = Empty
            Else
                Wadvic1(i1) = .Fields(1).Value
                .Seek Wkey(i1) + 1
                If .EOF Or Wkey(i1) = 98 Then
                    Wadvic2(i1) = Empty
                Else
                    Wadvic2(i1) = .Fields(1).Value
                End If
            End If
        Next i1
'       ウエイト・アドバイス
        For i1 = 1 To 5
            Wadvic3(i1) = Empty
        Next i1
        Wkey = Sheets("DBmirror").Range("cu2")
        If Wkey <> 0 Then
            .Seek Wkey
            If Not .EOF Then
                Wadvic3(1) = Mid(.Fields(1).Value, 1, 20)
                Wadvic3(2) = Mid(.Fields(1).Value, 21, 20)
                Wadvic3(3) = Mid(.Fields(1).Value, 41, 20)
                If Wkey = 2803 Then
                    .Seek Wkey + 1
                    If Not .EOF Then
                        Wadvic3(3) = Mid(.Fields(1).Value, 1, 20)
                        Wadvic3(4) = Mid(.Fields(1).Value, 21, 20)
                    End If
                End If
            End If
        End If
    End With
    Set Rst_Advic = Nothing
    
'   検査
    If Sheets("DBmirror").Range("ck2") = 3330 Then
        Wadvic2(4) = Empty
        For i1 = 3 To 10
            If Mid(Range("q2"), i1, 1) = "1" Then
                Select Case i1
                    Case 3: Wadvic2(4) = Wadvic2(4) & "胸部Ｘ線　呼吸機能　"
                    Case 4: Wadvic2(4) = Wadvic2(4) & "心電図　血圧　脈拍　"
                    Case 5: Wadvic2(4) = Wadvic2(4) & "消化器系　"
                    Case 6: Wadvic2(4) = Wadvic2(4) & "血液　"
                    Case 7: Wadvic2(4) = Wadvic2(4) & "痔　"
                    Case 8: Wadvic2(4) = Wadvic2(4) & "肝機能　"
                    Case 9: Wadvic2(4) = Wadvic2(4) & "耳鼻咽喉　"
                    Case 10: Wadvic2(4) = Wadvic2(4) & "眼科　"
                End Select
            End If
        Next i1
    End If
    
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("an13") = Left(Wadvic1(1), 18)      '食生活／習慣
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("an14") = Mid(Wadvic1(1), 19, 18)
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("an15") = Mid(Wadvic1(1), 37, 18)
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("aj16") = Left(Wadvic2(1), 22)
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("aj17") = Mid(Wadvic2(1), 23, 22)
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("an20") = Left(Wadvic1(2), 18)      '睡眠と休養
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("an21") = Mid(Wadvic1(2), 19, 18)
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("an22") = Mid(Wadvic1(2), 37, 18)
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("aj23") = Left(Wadvic2(2), 22)
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("aj24") = Mid(Wadvic2(2), 23, 22)
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("an27") = Left(Wadvic1(3), 18)      '運動
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("an28") = Mid(Wadvic1(3), 19, 18)
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("an29") = Mid(Wadvic1(3), 37, 18)
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("aj30") = Left(Wadvic2(3), 22)
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("aj31") = Mid(Wadvic2(3), 23, 22)
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("an34") = Left(Wadvic1(4), 18)      '健康状態
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("an35") = Mid(Wadvic1(4), 19, 18)
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("an36") = Mid(Wadvic1(4), 37, 18)
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("aj37") = Left(Wadvic1(5), 22)
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("aj38") = Mid(Wadvic1(5), 23, 22)
    
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("bc17") = Wadvic3(1)
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("bc18") = Wadvic3(2)
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("bc19") = Wadvic3(3)
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("bc20") = Wadvic3(4)
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("bc21") = Wadvic3(5)
    
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("bc63") = Left(Wadvic1(6), 18)
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("bc64") = Mid(Wadvic1(6), 19, 18)
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("bc65") = Left(Wadvic2(6), 18)
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("bc66") = Mid(Wadvic2(6), 19, 18)
    
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("u73") = Wadvic1(12)
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("u74") = Wadvic2(12)
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("u75") = Wadvic1(13)
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("u76") = Wadvic2(13)
End Function
'--------------------------------------------------------------------------------
'   01_705 コード＆日付
'--------------------------------------------------------------------------------
Function Eiyo01_705Footer()
Dim Wtext   As String
    Wtext = "(" & Sheets("基礎").Range("g3") & ":" & Format(Date, "yymmdd") & ")"
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("b80") = Wtext
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("bd75") = Sheets("DBmirror").Range("db2")
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("bd76") = Sheets("DBmirror").Range("dc2")
    Sheets("ｶｳﾝｾﾘﾝｸﾞｼｰﾄ").Range("bd77") = Sheets("DBmirror").Range("dd2")
End Function
'--------------------------------------------------------------------------------
'   01_810 基礎画面作成
'--------------------------------------------------------------------------------
Function Eiyo01_810基礎画面作成()
Const PgmName = "Eiyo01_基礎摂食入力.xls"
Const ShtName = "基礎"
Dim i1      As Long
Dim i2      As Long
Dim FldItem As Variant

    If ActiveWorkbook.Name <> PgmName Then
        MsgBox PgmName & " ではありません"
        End
    End If
    If ActiveSheet.Name <> ShtName Then
        MsgBox ShtName & " ではありません"
        End
    End If
    Call Eiyo01_000init
'   画面の作成
    Call Eiyo930Screen_Hold                 '画面抑止ほか
    While (ActiveSheet.Shapes.Count > 0)    'コマンドボタン取消
        ActiveSheet.Shapes(1).Cut
    Wend
    Cells.Delete Shift:=xlUp                '全消去
    Cells.NumberFormatLocal = "@"           '全画面文字列属性
    Cells.Select
    With Selection.Font                     '文字フォント
        .Name = "ＭＳ ゴシック"
        .Size = 11
    End With
    Selection.ColumnWidth = 1.75            '列幅
    Selection.Interior.ColorIndex = 40      '全画面背景色（淡燈）
    
'   表題
    Range("G1:AA1").Select
    Selection.MergeCells = True                 '表題セル連結
    Selection.HorizontalAlignment = xlCenter    '表題センタリング
    Selection.Interior.ColorIndex = 37          '表題色（ペールブルー）
    With Selection.Font                         'フォント
        .FontStyle = "太字"
        .Size = 16
    End With
    Range("G01") = "栄養計算（基礎）２７栄養素版"
    Range("A01").VerticalAlignment = xlTop
    Range("A01") = "v-01"
    Range("A03") = "コード"
    Range("A04") = "調査期間"
    Range("A05") = "氏名"
    Range("A06") = "性別"
    Range("i06") = "(0:男 1:女)"
    Range("A07") = "生年月日"
    Range("a08") = "身長"
    Range("j08") = "cm"
    Range("a09") = "体重"
    Range("j09") = "Kg"
    Range("a10") = "皮下脂肪"
    Range("j10") = "cm"
    Range("A11") = "郵便番号"
    Range("A12") = "住所ー１"
    Range("A13") = "住所ー２"
    Range("A14") = "地域"
    Range("a15") = "都道府県"
    Range("a16") = "3.食習"
    Range("a17") = "4.休養"
    Range("a18") = "5.運動"
    Range("a19") = "6.健康"
    Range("m08") = "7.職業"
    Range("m09") = "A.主婦"
    Range("m10") = "B.妊娠"
    Range("m11") = "C.糖尿"
    Range("m14") = "D.高血圧"
    Range("m15") = "E.ｽﾎﾟｰﾂ"
    Range("m16") = "F.運動部"
    Range("m17") = "*.喫煙"
    Range("m18") = "G.身障害"
    Range("m19") = "H.ｳｴｲﾄCT"
    Range("m20") = "ｴﾈﾙｷﾞｰ指定"
    Range("M20").Characters(Start:=7, Length:=2).Font.Size = 9
    Range("m21") = "ﾀﾝﾊﾟｸ 指定"
    Range("M21").Characters(Start:=7, Length:=2).Font.Size = 9
    Range("m22") = "(0:無 1:指定 2:実体重 3:標準体重)"
    Range("m23") = "ｶｳﾝｾﾗｰ1"
    Range("m24") = "ｶｳﾝｾﾗｰ2"    
    Range("m25") = "ｶｳﾝｾﾗｰ3"
    
    Range("x03") = "血液型"
    Range("x04") = "支社部CD"
    Range("x05") = "保健記号"
    Range("x06") = "      No"
    Range("x07") = "定期健診"
    Range("x08") = "腕(L,R)"
    Range("x09") = "検査日"
    Range("x10") = "赤血球数"
    Range("x11") = "血色素量"
    Range("x12") = "ﾍﾏﾄｸﾘｯﾄ"
    Range("x13") = "ｺﾚｽﾃﾛｰﾙ"
    Range("x14") = "HDL"
    Range("x15") = "中性脂肪"
    Range("x16") = "G.O.T."
    Range("x17") = "G.P.T."
    Range("x18") = "尿酸"
    Range("x19") = "血糖"
    Range("x20") = "血圧最高"
    Range("x21") = "    最低"
    
    Cells.Locked = True                             '全セルをロック
    For i1 = 0 To UBound(Fld_Adrs1)
        FldItem = Split(Fld_Adrs1(i1), ",")
        Range(Trim(FldItem(1))).Select
        Selection.MergeCells = True                 'セル結合
        Range(Left(FldItem(1), 4)).Name = Trim(FldItem(0))
        If FldItem(2) = "i" Then
            With Selection.Borders                      '入力項目の枠罫線
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .Weight = xlThin
            End With
            Selection.Interior.ColorIndex = xlNone      '入力項目の白抜き化
            Selection.Locked = False                    '入力項目の保護解除
        End If
        Select Case FldItem(4)
            Case "90": Selection.NumberFormatLocal = "G/標準"
            Case "91": Selection.NumberFormatLocal = "#0.0"
            Case "92": Selection.NumberFormatLocal = "#0.00"
            Case "Ds"
                Selection.NumberFormatLocal = "yyyy/mm/dd"
                Selection.HorizontalAlignment = xlLeft
            Case "Dw"
                Selection.NumberFormatLocal = "gy.m.d"
                Selection.HorizontalAlignment = xlLeft
            Case "J "
                With Selection.Validation           '漢字項目
                    .Delete
                    .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .InputTitle = ""
                    .ErrorTitle = ""
                    .InputMessage = ""
                    .ErrorMessage = ""
                    .IMEMode = xlIMEModeOn
                    .ShowInput = True
                    .ShowError = True
                End With
        End Select
        Selection.Value = FldItem(5)
    Next i1
    
    Range("Fsave").Font.ColorIndex = 40
    Range("Gyyyy").NumberFormatLocal = "gy"     '生年月日の和暦年表示
    Range("Gyyyy") = "=RC[-5]"
    Range("Age").NumberFormatLocal = "G/標準"
    Range("Age") = "=DATEDIF(RC[-7],R[-3]C[-7],""y"")"
    Range("p07") = "才"

    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=10, Top:=350, Width:=60, Height:=30)
        .Object.Caption = "画面消去"
        .Name = "クリア"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=100, Top:=350, Width:=60, Height:=30)
        .Object.Caption = "ﾃﾞｰﾀ呼出"
        .Name = "検索"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=170, Top:=350, Width:=60, Height:=30)
        .Object.Caption = "ﾃﾞｰﾀ登録"
        .Name = "更新"
    End With        
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=240, Top:=350, Width:=60, Height:=30)
        .Object.Caption = "ｶｳﾝｾﾘﾝｸﾞ" & vbLf & "ｼｰﾄ作表"
        .Name = "作表"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=330, Top:=450, Width:=60, Height:=30)
        .Object.Caption = "ﾃﾞｰﾀ取消"
        .Name = "取消"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=400, Top:=350, Width:=60, Height:=30)
        .Object.Caption = "終了"
        .Name = "終了"
    End With
    
    Range("Gmesg").Font.Bold = True                            'メッセージエリア
    Range("Gmesg").Font.ColorIndex = 3
    Cells.FormatConditions.Delete               'シート全体から条件付き書式を削除する
    Cells.FormatConditions.Add Type:=xlExpression, Formula1:="=AND(CELL(""row"")=ROW(),CELL(""col"")=COLUMN())"
    Cells.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Cells.FormatConditions(1).Interior.Color = 255
    
    Call Eiyo01_820操作ガイド
    Range("g4").Select
    Call Eiyo940Screen_Start    '画面描画ほか
End Function
'--------------------------------------------------------------------------------
'   01_820 操作ガイド
'--------------------------------------------------------------------------------
Function Eiyo01_820操作ガイド()
    Call Eiyo930Screen_Hold     '画面抑止ほか
    Columns("ah:hz").Delete Shift:=xlToLeft
    Range("ah01") = "1.人を検索する"
    Range("ah02") = "　画面のいづれかの項目に"
    Range("ah03") = "　入力後、「検索」を押下してください。"
    Range("ah04") = "　同名など複数該当者の場合は、右側の一覧から"
    Range("ah05") = "　コードをダブルクリックして選択します。"
    Range("ah07") = "　「検索」は原則『前方一致』です、"
    Range("ah08") = "　先頭に[%]を付けると『含む』になります。"
    Range("ah10") = "2.人を登録する"
    Range("ah11") = "　画面の各項目を入力し"
    Range("ah12") = "　「更新」を押下してください。"
    Range("ah14") = "3.人の変更・取消"
    Range("ah15") = "　人を検索し、"
    Range("ah16") = "　修正後に「更新」または「取消」を押下してください。"
    Range("ah18") = "4.摂食の登録"
    Range("ah19") = "　人の登録または照会後「摂食」シートに切り替えてください"
    Call Eiyo940Screen_Start    '画面描画ほか
End Function
'--------------------------------------------------------------------------------
'   01_830 摂食画面作成
'--------------------------------------------------------------------------------
Function Eiyo01_830摂食画面作成()
Const PgmName = "Eiyo01_基礎摂食入力.xls"
Const ShtName = "摂食"
Dim i1      As Long
Dim i2      As Long
Dim FldItem As Variant

    If ActiveWorkbook.Name <> PgmName Then
        MsgBox PgmName & " ではありません"
        End
    End If
    If ActiveSheet.Name <> ShtName Then
        MsgBox ShtName & " ではありません"
        End
    End If
    Call Eiyo01_000init
'   画面の作成
    Call Eiyo930Screen_Hold     '画面抑止ほか
    ActiveWindow.FreezePanes = False        'ウインド枠固定の解除
    
    While (ActiveSheet.Shapes.Count > 0)    'コマンドボタン取消
        ActiveSheet.Shapes(1).Cut
    Wend
    Cells.Delete Shift:=xlUp                '全消去
    Cells.Select
    With Selection.Font                     '文字フォント
        .Name = "ＭＳ ゴシック"
        .Size = 11
    End With
    Cells.Interior.ColorIndex = 34          '全画面背景色（淡緑）
    Columns("A:B").Interior.ColorIndex = xlNone
    Columns("d").Interior.ColorIndex = xlNone
    Columns("f").Interior.ColorIndex = xlNone
    Rows("1:4").Interior.ColorIndex = 34       '全画面背景色（淡緑）
    Cells.Select                               '罫線
    With Selection.Borders
        .LineStyle = xlContinuous
        .ColorIndex = 40
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 40
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 40
        .Weight = xlThin
    End With
    Rows("1:3").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
'    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
'   表題
    With Range("d1").Font                         'フォント
        .FontStyle = "太字"
        .Size = 16
    End With
    Range("D01") = "栄養計算（摂食）"
    Range("A01").VerticalAlignment = xlTop
    Range("A01") = "v-01"
'    Range("a2") = "=Fcode & ":" & Namej
    Columns("A:A").NumberFormatLocal = "yyyy/mm/dd"
    Columns("F:F").NumberFormatLocal = "0.0 "
    Range("A4") = "摂食日"
    Range("B4") = "食区分"
    Range("D4") = "食品CD"
    Range("E4") = "品名・材料　ｻﾌﾟﾘﾒﾝﾄ"
    Range("F4") = "摂取量"
    Range("k3") = "・食品CD欄をダブルクリックするとメニューに変わります"
    Range("k2") = "・摂取量がゼロの行は削除されます"
    Range("k3") = "・追加は最終行の後ろに入力してください"
    Range("k1:k2").Font.Size = 9
    Call Eiyo01_840食品マスタ(4, 6)    
    Columns("A:D").HorizontalAlignment = xlCenter
    Rows("1:2").HorizontalAlignment = xlGeneral
    Range("a1:a2,B4").HorizontalAlignment = xlGeneral
    Range("F4").HorizontalAlignment = xlCenter
    Columns("A:A").ColumnWidth = 10.88
    Columns("B:B").ColumnWidth = 2.25
    Columns("C:C").ColumnWidth = 3.25
    Columns("D:D").ColumnWidth = 7
    Columns("E:E").ColumnWidth = 20
    Columns("F:F").ColumnWidth = 7
    Columns("J:J").ColumnWidth = 4.25
    Columns("K:K").ColumnWidth = 20
    Range("G:I").EntireColumn.Hidden = True
    
    Cells.Locked = True                             '全セルをロック
    Range("A:B,D:D,F:F").Locked = False             '入力列の解除
    Rows("1:3").Locked = True                       '表題行のロック
    
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=300, Top:=5, Width:=50, Height:=25)
        .Object.Caption = "登録"
        .Name = "登録"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=650, Top:=5, Width:=50, Height:=25)
        .Object.Caption = "検証"
        .Name = "検証"
    End With

    Range("a3").Font.Bold = True                    'メッセージエリア
    Range("a3").Font.ColorIndex = 3
    Range("E5").Select
    ActiveWindow.FreezePanes = True                 'ウインド枠固定の設定
    
'    ActiveSheet.Protect UserInterfaceOnly:=True     '保護を有効にする
    Range("g4").Select
    Call Eiyo940Screen_Start                        '画面描画ほか
End Function
'--------------------------------------------------------------------------------
'   01_840 食品マスタ項題
'--------------------------------------------------------------------------------
Function Eiyo01_840食品マスタ(il As Long, ic As Long)
Dim Wtext   As String
Dim Warray  As Variant
Dim i1      As Long

    Wtext = Empty
    Wtext = Wtext & vbLf & "コード"
    Wtext = Wtext & vbLf & "食品名"
    Wtext = Wtext & vbLf & "読み（分類）"
    Wtext = Wtext & vbLf & "単位"           '入力単位
    Wtext = Wtext & vbLf & "コメント"
    Wtext = Wtext & vbLf & "登録単位"
    Wtext = Wtext & vbLf & "換算係数"
    Wtext = Wtext & vbLf & "ﾒﾆｭｰ位置１"
    Wtext = Wtext & vbLf & "ﾒﾆｭｰ位置２"
    Wtext = Wtext & vbLf & "食酒"           '0:食 1:酒 2:ｻﾌﾟﾘﾒﾝﾄ
    Wtext = Wtext & vbLf & "摂食範囲下限"
    Wtext = Wtext & vbLf & "摂食範囲上限"
    Wtext = Wtext & vbLf & "栄養素-01"
    Wtext = Wtext & vbLf & "栄養素-02"
    Wtext = Wtext & vbLf & "栄養素-03"
    Wtext = Wtext & vbLf & "栄養素-04"
    Wtext = Wtext & vbLf & "栄養素-05"
    Wtext = Wtext & vbLf & "栄養素-06"
    Wtext = Wtext & vbLf & "栄養素-07"
    Wtext = Wtext & vbLf & "栄養素-08"
    Wtext = Wtext & vbLf & "栄養素-09"
    Wtext = Wtext & vbLf & "栄養素-10"
    Wtext = Wtext & vbLf & "栄養素-11"
    Wtext = Wtext & vbLf & "栄養素-12"
    Wtext = Wtext & vbLf & "栄養素-13"
    Wtext = Wtext & vbLf & "栄養素-14"
    Wtext = Wtext & vbLf & "栄養素-15"
    Wtext = Wtext & vbLf & "栄養素-16"
    Wtext = Wtext & vbLf & "栄養素-17"
    Wtext = Wtext & vbLf & "栄養素-18"
    Wtext = Wtext & vbLf & "栄養素-19"
    Wtext = Wtext & vbLf & "栄養素-20"
    Wtext = Wtext & vbLf & "栄養素-21"
    Wtext = Wtext & vbLf & "栄養素-22"
    Wtext = Wtext & vbLf & "栄養素-23"
    Wtext = Wtext & vbLf & "栄養素-24"
    Wtext = Wtext & vbLf & "栄養素-25"
    Wtext = Wtext & vbLf & "栄養素-26"
    Wtext = Wtext & vbLf & "栄養素-27"
    Wtext = Wtext & vbLf & "ENE/C 01"
    Wtext = Wtext & vbLf & "ENE/C 02"
    Wtext = Wtext & vbLf & "ENE/C 03"
    Wtext = Wtext & vbLf & "ENE/C 04"
    Wtext = Wtext & vbLf & "ENE/C 05"
    Wtext = Wtext & vbLf & "ENE/C 06"
    Wtext = Wtext & vbLf & "ENE/C 07"
    Wtext = Wtext & vbLf & "ENE/C 08"
    Wtext = Wtext & vbLf & "ENE/C 09"
    Wtext = Wtext & vbLf & "ENE/C 10"
    Wtext = Wtext & vbLf & "ENE/C 11"
    Wtext = Wtext & vbLf & "ENE/C 12"
    Wtext = Wtext & vbLf & "ENE/C 13"
    Wtext = Wtext & vbLf & "ENE/C 14"
    Wtext = Wtext & vbLf & "ENE/C 15"
    Wtext = Wtext & vbLf & "ENE/W 01"
    Wtext = Wtext & vbLf & "ENE/W 02"
    Wtext = Wtext & vbLf & "ENE/W 03"
    Wtext = Wtext & vbLf & "ENE/W 04"
    Wtext = Wtext & vbLf & "ENE/W 05"
    Wtext = Wtext & vbLf & "ENE/W 06"
    Wtext = Wtext & vbLf & "ENE/W 07"
    Wtext = Wtext & vbLf & "ENE/W 08"
    Wtext = Wtext & vbLf & "ENE/W 09"
    Wtext = Wtext & vbLf & "ENE/W 10"
    Wtext = Wtext & vbLf & "ENE/W 11"
    Wtext = Wtext & vbLf & "ENE/W 12"
    Wtext = Wtext & vbLf & "ENE/W 13"
    Wtext = Wtext & vbLf & "ENE/W 14"
    Wtext = Wtext & vbLf & "ENE/W 15"
    Wtext = Wtext & vbLf & "CL 01"
    Wtext = Wtext & vbLf & "CL 02"
    Wtext = Wtext & vbLf & "CL 03"
    Wtext = Wtext & vbLf & "CL 04"
    Wtext = Wtext & vbLf & "CL 05"
    Wtext = Wtext & vbLf & "CL 06"
    Wtext = Wtext & vbLf & "CL 07"
    Wtext = Wtext & vbLf & "CL 08"
    Wtext = Wtext & vbLf & "CL 09"
    Wtext = Wtext & vbLf & "CL 10"
    Wtext = Wtext & vbLf & "CL 11"
    Wtext = Wtext & vbLf & "CL 12"
    Wtext = Wtext & vbLf & "CL 13"
    Wtext = Wtext & vbLf & "CL 14"
    Wtext = Wtext & vbLf & "CL 15"
    Wtext = Wtext & vbLf & "脂質・動物"
    Wtext = Wtext & vbLf & "脂質・魚介"
    Wtext = Wtext & vbLf & "脂質・植物"
    Warray = Split(Wtext, vbLf)
    For i1 = 1 To UBound(Warray)
        Cells(il, ic + i1) = Warray(i1)
    Next i1
End Function
'--------------------------------------------------------------------------------
'   01_900  毎朝一番にＤＢのコピーをとる。
'           タイムスタンプは前日の最終更新時刻となる
'--------------------------------------------------------------------------------
Function Eiyo01_900WorkbookOpen()
Dim F_name          As String   '検索したファイル名
Dim F_dbname_today  As String   'DB+本日
Dim F_dbname_min    As String   'DB+00000000
Dim F_dbname_max    As String   'DB+2週間前
Dim W_path          As String

    W_path = ThisWorkbook.Path & "BackUp"""
    F_dbname_today = W_path & "Eiyo_" & Format(Date, "yyyymmdd") & ".mdb"""
    F_dbname_min = "Eiyo_00000000.mdb"
    F_dbname_max = "Eiyo_" & Format(Date - 14, "yyyymmdd") & ".mdb"
    
    SetCurrentDirectory (W_path)            'Dir変更
    If Dir(F_dbname_today) = "" Then        '今日の保存ファイルが存在しない
        F_name = Dir("*", vbNormal)
        Do While F_name <> ""
            If (F_name > F_dbname_min And F_name < F_dbname_max) Then
               Kill F_name
            End If
            F_name = Dir                    ' 次のフォルダ名を返します。
        Loop
        FileCopy ThisWorkbook.Path & myFileName, F_dbname_today
    End If
End Function
'--------------------------------------------------------------------------------
'   03_030  クリアのボタン・クリック
'--------------------------------------------------------------------------------
Function Eiyo03_030クリアClick()
    Call Eiyo930Screen_Hold     '画面抑止ほか
    Range("b3:b11") = Empty
    Range("b12") = Empty
    Range("b13:b14") = Empty
    Range("g4:g30").ClearContents
    Range("j4:l18").ClearContents
    Range("j22:j24").ClearContents
    Range("a17") = Empty
    Columns("n:hz").Delete Shift:=xlToLeft

    Range("b3").Select
    Call Eiyo940Screen_Start    '画面描画ほか
End Function
'--------------------------------------------------------------------------------
'   03_100  検索_Click
'--------------------------------------------------------------------------------
Function Eiyo03_100検索Click()
Dim Wsql    As String
Dim i1      As Long

    Range("a17") = Empty
    For i1 = 3 To 14
        If Not IsEmpty(Cells(i1, 2)) Then: Exit For
    Next i1
    If i1 > 14 Then
        Range("a17") = "キーがありません"
        Exit Function
    End If
    Call Eiyo930Screen_Hold     '画面抑止ほか
    Columns("n:hz").Delete Shift:=xlToLeft
    
    Wsql = "SELECT * FROM " & Tbl_Food & " Where "
    Select Case i1
        Case 3:  Wsql = Wsql & "Foodc = " & StrConv(Range("b03"), vbNarrow)
        Case 4:  Wsql = Wsql & "Fname Like ""%" & Range("b04") & "%"""
        Case 5:  Wsql = Wsql & "Kyomi Like ""%" & Range("b05") & "%"""
        Case 6:  Wsql = Wsql & "Ftani = """ & Range("b06") & """"
        Case 7:  Wsql = Wsql & "Comme Like ""%" & Range("b07") & "%"""
        Case 8:  Wsql = Wsql & "Mtani = """ & Range("b08") & """"
        Case 9:  Wsql = Wsql & "Conve = " & Range("b09")
        Case 10: Wsql = Wsql & "Posi1 = """ & Range("b10") & """"
        Case 11: Wsql = Wsql & "Posi2 = """ & Range("b11") & """"
        Case 12: Wsql = Wsql & "Drink = """ & Range("b12") & """"
        Case 13: Wsql = Wsql & "Enlhl = " & Range("b13")
        Case 14: Wsql = Wsql & "Enlhh = " & Range("b14")
    End Select
    
    Call Eiyo91DB_Open      'DB Open
    Set Rst_Food = myCon.Execute(Wsql)
    If Rst_Food.EOF Then
        Range("a17") = "該当データはありません"
    Else
        With Rst_Food
            Range("n2").CopyFromRecordset Rst_Food  'レコード
            If IsEmpty(Range("n3")) Then            '該当が１件のとき
                For i1 = 1 To 12                    '画面項目の順次処理
                    Cells(i1 + 2, 2) = Cells(2, i1 + 13)
                Next i1
                For i1 = 13 To 39
                    Cells(i1 - 9, 7) = Cells(2, i1 + 13)
                Next i1
                For i1 = 40 To 54
                    Cells(i1 - 36, 10) = Cells(2, i1 + 13)
                    Cells(i1 - 36, 11) = Cells(2, i1 + 28)
                    Cells(i1 - 36, 12) = Cells(2, i1 + 43)
                Next i1
                Cells(22, 10) = Cells(2, 98)
                Cells(23, 10) = Cells(2, 99)
                Cells(24, 10) = Cells(2, 100)
                Columns("n:hz").Delete Shift:=xlToLeft
            Else
                For i1 = 1 To .Fields.Count                     'フィールド名
                    Cells(1, i1 + 13).Value = .Fields(i1 - 1).Name
                Next
                Columns("n:hz").EntireColumn.AutoFit           '幅
                i1 = Range("n1").End(xlDown).Row
                Range("N:N").Locked = False                     '入力列の解除
            End If
            .Close
        End With
    End If
    Set Rst_Food = Nothing              'オブジェクトの解放
    Call Eiyo920DB_Close                'DB Close
    Columns("J:L").EntireColumn.AutoFit
    Call Eiyo940Screen_Start            '画面描画ほか
End Function
'--------------------------------------------------------------------------------
'   03_110  前検索_Click
'--------------------------------------------------------------------------------
Function Eiyo03_110前検索Click()
Dim Wsql    As String
Dim Wkey    As Long

    Range("a17") = Empty
    Wkey = Range("b03")
    Call Eiyo930Screen_Hold             '画面抑止ほか
    Call Eiyo91DB_Open                  'DB Open
    Wsql = "SELECT Foodc FROM " & Tbl_Food & " Where Foodc < " & Wkey & " Order by Foodc DESC"
    Set Rst_Food = myCon.Execute(Wsql)
    If Rst_Food.EOF Then
        Range("a17") = "該当データはありません"
    Else
        With Rst_Food
            Range("b03") = .Fields(0).Value
            .Close
        End With
    End If
    Set Rst_Food = Nothing      'オブジェクトの解放
    Call Eiyo920DB_Close        'DB Close
    Call Eiyo03_100検索Click
End Function
'--------------------------------------------------------------------------------
'   03_120  次検索_Click
'--------------------------------------------------------------------------------
Function Eiyo03_120次検索Click()
Dim Wsql    As String
Dim Wkey    As Long

    Range("a17") = Empty
    Wkey = Range("b03")
    Call Eiyo930Screen_Hold             '画面抑止ほか
    Call Eiyo91DB_Open                  'DB Open
    Wsql = "SELECT Foodc FROM " & Tbl_Food & " Where Foodc > " & Wkey & " Order by Foodc"
    Set Rst_Food = myCon.Execute(Wsql)
    If Rst_Food.EOF Then
        Range("a17") = "該当データはありません"
    Else
        With Rst_Food
            Range("b03") = .Fields(0).Value
            .Close
        End With
    End If
    Set Rst_Food = Nothing      'オブジェクトの解放
    Call Eiyo920DB_Close        'DB Close
    Call Eiyo03_100検索Click
End Function
'--------------------------------------------------------------------------------
'   03_200  更新
'--------------------------------------------------------------------------------
Function Eiyo03_200更新Click()
Dim i1      As Long
Dim Wsw     As Long
Dim Wtemp   As Variant

    Wsw = 0
    Range("a17") = Empty
    If Range("b3") < 1 Then: Exit Function
    Call Eiyo91DB_Open                      'DB Open
    '準備ここまで
    With Rst_Food
        'インデックスの設定
        .Index = "PrimaryKey"
        'レコードセットを開く
        Rst_Food.Open Source:=Tbl_Food, ActiveConnection:=myCon, _
            CursorType:=adOpenKeyset, LockType:=adLockOptimistic, _
            Options:=adCmdTableDirect
        '番号が登録されているか検索する
        If Not .EOF Then .Seek Range("b3")
        If .EOF Then
            .AddNew
            For i1 = 1 To 87
                Select Case i1
                    Case 1 To 12: Wtemp = Cells(i1 + 2, 2)
                    Case 13 To 39: Wtemp = Cells(i1 - 9, 7)
                    Case 40 To 54: Wtemp = Cells(i1 - 36, 10)
                    Case 55 To 69: Wtemp = Cells(i1 - 51, 11)
                    Case 70 To 84: Wtemp = Cells(i1 - 66, 12)
                    Case 85 To 87: Wtemp = Cells(i1 - 63, 10)
                End Select
                .Fields(i1 - 1).Value = Wtemp
            Next i1
            .Update
            Range("a17") = "追加されました"
        Else
            For i1 = 2 To 87
                Select Case i1
                    Case 2 To 12: Wtemp = Cells(i1 + 2, 2)
                    Case 13 To 39: Wtemp = Cells(i1 - 9, 7)
                    Case 40 To 54: Wtemp = Cells(i1 - 36, 10)
                    Case 55 To 69: Wtemp = Cells(i1 - 51, 11)
                    Case 70 To 84: Wtemp = Cells(i1 - 66, 12)
                    Case 85 To 87: Wtemp = Cells(i1 - 63, 10)
                End Select
                If .Fields(i1 - 1).Value <> Wtemp Then
                    .Fields(i1 - 1).Value = Wtemp
                    Wsw = Wsw + 1
                End If
            Next i1
            If Wsw = 0 Then
                Range("a17") = "変更項目がありません"
            Else
                If MsgBox("変更項目は" & Wsw & "ヵ所です、更新してよろしいですか", vbOKCancel) = vbOK Then
                    .Update
                    Range("a17") = "更新されました"
                End If
            End If
        End If
'        .Close
    End With
    Set Rst_Food = Nothing      'オブジェクトの解放
    Call Eiyo920DB_Close        'DB Close
End Function
'--------------------------------------------------------------------------------
'   03_300  取消
'--------------------------------------------------------------------------------
Function Eiyo03_300取消Click()
    Range("a17") = Empty
    Call Eiyo91DB_Open      'DB Open
    '準備ここまで
    With Rst_Food
        'インデックスの設定
        .Index = "PrimaryKey"
        'レコードセットを開く
        Rst_Food.Open Source:=Tbl_Food, ActiveConnection:=myCon, _
            CursorType:=adOpenKeyset, LockType:=adLockOptimistic, _
            Options:=adCmdTableDirect
        '番号が登録されているか検索する
        If Not .EOF Then .Seek Range("b3")
        If .EOF Then
            Range("a17") = "キーが存在しません"
        Else
            If MsgBox("削除してよろしいですか", vbOKCancel) = vbOK Then
                .Delete
                Range("a17") = "取消されました"
            End If
        End If
        .Close
    End With
    Set Rst_Food = Nothing      'オブジェクトの解放
    Call Eiyo920DB_Close        'DB Close
End Function
'--------------------------------------------------------------------------------
'   03_400  食品コードデータの作成
'--------------------------------------------------------------------------------
Function Eiyo03_400コードClick()
Dim i1      As Long
Dim Wsql    As String

    Call Eiyo930Screen_Hold
    Call Eiyo91DB_Open      'DB Open
    
    Wsql = "SELECT Foodc,Fname FROM " & Tbl_Food & " Order by Foodc"
    Set Rst_Food = myCon.Execute(Wsql)
    If Rst_Food.EOF Then
        Range("a17") = "該当データはありません"
    Else
        With Rst_Food
            Range("AA1").CopyFromRecordset Rst_Food
            .Close
        End With
    End If
    Set Rst_Food = Nothing      'オブジェクトの解放
    Call Eiyo920DB_Close        'DB Close
    
    Open ThisWorkbook.Path & "食品コード.txt" For Output As #22
    For i1 = 1 To Range("aa60000").End(xlUp).Row
        Print #2, Cells(i1, 27) & vbTab & Cells(i1, 28)
    Next i1
    Close
    Columns("Aa:BZ").Delete Shift:=xlToLeft
    Call Eiyo940Screen_Start
End Function
'--------------------------------------------------------------------------------
'   03_810  シートの作成
'--------------------------------------------------------------------------------
Function Eiyo03_810食品sheet_make()
    Call Eiyo930Screen_Hold     '画面抑止ほか
    Call Eiyo03_811_init        'シートの初期化
    Call Eiyo03_812_zokusei     'キー、名称、属性
    Call Eiyo03_813_eiyoso      '栄養素
    Call Eiyo03_814_shokugun    '食品群
    Call Eiyo03_815_sisitu      '脂質
    Call Eiyo03_816_keisen      '罫線、列幅
    Call Eiyo03_817_button      'コマンド・ボタン
    Call Eiyo940Screen_Start    '画面描画ほか
End Function
'--------------------------------------------------------------------------------
'   03_811  シートの初期化
'--------------------------------------------------------------------------------
Function Eiyo03_811_init()
    Sheets("食品マスタ").Select
    While (ActiveSheet.Shapes.Count > 0)    'コマンドボタン取消
        ActiveSheet.Shapes(1).Cut
    Wend
    Cells.Delete Shift:=xlUp                '全消去
    Cells.NumberFormatLocal = "@"           '全画面文字列属性
    Range("e1") = "※　食品マスタ　照会・更新　※"
    Range("e1").Font.Size = 16
End Function
'--------------------------------------------------------------------------------
'   03_812  キー、名称、属性
'--------------------------------------------------------------------------------
Function Eiyo03_812_zokusei()
Dim i1      As Long
Dim Wtext   As String
Dim Warray  As Variant
'   属性ほか
    Wtext = Empty
    Wtext = Wtext & vbLf & "食品コード"
    Wtext = Wtext & vbLf & "食品名"
    Wtext = Wtext & vbLf & "読み(分類）"
    Wtext = Wtext & vbLf & "入力単位"
    Wtext = Wtext & vbLf & "摘要"
    Wtext = Wtext & vbLf & "登録単位"
    Wtext = Wtext & vbLf & "単位換算値"
    Wtext = Wtext & vbLf & "ﾒﾆｭｰ位置１"
    Wtext = Wtext & vbLf & "ﾒﾆｭｰ位置２"
    Wtext = Wtext & vbLf & "食品分類"
    Wtext = Wtext & vbLf & "摂食量下限"
    Wtext = Wtext & vbLf & "摂食量上限"
    Warray = Split(Wtext, vbLf)
    If UBound(Warray) <> 12 Then
        MsgBox "Program Error No.01 " & UBound(Warray)
        End
    End If
    For i1 = 1 To UBound(Warray)
        Cells(i1 + 2, 1) = Warray(i1)
    Next i1
    Range("c12") = "0:食品 1:酒類 2:ｻﾌﾟﾘﾒﾝﾄ"
'   セル結合
    Warray = Array(, 2, 3, 3, 2, 3, 2, 2, 2, 2, 1, 2, 2)
    For i1 = 1 To UBound(Warray)
        Range(Cells(i1 + 2, 2).Address & ":" & Cells(i1 + 2, Warray(i1) + 1).Address).Select
        Selection.MergeCells = True
        With Selection.Borders
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Interior                         '背景色
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        Selection.Locked = False                        '保護解除
        Selection.FormulaHidden = False
    Next i1
    Range("a17").Locked = False                        '保護解除
End Function
'--------------------------------------------------------------------------------
'   03_813  栄養素
'--------------------------------------------------------------------------------
Function Eiyo03_813_eiyoso()
Dim i1      As Long
Dim Wtext   As String
Dim Warray  As Variant
    Wtext = Empty
    Wtext = Wtext & vbLf & "01:エネルギー"
    Wtext = Wtext & vbLf & "02:たんぱく質"
    Wtext = Wtext & vbLf & "03:動物性たんぱく質"
    Wtext = Wtext & vbLf & "04:植物性たんぱく質"
    Wtext = Wtext & vbLf & "05:脂質"
    Wtext = Wtext & vbLf & "06:糖質"
    Wtext = Wtext & vbLf & "07:食物せんい"
    Wtext = Wtext & vbLf & "08:カルシウム"
    Wtext = Wtext & vbLf & "09:リン"
    Wtext = Wtext & vbLf & "10:鉄"
    Wtext = Wtext & vbLf & "11:ナトリウム"
    Wtext = Wtext & vbLf & "12:ビタミンＡ"
    Wtext = Wtext & vbLf & "13:ビタミンＢ１"
    Wtext = Wtext & vbLf & "14:ビタミンＢ２"
    Wtext = Wtext & vbLf & "15:ナイアシン"
    Wtext = Wtext & vbLf & "16:ビタミンＣ"
    Wtext = Wtext & vbLf & "17:ビタミンＢ６"
    Wtext = Wtext & vbLf & "18:パントテン酸"
    Wtext = Wtext & vbLf & "19:葉酸"
    Wtext = Wtext & vbLf & "20:ビタミンＥ"
    Wtext = Wtext & vbLf & "21:カリウム"
    Wtext = Wtext & vbLf & "22:マグネシウム"
    Wtext = Wtext & vbLf & "23:食塩"
    Wtext = Wtext & vbLf & "24:コレステロール"
    Wtext = Wtext & vbLf & "25:不飽和脂肪酸"
    Wtext = Wtext & vbLf & "26:飽和脂肪酸"
    Wtext = Wtext & vbLf & "27:砂糖"
    Warray = Split(Wtext, vbLf)
    If UBound(Warray) <> 27 Then
        MsgBox "Program Error No.02 " & UBound(Warray)
        End
    End If
    For i1 = 1 To UBound(Warray)
        Cells(i1 + 3, 6) = Warray(i1)
    Next i1
    Cells(3, 6) = "栄養素"
    Wtext = Empty
    Wtext = Wtext & vbLf & "Kcal"
    Wtext = Wtext & vbLf & "g"
    Wtext = Wtext & vbLf & "g"
    Wtext = Wtext & vbLf & "g"
    Wtext = Wtext & vbLf & "g"
    Wtext = Wtext & vbLf & "g"
    Wtext = Wtext & vbLf & "g"
    Wtext = Wtext & vbLf & "mg"
    Wtext = Wtext & vbLf & "mg"
    Wtext = Wtext & vbLf & "mg"
    Wtext = Wtext & vbLf & "mg"
    Wtext = Wtext & vbLf & "I.U."
    Wtext = Wtext & vbLf & "mg"
    Wtext = Wtext & vbLf & "mg"
    Wtext = Wtext & vbLf & "mg"
    Wtext = Wtext & vbLf & "mg"
    Wtext = Wtext & vbLf & "mg"
    Wtext = Wtext & vbLf & "mg"
    Wtext = Wtext & vbLf & "μg"
    Wtext = Wtext & vbLf & "mg"
    Wtext = Wtext & vbLf & "mg"
    Wtext = Wtext & vbLf & "mg"
    Wtext = Wtext & vbLf & "g"
    Wtext = Wtext & vbLf & "mg"
    Wtext = Wtext & vbLf & "mg"
    Wtext = Wtext & vbLf & "mg"
    Wtext = Wtext & vbLf & "g"
    Warray = Split(Wtext, vbLf)
    If UBound(Warray) <> 27 Then
        MsgBox "Program Error No.03 " & UBound(Warray)
        End
    End If
    For i1 = 1 To UBound(Warray)
        Cells(i1 + 3, 8) = Warray(i1)
    Next i1
End Function
'--------------------------------------------------------------------------------
'   03_814  食品群
'--------------------------------------------------------------------------------
Function Eiyo03_814_shokugun()
Dim i1      As Long
Dim Wtext   As String
Dim Warray  As Variant
    Wtext = Empty
    Wtext = Wtext & vbLf & "食品群"
    Wtext = Wtext & vbLf & "01:大豆製品"
    Wtext = Wtext & vbLf & "02:魚介類"
    Wtext = Wtext & vbLf & "03:肉　類"
    Wtext = Wtext & vbLf & "04:卵"
    Wtext = Wtext & vbLf & "05:海　草"
    Wtext = Wtext & vbLf & "06:乳製品"
    Wtext = Wtext & vbLf & "07:小　魚"
    Wtext = Wtext & vbLf & "08:緑黄色野菜"
    Wtext = Wtext & vbLf & "09:淡色野菜"
    Wtext = Wtext & vbLf & "10:果　物"
    Wtext = Wtext & vbLf & "11:穀　類"
    Wtext = Wtext & vbLf & "12:いも類"
    Wtext = Wtext & vbLf & "13:砂　糖"
    Wtext = Wtext & vbLf & "14:植物性油脂"
    Wtext = Wtext & vbLf & "15:動物性油脂"
    Warray = Split(Wtext, vbLf)
    If UBound(Warray) <> 16 Then
        MsgBox "Program Error No.04 " & UBound(Warray)
        End
    End If
    For i1 = 1 To UBound(Warray)
        Cells(i1 + 2, 9) = Warray(i1)
    Next i1
    Cells(3, 10) = "ｶﾛﾘｰ"
    Cells(3, 11) = "重量"
    Cells(3, 12) = "ｶﾙｼｳﾑ"
    Cells(19, 9) = "Total"
    Cells(20, 10) = "(kcal)"
    Cells(20, 11) = "(g)"
    Cells(20, 12) = "(g)"
    Range("j3:l3,i19,j20:l20").HorizontalAlignment = xlRight
End Function
'--------------------------------------------------------------------------------
'   03_815  脂質
'--------------------------------------------------------------------------------
Function Eiyo03_815_sisitu()
Dim i1      As Long
Dim Wtext   As String
Dim Warray  As Variant
    Wtext = Empty
    Wtext = Wtext & vbLf & "脂質(動物)"
    Wtext = Wtext & vbLf & "脂質(魚介)"
    Wtext = Wtext & vbLf & "脂質(植物)"
    Wtext = Wtext & vbLf & "g"
    Wtext = Wtext & vbLf & "g"
    Wtext = Wtext & vbLf & "g"
    Warray = Split(Wtext, vbLf)
    If UBound(Warray) <> 6 Then
        MsgBox "Program Error No.05 " & UBound(Warray)
        End
    End If
    For i1 = 1 To 3
        Cells(i1 + 21, 9) = Warray(i1)
        Cells(i1 + 21, 11) = Warray(i1 + 3)
    Next i1
    Range("I25") = "Total"
    Range("I22:I25").HorizontalAlignment = xlRight
End Function
'--------------------------------------------------------------------------------
'   03_816  罫線、列幅
'--------------------------------------------------------------------------------
Function Eiyo03_816_keisen()
Dim i1      As Long
Dim Wtext   As String
Dim Warray  As Variant
'   罫線
    Range("g4:g30,j4:l18,j22:j24").Select
    With Selection.Borders
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Interior                         '背景色
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Locked = False                        '保護解除
    Selection.FormulaHidden = False
'   列幅
    Columns("F:F").ShrinkToFit = True
    Cells.EntireColumn.AutoFit
    Warray = Array(, 0, 1.75, 5, 18, 2, 15, 0, 7)
    For i1 = 1 To UBound(Warray)
        If Warray(i1) > 0 Then: Columns(i1).ColumnWidth = Warray(i1)
    Next i1
'    Range("B4:C4,B6:C6,B11:C11,B12:C12").NumberFormatLocal = "#,##0.00;[赤]-#,##0.00"
    Range("B9,B13,B14").NumberFormatLocal = "#,##0.00;[赤]-#,##0.00"
    Range("g4:g30,j4:l19,j22:j25").NumberFormatLocal = "#,##0.00;[赤]-#,##0.00"
'   カーソル位置を明確化する
    Cells.FormatConditions.Delete               'シート全体から条件付き書式を削除する
    Cells.FormatConditions.Add Type:=xlExpression, Formula1:="=AND(CELL(""row"")=ROW(),CELL(""col"")=COLUMN())"
    Cells.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Cells.FormatConditions(1).Interior.Color = 255
'   合計の式と条件付き書式
    Range("J19") = "=SUM(J4:J18)"
    Range("k19") = "=SUM(k4:k18)"
    Range("l19") = "=SUM(l4:l18)"
    Range("j25") = "=SUM(j22:j24)"
    
    Wtext = Empty
    Wtext = Wtext & vbLf & "=G04<>J19"
    Wtext = Wtext & vbLf & "=G11<>L19"
    Wtext = Wtext & vbLf & "=G08<>J25"
    Wtext = Wtext & vbLf & "=J19<>G04"
    Wtext = Wtext & vbLf & "=L19<>G11"
    Wtext = Wtext & vbLf & "=J25<>G08"
    Warray = Split(Wtext, vbLf)
    For i1 = 1 To UBound(Warray)
        Range(Right(Warray(i1), 3)).Select
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:=Warray(i1)
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        Selection.FormatConditions(1).Interior.ColorIndex = 6
    Next i1

    Range("a17").Font.Bold = True           'メッセージエリア（赤太文字）
    Range("a17").Font.ColorIndex = 3
End Function
'--------------------------------------------------------------------------------
'   03_817 コマンド・ボタン
'--------------------------------------------------------------------------------
Function Eiyo03_817_button()
    While (ActiveSheet.Shapes.Count > 0)    'コマンドボタン取消
        ActiveSheet.Shapes(1).Cut
    Wend
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=10, Top:=250, Width:=50, Height:=30)
        .Object.Caption = "クリア"
        .Name = "クリア"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=70, Top:=250, Width:=50, Height:=30)
        .Object.Caption = "検索"
        .Name = "検索"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=130, Top:=250, Width:=50, Height:=30)
        .Object.Caption = "更新"
        .Name = "更新"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=10, Top:=300, Width:=50, Height:=30)
        .Object.Caption = "前検索"
        .Name = "前検索"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=70, Top:=300, Width:=50, Height:=30)
        .Object.Caption = "次検索"
        .Name = "次検索"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=130, Top:=300, Width:=50, Height:=30)
        .Object.Caption = "取消"
        .Name = "取消"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=70, Top:=350, Width:=50, Height:=30)
        .Object.Caption = "終了"
        .Name = "終了"
    End With
End Function
'--------------------------------------------------------------------------------
'   共通処理　Eiyo.mdb のオープン
'--------------------------------------------------------------------------------
Function Eiyo91DB_Open()
    myCon.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
               "Data Source=" & ThisWorkbook.Path & "Eiyo.mdb;"""
End Function
'--------------------------------------------------------------------------------
'   共通処理　Eiyo.mdb のクローズ
'--------------------------------------------------------------------------------
Function Eiyo920DB_Close()
    myCon.Close
    Set myCon = Nothing
End Function
'--------------------------------------------------------------------------------
'   共通処理　画面抑止ほか
'--------------------------------------------------------------------------------
Function Eiyo930Screen_Hold()
    Application.ScreenUpdating = False      '画面描画抑止
    Application.EnableEvents = False        'イベント発生抑止
    ActiveSheet.Unprotect                   'シートの保護を解除
End Function
'--------------------------------------------------------------------------------
'   共通処理　画面描画ほか
'--------------------------------------------------------------------------------
Function Eiyo940Screen_Start()
    Application.ScreenUpdating = True           '画面描画の復活
    Application.EnableEvents = True             'イベント発生再開
    ActiveSheet.Protect UserInterfaceOnly:=True '保護を有効にする
End Function
'--------------------------------------------------------------------------------
'   共通処理　ボタン作成
'--------------------------------------------------------------------------------
Function Eiyo950Button_Add(in_L As Long, in_t As Long, in_W As Long, in_H As Long, in_text As String)
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=in_L, Top:=in_t, Width:=in_W, Height:=in_H)
        .Object.Caption = in_text
        .Name = in_text
    End With
End Function
'--------------------------------------------------------------------------------
'   共通処理　指定シート削除
'--------------------------------------------------------------------------------
Function Eiyo99_指定シート削除(Sname As String)
    Application.DisplayAlerts = False                                   '確認抑止
    If Not IsError(Evaluate(Sname & "!a1")) Then: Sheets(Sname).Delete  'シート削除
    Application.DisplayAlerts = True                                    '確認復活
End Function

