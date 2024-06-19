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
'   Eiyo01_000 ‰æ–Ê€–Ú’è‹`
'   0:Field-Name
'   1:ƒZƒ‹”ÍˆÍ
'   2:i/o “ü—Í‰Â”Û‚ğ–¾Šm‰»
'   3:Field (D:DB G:Guid)
'   4:Type
'   5:Sample
'--------------------------------------------------------------------------------
Function Eiyo01_000init()
Dim Wtext   As String
    Wtext = "Gmesg,a025:a025,o,00,X ,Message"
    Wtext = Wtext & vbLf & "Fcode,g003:k003,i,D,X,1234567890"    'Fcode
    Wtext = Wtext & vbLf & "Fsave,l003:p003,o,G,90,1234567890"   'Fcode save
    Wtext = Wtext & vbLf & "Date1,g004:k004,i,D,Ds,2008/10/10"   '’²¸ŠúŠÔ©
    Wtext = Wtext & vbLf & "Nissu,l004:l004,i,D,90,1"            'ŠúŠÔ
    Wtext = Wtext & vbLf & "Namej,g005:o005,i,D,J ,–¼‚P–¼‚Q–¼‚R"
    Wtext = Wtext & vbLf & "Sex  ,g006:g006,i,D,X ,1"            '«•Ê"
    Wtext = Wtext & vbLf & "Birth,g007:k007,i,D,Ds,2001/1/1"     '¶”NŒ“ú"
    Wtext = Wtext & vbLf & "Gyyyy,l007:m007,o,G,X ,"             '˜a—ï¶”N
    Wtext = Wtext & vbLf & "Age  ,n007:o007,o,D,90,"             '”N—î"
    Wtext = Wtext & vbLf & "Hight,g008:i008,i,D,91,123.4"        'g’·
    Wtext = Wtext & vbLf & "Weght,g009:i009,i,D,91,123.4"        '‘Ìd
    Wtext = Wtext & vbLf & "Sibou,g010:i010,i,D,91,123.4"        '”ç‰º‰–b
    Wtext = Wtext & vbLf & "Adrno,g011:j011,i,D,X ,123-4567"     '—X•Ö”Ô†
    Wtext = Wtext & vbLf & "Adrs1,g012:v012,i,D,J ,ZŠ[‚P”ZŠ[‚P”ZŠ[‚P”ZŠ["
    Wtext = Wtext & vbLf & "Adrs2,g013:v013,i,D,J ,ZŠ[‚Q”ZŠ[‚Q”ZŠ[‚Q”ZŠ["
    Wtext = Wtext & vbLf & "Area1,g014:h014,o,D,X ,12"           '’n‹æ
    Wtext = Wtext & vbLf & "Gare1,i014:i014,o,G,X ,’nˆæ–¼"       '’nˆæ
    Wtext = Wtext & vbLf & "Area2,g015:h015,o,D,X ,12"           '’n‹æ
    Wtext = Wtext & vbLf & "Gare2,i015:i015,o,G,X ,“s•{Œ§"       '’nˆæ
    Wtext = Wtext & vbLf & "Q3rec,g016:k016,i,D,X ,1234567890"   'Q3.HK
    Wtext = Wtext & vbLf & "Q4rec,g017:i017,i,D,X ,12345"        'Q4.‹x—{
    Wtext = Wtext & vbLf & "Q5rec,g018:h018,i,D,X ,123"          'Q5.‰^“®
    Wtext = Wtext & vbLf & "Q6r_a,g019:k019,i,D,X ,1234567890"   'Q6.Œ’N-1
    Wtext = Wtext & vbLf & "Q6r_b,g020:k020,i,D,X ,1234567890"   'Q6.Œ’N-2
    Wtext = Wtext & vbLf & "Q6r_c,g021:k021,i,D,X ,1234567890"   'Q6.Œ’N-3
    Wtext = Wtext & vbLf & "Q6r_d,g022:k022,i,D,X ,1234567890"   'Q6.Œ’N-4
    Wtext = Wtext & vbLf & "Q6r_e,g023:k023,i,D,X ,1234567890"   'Q6.Œ’N-5
    Wtext = Wtext & vbLf & "Qjob1,q008:r008,i,D,X ,1234"         'Q7.E‹Æ-1
    Wtext = Wtext & vbLf & "Qjob5,s008:s008,i,D,X ,1"            'Q7.E‹Æ-2
    Wtext = Wtext & vbLf & "Qsyuf,q009:q009,i,D,X ,1"            'Qa.å•w
    Wtext = Wtext & vbLf & "Qcnd1,q010:q010,i,D,X ,1"            'Qb.”DP
    Wtext = Wtext & vbLf & "Qtony,q011:q011,i,G,X ,1"            'Qc.“œ”A
    Wtext = Wtext & vbLf & "Qill1,r011:r011,o,D,X ,123456"       'Qc.“œ”A
    Wtext = Wtext & vbLf & "Qkoke,q014:q014,i,G,X ,1"            'Qd.‚ŒŒˆ³
    Wtext = Wtext & vbLf & "Qill2,r014:r014,o,D,X ,123456"       'Qd.‚ŒŒˆ³
    Wtext = Wtext & vbLf & "Qsrmr,q015:r015,i,D,90,123"          'Qe.Spot-1
    Wtext = Wtext & vbLf & "Qsmin,s015:t015,i,D,90,123"          'Qe.Spot-2
    Wtext = Wtext & vbLf & "Qclab,q016:q016,i,D,90,1"            'Qf.‰^“®•”
    Wtext = Wtext & vbLf & "Qtobc,q017:q017,i,D,90,1"            'Qt.‹i‰Œ
    Wtext = Wtext & vbLf & "Qsyog,q018:r018,i,D,90,12"           'Qg.g‘ÌáŠQ
    Wtext = Wtext & vbLf & "Qwcnt,q019:q019,i,D,90,1"            'Qh.³´²ÄCT
    Wtext = Wtext & vbLf & "Tenes,q020:q020,i,D,90,1"            'Qi.´ÈÙ·Ş°w’è-1
    Wtext = Wtext & vbLf & "Tenee,r020:u020,i,D,92,12345.67"     'Qi.´ÈÙ·Ş°w’è-2
    Wtext = Wtext & vbLf & "Tanps,q021:q021,i,D,90,1"            'Qj.ÀİÊß¸ w’è-1
    Wtext = Wtext & vbLf & "Tanpe,r021:u021,i,D,92,12345.67"     'Qj.ÀİÊß¸ w’è-2
    Wtext = Wtext & vbLf & "¶³İ¾×1,q023:af23,i,D,J ,¶³İ¾×1"      '
    Wtext = Wtext & vbLf & "¶³İ¾×2,q024:af24,i,D,J ,¶³İ¾×2"      '
    Wtext = Wtext & vbLf & "¶³İ¾×3,q025:af25,i,D,J ,¶³İ¾×3"      '
    Wtext = Wtext & vbLf & "Blood,ab03:ac03,i,D,X ,12"           'B1.ŒŒ‰tŒ^
    Wtext = Wtext & vbLf & "Bscd1,ab04:ac04,i,D,X ,123"          'B2.xĞ•”-1
    Wtext = Wtext & vbLf & "Bscd2,ad04:ae04,i,D,X ,12"           'B2.xĞ•”-2
    Wtext = Wtext & vbLf & "Bhok1,ab05:ae05,i,D,X ,12345678"     'B3.•ÛŒ’‹L†
    Wtext = Wtext & vbLf & "Bhok2,ab06:ae06,i,D,X ,12345678"     'B4.•ÛŒ’”Ô†
    Wtext = Wtext & vbLf & "Bhant,ab07:ac07,i,D,X ,12"           'B5.’èŠúŒ’f”»’è
    Wtext = Wtext & vbLf & "Barm ,ab08:ab08,i,D,X ,1"            'B6.ŒŸ¸˜r
    Wtext = Wtext & vbLf & "Bdate,ab09:af09,i,D,Ds,2008/10/10"   'B7.ŒŸ¸“ú
    Wtext = Wtext & vbLf & "Bbl01,ab10:ad10,i,D,91,123.41"       'B8.ÔŒŒ‹…”
    Wtext = Wtext & vbLf & "Bbl02,ab11:ad11,i,D,91,123.41"       'B8.ŒŒF‘f—Ê
    Wtext = Wtext & vbLf & "Bbl03,ab12:ad12,i,D,91,123.41"       'B8.ÍÏÄ¸Ø¯Ä
    Wtext = Wtext & vbLf & "Bbl04,ab13:ad13,i,D,91,123.41"       'B8.ºÚ½ÃÛ°Ù
    Wtext = Wtext & vbLf & "Bbl05,ab14:ad14,i,D,91,123.41"       'B8.HDL
    Wtext = Wtext & vbLf & "Bbl06,ab15:ad15,i,D,91,123.41"       'B8.’†«‰–b
    Wtext = Wtext & vbLf & "Bbl07,ab16:ad16,i,D,91,123.41"       'B8.G.O.T.
    Wtext = Wtext & vbLf & "Bbl08,ab17:ad17,i,D,91,123.41"       'B8.G.P.T.
    Wtext = Wtext & vbLf & "Bbl09,ab18:ad18,i,D,91,123.41"       'B8.”A_
    Wtext = Wtext & vbLf & "Bbl10,ab19:ad19,i,D,91,123.41"       'B8.ŒŒ“œ
    Wtext = Wtext & vbLf & "Bbl11,ab20:ad20,i,D,91,123.41"       'B8.ŒŒˆ³Å‚
    Wtext = Wtext & vbLf & "Bbl12,ab21:ad21,i,D,91,123.41"       'B8.ŒŒˆ³Å’á
    Fld_Adrs1 = Split(Wtext, vbLf)

    Wtext = "100,ŠÖ“Œ…Ÿ,100" & vbLf & "101,ŠÖ“Œ… ,101" & vbLf & "102,–k@—¤,102"
    Wtext = Wtext & vbLf & "103,“Œ@ŠC,103" & vbLf & "104,‹ß‹E…Ÿ,104" & vbLf & "105,‹ß‹E… ,105"
    Wtext = Wtext & vbLf & "106,’†@‘,106" & vbLf & "107,l@‘,107" & vbLf & "108,–k‹ãB,108"
    Wtext = Wtext & vbLf & "109,“ì‹ãB,109" & vbLf & "110,–kŠC“¹,110" & vbLf & "111,“Œ@–k,111"
    Wtext = Wtext & vbLf & "201,–kŠC“¹,110" & vbLf & "202,ÂX@,111" & vbLf & "203,Šâè@,111"
    Wtext = Wtext & vbLf & "204,‹{é@,111" & vbLf & "205,H“c@,111" & vbLf & "206,RŒ`@,111"
    Wtext = Wtext & vbLf & "207,•Ÿ“‡@,111" & vbLf & "208,ˆïé@,101" & vbLf & "209,“È–Ø@,101"
    Wtext = Wtext & vbLf & "210,ŒQ”n@,101" & vbLf & "211,é‹Ê@,100" & vbLf & "212,ç—t@,100"
    Wtext = Wtext & vbLf & "213,“Œ‹@,100" & vbLf & "214,_“Şì,100" & vbLf & "215,VŠƒ@,102"
    Wtext = Wtext & vbLf & "216,•xR@,102" & vbLf & "217,Îì@,102" & vbLf & "218,•Ÿˆä@,102"
    Wtext = Wtext & vbLf & "219,R—œ@,101" & vbLf & "220,’·–ì@,101" & vbLf & "221,Šò•Œ@,103"
    Wtext = Wtext & vbLf & "222,Ã‰ª@,103" & vbLf & "223,ˆ¤’m@,103" & vbLf & "224,Od@,103"
    Wtext = Wtext & vbLf & "225, ‰ê@,105" & vbLf & "226,‹“s@,104" & vbLf & "227,‘åã@,104"
    Wtext = Wtext & vbLf & "228,•ºŒÉ@,104" & vbLf & "229,“Ş—Ç@,105" & vbLf & "230,˜a‰ÌR,105"
    Wtext = Wtext & vbLf & "231,’¹æ@,106" & vbLf & "232,“‡ª@,106" & vbLf & "233,‰ªR@,106"
    Wtext = Wtext & vbLf & "234,L“‡@,106" & vbLf & "235,RŒû@,106" & vbLf & "236,“¿“‡@,107"
    Wtext = Wtext & vbLf & "237,ì@,107" & vbLf & "238,ˆ¤•Q@,107" & vbLf & "239,‚’m@,107"
    Wtext = Wtext & vbLf & "240,•Ÿ‰ª@,108" & vbLf & "241,²‰ê@,108" & vbLf & "242,’·è@,108"
    Wtext = Wtext & vbLf & "243,ŒF–{@,109" & vbLf & "244,‘å•ª@,108" & vbLf & "245,‹{è@,109"
    Wtext = Wtext & vbLf & "246,­™“‡,109" & vbLf & "247,‰«“ê@,109"
    Fld_Area = Split(Wtext, vbLf)
End Function
'--------------------------------------------------------------------------------
'   01_010 ÛH‰æ–Ê‚Ìƒ[ƒNƒV[ƒg‚ªƒAƒNƒeƒBƒu‚É‚È‚Á‚½
'--------------------------------------------------------------------------------
Function Eiyo01_010ÛH_Activate()
    ActiveSheet.Unprotect                           'ƒV[ƒg‚Ì•ÛŒì‚ğ‰ğœ
'    ActiveSheet.Protect UserInterfaceOnly:=True     '•ÛŒì‚ğ—LŒø‚É‚·‚é
End Function
'--------------------------------------------------------------------------------
'   01_020 Šî‘b‰æ–Ê‚Ìƒ_ƒuƒ‹ƒNƒŠƒbƒN
'   AA—ñiŒŸõŠY“–•¡”j‚Ìƒ_ƒuƒ‹ƒNƒŠƒbƒN‚ÍŠY“–”Ô†‚ğƒZƒ‹[G3]‚Éİ’è
'--------------------------------------------------------------------------------
Function Eiyo01_020Šî‘b_BeforedoubleClick()
Dim Wadrs   As String
Dim Wcoul   As String
Dim Wtext   As String
Dim i1      As Long     'Fld_Adrs Index
Dim i3      As Long     'ƒ_ƒuƒ‹ƒNƒŠƒbƒN‚Ìs”Ô†

End Function
'--------------------------------------------------------------------------------
'   01_030 ƒNƒŠƒA_Click
'   “ü—Í€–Ú‚ÌÁ‹A’ •[EŒŸØƒV[ƒg‚Ìíœ
'--------------------------------------------------------------------------------
Function Eiyo01_030ƒNƒŠƒAClick()
Dim i1      As Long
Dim FldItem As Variant
Dim Lmax    As Long

    Call Eiyo01_000init
    Call Eiyo930Screen_Hold     '‰æ–Ê—}~‚Ù‚©
    
    For i1 = 0 To UBound(Fld_Adrs1)
        FldItem = Split(Fld_Adrs1(i1), ",")
        If FldItem(0) = "Gyyyy" Or _
           FldItem(0) = "Age  " Or _
           IsEmpty(Range(Trim(FldItem(0)))) Then
        Else
           Range(Trim(FldItem(0))) = Empty
        End If
    Next i1
    Call Eiyo01_820‘€ìƒKƒCƒh
    Lmax = Sheets("ÛH").UsedRange.Rows.Count
    If Lmax > 4 Then: Sheets("ÛH").Rows("5:" & Lmax).Delete Shift:=xlUp
    Call Eiyo99_w’èƒV[ƒgíœ("ŒŸØ")
    Call Eiyo99_w’èƒV[ƒgíœ("ŒŸØ2")
    Call Eiyo99_w’èƒV[ƒgíœ("DBmirror")
    Call Eiyo99_w’èƒV[ƒgíœ("¶³İ¾Øİ¸Ş¼°Ä")
    Range("Fcode").Select
    Call Eiyo940Screen_Start    '‰æ–Ê•`‰æ‚Ù‚©
End Function
'--------------------------------------------------------------------------------
'   01_100 ŒŸõ_Click
'       Šî‘bî•ñ‚ÌŒŸõA“Á’è‚³‚ê‚½ê‡‚ÉÛHî•ñ‚àæ“¾‚·‚é
'--------------------------------------------------------------------------------
Function Eiyo01_100ŒŸõClick()
Dim FldItem     As Variant
Dim i1          As Long

    Call Eiyo930Screen_Hold     '‰æ–Ê—}~‚Ù‚©
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
        Range("Gmesg") = "ŒŸõƒL[‚ª‚ ‚è‚Ü‚¹‚ñ"
    Else
        Call Eiyo01_110ŒŸõ(i1)
    End If
    
    If IsEmpty(Range("Fcode")) = False And _
       Range("Fcode") = Range("Fsave") Then         '“Á’è‚³‚ê‚½ê‡‚ÍÛHî•ñ
        Application.ScreenUpdating = False          '‰æ–Ê•`‰æ—}~
        Call Eiyo01_130MealGet
        Sheets("Šî‘b").Select
    End If
    Range("Fcode").Select
    Call Eiyo940Screen_Start                        '‰æ–Ê•`‰æ‚Ù‚©
End Function
'--------------------------------------------------------------------------------
'   01_110 ‚c‚aŒŸõˆ—     F-024
'--------------------------------------------------------------------------------
Function Eiyo01_110ŒŸõ(i1 As Long)
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
        
    'SQL‚Å“Ç‚İ‚Şƒf[ƒ^‚ğw’è‚·‚é
    in_key = Range(Trim(FldItem(0))).Text
    If Left(in_key, 1) = "“" Then: in_key = "%" & Right(in_key, Len(in_key) - 1)
    Call Eiyo91DB_Open      'DB Open
    If FldItem(0) = "Fcode" Then
        mySqlStr = "SELECT * FROM " & Tbl_Kiso & " Where Fcode = """ & in_key & """"
    Else
        mySqlStr = "SELECT * FROM " & Tbl_Kiso & " Where " & _
                   Trim(FldItem(0)) & " like """ & in_key & "%"""
    End If
    Set Rst_Kiso = myCon.Execute(mySqlStr)
    If Rst_Kiso.EOF Then
        Range("Gmesg") = "ŠY“–ƒf[ƒ^‚Í‚ ‚è‚Ü‚¹‚ñ"
        Range("Fcode").Select
    Else
        With Rst_Kiso
            Range("Ah2").CopyFromRecordset Rst_Kiso           'ƒŒƒR[ƒh
            If Range("Ah3") = Empty Then                        'ŠY“–‚ª‚PŒ‚Ì‚Æ‚«
                For i1 = 1 To UBound(Fld_Adrs1)                 '‰æ–Ê€–Ú‚Ì‡Ÿˆ—
                    FldItem = Split(Fld_Adrs1(i1), ",")
                    If FldItem(3) = "D" Then
                        For i2 = 0 To .Fields.Count - 1             'ƒtƒB[ƒ‹ƒh–¼
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
                Range("Gare1") = Eiyo01_120’nˆæ("1" & Range("area1"))
                Range("Gare2") = Eiyo01_120’nˆæ("2" & Range("area2"))
                Range("Fsave") = Range("Fcode")
                Call Eiyo01_820‘€ìƒKƒCƒh
            Else
                For i1 = 1 To .Fields.Count                     'ƒtƒB[ƒ‹ƒh–¼
                    Cells(1, i1 + 33).Value = .Fields(i1 - 1).Name
                Next
                Columns("ah:hz").EntireColumn.AutoFit           '•
                i1 = Range("ah1").End(xlDown).Row
                Range("Ah2:ah" & i1).Locked = False             '“ü—Í‰Â
                Range("Ah2:ah" & i1).Interior.ColorIndex = 34
            End If
            .Close
        End With
    End If
    Set Rst_Kiso = Nothing                        'ƒIƒuƒWƒFƒNƒg‚Ì‰ğ•ú
    Call Eiyo920DB_Close    'DB Close
End Function
'--------------------------------------------------------------------------------
'   01_120 ’nˆæE“s“¹•{Œ§•\¦
'--------------------------------------------------------------------------------
Function Eiyo01_120’nˆæ(in_code As String) As String
Dim i1      As Long
Dim Witem   As Variant

    Eiyo01_120’nˆæ = Empty
    For i1 = 0 To UBound(Fld_Area)
        Witem = Split(Fld_Area(i1), ",")
        If Witem(0) = in_code Then
            Eiyo01_120’nˆæ = Witem(1)
            Exit For
        End If
    Next i1
End Function
'--------------------------------------------------------------------------------
'   01_130@ÛHæ“¾
'--------------------------------------------------------------------------------
Function Eiyo01_130MealGet()
Dim mySqlStr    As String
Dim Lmax        As Long
Dim i1          As Long

    Sheets("ÛH").Select
    Application.EnableEvents = False                'ƒCƒxƒ“ƒg”­¶—}~
'    ActiveSheet.Unprotect                           'ƒV[ƒg‚Ì•ÛŒì‚ğ‰ğœ
    Range("b1") = Empty
    Range("a2") = Range("Fcode") & ":" & Range("Namej")
    Lmax = ActiveSheet.UsedRange.Rows.Count
    If Lmax > 4 Then: Rows("5:" & Lmax).Delete Shift:=xlUp
        
    'SQL‚Å“Ç‚İ‚Şƒf[ƒ^‚ğw’è‚·‚é
    Call Eiyo91DB_Open      'DB Open
    mySqlStr = "SELECT Sdate,Ekubn,Foodc,Suryo FROM " & Tbl_Meal & " Where Fcode = """ & Range("Fcode") & """"
    Set Rst_Meal = myCon.Execute(mySqlStr)
    If Rst_Meal.EOF Then
        Lmax = 0
    Else
        Range("A5").CopyFromRecordset Rst_Meal           'ƒŒƒR[ƒh
    End If
    
    Lmax = ActiveSheet.UsedRange.Rows.Count
    For i1 = 5 To Lmax
        Cells(i1, 6) = Cells(i1, 4)
        Cells(i1, 4) = Cells(i1, 3)
        Cells(i1, 3) = Eiyo01_401H–‹æ•ª(Cells(i1, 2))
        Call Eiyo01_402H•iƒ}ƒXƒ^(i1)
    Next i1
    Set Rst_Meal = Nothing                    'ƒIƒuƒWƒFƒNƒg‚Ì‰ğ•ú
    Call Eiyo920DB_Close    'DB Close
End Function
'--------------------------------------------------------------------------------
'   01_200 XV_Click
'--------------------------------------------------------------------------------
Function Eiyo01_200XVClick()
Dim Rtn As Long
    Call Eiyo930Screen_Hold                     '‰æ–Ê—}~‚Ù‚©
    Call Eiyo01_000init
    Rtn = Eiyo01_210KeyCheck                    'ƒL[ƒ`ƒFƒbƒN
    If Rtn = 0 Then: Rtn = Eiyo01_220€–ÚCheck  '€–Úƒ`ƒFƒbƒN
    If Rtn = 0 Then: Rtn = Eiyo01_230DBXV     'DBXV
    Call Eiyo940Screen_Start                    '‰æ–Ê•`‰æ‚Ù‚©
End Function
'--------------------------------------------------------------------------------
'   01_210 ƒL[ƒ`ƒFƒbƒN
'--------------------------------------------------------------------------------
Function Eiyo01_210KeyCheck() As Long
Dim mySqlStr    As String

    Call Eiyo91DB_Open      'DB Open
    mySqlStr = "SELECT * FROM " & Tbl_Kiso & " Where Fcode = """ & Range("Fcode") & """"
    Set Rst_Kiso = myCon.Execute(mySqlStr)
    If Rst_Kiso.EOF Then
        If Range("Fcode") = Range("Fsave") Then
            Range("Gmesg") = "Program Error Non Key & Save Key Same"    '~FƒL[‚È‚µASave“¯‚¶
            Eiyo01_210KeyCheck = 1
        Else
            Eiyo01_210KeyCheck = 0                                      '›FƒL[‚È‚µASaveˆÙ‚È‚é(V‹K)
        End If
    Else
        If Range("Fcode") = Range("Fsave") Then
            Eiyo01_210KeyCheck = 0                                      '›FƒL[‚ ‚èASave“¯‚¶(XV)
        Else
            Range("Gmesg") = "ƒR[ƒh‚ªd•¡‚µ‚Ä‚¢‚Ü‚·"                   '~FƒL[‚ ‚èASaveˆÙ‚È‚é
            Eiyo01_210KeyCheck = 1
        End If
    End If
    Set Rst_Kiso = Nothing                        'ƒIƒuƒWƒFƒNƒg‚Ì‰ğ•ú
    Call Eiyo920DB_Close    'DB Close
End Function
'--------------------------------------------------------------------------------
'   01_220 €–Úƒ`ƒFƒbƒN
'--------------------------------------------------------------------------------
Function Eiyo01_220€–ÚCheck() As Long
Dim Witem   As Variant
Dim Wlen    As Long
Dim i1      As Long
Dim Wtemp   As String

    Eiyo01_220€–ÚCheck = 1
    Range("Gmesg") = Empty
'   ƒR[ƒh
'    Witem = Range("Fcode")
'    If IsNumeric(Witem) = True And Len(Witem) <= 10 Then
'    Else
'        Range("Gmesg") = "ƒR[ƒh‚Í‚P‚OŒ…ˆÈ“à‚Ì”š‚É‚µ‚Ä‚­‚¾‚³‚¢@" & Len(Witem)
'        Range("Fcode").Activate
'        Exit Function
'    End If
'   ’²¸ŠúŠÔŠJn“ú
    If IsDate(Range("Date1")) Then
    Else
        Range("Gmesg") = "’²¸ŠúŠÔŠJn“ú‚ğÀİ“ú‚É‚µ‚Ä‚­‚¾‚³‚¢"
        Range("Date1").Activate
        Exit Function
    End If
'   ’²¸ŠúŠÔ“ú”
    Witem = Range("Nissu")
    If IsNumeric(Witem) = True And Len(Witem) = 1 Then
    Else
        Range("Gmesg") = "’²¸ŠúŠÔ“ú”‚Í‚PŒ…‚Ì”š‚É‚µ‚Ä‚­‚¾‚³‚¢"
        Range("Nissu").Activate
        Exit Function
    End If
'   –¼
    If Eiyo01_221Œ…”check("Namej", "–¼", 10) = 1 Then: Exit Function
'   «•Ê
    Witem = Range("sex")
    If Witem = "" Or Witem = "0" Or Witem = "1" Then
    Else
        Range("Gmesg") = "«•Ê‚Í‚PŒ…‚Ì”š‚É‚µ‚Ä‚­‚¾‚³‚¢"
        Range("sex").Activate
        Exit Function
    End If
'   ¶”NŒ“ú
    If IsDate(Range("Birth")) Then
    Else
        Range("Gmesg") = "¶”NŒ“ú‚ğÀİ“ú‚É‚µ‚Ä‚­‚¾‚³‚¢"
        Range("Birth").Activate
        Exit Function
    End If
    
    If Eiyo01_223”’lcheck("Hight", "g’·", 3, 1, 300) = 1 Then: Exit Function
    If Eiyo01_223”’lcheck("Weght", "‘Ìd", 3, 1, 300) = 1 Then: Exit Function
    If Eiyo01_223”’lcheck("Sibou", "”ç‰º‰–b", 2, 1, 50) = 1 Then: Exit Function
    
    If Eiyo01_221Œ…”check("Adrno", "—X•Ö”Ô†", 18) = 1 Then: Exit Function
    If Eiyo01_221Œ…”check("Adrs1", "ZŠ[‚P", 18) = 1 Then: Exit Function
    If Eiyo01_221Œ…”check("Adrs2", "ZŠ[‚Q", 18) = 1 Then: Exit Function
'   ’n‹æE’nˆæ
    Wtemp = Left(Range("adrs1"), 2)
    For i1 = 0 To UBound(Fld_Area)
        Witem = Split(Fld_Area(i1), ",")
        If Left(Witem(0), 1) = "2" And _
           Left(Witem(1), 2) = Wtemp Then
            Range("Area1") = Right(Witem(2), 2)
            Range("Gare1") = Eiyo01_120’nˆæ("1" & Range("Area1"))
            Range("Area2") = Right(Witem(0), 2)
            Range("Gare2") = Witem(1)
            Exit For
        End If
    Next i1
    If Eiyo01_222”šcheck("Q3rec", "Q3.HKŠµ", 10) = 1 Then: Exit Function
    If Eiyo01_222”šcheck("Q4rec", "Q4.‹x—{", 5) = 1 Then: Exit Function
    If Eiyo01_222”šcheck("Q5rec", "Q5.‰^“®", 3) = 1 Then: Exit Function
    If Eiyo01_222”šcheck("Q6r_a", "Q6.Œ’N‚P", 10) = 1 Then: Exit Function
    If Eiyo01_222”šcheck("Q6r_b", "Q6.Œ’N‚Q", 10) = 1 Then: Exit Function
    If Eiyo01_222”šcheck("Q6r_c", "Q6.Œ’N‚R", 10) = 1 Then: Exit Function
    If Eiyo01_222”šcheck("Q6r_d", "Q6.Œ’N‚S", 10) = 1 Then: Exit Function
    If Eiyo01_222”šcheck("Q6r_e", "Q6.Œ’N‚T", 10) = 1 Then: Exit Function
'   E‹Æ
    Range("Qjob1") = UCase(Range("Qjob1"))
    If Len(Range("Qjob1")) = 4 Then
    Else
        Range("Gmesg") = "E‹Æ‚Í‚SŒ…‚Æ‚µ‚Ä‚­‚¾‚³‚¢@" & Len(Range("Qjob1"))
        Range("Qjob1").Activate
    End If

    If Eiyo01_222”šcheck("Qsyuf", "QA.å•w", 1) = 1 Then: Exit Function
    If Eiyo01_222”šcheck("Qcnd1", "QB.”DP", 1) = 1 Then: Exit Function
    If Eiyo01_222”šcheck("Qtony", "QC.“œ”A", 1) = 1 Then: Exit Function
    If Range("Qtony") = "0" Then
        Range("Qill1") = "000000"
    Else
        Range("Qill1") = "000321"
    End If
    If Eiyo01_222”šcheck("Qkoke", "QC.“œ”A", 1) = 1 Then: Exit Function
    If Range("Qkoke") = "0" Then
        Range("Qill2") = "000000"
    Else
        Range("Qill2") = "000313"
    End If
    If Eiyo01_223”’lcheck("Qsrmr", "QE.½Îß°Â1", 3, 0, 1000) = 1 Then: Exit Function
    If Eiyo01_223”’lcheck("Qsmin", "QE.½Îß°Â2", 3, 0, 1000) = 1 Then: Exit Function
    If Eiyo01_222”šcheck("Qclab", "QF.‰^“®•”", 1) = 1 Then: Exit Function
    If Eiyo01_222”šcheck("Qclab", "Q .‹i‰Œ", 1) = 1 Then: Exit Function
    If Eiyo01_223”’lcheck("Qsyog", "QG.gáŠQ", 2, 0, 100) = 1 Then: Exit Function
    If Eiyo01_222”šcheck("Qwcnt", "QG.³´²ÄCT", 1) = 1 Then: Exit Function
    If Eiyo01_222”šcheck("Tenes", "´ÈÙ·Şw’è", 1) = 1 Then: Exit Function
    If Eiyo01_222”šcheck("Tanps", "ÀİÊß¸w’è", 1) = 1 Then: Exit Function
    If Eiyo01_223”’lcheck("Tenee", "´ÈÙ·Şw’è", 5, 2, 100000) = 1 Then: Exit Function
    If Eiyo01_223”’lcheck("Tanpe", "ÀİÊß¸w’è", 5, 2, 100000) = 1 Then: Exit Function
'
    Range("Blood") = UCase(Range("Blood"))
    Wtemp = Range("Blood")
    If Wtemp = "" Or Wtemp = "A" Or Wtemp = "B" Or Wtemp = "O" Or Wtemp = "AB" Then
    Else
        Range("Gmesg") = "ŒŒ‰tŒ^‚ª•s³‚Å‚·"
        Range("Blood").Activate
        Exit Function
    End If

    If Eiyo01_221Œ…”check("Bscd1", "x“X", 3) = 1 Then: Exit Function
    If Eiyo01_221Œ…”check("Bscd2", "x•”", 2) = 1 Then: Exit Function
    If Eiyo01_221Œ…”check("Bhok1", "•ÛŒ¯Ø‹L†", 8) = 1 Then: Exit Function
    If Eiyo01_221Œ…”check("Bhok2", "•ÛŒ¯ØNo", 8) = 1 Then: Exit Function
    If Eiyo01_221Œ…”check("Bhant", "’èŠúŒ’f", 2) = 1 Then: Exit Function
'
    Range("Barm") = UCase(Range("Barm"))
    Wtemp = Range("Barm")
    If Wtemp = "" Or Wtemp = "L" Or Wtemp = "R" Then
    Else
        Range("Gmesg") = "ŒŸ¸˜r‚ª•s³‚Å‚·"
        Range("Barm").Activate
        Exit Function
    End If
'   ŒŒ‰tŒŸ¸“ú
    If IsEmpty(Range("Bdate")) Or IsDate(Range("Bdate")) Then
    Else
        Range("Gmesg") = "ŒŒ‰tŒŸ¸“ú‚ğÀİ“ú‚É‚µ‚Ä‚­‚¾‚³‚¢"
        Range("Bdate").Activate
        Exit Function
    End If
    If Eiyo01_223”’lcheck("Bbl01", "ÔŒŒ‹…”", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223”’lcheck("Bbl02", "ŒŒF‘f—Ê", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223”’lcheck("Bbl03", "ÍÏÄ¸Ø¯Ä", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223”’lcheck("Bbl04", "ºÚ½ÃÛ°Ù", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223”’lcheck("Bbl05", "HDL", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223”’lcheck("Bbl06", "’†«‰–b", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223”’lcheck("Bbl07", "G.O.T.", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223”’lcheck("Bbl08", "G.P.T.", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223”’lcheck("Bbl09", "”A_", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223”’lcheck("Bbl10", "ŒŒ“œ", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223”’lcheck("Bbl11", "ŒŒˆ³Å‚", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223”’lcheck("Bbl12", "ŒŒˆ³Å’á", 3, 1, 10000) = 1 Then: Exit Function
    Eiyo01_220€–ÚCheck = 0
End Function
'--------------------------------------------------------------------------------
'   01_221 Œ…”ƒ`ƒFƒbƒN
'--------------------------------------------------------------------------------
Function Eiyo01_221Œ…”check(Ifld As String, Iname As String, Ilen As Long) As Long
    If Len(Range(Ifld)) > Ilen Then
        Range("Gmesg") = Iname & "‚Í" & Ilen & "Œ…ˆÈ“à‚É‚µ‚Ä‚­‚¾‚³‚¢@" & Len(Range(Ifld))
        Range(Ifld).Activate
        Eiyo01_221Œ…”check = 1
    Else
        Eiyo01_221Œ…”check = 0
    End If
End Function
'--------------------------------------------------------------------------------
'   01_222 ŒÅ’èŒ…”š€–Úƒ`ƒFƒbƒN
'--------------------------------------------------------------------------------
Function Eiyo01_222”šcheck(Ifld As String, Iname As String, Ilen As Long) As Long
Dim Witem   As Variant
Dim Wlen    As Long

    If Range(Ifld) = Empty Then: Range(Ifld) = String(Ilen, "0")
    Witem = Range(Ifld)
    Wlen = Len(Witem)
    If IsNumeric(Witem) And Wlen = Ilen Then
        Eiyo01_222”šcheck = 0
    Else
        Range("Gmesg") = Iname & "‚Í" & Ilen & "Œ…‚Ì”š‚É‚µ‚Ä‚­‚¾‚³‚¢@" & Wlen
        Range(Ifld).Activate
        Eiyo01_222”šcheck = 1
    End If
End Function
'--------------------------------------------------------------------------------
'   01_223 ”’l€–Úƒ`ƒFƒbƒN
'--------------------------------------------------------------------------------
Function Eiyo01_223”’lcheck(Ifld As String, Iname As String, _
                              Ilen1 As Long, Ilen2 As Long, Imax As Long) As Long
Dim Witem   As Variant
    
    Witem = Range(Ifld)
    If IsNumeric(Witem) And Witem < Imax Then
        Eiyo01_223”’lcheck = 0
    Else
        Range("Gmesg") = Iname & "‚Íã" & Ilen1 & "Œ…‰º" & Ilen2 & "Œ…ˆÈ“à‚Ì”’l‚É‚µ‚Ä‚­‚¾‚³‚¢"
        Range(Ifld).Activate
        Eiyo01_223”’lcheck = 1
    End If
End Function
'--------------------------------------------------------------------------------
'   01_230 ‚c‚aXV                                     F-026
'   Microsoft ActiveX Data Objects 2.X Library QÆİ’è
'--------------------------------------------------------------------------------
Function Eiyo01_230DBXV() As Long
Dim FldItem     As Variant
Dim FldName     As String
Dim i1          As Long

    Call Eiyo91DB_Open      'DB Open
    '€”õ‚±‚±‚Ü‚Å
    With Rst_Kiso
        'ƒCƒ“ƒfƒbƒNƒX‚Ìİ’è
        .Index = "PrimaryKey"
        'ƒŒƒR[ƒhƒZƒbƒg‚ğŠJ‚­
        Rst_Kiso.Open Source:=Tbl_Kiso, ActiveConnection:=myCon, _
            CursorType:=adOpenKeyset, LockType:=adLockOptimistic, _
            Options:=adCmdTableDirect
        '”Ô†‚ª“o˜^‚³‚ê‚Ä‚¢‚é‚©ŒŸõ‚·‚é
        If Not .EOF Then .Seek Range("Fcode")
        If .EOF Then
            .AddNew
            Range("Gmesg") = "’Ç‰Á“o˜^‚³‚ê‚Ü‚µ‚½B"
            Range("Fsave") = Range("Fcode")
        Else
            Range("Gmesg") = "XV‚³‚ê‚Ü‚µ‚½B"
        End If
        For i1 = 1 To UBound(Fld_Adrs1)                 '‰æ–Ê€–Ú‚Ì‡Ÿˆ—
            FldItem = Split(Fld_Adrs1(i1), ",")         '
            If FldItem(3) = "D" Then
                FldName = Trim(FldItem(0))
                .Fields(FldName).Value = Range(FldName).Value
            End If
        Next i1
        .Update
        .Close
    End With
    Set Rst_Kiso = Nothing      'ƒIƒuƒWƒFƒNƒg‚Ì‰ğ•ú
    Call Eiyo920DB_Close        'DB Close
    Eiyo01_230DBXV = 0
End Function
'--------------------------------------------------------------------------------
'   01_300@æÁ_Click
'--------------------------------------------------------------------------------
Function Eiyo01_300æÁClick()
    If Range("Fcode") = Range("Fsave") And _
        IsEmpty(Range("Fcode")) = False Then
        Call Eiyo91DB_Open      'DB Open
        myCon.Execute "DELETE FROM " & Tbl_Kiso & " Where Fcode = """ & Range("Fcode") & """"
        myCon.Execute "DELETE FROM " & Tbl_Meal & " Where Fcode = """ & Range("Fcode") & """"
        Range("Gmesg") = "æÁíœ‚³‚ê‚Ü‚µ‚½B"
        Range("Fsave") = Empty
        Call Eiyo920DB_Close    'DB Close
    Else
        Range("Gmesg") = "ŒŸõ‚³‚ê‚Ä‚¢‚Ü‚¹‚ñB"
    End If
End Function
'--------------------------------------------------------------------------------
'   01_400@ÛH•\¦
'--------------------------------------------------------------------------------
Function Eiyo01_400MealDisp()
Dim Rtn     As Long
Dim Wmsg    As String
    
    Range("a2") = Range("Fcode") & ":" & Range("Namej")
    Wmsg = "Šî‘bî•ñ‚ÌŒŸõ‚ªs‚í‚ê‚Ä‚¢‚Ü‚¹‚ñ"
    If IsEmpty(Range("Fcode")) Or Range("Fcode") <> Range("Fsave") Then
        Rtn = CreateObject("WScript.Shell").Popup(Wmsg, 3, "Microsoft Excel", 0)
        Sheets("Šî‘b").Select
    End If
End Function
'--------------------------------------------------------------------------------
'   01_401@H–‹æ•ª
'--------------------------------------------------------------------------------
Function Eiyo01_401H–‹æ•ª(kbn As Long) As String
    Select Case kbn
        Case 1: Eiyo01_401H–‹æ•ª = "’©"
        Case 2: Eiyo01_401H–‹æ•ª = "’‹"
        Case 3: Eiyo01_401H–‹æ•ª = "—["
        Case 4: Eiyo01_401H–‹æ•ª = "–é"
        Case 5: Eiyo01_401H–‹æ•ª = "ŠÔ"
        Case Else: Eiyo01_401H–‹æ•ª = Empty
    End Select
End Function
'--------------------------------------------------------------------------------
'   01_402@H•iƒ}ƒXƒ^æ“¾
'--------------------------------------------------------------------------------
Function Eiyo01_402H•iƒ}ƒXƒ^(in_line As Long)
Dim mySqlStr    As String
    If IsEmpty(Cells(in_line, 4)) Then
        Cells(in_line, 5) = Empty
        Range("g" & in_line & ":z" & in_line) = Empty
        Exit Function
    End If
        
    mySqlStr = "SELECT * FROM " & Tbl_Food & " Where Foodc = " & Cells(in_line, 4)
    Set Rst_Food = myCon.Execute(mySqlStr)
    If Rst_Food.EOF Then
        Cells(in_line, 5) = "ƒL[‚È‚µ"
        Range("g" & in_line & ":z" & in_line) = Empty
    Else
        Cells(in_line, 7).CopyFromRecordset Rst_Food
        Cells(in_line, 5) = Cells(in_line, 8)
    End If
    Rst_Food.Close
    Set Rst_Food = Nothing
End Function
'--------------------------------------------------------------------------------
'   01_410@ÛH‰æ–Ê‚ª•ÏX‚³‚ê‚½
'--------------------------------------------------------------------------------
Function Eiyo01_410MealChange(ChangeCell As String)
Dim Wl      As Long
Dim Wc      As Long

    Wl = Range(ChangeCell).Row
    Wc = Range(ChangeCell).Column
'    ActiveSheet.Unprotect                           'ƒV[ƒg‚Ì•ÛŒì‚ğ‰ğœ
    Select Case Wc
        Case 1: Cells(Wl, 2).Select
        Case 2
            Cells(Wl, 3) = Eiyo01_401H–‹æ•ª(Cells(Wl, 2))
            Cells(Wl, 4).Select
        Case 4
            'SQL‚Å“Ç‚İ‚Şƒf[ƒ^‚ğw’è‚·‚é
            Call Eiyo91DB_Open      'DB Open
            Call Eiyo01_402H•iƒ}ƒXƒ^(Wl)
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
'    ActiveSheet.Protect UserInterfaceOnly:=True     '•ÛŒì‚ğ—LŒø‚É‚·‚é
End Function
'--------------------------------------------------------------------------------
'   01_420@ƒƒjƒ…[‚ª‘I‘ğ‚³‚ê‚½
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
    If IsEmpty(Sheets("ÛH").Range("b1")) Then
        Wmsg = Wcode & ":" & Wname & "‚ª‘I‘ğ‚³‚ê‚Ü‚µ‚½B"
        Rtn = CreateObject("WScript.Shell").Popup(Wmsg, 3, "Microsoft Excel", 0)
    Else
        Application.EnableEvents = False            'ƒCƒxƒ“ƒg”­¶—}~
        Sheets("ÛH").Select
        Wcell = Range("b1")
        Range("b1") = Empty
        Application.EnableEvents = True             'ƒCƒxƒ“ƒg”­¶ÄŠJ
        Range(Wcell) = Wcode
    End If
End Function
'--------------------------------------------------------------------------------
'   01_500@“o˜^ Click
'--------------------------------------------------------------------------------
Function Eiyo01_500MealCalc(in_Func As Long)
    Call Eiyo930Screen_Hold                                 '‰æ–Ê—}~‚Ù‚©
    Call Eiyo91DB_Open                                      'DB Open
    If Eiyo01_501MealEntry = 1 Then: GoTo Eiyo01_503Exit    '–¢“ü—Íƒ`ƒFƒbƒN
    If Eiyo01_502Mealscope = 1 Then: GoTo Eiyo01_503Exit    'ÛH—Ê‚Ì”ÍˆÍƒ`ƒFƒbƒN
    If Eiyo01_503Mealzerod = 1 Then: GoTo Eiyo01_503Exit    'ÛH—Êƒ[ƒ‚Ìíœ
    If Eiyo01_504MealDoubl = 1 Then: GoTo Eiyo01_503Exit    'ÛH‚Ìd•¡“ü—Í
    If Eiyo01_510MealUdate = 1 Then: GoTo Eiyo01_503Exit    '‚c‚aXV
    If Eiyo01_511MealFldgt = 1 Then: GoTo Eiyo01_501Exit    '€–Ú—v‘fæ“¾
    If Eiyo01_512MealSheet = 1 Then: GoTo Eiyo01_501Exit    'ÛHŒvZƒV[ƒg
    If Eiyo01_513kenso2sht = 1 Then: GoTo Eiyo01_501Exit    'ŒŸØ‚QƒV[ƒgì¬
    If Eiyo01_514MealCalc1 = 1 Then: GoTo Eiyo01_501Exit    'ÛHŒvZ
    If Eiyo01_515MealTotal = 1 Then: GoTo Eiyo01_501Exit    'ÛH—Ê‡Œv
    If Eiyo01_521CalcDbGet(1) = 1 Then: GoTo Eiyo01_501Exit 'ÛH—Ê‡Œv
    If Eiyo01_522Mealcalc2 = 1 Then: GoTo Eiyo01_501Exit    '•W€‘Ìd‚Ù‚©
    If Eiyo01_525MealDiffe = 1 Then: GoTo Eiyo01_501Exit    '‰ß•s‘«ƒAƒhƒoƒCƒX
    If Eiyo01_528Eiyohirit = 1 Then: GoTo Eiyo01_501Exit    '‰h—{”ä—¦
    If in_Func = 2 Then
        If Eiyo01_540Old_Check = 1 Then: GoTo Eiyo01_501Exit    '‹ŒŒvZ’l
'    Else
'        Call Eiyo99_w’èƒV[ƒgíœ("DBmirror")
    End If
Eiyo01_501Exit:
    Call Eiyo01_550RstClose
Eiyo01_503Exit:
    Call Eiyo920DB_Close                'DB Close
    Sheets("Šî‘b").Select
    Call Eiyo940Screen_Start            '‰æ–Ê•`‰æ‚Ù‚©
End Function
'--------------------------------------------------------------------------------
'   01_501@ÛHî•ñ–¢“ü—Íƒ`ƒFƒbƒN
'--------------------------------------------------------------------------------
Function Eiyo01_501MealEntry() As Long
Dim Lmax    As Long
Dim i1      As Long
Dim Wnon    As Long

    Eiyo01_501MealEntry = 1
    Lmax = ActiveSheet.UsedRange.Rows.Count
    If Lmax < 5 Then
        MsgBox "ƒf[ƒ^‚ª‚ ‚è‚Ü‚¹‚ñ"
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
        MsgBox ("Œë‚è‚Ì€–Ú‚ğC³‚µ‚Ä‚­‚¾‚³‚¢B")
    Else
        Eiyo01_501MealEntry = 0
    End If
End Function
'--------------------------------------------------------------------------------
'   01_502@ÛH—Ê‚Ì”ÍˆÍƒ`ƒFƒbƒN
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
        Wmsg = "ÛH—Ê‚ÌˆÙí’l‚ª" & Wover & "ƒ•Š‚ ‚è‚Ü‚·"
        i1 = CreateObject("WScript.Shell").Popup(Wmsg, 1, "Microsoft Excel", 0)
        Eiyo01_502Mealscope = 0
    End If
    Eiyo01_502Mealscope = 0
End Function
'--------------------------------------------------------------------------------
'   01_503@ÛH—Êƒ[ƒ‚Ìíœ
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
        Wmsg = "ÛH—Êƒ[ƒ‚Ì" & Wzero / 2 & "s‚ğíœ‚µ‚Ü‚µ‚½B"
        i1 = CreateObject("WScript.Shell").Popup(Wmsg, 1, "Microsoft Excel", 0)
    End If
    Eiyo01_503Mealzerod = 0
End Function
'--------------------------------------------------------------------------------
'   01_504@ÛH‚Ìd•¡ƒ`ƒFƒbƒN
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
'   01_510@ÛHDB“o˜^
'--------------------------------------------------------------------------------
Function Eiyo01_510MealUdate() As Long
Dim Lmax    As Long
Dim i1      As Long
Dim Wkey    As Variant

    Lmax = Range("a4").End(xlDown).Row
    myCon.Execute "DELETE FROM " & Tbl_Meal & " Where Fcode = """ & Range("Fcode") & """"
    '€”õ‚±‚±‚Ü‚Å
    With Rst_Meal
        'ƒCƒ“ƒfƒbƒNƒX‚Ìİ’è
        .Index = "PrimaryKey"
        'ƒŒƒR[ƒhƒZƒbƒg‚ğŠJ‚­
        Rst_Meal.Open Source:=Tbl_Meal, ActiveConnection:=myCon, _
            CursorType:=adOpenKeyset, LockType:=adLockOptimistic, _
            Options:=adCmdTableDirect
        For i1 = 5 To Lmax
        '”Ô†‚ª“o˜^‚³‚ê‚Ä‚¢‚é‚©ŒŸõ‚·‚é
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
    Set Rst_Meal = Nothing                    'ƒIƒuƒWƒFƒNƒg‚Ì‰ğ•ú
    Eiyo01_510MealUdate = 0
End Function
'--------------------------------------------------------------------------------
'   01_511@‰h—{‘f€–Ú‚ÌŠeíî•ñæ“¾    F-018
'--------------------------------------------------------------------------------
Function Eiyo01_511MealFldgt() As Long
    
    Sheets.Add After:=Sheets(Sheets.Count)      'ƒV[ƒg’Ç‰Á
    'ƒŒƒR[ƒhƒZƒbƒg‚ğŠJ‚­
    Rst_Field.Open Source:=Tbl_Field, _
                ActiveConnection:=myCon, _
                CursorType:=adOpenForwardOnly, _
                LockType:=adLockReadOnly, _
                Options:=adCmdTableDirect
    'ƒŒƒR[ƒh
    Range("a1").CopyFromRecordset Rst_Field
    Fld_Field = ActiveSheet.UsedRange
    Rst_Field.Close
    Set Rst_Field = Nothing    'ƒIƒuƒWƒFƒNƒg‚Ì‰ğ•ú
    Application.DisplayAlerts = False               'Šm”F—}~
    ActiveSheet.Delete
    Application.DisplayAlerts = True                'Šm”F•œŠˆ
    Eiyo01_511MealFldgt = 0
End Function
'--------------------------------------------------------------------------------
'   01_512@ÛHŒvZƒV[ƒgì¬
'--------------------------------------------------------------------------------
Function Eiyo01_512MealSheet() As Long
Dim i1      As Long     'sIndex
Dim i2      As Long     '—“Index
Dim Wno     As String
Dim Wtext   As String

    Call Eiyo99_w’èƒV[ƒgíœ("ŒŸØ")
    Sheets.Add After:=Sheets(Sheets.Count)      'ƒV[ƒg’Ç‰Á
    ActiveSheet.Name = "ŒŸØ"
    Range("d1") = "‰h—{ŒvZ@ÛHŒŸØ"
    Range("a2") = Sheets("ÛH").Range("a2")
    Wtext = Empty
    For i1 = 1 To 27                            '‰h—{‘f
        Wno = Format(i1, "00")
        Wtext = Wtext & "Ûæ—Ê" & Wno & vbTab
        Wtext = Wtext & "”M‘¹Œã" & Wno & vbTab
    Next i1
    For i1 = 1 To 15
        Wtext = Wtext & "´ÈÙ·ŞC" & Format(i1, "00") & vbTab
    Next i1
    For i1 = 1 To 15
        Wtext = Wtext & "´ÈÙ·ŞW" & Format(i1, "00") & vbTab
    Next i1
    For i1 = 1 To 15
        Wtext = Wtext & "¶Ù¼³Ñ1" & Format(i1, "00") & vbTab
    Next i1
    For i1 = 1 To 15
        Wtext = Wtext & "¶Ù¼³Ñ2" & Format(i1, "00") & vbTab
    Next i1
    Wtext = Wtext & "‰¿“®•¨" & vbTab
    Wtext = Wtext & "‰¿‹›‰î" & vbTab
    Wtext = Wtext & "‰¿A•¨" & vbTab
    Wtext = Wtext & "”M‘¹“®•¨" & vbTab
    Wtext = Wtext & "”M‘¹‹›‰î" & vbTab
    Wtext = Wtext & "”M‘¹A•¨"
    Range("a4:dp4") = Split(Wtext, vbTab)
    ActiveWindow.FreezePanes = False        'ƒEƒCƒ“ƒh˜gŒÅ’è‚Ì‰ğœ
    Range("a5").Select
    ActiveWindow.FreezePanes = True         'ƒEƒCƒ“ƒh˜gŒÅ’è‚Ìİ’è
    Cells.NumberFormatLocal = "#,##0.00;[Ô]-#,##0.00"
    Eiyo01_512MealSheet = 0
End Function
'--------------------------------------------------------------------------------
'   01_513@ŒŸØ‚QƒV[ƒgì¬
'--------------------------------------------------------------------------------
Function Eiyo01_513kenso2sht() As Long
Dim i1      As Long
    Call Eiyo99_w’èƒV[ƒgíœ("ŒŸØ2")
    Sheets.Add After:=Sheets(Sheets.Count)      'ƒV[ƒg’Ç‰Á
    ActiveSheet.Name = "ŒŸØ2"
    Cells.Interior.ColorIndex = 36              '‘S‰æ–Ê”wŒiF
'   •\‘è
    Range("C1:F1").Select
    Selection.MergeCells = True                 '•\‘èƒZƒ‹˜AŒ‹
    Selection.HorizontalAlignment = xlCenter    '•\‘èƒZƒ“ƒ^ƒŠƒ“ƒO
    Selection.Interior.ColorIndex = 37          '•\‘èFiƒy[ƒ‹ƒuƒ‹[j
    With Selection.Font                         'ƒtƒHƒ“ƒg
        .FontStyle = "‘¾š"
        .Size = 16
    End With
    Range("C1") = "‰h—{ŒvZ@ŒŸØ‘—¿‚Q"
    
'   ‰h—{‘f–¼
    Range("a4") = "No.‰h—{‘f–¼"
    Range("b4") = "’PˆÊ"
    For i1 = 1 To 27
        Cells(4 + i1, 1) = Format(i1, "00") & "." & Fld_Field(i1, 4)
        Cells(4 + i1, 2) = Fld_Field(i1, 5)
    Next i1
'   Ûæ—Ê
    Range("c3") = "<========= ”M‘¹ŒãÛæ—Ê ==========>"
    Range("c4") = "‘—Ê"
    Range("d4") = "^“ú"
    Range("e4") = "®"
    Range("f4") = "•â³Œã"
    Range("c4:p4").HorizontalAlignment = xlCenter
    Range("c5:d31,f5:f31").Interior.ColorIndex = xlNone      '”’”²‚«‰»
    With Range("c5:d31,f5:f31").Borders                      '˜gŒrü
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .Weight = xlThin
    End With
    Range("c5:d31,f5:f31").NumberFormatLocal = "#,##0.00;[Ô]-#,##0.00"
    Range("c5").Name = "ks2_eiyoso"
'   Ûæ—Ê‚Ì•â³ğŒ
    Range("e:e,o:o").HorizontalAlignment = xlCenter '‰¡’†‰›
    Range("e15,e20,e24,e27").Interior.ColorIndex = xlNone   '”’”²‚«‰»
    With Range("e15,e20,e24,e27").Borders                   '˜gŒrü
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .Weight = xlThin
    End With
    Range("e15").Name = "ks2_hosei11"
    Range("e20").Name = "ks2_hosei16"
    Range("e24").Name = "ks2_hosei20"
    Range("e27").Name = "ks2_hosei23"
'   Šî‘bî•ñ
    Range("i4") = "À‘Ìd"
    Range("j4") = "•W€‘Ìd"
    Range("h5") = "a.‘Ìd"
    Range("H6") = "b.‘Ì•\–ÊÏ"
    Range("H8") = "Šî‘b‘ãÓ"
    Range("h9") = "c.¶Šˆw”"
    Range("H10") = "d.ƒR[ƒh"
    Range("H11") = "e.–ÊÏ“–‚è"
    Range("H12") = "f.^“ú"
    Range("H13") = "g.^•ª"
    Range("H15") = "ƒGƒlƒ‹ƒM["
    Range("H16") = "h.•W€—Ê"
    Range("H17") = "i.“K—pğŒ"
    Range("H18") = "j.´ÈÙ·Ş°1"
    Range("H19") = "k.´ÈÙ·Ş°2"
    Range("i21") = "f = b * e * 24"
    Range("i22") = "h = f * (1+c) * 1.1"
    Range("F05").Copy Range("I5:j6,i9:i11,i12:j13,i16:j16,i18:i19")
    Range("E15").Copy Range("I10,i17")
    Range("i05").Name = "ks2_weght"     '‘Ìd
    Range("i06").Name = "ks2_Aansa"     '‘Ì•\–ÊÏ
    Range("i09").Name = "ks2_Aansx"     '¶Šˆw”
    Range("i10").Name = "ks2_kisocd"    'Šî‘bƒR[ƒh
    Range("i11").Name = "ks2_kisot"     '–ÊÏ“–‚èŠî‘b‘ãÓ
    Range("i12").Name = "ks2_Aansb"     'Šî‘b‘ãÓ^“ú
    Range("i13").Name = "ks2_Aansc"     'Šî‘b‘ãÓ^“ú
    Range("i16").Name = "ks2_Aansd"     '´ÈÙ·Ş°•W€—Ê
    Range("i17").Name = "ks2_energ"     'Š—v—Ê´ÈÙ·Ş°ğŒ
'   Š—v—Ê(´ÈÙ·Ş°ˆÈŠO)
    Range("l3") = "Š—vTBL"
    Range("m3") = "¶Šˆ‹­“x"
    Range("n3") = "”DP•â³"
    Range("o3") = "®"
    Range("p3") = "Š—v—Ê"
    Range("q3") = "“E—v"
    Range("E15").Copy Range("l4:n4")
    Range("F05").Copy Range("l6:n31")
    Range("l4").Name = "ks2_syoyo"
    Range("F05").Copy Range("p5:p31")

    With ActiveSheet.PageSetup
        .Orientation = xlLandscape                      '‰¡’·
'        .PrintHeadings = True                           's—ñ”Ô†
        .LeftMargin = Application.InchesToPoints(0.4)   '¶—]”’
        .RightMargin = Application.InchesToPoints(0.2)  '‰E—]”’
        .Zoom = False
        .FitToPagesWide = 1                             '‰¡‚P•Å
        .FitToPagesTall = 1                             'c‚P•Å
    End With
    Cells.EntireColumn.AutoFit                          '—ñ•
    Range("C:D,F:F,I:J,L:N,P:P").ColumnWidth = 10
    Range("G:G,K:K").ColumnWidth = 4
    Range("a01") = Range("Namej")
    Range("a02") = "[" & Range("Fcode") & "]"
    Range("q05") = "´ÈÙ·Ş°1"
    Range("q06") = "”N—ß«•Ê‚Ù‚©"
    Range("q07") = "‚½‚ñ‚Ï‚­¿ 1/2"
    Range("q08") = "‚½‚ñ‚Ï‚­¿ 1/2"
    Range("q09") = "´ÈÙ·Ş°1 & ¶Šˆ‹­“x"
    Range("q10") = "´ÈÙ·Ş°1‚©‚çŒvZ"
    Range("q11") = "´ÈÙ·Ş°1 * 0.0099"
    Range("q12") = "À‘Ìd‚ÆŒW”"
    Range("q13") = "¶Ù¼³Ñ“¯’l"
    Range("q14") = "”N—ß"
    Range("q15") = "TBL’l(‚ŒŒˆ³‚Íw’è’l)"
    Range("q16") = "TBL’l"
    Range("q17") = "´ÈÙ·Ş°2 * 0.0004"
    Range("q18") = "´ÈÙ·Ş°2 * 0.00055"
    Range("q19") = "´ÈÙ·Ş°2 * 0.0066"
    Range("q20") = "TBL’l"
    Range("q21") = "(TBL’l)"
    Range("q22") = "(TBL’l)"
    Range("q23") = "(TBL’l)"
    Range("q24") = "•s–O˜a‰–b_ * 0.6"
    Range("q25") = "ÅÄØ³Ñ“¯’l"
    Range("q26") = "¶Ù¼³Ñ 1/2"
    Range("q27") = "‚ŒŒˆ³‚Íw’è’l"
    Range("q28") = "ˆê—¥"
    Range("q29") = "‰¿ 66%"
    Range("q30") = "‰¿ 34%"
    Range("q31") = "“œ”A/³´²Ä¥ºİÄÛ°Ù/ˆê”Ê"
    Sheets("ŒŸØ").Select
    Eiyo01_513kenso2sht = 0
End Function
'--------------------------------------------------------------------------------
'   01_514@ÛHŒvZ
'--------------------------------------------------------------------------------
Function Eiyo01_514MealCalc1() As Long
Dim aa      As Variant
Dim bb      As Worksheet
Dim i1      As Long     'sIndex
Dim i2      As Long     '—“Index
Dim Lmax    As Long     'sMax
Dim Wtemp1  As Double
Dim wtemp2  As Double
    
    aa = Sheets("ÛH").UsedRange
    Set bb = Sheets("ŒŸØ")
    Lmax = UBound(aa, 1)
    For i1 = 5 To Lmax
        For i2 = 1 To 27
            If i1 = 10 And i2 = 2 Then
                Wtemp1 = 1
            End If
            
'           ‰h—{‘fŒvZ =   Ûæ—Ê(F) * ‰h—{‘f(S)       * Š·Z’l(M)
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
'           ´ÈÙ·Ş°C              =      Ûæ—Ê(F) * ´ÈÙ·Ş°C(AT)     * Š·Z’l(M)
            Cells(i1, i2 + 54) = WorksheetFunction.Round(aa(i1, 6) * aa(i1, i2 + 45) * aa(i1, 13), 2)
'           ´ÈÙ·Ş°W              =      Ûæ—Ê(F) * ´ÈÙ·Ş°W(BI)     * Š·Z’l(M)
            Cells(i1, i2 + 69) = WorksheetFunction.Round(aa(i1, 6) * aa(i1, i2 + 60) * aa(i1, 13), 2)
'           ¶Ù¼³Ñ  =       Ûæ—Ê(F) * ¶Ù¼³Ñ           * Š·Z’l
            Wtemp1 = WorksheetFunction.Round(aa(i1, 6) * aa(i1, i2 + 75) * aa(i1, 13), 2)
            If aa(i1, 16) = 2 Then
                wtemp2 = Wtemp1
            Else
                wtemp2 = WorksheetFunction.Round(Wtemp1 * (100 - Fld_Field(8, 20)) / 100, 2)  '”M‘¹‚Í(8)
            End If
            Cells(i1, i2 + 84) = Wtemp1
            Cells(i1, i2 + 99) = wtemp2
        Next i2
        For i2 = 1 To 3
'           ‰¿    =      Ûæ—Ê(F) * ‰¿(CM)        * Š·Z’l(M)
            Wtemp1 = WorksheetFunction.Round(aa(i1, 6) * aa(i1, i2 + 90) * aa(i1, 13), 2)
            If aa(i1, 16) = 2 Then
                wtemp2 = Wtemp1
            Else
                wtemp2 = WorksheetFunction.Round(Wtemp1 * (100 - Fld_Field(i2, 20)) / 100, 2) '”M‘¹?
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
'   01_515@ÛH—Ê‡Œv
'--------------------------------------------------------------------------------
Function Eiyo01_515MealTotal() As Long
Dim Lmax    As Long     'sMax
Dim i1      As Long     'sIndex
Dim i2      As Long     '—“Index
Dim Wnissu  As Long

    Wnissu = Range("Nissu")
    Lmax = Sheets("ÛH").UsedRange.Rows.Count
    i1 = Lmax + 2
    For i2 = 1 To 120
        Cells(i1, i2) = "=SUM(R5C:R[-2]C)"
        Cells(i1 + 1, i2) = "=round(R[-1]C/" & Wnissu & ",2)"
    Next i2
    For i2 = 2 To 54 Step 2
        Range("ks2_eiyoso").Offset(i2 / 2 - 1, 0) = Cells(i1, i2)
        Range("ks2_eiyoso").Offset(i2 / 2 - 1, 1) = Cells(i1 + 1, i2)
    Next i2
    
    i1 = i1 + 1                             'ˆê“ú“–‚½‚è‡ŒvsIndex
    If Mid(Range("Q3rec"), 7, 1) = "3" Then '”––¡‚É‚æ‚é•â³
        If Cells(i1, 45) > 17 Then                           '¼µ16G
            Cells(i1, 45) = Cells(i1, 45) - 8                 ' -8G
            Cells(i1, 21) = Cells(i1, 21) - 3149              'ÅÄØ³Ñ
        ElseIf Cells(i1, 45) > 9 Then                        ' 9G²¶
            Cells(i1, 21) = Cells(i1, 21) - _
                          WorksheetFunction.RoundDown( _
                          (Cells(i1, 45) - 9) / 0.00254, 2)
            Cells(i1, 45) = 9                                 '²ÁØÂ
        End If
        If Cells(i1, 46) > 17 Then                           '¼µ16G
            Cells(i1, 46) = Cells(i1, 46) - 8                ' -8G
            Cells(i1, 22) = Cells(i1, 22) - 3149             'ÅÄØ³Ñ
            Range("ks2_hosei23") = 1
        ElseIf Cells(i1, 46) > 9 Then                        ' 9G²¶
            Cells(i1, 22) = Cells(i1, 22) - _
                          WorksheetFunction.RoundDown( _
                          (Cells(i1, 46) - 9) / 0.00254, 2)
            Cells(i1, 46) = 9                                 '²ÁØÂ
            Range("ks2_hosei23") = 2
        Else
            Range("ks2_hosei23") = 3
        End If
    Else
        Range("ks2_hosei23") = 4
    End If
    Range("ks2_hosei11") = Range("ks2_hosei23")
    
    
    If Cells(i1, 32) < 40 Then              'VC•â³(16)
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
    If Cells(i1, 40) < 3 Then               'VE•â³(20)
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
'   01_521@Šî‘bî•ñ‚Ù‚©æ“¾
'           ˆø” Func 1:XV‚ ‚è 2:XV‚È‚µ
'--------------------------------------------------------------------------------
Function Eiyo01_521CalcDbGet(Func As Long) As Long
Dim mySqlStr    As String
Dim i1          As Long
Dim i2          As Long

    Call Eiyo99_w’èƒV[ƒgíœ("DBmirror")
    Sheets.Add After:=Sheets(Sheets.Count)      'ƒV[ƒg’Ç‰Á
    ActiveSheet.Name = "DBmirror"
    Sheets("DBmirror").Range("m2:m3").NumberFormatLocal = "@"     'ZŠ[‚Q
'   Šî‘bî•ñæ“¾
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
'   Š—v—Êæ“¾
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
'   ƒGƒlƒ‹ƒM[^ƒJƒƒŠ[
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
        i2 = Sheets("ŒŸØ").UsedRange.Rows.Count
'       ‰¿‚Ì–ß‚µ
        Rst_Kiso.Fields("Sfods1").Value = Sheets("ŒŸØ").Cells(i2, 115)
        Rst_Kiso.Fields("Sfods2").Value = Sheets("ŒŸØ").Cells(i2, 116)
        Rst_Kiso.Fields("Sfods3").Value = Sheets("ŒŸØ").Cells(i2, 117)
        Rst_Kiso.Fields("Sfodh1").Value = Sheets("ŒŸØ").Cells(i2, 118)
        Rst_Kiso.Fields("Sfodh2").Value = Sheets("ŒŸØ").Cells(i2, 119)
        Rst_Kiso.Fields("Sfodh3").Value = Sheets("ŒŸØ").Cells(i2, 120)
'       Š—v—Ê‚Ì–ß‚µ
        For i1 = 1 To 27
            Rst_Syoyo.Fields(i1 * 5 - 3).Value = Sheets("ŒŸØ").Cells(i2, i1 * 2 - 1)
            Rst_Syoyo.Fields(i1 * 5 - 2).Value = Sheets("ŒŸØ").Cells(i2, i1 * 2)
        Next i1
'       ƒGƒlƒ‹ƒM[^ƒJƒƒŠ[‚Ì–ß‚µ
        For i1 = 1 To 15
            Rst_Energ.Fields(i1 + 1).Value = Sheets("ŒŸØ").Cells(i2, i1 + 54)
            Rst_Energ.Fields(i1 + 16).Value = Sheets("ŒŸØ").Cells(i2, i1 + 69)
            Rst_Energ.Fields(i1 + 31).Value = WorksheetFunction.Round(Sheets("ŒŸØ").Cells(i2, i1 + 54) / 80, 2)
            Rst_Energ.Fields(i1 + 58).Value = Sheets("ŒŸØ").Cells(i2, i1 + 84)
            Rst_Energ.Fields(i1 + 73).Value = Sheets("ŒŸØ").Cells(i2, i1 + 99)
        Next i1
    End If
    
    Eiyo01_521CalcDbGet = 0
End Function
'--------------------------------------------------------------------------------
'   01_522 •W€‘Ìd‚Ù‚©
'--------------------------------------------------------------------------------
Function Eiyo01_522Mealcalc2() As Long
Dim mySqlStr    As String
Dim À‘Ìd      As Double
Dim •W€‘Ìd    As Double
Dim ˜Jì        As String   '˜Jì‹­“x@Q7.E‹Æ‚Ì‚SŒ…–Ú
Dim ”DP        As Long
Dim Wtemp       As Double
Dim i1          As Long
Dim Wcondition  As Long     'Š—v—ÊƒGƒlƒ‹ƒM[“K—pğŒ
Dim Wenerg1     As Double   'w’è‚ ‚è‚Ì’²®´ÈÙ·Ş
Dim Wenerg2     As Double   'w’èœŠO‚Ì’²®´ÈÙ·Ş
Dim KisoCd1     As Long     '•K—v—Êƒ}ƒXƒ^‚ÌŠî‘bƒR[ƒh  Šî‘b‘ãÓAŠˆ“®‘ãÓ
Dim KisoCd2     As Long     '•K—v—Êƒ}ƒXƒ^‚ÌŠî‘bƒR[ƒh@Š—v—Ê
Dim KisoCd3     As Long     '•K—v—Êƒ}ƒXƒ^‚ÌŠî‘bƒR[ƒh@”DPö“û
Dim ”N—î        As Long
Dim Warray      As Variant
Dim Wtext       As String

    À‘Ìd = Range("Weght")
    ”N—î = Range("Age")
    ˜Jì = Mid(Range("Qjob1"), 4, 1)
    ”DP = Range("Qcnd1")
'   •W€‘Ìd
    If ”N—î <= 12 Then
        •W€‘Ìd = À‘Ìd
    ElseIf Range("Hight") <= 150 Then
        •W€‘Ìd = Range("Hight") - 100
    ElseIf Range("Hight") <= 165 Then
        •W€‘Ìd = WorksheetFunction.Round((Range("Hight") - 100) * 0.9, 1)
    Else
        •W€‘Ìd = Range("Hight") - 110
    End If
    Rst_Kiso.Fields("Aans1").Value = •W€‘Ìd
    Range("ks2_weght") = À‘Ìd
    Range("ks2_weght").Offset(0, 1) = •W€‘Ìd
'   ”ì–“x
    Rst_Kiso.Fields("Himanp").Value = WorksheetFunction.Round( _
                                 (À‘Ìd - •W€‘Ìd) / À‘Ìd * 100, 0)
'   ‘ÌŠiw”
    If ”N—î <= 2 Then
        Wtemp = À‘Ìd / (Range("Hight") ^ 2) * 10 ^ 4              '¶³Ìßw”
    ElseIf ”N—î <= 12 Then
        Wtemp = À‘Ìd / (Range("Hight") ^ 3) * 10 ^ 7              'Û°ÚÙw”
    Else
        Wtemp = WorksheetFunction.Round(À‘Ìd / •W€‘Ìd * 100, 0) 'ÌŞÛ°¶°w”
    End If
    Rst_Kiso.Fields("Taiis").Value = Wtemp
'   ‘Ì•\–ÊÏ
    Rst_Kiso.Fields("Aansa").Value = Eiyo01_523_taihyou(À‘Ìd)
    Rst_Kiso.Fields("Bansa").Value = Eiyo01_523_taihyou(•W€‘Ìd)
    Range("ks2_Aansa") = Rst_Kiso.Fields("Aansa").Value
    Range("ks2_Aansa").Offset(0, 1) = Rst_Kiso.Fields("Bansa").Value
'   ¶Šˆw”
    Select Case ˜Jì
        Case "A":  Wtemp = 0.35
        Case "B":  Wtemp = 0.5
        Case "C":  Wtemp = 0.75
        Case Else: Wtemp = 1#
    End Select
    Rst_Kiso.Fields("Aansx").Value = Wtemp
    Range("ks2_Aansx") = Wtemp
'   ’PˆÊ•\–ÊÏ‚ ‚½‚è‚ÌŠî‘b‘ãÓ
    If ”N—î < 20 Then
        KisoCd1 = ”N—î
    ElseIf ”N—î < 80 Then
        KisoCd1 = Int(”N—î / 10) * 10
    Else
        KisoCd1 = 80
    End If
    If Range("Sex") = 1 Then: KisoCd1 = KisoCd1 + 100
    mySqlStr = "SELECT Šî‘b‘ãÓ,Šˆ“®‘ãÓ FROM " & Tbl_Need & " Where Ncode = " & KisoCd1
    Set Rst_Need = myCon.Execute(mySqlStr)
    If Rst_Need.EOF Then
        MsgBox "•K—v—Êƒ}ƒXƒ^‚ÌƒL[‚È‚µ:" & KisoCd1
    End If
    Rst_Kiso.Fields("Aans3").Value = Rst_Need.Fields("Šî‘b‘ãÓ").Value
    Range("ks2_kisocd") = KisoCd1
    Range("ks2_kisot") = Rst_Kiso.Fields("Aans3").Value
'   Šî‘b‘ãÓ
    Rst_Kiso.Fields("Aansb").Value = WorksheetFunction.Round(Rst_Kiso.Fields("Aans3").Value _
                                                           * Rst_Kiso.Fields("Aansa").Value * 24, 2)    'À‘Ìd‚ÌŠî‘b‘ãÓ^“ú
    Rst_Kiso.Fields("Aansc").Value = WorksheetFunction.Round(Rst_Kiso.Fields("Aansb").Value / 1440, 2)  'À‘Ìd‚ÌŠî‘b‘ãÓ^•ª
    Rst_Kiso.Fields("Aansd").Value = Eiyo01_524ansd(Rst_Kiso.Fields("Aansb").Value)                     'À‘Ìd‚ÌE•W€—Ê
    
    Rst_Kiso.Fields("Bansb").Value = WorksheetFunction.Round(Rst_Kiso.Fields("Aans3").Value _
                                                           * Rst_Kiso.Fields("Bansa").Value * 24, 2)    '•W€‘Ìd‚ÌŠî‘b‘ãÓ^“ú
    Rst_Kiso.Fields("Bansc").Value = WorksheetFunction.Round(Rst_Kiso.Fields("Bansb").Value / 1440, 2)  '•W€‘Ìd‚ÌŠî‘b‘ãÓ^•ª
    Rst_Kiso.Fields("Bansd").Value = Eiyo01_524ansd(Rst_Kiso.Fields("Bansb").Value)                     '•W€‘Ìd‚ÌE•W€—Ê
    Range("ks2_Aansb") = Rst_Kiso.Fields("Aansb").Value                 'À‘Ìd‚ÌŠî‘b‘ãÓ^“ú
    Range("ks2_Aansc") = Rst_Kiso.Fields("Aansc").Value                 'À‘Ìd‚ÌŠî‘b‘ãÓ^•ª
    Range("ks2_Aansd") = Rst_Kiso.Fields("Aansd").Value                 'À‘Ìd‚ÌE•W€—Ê
    Range("ks2_Aansb").Offset(0, 1) = Rst_Kiso.Fields("Bansb").Value    '•W€‘Ìd‚ÌŠî‘b‘ãÓ^“ú
    Range("ks2_Aansc").Offset(0, 1) = Rst_Kiso.Fields("Bansc").Value    '•W€‘Ìd‚ÌŠî‘b‘ãÓ^•ª
    Range("ks2_Aansd").Offset(0, 1) = Rst_Kiso.Fields("Bansd").Value    '•W€‘Ìd‚ÌE•W€—Ê

'   Š—v—Ê@ƒGƒlƒ‹ƒM[  -------------------------------------------------------------------------------------
    If Rst_Kiso.Fields("Tenes").Value = 1 Then              'ƒGƒlƒ‹ƒM[w’èE’l
        Wenerg1 = Rst_Kiso.Fields("Tenee").Value
        Wenerg2 = Rst_Kiso.Fields("Aansd").Value
        Wcondition = 1
    ElseIf Rst_Kiso.Fields("Tenes").Value = 2 Then          'ƒGƒlƒ‹ƒM[w’èEÀ‘Ìd
        Wenerg1 = WorksheetFunction.RoundDown(Rst_Kiso.Fields("Tenee").Value * À‘Ìd, 2)
        Wenerg2 = Rst_Kiso.Fields("Aansd").Value
        Wcondition = 2
    ElseIf Rst_Kiso.Fields("Tenes").Value = 3 Then          'ƒGƒlƒ‹ƒM[w’èE•W€‘Ìd
        Wenerg1 = WorksheetFunction.RoundDown(Rst_Kiso.Fields("Tenee").Value * •W€‘Ìd, 2)
        Wenerg2 = Rst_Kiso.Fields("Aansd").Value
        Wcondition = 3
    ElseIf ”DP = 1 Then                                    '”DP‘OŠú
        Wenerg1 = Rst_Kiso.Fields("Bansd").Value + 150
        Wenerg2 = Rst_Kiso.Fields("Aansd").Value + 150
        Wcondition = 4
    ElseIf ”DP = 2 Then                                    '”DPŒãŠú
        Wenerg1 = Rst_Kiso.Fields("Bansd").Value + 350
        Wenerg2 = Rst_Kiso.Fields("Aansd").Value + 350
        Wcondition = 5
    ElseIf ”DP = 3 Then                                    'ö“ûŠú
        Wenerg1 = Rst_Kiso.Fields("Bansd").Value + 700
        Wenerg2 = Rst_Kiso.Fields("Aansd").Value + 700
        Wcondition = 6
    ElseIf Rst_Kiso.Fields("Qill1").Value <> 0 Then         '“œ”A•a
        Wenerg1 = Rst_Kiso.Fields("Bansd").Value - 200
        Wenerg2 = Rst_Kiso.Fields("Aansd").Value
        Wcondition = 7
    ElseIf Rst_Kiso.Fields("Qsrmr").Value <> 0 Then         'ƒXƒ|[ƒc
        Wenerg1 = Rst_Kiso.Fields("Aansb").Value _
                + WorksheetFunction.RoundDown((Rst_Kiso.Fields("Qsrmr").Value + 1.2) _
                                             * Rst_Kiso.Fields("Qsmin").Value _
                                             * Rst_Need.Fields("Šˆ“®‘ãÓ").Value _
                                             * Rst_Kiso.Fields("Aansc").Value, 2)
        Wenerg2 = Wenerg1
        Wcondition = 8
    ElseIf Rst_Kiso.Fields("Taiis").Value <= 90 Or _
           Rst_Kiso.Fields("Taiis").Value >= 120 Then      '”ì–
        Wenerg1 = Rst_Kiso.Fields("Bansd").Value
        Wenerg2 = Rst_Kiso.Fields("Aansd").Value
        Wcondition = 9
    Else                                                    '‚»‚Ì‘¼ˆê”Ê
        Wenerg1 = Rst_Kiso.Fields("Aansd").Value
        Wenerg2 = Rst_Kiso.Fields("Aansd").Value
        Wcondition = 10
    End If
    Rst_Syoyo.Fields("Syoyo01").Value = Wenerg1
    Range("ks2_energ") = Wcondition
    Range("ks2_energ").Offset(1, 0) = Wenerg1
    Range("ks2_energ").Offset(2, 0) = Wenerg2
    
'   Š—v—Ê@‚»‚Ì‘¼  ------------------------------------------------------------------------------------------
    KisoCd2 = KisoCd1 + 1000
    mySqlStr = "SELECT * FROM " & Tbl_Need & " Where Ncode = " & KisoCd2
    Set Rst_Need = myCon.Execute(mySqlStr)
    If Rst_Need.EOF Then
        MsgBox "•K—v—Êƒ}ƒXƒ^‚ÌƒL[‚È‚µ:" & KisoCd2
    End If
    Range("ks2_syoyo") = KisoCd2
    For i1 = 2 To 27
        Rst_Syoyo.Fields(i1 * 5 - 1).Value = Rst_Need.Fields(i1 + 1).Value
        Range("ks2_syoyo").Offset(i1, 0) = Rst_Need.Fields(i1 + 1).Value
    Next i1
    If ˜Jì = "B" Or ”N—î < 15 Then
    Else
        Select Case ˜Jì
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
            MsgBox "•K—v—Êƒ}ƒXƒ^‚ÌƒL[‚È‚µ:" & KisoCd2
        End If
        Range("ks2_syoyo").Offset(0, 1) = KisoCd2
        For i1 = 2 To 27
            Rst_Syoyo.Fields(i1 * 5 - 1).Value = _
            Rst_Syoyo.Fields(i1 * 5 - 1).Value + Rst_Need.Fields(i1 + 1).Value
            Range("ks2_syoyo").Offset(i1, 1) = Rst_Need.Fields(i1 + 1).Value
        Next i1
    End If
    Select Case ”DP
        Case 1:    KisoCd3 = 1401   '”DP‘OŠú
        Case 2:    KisoCd3 = 1402   '”DPŒãŠú
        Case 3:    KisoCd3 = 1403   'ö“ûŠú
        Case Else: KisoCd3 = 0
    End Select
    If KisoCd3 > 0 Then
        mySqlStr = "SELECT * FROM " & Tbl_Need & " Where Ncode = " & KisoCd3
        Set Rst_Need = myCon.Execute(mySqlStr)
        If Rst_Need.EOF Then
            MsgBox "•K—v—Êƒ}ƒXƒ^‚ÌƒL[‚È‚µ:" & KisoCd3
        End If
        Range("ks2_syoyo").Offset(0, 2) = KisoCd3
        For i1 = 2 To 27
            Rst_Syoyo.Fields(i1 * 5 - 1).Value = _
            Rst_Syoyo.Fields(i1 * 5 - 1).Value + Rst_Need.Fields(i1 + 1).Value
            Range("ks2_syoyo").Offset(i1, 2) = Rst_Need.Fields(i1 + 1).Value
        Next i1
    End If
'   Š—v—Ê@‚½‚ñ‚Ï‚­¿  --------------------------------------------------------------------------------------
    Wcondition = 0
    If Rst_Kiso.Fields("Tanps").Value = 1 Then              '‚½‚ñ‚Ï‚­w’èE’l
        Wtemp = Rst_Kiso.Fields("Tanpe").Value
        Wcondition = 1
    ElseIf Rst_Kiso.Fields("Tanps").Value = 2 Then          '‚½‚ñ‚Ï‚­w’èEÀ‘Ìd
        Wtemp = WorksheetFunction.RoundDown(Rst_Kiso.Fields("Tanpe").Value * À‘Ìd, 2)
        Wcondition = 2
    ElseIf Rst_Kiso.Fields("Tanps").Value = 3 Then          '‚½‚ñ‚Ï‚­w’èE•W€‘Ìd
        Wtemp = WorksheetFunction.RoundDown(Rst_Kiso.Fields("Tanpe").Value * •W€‘Ìd, 2)
        Wcondition = 3
    ElseIf Rst_Kiso.Fields("Tanps").Value = 4 Then          '‚½‚ñ‚Ï‚­w’èEƒGƒlƒ‹ƒM[”ä
        Wtemp = WorksheetFunction.RoundDown(Rst_Kiso.Fields("Tanpe").Value _
                                          * Wenerg1 / 400, 2)
        Wcondition = 4
    ElseIf Rst_Kiso.Fields("Qsrmr").Value <> 0 Then         'ƒXƒ|[ƒc
        Wtemp = À‘Ìd * 1.4
        Wcondition = 5
    ElseIf Rst_Kiso.Fields("Qwcnt").Value <> 0 Then         '³´²Ä¥ºİÄÛ°Ù
        If Wenerg1 < 1412 Then
            Wtemp = 60
            Wcondition = 6
        Else
            Wtemp = Wenerg1 * 17 / 400
            Wcondition = 7
        End If
    Else
        If ”N—î < 21 Then
            If Range("Sex") = 0 Then
'                          3   4   5   6   7   8   9  10  11  12  13  14  15  16  17  18  19  20Ë
                Wtext = "117,116,122,124,128,135,138,144,140,138,136,129,125,121,117,117,113,113"
            Else
                Wtext = "117,120,126,132,137,144,149,148,144,146,139,133,131,129,129,124,121,121"
            End If
            Warray = Split(Wtext, ",")
            Wtemp = Warray(”N—î - 3)
            If Rst_Kiso.Fields("Qsyog").Value = 1 Then: Wtemp = Wtemp + 20  'áŠQÒ
            Wcondition = 8
        Else
            i1 = Int(”N—î / 10) - 2
            If i1 > 6 Then: i1 = 6
            Select Case ˜Jì       ' 20  30  40  50  60  70  80      Ë‘ã
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
        Select Case ”DP
            Case 1: Wtemp = Wtemp + 10  '”DP‘OŠú
            Case 2: Wtemp = Wtemp + 20  '”DPŒãŠú
            Case 3: Wtemp = Wtemp + 20  'ö“ûŠú
        End Select
    End If
    Wtemp = WorksheetFunction.RoundDown(Wtemp, 2)
    Rst_Syoyo.Fields("Syoyo02").Value = WorksheetFunction.RoundDown(Wtemp, 2)       '‚½‚ñ‚Ï‚­¿  (02)
    Rst_Syoyo.Fields("Syoyo03").Value = WorksheetFunction.RoundDown(Wtemp / 2, 2)   '“®•¨‚½‚ñ‚Ï‚­(03)
    Rst_Syoyo.Fields("Syoyo04").Value = WorksheetFunction.RoundDown(Wtemp / 2, 2)   'A•¨‚½‚ñ‚Ï‚­(04)
    Range("ks2_syoyo").Offset(2, 3) = Wcondition
'   Š—v—Ê@‰¿  --------------------------------------------------------------------------
    If ”DP > 0 Then
        Wtemp = 275
    ElseIf ”N—î < 21 Then
        If ˜Jì = "A" Then
            Wtemp = 225
        Else
            Wtemp = 275
        End If
    Else
        Select Case ˜Jì
            Case "A", "B": Wtemp = 225
            Case Else:     Wtemp = 275
        End Select
    End If
    Wtemp = WorksheetFunction.RoundDown(Wenerg1 * Wtemp / 9000, 2)
    Rst_Syoyo.Fields("Syoyo05").Value = Wtemp                                       '‰¿   (05)
    Rst_Syoyo.Fields("Syoyo26").Value = WorksheetFunction.Round(Wtemp * 0.34, 2)    'S      (24)
    Rst_Syoyo.Fields("Syoyo25").Value = WorksheetFunction.Round(Wtemp * 0.66, 2)    'P      (25)
    Rst_Syoyo.Fields("Syoyo24").Value = 300                                         'ºÚ½ÃÛ°Ù(24)
    
    Rst_Syoyo.Fields("Syoyo06").Value = WorksheetFunction.RoundDown((Wenerg1 _
                                      - Rst_Syoyo.Fields("Syoyo02").Value * 4 _
                                      - Rst_Syoyo.Fields("Syoyo05").Value * 9) / 4, 2)  '“œ¿(06)
    If Rst_Kiso.Fields("Qill1").Value <> 0 Then                                         '»“œ(27)
        Rst_Syoyo.Fields("Syoyo27").Value = 10          '“œ”A•a
    ElseIf Rst_Kiso.Fields("Qwcnt").Value <> 0 Then
        Rst_Syoyo.Fields("Syoyo27").Value = 10          '³´²Ä¥ºİÄÛ°Ù
    Else
        Rst_Syoyo.Fields("Syoyo27").Value = 30          '‚»‚Ì‘¼(ˆê”Ê)
    End If
    Rst_Syoyo.Fields("Syoyo07").Value = WorksheetFunction.RoundDown(Wenerg1 * 0.0099, 2)    'H•¨‚¹‚ñ‚¢(07)
'   ¶Ù¼³Ñ(08) ------------------------------------------------------------------------------------------------
    Wcondition = 0
    If ”N—î < 21 Then
        If Range("Sex") = 0 Then
'                      3   4   5   6   7   8   9  10  11  12  13  14  15  16  17  18  19  20Ë
            Wtext = "171,168,169,173,176,165,169,177,187,188,177,156,134,124,115,109,103,103"
        Else
            Wtext = "173,169,169,174,182,177,184,184,175,158,142,133,119,108,100,100,100,100"
        End If
        Warray = Split(Wtext, ",")
        Wtemp = WorksheetFunction.RoundDown(À‘Ìd * Warray(”N—î - 3) / 10, 2)
        Wcondition = 1
    ElseIf ”N—î < 60 Then
        Wtemp = À‘Ìd * 10
        Wcondition = 2
    Else
        Wtemp = 600
        Wcondition = 3
    End If
    Select Case ”DP
        Case 1, 2
            Wtemp = Wtemp + 400 '”DP‘OŒãŠú
            Wcondition = 4
        Case 3
            Wtemp = Wtemp + 500 'ö“ûŠú
            Wcondition = 5
    End Select
    Rst_Syoyo.Fields("Syoyo08").Value = Wtemp   '¶Ù¼³Ñ(08)
    Rst_Syoyo.Fields("Syoyo09").Value = Wtemp   'ƒŠƒ“ (09)
    Range("ks2_syoyo").Offset(8, 3) = Wcondition
'   “S  ------------------------------------------------------------------------------------------------------
    Select Case ”DP
        Case 1:    Wtemp = 15               '”DP‘OŠú
        Case 2, 3: Wtemp = 20               '”DPŒãŠúEö“ûŠú
        Case Else
            Select Case ”N—î
                Case 1 To 5:   Wtemp = 8    '    ‚TËˆÈ‰º
                Case 6 To 8:   Wtemp = 9    ' 6` 8Ë
                Case 9 To 11:  Wtemp = 10   ' 9`11Ë
                Case 12 To 19: Wtemp = 12   '12`19Ë
                Case 20 To 49
                    Select Case Range("Sex")
                        Case 0:    Wtemp = 10   '20`49Ë‚Ì’j
                        Case Else: Wtemp = 12   '20`49Ë‚Ì—
                    End Select
                Case Else: Wtemp = 10           '50ËˆÈã
            End Select
    End Select
    Rst_Syoyo.Fields("Syoyo10").Value = Wtemp
'   VB1/VB2/Å²±¼İ  -------------------------------------------------------------------------------------------
    Rst_Syoyo.Fields("Syoyo13").Value = WorksheetFunction.RoundDown(Wenerg2 * 0.0004, 2)
    Rst_Syoyo.Fields("Syoyo14").Value = WorksheetFunction.RoundDown(Wenerg2 * 0.00055, 2)
    Rst_Syoyo.Fields("Syoyo15").Value = WorksheetFunction.RoundDown(Wenerg2 * 0.0066, 2)
    Select Case ”DP
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
        Rst_Syoyo.Fields("Syoyo23").Value = 6       '¼µ
        Rst_Syoyo.Fields("Syoyo11").Value = 2800    'ÅÄØ³Ñ
        Range("ks2_syoyo").Offset(11, 3) = 1
        Range("ks2_syoyo").Offset(23, 3) = 1
    Else
        Rst_Syoyo.Fields("Syoyo23").Value = 10      '¼µ
        Range("ks2_syoyo").Offset(23, 3) = 2
    End If
    Rst_Syoyo.Fields("Syoyo20").Value = WorksheetFunction.Round(Rst_Syoyo.Fields("Syoyo25").Value * 0.6, 2)     'VE
    Rst_Syoyo.Fields("Syoyo21").Value = Rst_Syoyo.Fields("Syoyo11").Value                                       '¶Ø³Ñ <= ÅÄØ³Ñ
    Rst_Syoyo.Fields("Syoyo22").Value = WorksheetFunction.RoundDown(Rst_Syoyo.Fields("Syoyo08").Value / 2, 2)   'Mg=Ca/2
'   ‰h—{‘f”ä—¦
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

'   XVŒ‹‰Ê•\¦
    For i1 = 1 To Rst_Kiso.Fields.Count
        Cells(3, i1).Value = Rst_Kiso.Fields(i1 - 1).Value
    Next
    For i1 = 1 To 27
        Range("ks2_syoyo").Offset(i1, 4) = Rst_Syoyo.Fields(i1 * 5 - 1).Value
    Next i1
End Function
'--------------------------------------------------------------------------------
'   01_523 ‘Ì•\–ÊÏ
'       ‚TÎˆÈ‰º    ‘Ìd^0.423 * g’·^0.362 * 382.89 / 10000
'       ‚UÎˆÈã    ‘Ìd^0.444 * g’·^0.663 *  88.83 / 10000
'--------------------------------------------------------------------------------
Function Eiyo01_523_taihyou(‘Ìd As Double) As Double
Dim Wtemp   As Double

    If Range("Age") < 6 Then
        Wtemp = WorksheetFunction.Round(‘Ìd ^ 0.423 * Range("hight") ^ 0.362 * 382.89 / 10000, 2)
    Else
        Wtemp = WorksheetFunction.Round(‘Ìd ^ 0.444 * Range("hight") ^ 0.663 * 88.83 / 10000, 2)
    End If
    Eiyo01_523_taihyou = Wtemp
End Function
'--------------------------------------------------------------------------------
'   01_524 ƒGƒlƒ‹ƒM[•W€—Ê@¶ŠˆŠˆ“®‹­“x•â³
'       áŠQÒ      50%
'       ‚U‚OÎ‘ã    90%
'       ‚V‚OÎ‘ã    80%
'       ‚W‚OÎˆÈã  70%
'       Šî‘b‘ãÓ^“ú * (•â³¶ŠˆŠˆ“®‹­“x+1) * 1.1
'--------------------------------------------------------------------------------
Function Eiyo01_524ansd(Šî‘b‘ãÓ As Double) As Double
Dim Wtemp   As Double
    
    If Rst_Kiso.Fields("Qsyog").Value = 1 Then                      'áŠQÒ
        Wtemp = WorksheetFunction.Round(Rst_Kiso.Fields("Aansx").Value * 0.5, 2)
    ElseIf Range("Age") < 60 Then                                   '‚U‚OÎ–¢–
        Wtemp = Rst_Kiso.Fields("Aansx").Value
    ElseIf Range("Age") >= 60 And Range("Age") <= 69 Then           '‚U‚OÎ‘ã
        Wtemp = WorksheetFunction.Round(Rst_Kiso.Fields("Aansx").Value * 0.9, 2)
    ElseIf Range("Age") >= 70 And Range("Age") <= 79 Then           '‚V‚OÎ‘ã
        Wtemp = WorksheetFunction.Round(Rst_Kiso.Fields("Aansx").Value * 0.8, 2)
    Else                                                            '‚W‚OÎˆÈã
        Wtemp = WorksheetFunction.Round(Rst_Kiso.Fields("Aansx").Value * 0.7, 2)
    End If
    Wtemp = WorksheetFunction.Round(Šî‘b‘ãÓ * (1 + Wtemp) * 1.1, 2)
    Eiyo01_524ansd = Wtemp
End Function
'--------------------------------------------------------------------------------
'   01_525 ‰ß•s‘«ŒvZAƒAƒhƒoƒCƒX
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
                                          / Rst_Syoyo.Fields(i1 * 5 - 1).Value * 100, 0) - 100   '‰ß•s‘«—¦
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
'   ƒAƒhƒoƒCƒX
    Wtext = Rst_Kiso.Fields("Q3rec").Value      'HKŠµ
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
        Wans1 = 98  '2008/4/25 3050‚ğ0098‚É•ÏX
    End If
    Rst_Kiso.Fields("Badv1").Value = Wans1
    
    Wtext = Rst_Kiso.Fields("Q4rec").Value      '‹x—{
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
        Wans1 = 98  '2008/4/25 3150‚ğ0098‚É•ÏX
    End If
    Rst_Kiso.Fields("Badv2").Value = Wans1
    
    Wtext = Rst_Kiso.Fields("Q5rec").Value      '‰^“®
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
        Wans1 = 98  '2008/4/25 3250‚ğ0098‚É•ÏX
    End If
    Rst_Kiso.Fields("Badv3").Value = Wans1
    
    Wtext = Rst_Kiso.Fields("Q6r_a").Value & Rst_Kiso.Fields("Q6r_b").Value _
          & Rst_Kiso.Fields("Q6r_c").Value & Rst_Kiso.Fields("Q6r_d").Value _
          & Rst_Kiso.Fields("Q6r_e").Value                              'Œ’N’²¸
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
    
    If Rst_Kiso.Fields("Qsrmr").Value <> 0 Then            '  ³´²Ä ±ÄŞÊŞ²½
       Wans1 = 2801
    ElseIf Rst_Kiso.Fields("age").Value <= 12 Then
       Wans1 = 2806
    ElseIf Rst_Kiso.Fields("Taiis").Value < 120 Or _
           Rst_Kiso.Fields("Qcnd1").Value = 1 Or _
           Rst_Kiso.Fields("Qcnd1").Value = 2 Then      ' 120%ĞÏİ OR Æİ¼İ
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
    If Rst_Syoyo.Fields("Syort20").Value < -37 Then     'VE Ì¿¸
       If Rst_Kiso.Fields("age").Value < 40 And _
          Rst_Kiso.Fields("Sex").Value = 1 Then
           Call Eiyo01_526Cadvs(3630)                   '39²¶  µİÅ
       Else
           Call Eiyo01_526Cadvs(3610)                   'µÄº & 40²¼Ş®³ µİÅ
       End If
    End If
    If Rst_Syoyo.Fields("Syort12") < -37 Then: Call Eiyo01_526Cadvs(3640)   'VA
    If Rst_Syoyo.Fields("Syort13") < -37 Then: Call Eiyo01_526Cadvs(3620)   'VB1
    If Rst_Syoyo.Fields("Syort08") < -37 Then: Call Eiyo01_526Cadvs(3650)   'CA
    If Rst_Syoyo.Fields("Syort07") < -37 Then: Call Eiyo01_526Cadvs(3660)   '¾İ²
    
    If Rst_Syoyo.Fields("Syort23") > 12 Then: Call Eiyo01_527Cadvs(3730)    '¼µ
    If Rst_Energ.Fields("Enet08") < 85 Then: Call Eiyo01_527Cadvs(3720)     '3¸Şİ ´²Ö³ ¾¯¼­
    If Rst_Syoyo.Fields("Syort08") < -37 Then: Call Eiyo01_527Cadvs(3760)   'CA
    If Rst_Energ.Fields("Enet08") < 85 Or _
       Rst_Energ.Fields("Enet09") < 85 Then: Call Eiyo01_527Cadvs(3740)     '4¸Şİ
    If Rst_Kiso.Fields("PER04") > 50 Then: Call Eiyo01_527Cadvs(3710)      'ÄŞ³ÌŞÂ ÀİÊß¸¼Â Ë
    
    
'                     ---- 0.35 ----  ---- 0.5 -----
    Wtext = Empty   '<=-110=><=111-=><=-110=><=111-=>
    Wtext = Wtext & "00012023000120270001021500010211"  '   -20 µÄº
    Wtext = Wtext & "00012028000120210001021600010214"  ' 21-30
    Wtext = Wtext & "00013335000133340001202800012029"  ' 31-40
    Wtext = Wtext & "00013337000133360001202100012025"  ' 41-50
    Wtext = Wtext & "00012025000120290001021100010214"  ' 51-60
    Wtext = Wtext & "00012030000120260001021700010218"  ' 61-70
    Wtext = Wtext & "00012032000120310001020900010219"  ' 71-
    Wtext = Wtext & "00010204000102030001021000010212"  '   -20 µİÅ
    Wtext = Wtext & "00012021000120220001020300010211"  ' 21-30
    Wtext = Wtext & "00012024000120230001021300010212"  ' 31-40
    Wtext = Wtext & "00012026000120250001020700010214"  ' 41-50
    Wtext = Wtext & "00010205000120260001020500010208"  ' 51-60
    Wtext = Wtext & "00010206000102050001020600010209"  ' 61-70
    Wtext = Wtext & "00010209000102080001020900010208"  ' 71-
    If Rst_Kiso.Fields("Aansx") > 50 Then
        Wtext2 = "38394041"
    Else
        If Rst_Kiso.Fields("Taiis").Value < 111 Then   'À²¶¸ ¼½³
            Wans1 = 0
        Else
            Wans1 = 1
        End If
        If Rst_Kiso.Fields("Aansx") = 0.5 Then: Wans1 = Wans1 + 2        '¾²¶Â ¼½³
        Select Case Rst_Kiso.Fields("Age")
            Case 0 To 20:
            Case 21 To 30: Wans1 = Wans1 + 4
            Case 31 To 40: Wans1 = Wans1 + 8
            Case 41 To 50: Wans1 = Wans1 + 12
            Case 51 To 60: Wans1 = Wans1 + 16
            Case 61 To 70: Wans1 = Wans1 + 20
            Case Else:     Wans1 = Wans1 + 24
        End Select
        If Rst_Kiso.Fields("Sex") = 1 Then: Wans1 = Wans1 + 28        'µİÅ
        Wtext2 = Mid(Wtext, Wans1 * 8 + 1, 8)
    End If
    Rst_Kiso.Fields("Dadv1").Value = Left(Wtext2, 2)
    Rst_Kiso.Fields("Dadv2").Value = Mid(Wtext2, 3, 2)
    Rst_Kiso.Fields("Dadv3").Value = Mid(Wtext2, 5, 2)
    Rst_Kiso.Fields("Dadv4").Value = Mid(Wtext2, 7, 2)
    
    Eiyo01_525MealDiffe = 0
End Function
'--------------------------------------------------------------------------------
'   01_526@CƒAƒhƒoƒCƒX‚P
'--------------------------------------------------------------------------------
Function Eiyo01_526Cadvs(advc As Long)
    If Rst_Kiso.Fields("Cadv1").Value = 98 Then
        Rst_Kiso.Fields("Cadv1").Value = advc
    ElseIf Rst_Kiso.Fields("Cadv2").Value = 98 Then
        Rst_Kiso.Fields("Cadv2").Value = advc
    End If
End Function
'--------------------------------------------------------------------------------
'   01_527@CƒAƒhƒoƒCƒX‚Q
'--------------------------------------------------------------------------------
Function Eiyo01_527Cadvs(advc As Long)
    If Rst_Kiso.Fields("Cadv3").Value = 98 Then
        Rst_Kiso.Fields("Cadv3").Value = advc
    ElseIf Rst_Kiso.Fields("Cadv4").Value = 98 Then
        Rst_Kiso.Fields("Cadv4").Value = advc
    End If
End Function
'--------------------------------------------------------------------------------
'   01_528@‰h—{”ä—¦
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
'   01_540@‹ŒÛHŒvZ’l‚Ì”äŠr—p
'--------------------------------------------------------------------------------
Function Eiyo01_540Old_Check() As Long
Dim mySqlStr    As String
Dim i1          As Long
Dim i2          As Long
Dim Lmax1       As Long
Dim Lmax2       As Long
Dim Lmax3       As Long
Dim Errcnt      As Long

'   XVŒ‹‰Ê•\¦
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
    If Errcnt > 0 Then: MsgBox "•sˆê’v " & Errcnt
    
End Function
'--------------------------------------------------------------------------------
'   01_541@”äŠr
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
'   01_550@Šî‘bî•ñ‚Ù‚©Close
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
'   01_700@ƒJƒEƒ“ƒZƒŠƒ“ƒOƒV[ƒgì•\
'--------------------------------------------------------------------------------
Function Eiyo01_700ì•\Click()
    
    If IsEmpty(Range("Fcode")) Or _
       Range("Fcode") <> Range("Fsave") Then
        MsgBox "Šî‘bî•ñ‚ÌŒŸõ‚ªs‚í‚ê‚Ä‚¢‚Ü‚¹‚ñ"
        Exit Function
    End If
    Application.ScreenUpdating = False  '‰æ–Ê•`‰æ—}~
    Call Eiyo91DB_Open                  'DB Open
    Call Eiyo01_511MealFldgt            '€–Ú—v‘fæ“¾
    Call Eiyo01_701Sheet                '¶³İ¾Øİ¸Ş¼°Ä’Ç‰Á
    Call Eiyo01_702DbGet                'DB Get(521)
    Call Eiyo01_703Pset                 'ˆóü€–Ú‚Ìİ’è
    Call Eiyo01_704Advic                'ƒAƒhƒoƒCƒX
    Call Eiyo01_705Footer               'ƒR[ƒh•“ú•tAƒJƒ“ƒEƒZƒ‰[
    Call Eiyo920DB_Close                'DB Close
'    Call Eiyo99_w’èƒV[ƒgíœ("DBmirror")
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Select
End Function
'--------------------------------------------------------------------------------
'   01_701@ƒV[ƒg’Ç‰Á
'--------------------------------------------------------------------------------
Function Eiyo01_701Sheet()
Const ShtName = "¶³İ¾Øİ¸Ş¼°Ä"
Const Eiyo01Bk = "Eiyo01_Šî‘bÛH“ü—Í.xls"
Const Eiyo02Bk = "Eiyo02_¶³İ¾Øİ¸Ş¼°Ä.xls"
    
    Call Eiyo99_w’èƒV[ƒgíœ(ShtName)
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
'   01_703@ˆóü€–Ú‚Ìİ’è
'--------------------------------------------------------------------------------
Function Eiyo01_703Pset()
Dim aa  As Worksheet
Dim bb  As Worksheet
Dim i1  As Long
Dim i2  As Long
Dim i3  As Long
Dim i4  As Long

    Set aa = Sheets("DBmirror")
    Set bb = Sheets("¶³İ¾Øİ¸Ş¼°Ä")
'   ’²¸“úEŠúŠÔ
    bb.Range("p_date1") = Format(aa.Range("b2"), " yyyy""”N"" mm""Œ"" dd""“ú‚©‚ç") & _
                     Format(aa.Range("b2") + aa.Range("c2") - 1, " mm""Œ"" dd""“ú‚Ü‚Å(") & _
                     aa.Range("c2") & "“úŠÔ)"
'   «•Ê
    If aa.Range("e2") = 0 Then
        bb.Range("P_sex") = "’j"
    Else
        bb.Range("P_sex") = "—"
    End If
'
    bb.Range("P_age") = aa.Range("g2")              '”N—î
    bb.Range("P_adrno") = aa.Range("k2")            '—X•Ö”Ô†
    bb.Range("P_adrs1") = aa.Range("l2")            'ZŠ[‚P
    bb.Range("P_adrs2") = "'" & aa.Range("m2")      'ZŠ[‚Q
    bb.Range("P_namej") = aa.Range("d2") & "@—l"   '–¼
    bb.Range("P_fcode") = aa.Range("a2")            'Fcode
    bb.Range("P_hok1") = aa.Range("at2")            '•ÛŒ¯Ø‹L†
    bb.Range("P_hok2") = aa.Range("au2")            '•ÛŒ¯Ø‚m‚n
'   ‘ÌˆÊ
    bb.Range("P_hight") = aa.Range("h2")            'g’·
    bb.Range("P_weght") = aa.Range("i2")            '‘Ìd
    If aa.Range("g2") > 12 Or _
       aa.Range("ad2") = 0 Or _
       aa.Range("ag2") = 0 Then
        bb.Range("P_aans1") = aa.Range("bl2")       '•W€‘Ìd
    Else
        bb.Range("P_aans1") = Empty
    End If
    If aa.Range("j2") = 0 Then                      '”ç‰º‰–b
        bb.Range("P_sibou") = Empty
    Else
        bb.Range("P_sibou") = aa.Range("j2")
    End If
    If aa.Range("ad2") = 0 And aa.Range("ag2") = 0 Then '”DP/½Îß°Â
        bb.Range("P_taii") = aa.Range("bx2")        '‘ÌˆÊw”
        If aa.Range("g2") < 3 Then
            bb.Range("P_tsisu") = "(ƒJƒEƒvw”)"    '‚QËˆÈ‰º
        ElseIf aa.Range("g2") < 13 Then
            bb.Range("P_tsisu") = "(ƒ[ƒŒƒ‹w”)"  '‚R`‚P‚QË
        ElseIf aa.Range("i2") < 150 Then
            bb.Range("P_tsisu") = "(ƒuƒ[ƒJ[w”•Ï–@)"  'g’·150cm–¢–
        Else
            bb.Range("P_tsisu") = "(ƒuƒ[ƒJ[w”•Ï–@)"  'g’·150cmˆÈã
        End If
    Else
        bb.Range("P_taii") = Empty
        bb.Range("P_tsisu") = Empty
    End If
'   H•iÛæƒoƒ‰ƒ“ƒX
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
'   ŒŒ‰tŒŸ¸
    bb.Range("P_bdate") = aa.Cells(2, 50)
    For i1 = 1 To 12
        bb.Range("P_bbl01").Offset(i1 - 1, 0) = aa.Cells(2, i1 + 50)
    Next i1
'   ‰h—{‘fÛæƒoƒ‰ƒ“ƒX  i1:‰h—{‘fIndex 1`27  i2:sIndex 1`24
    For i1 = 1 To 27
        i2 = Fld_Field(i1, 24)
        If i2 > 0 Then
            bb.Range("i44").Offset(i2, 0) = aa.Range("e7").Offset(0, (i1 * 5 - 5))
            bb.Range("l44").Offset(i2, 0) = aa.Range("d7").Offset(0, (i1 * 5 - 5))
            bb.Range("o44").Offset(i2, 0) = bb.Range("l44").Offset(i2, 0) - bb.Range("i44").Offset(i2, 0)
            i3 = aa.Range("f7").Offset(0, (i1 * 5 - 5))
            If i3 > -62.5 Then
'                bb.Range("r44").Offset(i2, 0) = String((i3 + 62.5) * 52 / 125, "*")
                bb.Range("r44").Offset(i2, 0) = String((i3 + 62.5) * 52 / 250, "¡")
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
'   ‰h—{‘f”ä—¦
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
'   01_704@ƒAƒhƒoƒCƒX€–Ú‚Ìİ’è
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
'       ƒEƒGƒCƒgEƒAƒhƒoƒCƒX
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
    
'   ŒŸ¸
    If Sheets("DBmirror").Range("ck2") = 3330 Then
        Wadvic2(4) = Empty
        For i1 = 3 To 10
            If Mid(Range("q2"), i1, 1) = "1" Then
                Select Case i1
                    Case 3: Wadvic2(4) = Wadvic2(4) & "‹¹•”‚wü@ŒÄ‹z‹@”\@"
                    Case 4: Wadvic2(4) = Wadvic2(4) & "S“d}@ŒŒˆ³@–¬”@"
                    Case 5: Wadvic2(4) = Wadvic2(4) & "Á‰»ŠíŒn@"
                    Case 6: Wadvic2(4) = Wadvic2(4) & "ŒŒ‰t@"
                    Case 7: Wadvic2(4) = Wadvic2(4) & "¤@"
                    Case 8: Wadvic2(4) = Wadvic2(4) & "ŠÌ‹@”\@"
                    Case 9: Wadvic2(4) = Wadvic2(4) & "¨•@ˆôA@"
                    Case 10: Wadvic2(4) = Wadvic2(4) & "Šá‰È@"
                End Select
            End If
        Next i1
    End If
    
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("an13") = Left(Wadvic1(1), 18)      'H¶Šˆ^KŠµ
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("an14") = Mid(Wadvic1(1), 19, 18)
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("an15") = Mid(Wadvic1(1), 37, 18)
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("aj16") = Left(Wadvic2(1), 22)
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("aj17") = Mid(Wadvic2(1), 23, 22)
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("an20") = Left(Wadvic1(2), 18)      '‡–°‚Æ‹x—{
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("an21") = Mid(Wadvic1(2), 19, 18)
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("an22") = Mid(Wadvic1(2), 37, 18)
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("aj23") = Left(Wadvic2(2), 22)
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("aj24") = Mid(Wadvic2(2), 23, 22)
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("an27") = Left(Wadvic1(3), 18)      '‰^“®
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("an28") = Mid(Wadvic1(3), 19, 18)
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("an29") = Mid(Wadvic1(3), 37, 18)
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("aj30") = Left(Wadvic2(3), 22)
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("aj31") = Mid(Wadvic2(3), 23, 22)
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("an34") = Left(Wadvic1(4), 18)      'Œ’Nó‘Ô
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("an35") = Mid(Wadvic1(4), 19, 18)
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("an36") = Mid(Wadvic1(4), 37, 18)
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("aj37") = Left(Wadvic1(5), 22)
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("aj38") = Mid(Wadvic1(5), 23, 22)
    
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("bc17") = Wadvic3(1)
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("bc18") = Wadvic3(2)
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("bc19") = Wadvic3(3)
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("bc20") = Wadvic3(4)
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("bc21") = Wadvic3(5)
    
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("bc63") = Left(Wadvic1(6), 18)
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("bc64") = Mid(Wadvic1(6), 19, 18)
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("bc65") = Left(Wadvic2(6), 18)
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("bc66") = Mid(Wadvic2(6), 19, 18)
    
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("u73") = Wadvic1(12)
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("u74") = Wadvic2(12)
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("u75") = Wadvic1(13)
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("u76") = Wadvic2(13)
End Function
'--------------------------------------------------------------------------------
'   01_705@ƒR[ƒh•“ú•t
'--------------------------------------------------------------------------------
Function Eiyo01_705Footer()
Dim Wtext   As String
    Wtext = "(" & Sheets("Šî‘b").Range("g3") & ":" & Format(Date, "yymmdd") & ")"
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("b80") = Wtext
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("bd75") = Sheets("DBmirror").Range("db2")
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("bd76") = Sheets("DBmirror").Range("dc2")
    Sheets("¶³İ¾Øİ¸Ş¼°Ä").Range("bd77") = Sheets("DBmirror").Range("dd2")
End Function
'--------------------------------------------------------------------------------
'   01_810@Šî‘b‰æ–Êì¬
'--------------------------------------------------------------------------------
Function Eiyo01_810Šî‘b‰æ–Êì¬()
Const PgmName = "Eiyo01_Šî‘bÛH“ü—Í.xls"
Const ShtName = "Šî‘b"
Dim i1      As Long
Dim i2      As Long
Dim FldItem As Variant

    If ActiveWorkbook.Name <> PgmName Then
        MsgBox PgmName & " ‚Å‚Í‚ ‚è‚Ü‚¹‚ñ"
        End
    End If
    If ActiveSheet.Name <> ShtName Then
        MsgBox ShtName & " ‚Å‚Í‚ ‚è‚Ü‚¹‚ñ"
        End
    End If
    Call Eiyo01_000init
'   ‰æ–Ê‚Ìì¬
    Call Eiyo930Screen_Hold                 '‰æ–Ê—}~‚Ù‚©
    While (ActiveSheet.Shapes.Count > 0)    'ƒRƒ}ƒ“ƒhƒ{ƒ^ƒ“æÁ
        ActiveSheet.Shapes(1).Cut
    Wend
    Cells.Delete Shift:=xlUp                '‘SÁ‹
    Cells.NumberFormatLocal = "@"           '‘S‰æ–Ê•¶š—ñ‘®«
    Cells.Select
    With Selection.Font                     '•¶šƒtƒHƒ“ƒg
        .Name = "‚l‚r ƒSƒVƒbƒN"
        .Size = 11
    End With
    Selection.ColumnWidth = 1.75            '—ñ•
    Selection.Interior.ColorIndex = 40      '‘S‰æ–Ê”wŒiFi’W“•j
    
'   •\‘è
    Range("G1:AA1").Select
    Selection.MergeCells = True                 '•\‘èƒZƒ‹˜AŒ‹
    Selection.HorizontalAlignment = xlCenter    '•\‘èƒZƒ“ƒ^ƒŠƒ“ƒO
    Selection.Interior.ColorIndex = 37          '•\‘èFiƒy[ƒ‹ƒuƒ‹[j
    With Selection.Font                         'ƒtƒHƒ“ƒg
        .FontStyle = "‘¾š"
        .Size = 16
    End With
    Range("G01") = "‰h—{ŒvZiŠî‘bj‚Q‚V‰h—{‘f”Å"
    Range("A01").VerticalAlignment = xlTop
    Range("A01") = "v-01"
    Range("A03") = "ƒR[ƒh"
    Range("A04") = "’²¸ŠúŠÔ"
    Range("A05") = "–¼"
    Range("A06") = "«•Ê"
    Range("i06") = "(0:’j 1:—)"
    Range("A07") = "¶”NŒ“ú"
    Range("a08") = "g’·"
    Range("j08") = "cm"
    Range("a09") = "‘Ìd"
    Range("j09") = "Kg"
    Range("a10") = "”ç‰º‰–b"
    Range("j10") = "cm"
    Range("A11") = "—X•Ö”Ô†"
    Range("A12") = "ZŠ[‚P"
    Range("A13") = "ZŠ[‚Q"
    Range("A14") = "’nˆæ"
    Range("a15") = "“s“¹•{Œ§"
    Range("a16") = "3.HK"
    Range("a17") = "4.‹x—{"
    Range("a18") = "5.‰^“®"
    Range("a19") = "6.Œ’N"
    Range("m08") = "7.E‹Æ"
    Range("m09") = "A.å•w"
    Range("m10") = "B.”DP"
    Range("m11") = "C.“œ”A"
    Range("m14") = "D.‚ŒŒˆ³"
    Range("m15") = "E.½Îß°Â"
    Range("m16") = "F.‰^“®•”"
    Range("m17") = "*.‹i‰Œ"
    Range("m18") = "G.gáŠQ"
    Range("m19") = "H.³´²ÄCT"
    Range("m20") = "´ÈÙ·Ş°w’è"
    Range("M20").Characters(Start:=7, Length:=2).Font.Size = 9
    Range("m21") = "ÀİÊß¸ w’è"
    Range("M21").Characters(Start:=7, Length:=2).Font.Size = 9
    Range("m22") = "(0:–³ 1:w’è 2:À‘Ìd 3:•W€‘Ìd)"
    Range("m23") = "¶³İ¾×°1"
    Range("m24") = "¶³İ¾×°2"
    Range("m25") = "¶³İ¾×°3"
    
    Range("x03") = "ŒŒ‰tŒ^"
    Range("x04") = "xĞ•”CD"
    Range("x05") = "•ÛŒ’‹L†"
    Range("x06") = "      No"
    Range("x07") = "’èŠúŒ’f"
    Range("x08") = "˜r(L,R)"
    Range("x09") = "ŒŸ¸“ú"
    Range("x10") = "ÔŒŒ‹…”"
    Range("x11") = "ŒŒF‘f—Ê"
    Range("x12") = "ÍÏÄ¸Ø¯Ä"
    Range("x13") = "ºÚ½ÃÛ°Ù"
    Range("x14") = "HDL"
    Range("x15") = "’†«‰–b"
    Range("x16") = "G.O.T."
    Range("x17") = "G.P.T."
    Range("x18") = "”A_"
    Range("x19") = "ŒŒ“œ"
    Range("x20") = "ŒŒˆ³Å‚"
    Range("x21") = "    Å’á"
    
    Cells.Locked = True                             '‘SƒZƒ‹‚ğƒƒbƒN
    For i1 = 0 To UBound(Fld_Adrs1)
        FldItem = Split(Fld_Adrs1(i1), ",")
        Range(Trim(FldItem(1))).Select
        Selection.MergeCells = True                 'ƒZƒ‹Œ‹‡
        Range(Left(FldItem(1), 4)).Name = Trim(FldItem(0))
        If FldItem(2) = "i" Then
            With Selection.Borders                      '“ü—Í€–Ú‚Ì˜gŒrü
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .Weight = xlThin
            End With
            Selection.Interior.ColorIndex = xlNone      '“ü—Í€–Ú‚Ì”’”²‚«‰»
            Selection.Locked = False                    '“ü—Í€–Ú‚Ì•ÛŒì‰ğœ
        End If
        Select Case FldItem(4)
            Case "90": Selection.NumberFormatLocal = "G/•W€"
            Case "91": Selection.NumberFormatLocal = "#0.0"
            Case "92": Selection.NumberFormatLocal = "#0.00"
            Case "Ds"
                Selection.NumberFormatLocal = "yyyy/mm/dd"
                Selection.HorizontalAlignment = xlLeft
            Case "Dw"
                Selection.NumberFormatLocal = "gy.m.d"
                Selection.HorizontalAlignment = xlLeft
            Case "J "
                With Selection.Validation           'Š¿š€–Ú
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
    Range("Gyyyy").NumberFormatLocal = "gy"     '¶”NŒ“ú‚Ì˜a—ï”N•\¦
    Range("Gyyyy") = "=RC[-5]"
    Range("Age").NumberFormatLocal = "G/•W€"
    Range("Age") = "=DATEDIF(RC[-7],R[-3]C[-7],""y"")"
    Range("p07") = "Ë"

    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=10, Top:=350, Width:=60, Height:=30)
        .Object.Caption = "‰æ–ÊÁ‹"
        .Name = "ƒNƒŠƒA"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=100, Top:=350, Width:=60, Height:=30)
        .Object.Caption = "ÃŞ°ÀŒÄo"
        .Name = "ŒŸõ"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=170, Top:=350, Width:=60, Height:=30)
        .Object.Caption = "ÃŞ°À“o˜^"
        .Name = "XV"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=240, Top:=350, Width:=60, Height:=30)
        .Object.Caption = "¶³İ¾Øİ¸Ş" & vbLf & "¼°Äì•\"
        .Name = "ì•\"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=330, Top:=450, Width:=60, Height:=30)
        .Object.Caption = "ÃŞ°ÀæÁ"
        .Name = "æÁ"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=400, Top:=350, Width:=60, Height:=30)
        .Object.Caption = "I—¹"
        .Name = "I—¹"
    End With
    
    Range("Gmesg").Font.Bold = True                            'ƒƒbƒZ[ƒWƒGƒŠƒA
    Range("Gmesg").Font.ColorIndex = 3
    Cells.FormatConditions.Delete               'ƒV[ƒg‘S‘Ì‚©‚çğŒ•t‚«‘®‚ğíœ‚·‚é
    Cells.FormatConditions.Add Type:=xlExpression, Formula1:="=AND(CELL(""row"")=ROW(),CELL(""col"")=COLUMN())"
    Cells.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Cells.FormatConditions(1).Interior.Color = 255
    
    Call Eiyo01_820‘€ìƒKƒCƒh
    Range("g4").Select
    Call Eiyo940Screen_Start    '‰æ–Ê•`‰æ‚Ù‚©
End Function
'--------------------------------------------------------------------------------
'   01_820 ‘€ìƒKƒCƒh
'--------------------------------------------------------------------------------
Function Eiyo01_820‘€ìƒKƒCƒh()
    Call Eiyo930Screen_Hold     '‰æ–Ê—}~‚Ù‚©
    Columns("ah:hz").Delete Shift:=xlToLeft
    Range("ah01") = "1.l‚ğŒŸõ‚·‚é"
    Range("ah02") = "@‰æ–Ê‚Ì‚¢‚Ã‚ê‚©‚Ì€–Ú‚É"
    Range("ah03") = "@“ü—ÍŒãAuŒŸõv‚ğ‰Ÿ‰º‚µ‚Ä‚­‚¾‚³‚¢B"
    Range("ah04") = "@“¯–¼‚È‚Ç•¡”ŠY“–Ò‚Ìê‡‚ÍA‰E‘¤‚Ìˆê——‚©‚ç"
    Range("ah05") = "@ƒR[ƒh‚ğƒ_ƒuƒ‹ƒNƒŠƒbƒN‚µ‚Ä‘I‘ğ‚µ‚Ü‚·B"
    Range("ah07") = "@uŒŸõv‚ÍŒ´‘¥w‘O•ûˆê’vx‚Å‚·A"
    Range("ah08") = "@æ“ª‚É[%]‚ğ•t‚¯‚é‚ÆwŠÜ‚Şx‚É‚È‚è‚Ü‚·B"
    Range("ah10") = "2.l‚ğ“o˜^‚·‚é"
    Range("ah11") = "@‰æ–Ê‚ÌŠe€–Ú‚ğ“ü—Í‚µ"
    Range("ah12") = "@uXVv‚ğ‰Ÿ‰º‚µ‚Ä‚­‚¾‚³‚¢B"
    Range("ah14") = "3.l‚Ì•ÏXEæÁ"
    Range("ah15") = "@l‚ğŒŸõ‚µA"
    Range("ah16") = "@C³Œã‚ÉuXVv‚Ü‚½‚ÍuæÁv‚ğ‰Ÿ‰º‚µ‚Ä‚­‚¾‚³‚¢B"
    Range("ah18") = "4.ÛH‚Ì“o˜^"
    Range("ah19") = "@l‚Ì“o˜^‚Ü‚½‚ÍÆ‰ïŒãuÛHvƒV[ƒg‚ÉØ‚è‘Ö‚¦‚Ä‚­‚¾‚³‚¢"
    Call Eiyo940Screen_Start    '‰æ–Ê•`‰æ‚Ù‚©
End Function
'--------------------------------------------------------------------------------
'   01_830@ÛH‰æ–Êì¬
'--------------------------------------------------------------------------------
Function Eiyo01_830ÛH‰æ–Êì¬()
Const PgmName = "Eiyo01_Šî‘bÛH“ü—Í.xls"
Const ShtName = "ÛH"
Dim i1      As Long
Dim i2      As Long
Dim FldItem As Variant

    If ActiveWorkbook.Name <> PgmName Then
        MsgBox PgmName & " ‚Å‚Í‚ ‚è‚Ü‚¹‚ñ"
        End
    End If
    If ActiveSheet.Name <> ShtName Then
        MsgBox ShtName & " ‚Å‚Í‚ ‚è‚Ü‚¹‚ñ"
        End
    End If
    Call Eiyo01_000init
'   ‰æ–Ê‚Ìì¬
    Call Eiyo930Screen_Hold     '‰æ–Ê—}~‚Ù‚©
    ActiveWindow.FreezePanes = False        'ƒEƒCƒ“ƒh˜gŒÅ’è‚Ì‰ğœ
    
    While (ActiveSheet.Shapes.Count > 0)    'ƒRƒ}ƒ“ƒhƒ{ƒ^ƒ“æÁ
        ActiveSheet.Shapes(1).Cut
    Wend
    Cells.Delete Shift:=xlUp                '‘SÁ‹
    Cells.Select
    With Selection.Font                     '•¶šƒtƒHƒ“ƒg
        .Name = "‚l‚r ƒSƒVƒbƒN"
        .Size = 11
    End With
    Cells.Interior.ColorIndex = 34          '‘S‰æ–Ê”wŒiFi’W—Îj
    Columns("A:B").Interior.ColorIndex = xlNone
    Columns("d").Interior.ColorIndex = xlNone
    Columns("f").Interior.ColorIndex = xlNone
    Rows("1:4").Interior.ColorIndex = 34       '‘S‰æ–Ê”wŒiFi’W—Îj
    Cells.Select                               'Œrü
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

'   •\‘è
    With Range("d1").Font                         'ƒtƒHƒ“ƒg
        .FontStyle = "‘¾š"
        .Size = 16
    End With
    Range("D01") = "‰h—{ŒvZiÛHj"
    Range("A01").VerticalAlignment = xlTop
    Range("A01") = "v-01"
'    Range("a2") = "=Fcode & ":" & Namej
    Columns("A:A").NumberFormatLocal = "yyyy/mm/dd"
    Columns("F:F").NumberFormatLocal = "0.0 "
    Range("A4") = "ÛH“ú"
    Range("B4") = "H‹æ•ª"
    Range("D4") = "H•iCD"
    Range("E4") = "•i–¼EŞ—¿@»ÌßØÒİÄ"
    Range("F4") = "Ûæ—Ê"
    Range("k3") = "EH•iCD—“‚ğƒ_ƒuƒ‹ƒNƒŠƒbƒN‚·‚é‚Æƒƒjƒ…[‚É•Ï‚í‚è‚Ü‚·"
    Range("k2") = "EÛæ—Ê‚ªƒ[ƒ‚Ìs‚Ííœ‚³‚ê‚Ü‚·"
    Range("k3") = "E’Ç‰Á‚ÍÅIs‚ÌŒã‚ë‚É“ü—Í‚µ‚Ä‚­‚¾‚³‚¢"
    Range("k1:k2").Font.Size = 9
    Call Eiyo01_840H•iƒ}ƒXƒ^(4, 6)
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
    
    Cells.Locked = True                             '‘SƒZƒ‹‚ğƒƒbƒN
    Range("A:B,D:D,F:F").Locked = False             '“ü—Í—ñ‚Ì‰ğœ
    Rows("1:3").Locked = True                       '•\‘ès‚ÌƒƒbƒN
    
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=300, Top:=5, Width:=50, Height:=25)
        .Object.Caption = "“o˜^"
        .Name = "“o˜^"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=650, Top:=5, Width:=50, Height:=25)
        .Object.Caption = "ŒŸØ"
        .Name = "ŒŸØ"
    End With
    
    Range("a3").Font.Bold = True                    'ƒƒbƒZ[ƒWƒGƒŠƒA
    Range("a3").Font.ColorIndex = 3
    Range("E5").Select
    ActiveWindow.FreezePanes = True                 'ƒEƒCƒ“ƒh˜gŒÅ’è‚Ìİ’è
    
'    ActiveSheet.Protect UserInterfaceOnly:=True     '•ÛŒì‚ğ—LŒø‚É‚·‚é
    Range("g4").Select
    Call Eiyo940Screen_Start                        '‰æ–Ê•`‰æ‚Ù‚©
End Function
'--------------------------------------------------------------------------------
'   01_840@H•iƒ}ƒXƒ^€‘è
'--------------------------------------------------------------------------------
Function Eiyo01_840H•iƒ}ƒXƒ^(il As Long, ic As Long)
Dim Wtext   As String
Dim Warray  As Variant
Dim i1      As Long

    Wtext = Empty
    Wtext = Wtext & vbLf & "ƒR[ƒh"
    Wtext = Wtext & vbLf & "H•i–¼"
    Wtext = Wtext & vbLf & "“Ç‚İi•ª—Şj"
    Wtext = Wtext & vbLf & "’PˆÊ"           '“ü—Í’PˆÊ
    Wtext = Wtext & vbLf & "ƒRƒƒ“ƒg"
    Wtext = Wtext & vbLf & "“o˜^’PˆÊ"
    Wtext = Wtext & vbLf & "Š·ZŒW”"
    Wtext = Wtext & vbLf & "ÒÆ­°ˆÊ’u‚P"
    Wtext = Wtext & vbLf & "ÒÆ­°ˆÊ’u‚Q"
    Wtext = Wtext & vbLf & "Hğ"           '0:H 1:ğ 2:»ÌßØÒİÄ
    Wtext = Wtext & vbLf & "ÛH”ÍˆÍ‰ºŒÀ"
    Wtext = Wtext & vbLf & "ÛH”ÍˆÍãŒÀ"
    Wtext = Wtext & vbLf & "‰h—{‘f-01"
    Wtext = Wtext & vbLf & "‰h—{‘f-02"
    Wtext = Wtext & vbLf & "‰h—{‘f-03"
    Wtext = Wtext & vbLf & "‰h—{‘f-04"
    Wtext = Wtext & vbLf & "‰h—{‘f-05"
    Wtext = Wtext & vbLf & "‰h—{‘f-06"
    Wtext = Wtext & vbLf & "‰h—{‘f-07"
    Wtext = Wtext & vbLf & "‰h—{‘f-08"
    Wtext = Wtext & vbLf & "‰h—{‘f-09"
    Wtext = Wtext & vbLf & "‰h—{‘f-10"
    Wtext = Wtext & vbLf & "‰h—{‘f-11"
    Wtext = Wtext & vbLf & "‰h—{‘f-12"
    Wtext = Wtext & vbLf & "‰h—{‘f-13"
    Wtext = Wtext & vbLf & "‰h—{‘f-14"
    Wtext = Wtext & vbLf & "‰h—{‘f-15"
    Wtext = Wtext & vbLf & "‰h—{‘f-16"
    Wtext = Wtext & vbLf & "‰h—{‘f-17"
    Wtext = Wtext & vbLf & "‰h—{‘f-18"
    Wtext = Wtext & vbLf & "‰h—{‘f-19"
    Wtext = Wtext & vbLf & "‰h—{‘f-20"
    Wtext = Wtext & vbLf & "‰h—{‘f-21"
    Wtext = Wtext & vbLf & "‰h—{‘f-22"
    Wtext = Wtext & vbLf & "‰h—{‘f-23"
    Wtext = Wtext & vbLf & "‰h—{‘f-24"
    Wtext = Wtext & vbLf & "‰h—{‘f-25"
    Wtext = Wtext & vbLf & "‰h—{‘f-26"
    Wtext = Wtext & vbLf & "‰h—{‘f-27"
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
    Wtext = Wtext & vbLf & "‰¿E“®•¨"
    Wtext = Wtext & vbLf & "‰¿E‹›‰î"
    Wtext = Wtext & vbLf & "‰¿EA•¨"
    Warray = Split(Wtext, vbLf)
    For i1 = 1 To UBound(Warray)
        Cells(il, ic + i1) = Warray(i1)
    Next i1
End Function
'--------------------------------------------------------------------------------
'   01_900  –ˆ’©ˆê”Ô‚É‚c‚a‚ÌƒRƒs[‚ğ‚Æ‚éB
'           ƒ^ƒCƒ€ƒXƒ^ƒ“ƒv‚Í‘O“ú‚ÌÅIXV‚Æ‚È‚é
'--------------------------------------------------------------------------------
Function Eiyo01_900WorkbookOpen()
Dim F_name          As String   'ŒŸõ‚µ‚½ƒtƒ@ƒCƒ‹–¼
Dim F_dbname_today  As String   'DB+–{“ú
Dim F_dbname_min    As String   'DB+00000000
Dim F_dbname_max    As String   'DB+2TŠÔ‘O
Dim W_path          As String

    W_path = ThisWorkbook.Path & "BackUp"""
    F_dbname_today = W_path & "Eiyo_" & Format(Date, "yyyymmdd") & ".mdb"""
    F_dbname_min = "Eiyo_00000000.mdb"
    F_dbname_max = "Eiyo_" & Format(Date - 14, "yyyymmdd") & ".mdb"
    
    SetCurrentDirectory (W_path)            'Dir•ÏX
    If Dir(F_dbname_today) = "" Then        '¡“ú‚Ì•Û‘¶ƒtƒ@ƒCƒ‹‚ª‘¶İ‚µ‚È‚¢
        F_name = Dir("*", vbNormal)
        Do While F_name <> ""
            If (F_name > F_dbname_min And F_name < F_dbname_max) Then
               Kill F_name
            End If
            F_name = Dir                    ' Ÿ‚ÌƒtƒHƒ‹ƒ_–¼‚ğ•Ô‚µ‚Ü‚·B
        Loop
        FileCopy ThisWorkbook.Path & myFileName, F_dbname_today
    End If
End Function
'--------------------------------------------------------------------------------
'   03_030  ƒNƒŠƒA‚Ìƒ{ƒ^ƒ“EƒNƒŠƒbƒN
'--------------------------------------------------------------------------------
Function Eiyo03_030ƒNƒŠƒAClick()
    Call Eiyo930Screen_Hold     '‰æ–Ê—}~‚Ù‚©
    Range("b3:b11") = Empty
    Range("b12") = Empty
    Range("b13:b14") = Empty
    Range("g4:g30").ClearContents
    Range("j4:l18").ClearContents
    Range("j22:j24").ClearContents
    Range("a17") = Empty
    Columns("n:hz").Delete Shift:=xlToLeft
    
    Range("b3").Select
    Call Eiyo940Screen_Start    '‰æ–Ê•`‰æ‚Ù‚©
End Function
'--------------------------------------------------------------------------------
'   03_100  ŒŸõ_Click
'--------------------------------------------------------------------------------
Function Eiyo03_100ŒŸõClick()
Dim Wsql    As String
Dim i1      As Long

    Range("a17") = Empty
    For i1 = 3 To 14
        If Not IsEmpty(Cells(i1, 2)) Then: Exit For
    Next i1
    If i1 > 14 Then
        Range("a17") = "ƒL[‚ª‚ ‚è‚Ü‚¹‚ñ"
        Exit Function
    End If
    Call Eiyo930Screen_Hold     '‰æ–Ê—}~‚Ù‚©
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
        Range("a17") = "ŠY“–ƒf[ƒ^‚Í‚ ‚è‚Ü‚¹‚ñ"
    Else
        With Rst_Food
            Range("n2").CopyFromRecordset Rst_Food  'ƒŒƒR[ƒh
            If IsEmpty(Range("n3")) Then            'ŠY“–‚ª‚PŒ‚Ì‚Æ‚«
                For i1 = 1 To 12                    '‰æ–Ê€–Ú‚Ì‡Ÿˆ—
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
                For i1 = 1 To .Fields.Count                     'ƒtƒB[ƒ‹ƒh–¼
                    Cells(1, i1 + 13).Value = .Fields(i1 - 1).Name
                Next
                Columns("n:hz").EntireColumn.AutoFit           '•
                i1 = Range("n1").End(xlDown).Row
                Range("N:N").Locked = False                     '“ü—Í—ñ‚Ì‰ğœ
            End If
            .Close
        End With
    End If
    Set Rst_Food = Nothing              'ƒIƒuƒWƒFƒNƒg‚Ì‰ğ•ú
    Call Eiyo920DB_Close                'DB Close
    Columns("J:L").EntireColumn.AutoFit
    Call Eiyo940Screen_Start            '‰æ–Ê•`‰æ‚Ù‚©
End Function
'--------------------------------------------------------------------------------
'   03_110  ‘OŒŸõ_Click
'--------------------------------------------------------------------------------
Function Eiyo03_110‘OŒŸõClick()
Dim Wsql    As String
Dim Wkey    As Long

    Range("a17") = Empty
    Wkey = Range("b03")
    Call Eiyo930Screen_Hold             '‰æ–Ê—}~‚Ù‚©
    Call Eiyo91DB_Open                  'DB Open
    Wsql = "SELECT Foodc FROM " & Tbl_Food & " Where Foodc < " & Wkey & " Order by Foodc DESC"
    Set Rst_Food = myCon.Execute(Wsql)
    If Rst_Food.EOF Then
        Range("a17") = "ŠY“–ƒf[ƒ^‚Í‚ ‚è‚Ü‚¹‚ñ"
    Else
        With Rst_Food
            Range("b03") = .Fields(0).Value
            .Close
        End With
    End If
    Set Rst_Food = Nothing      'ƒIƒuƒWƒFƒNƒg‚Ì‰ğ•ú
    Call Eiyo920DB_Close        'DB Close
    Call Eiyo03_100ŒŸõClick
End Function
'--------------------------------------------------------------------------------
'   03_120  ŸŒŸõ_Click
'--------------------------------------------------------------------------------
Function Eiyo03_120ŸŒŸõClick()
Dim Wsql    As String
Dim Wkey    As Long

    Range("a17") = Empty
    Wkey = Range("b03")
    Call Eiyo930Screen_Hold             '‰æ–Ê—}~‚Ù‚©
    Call Eiyo91DB_Open                  'DB Open
    Wsql = "SELECT Foodc FROM " & Tbl_Food & " Where Foodc > " & Wkey & " Order by Foodc"
    Set Rst_Food = myCon.Execute(Wsql)
    If Rst_Food.EOF Then
        Range("a17") = "ŠY“–ƒf[ƒ^‚Í‚ ‚è‚Ü‚¹‚ñ"
    Else
        With Rst_Food
            Range("b03") = .Fields(0).Value
            .Close
        End With
    End If
    Set Rst_Food = Nothing      'ƒIƒuƒWƒFƒNƒg‚Ì‰ğ•ú
    Call Eiyo920DB_Close        'DB Close
    Call Eiyo03_100ŒŸõClick
End Function
'--------------------------------------------------------------------------------
'   03_200  XV
'--------------------------------------------------------------------------------
Function Eiyo03_200XVClick()
Dim i1      As Long
Dim Wsw     As Long
Dim Wtemp   As Variant

    Wsw = 0
    Range("a17") = Empty
    If Range("b3") < 1 Then: Exit Function
    Call Eiyo91DB_Open                      'DB Open
    '€”õ‚±‚±‚Ü‚Å
    With Rst_Food
        'ƒCƒ“ƒfƒbƒNƒX‚Ìİ’è
        .Index = "PrimaryKey"
        'ƒŒƒR[ƒhƒZƒbƒg‚ğŠJ‚­
        Rst_Food.Open Source:=Tbl_Food, ActiveConnection:=myCon, _
            CursorType:=adOpenKeyset, LockType:=adLockOptimistic, _
            Options:=adCmdTableDirect
        '”Ô†‚ª“o˜^‚³‚ê‚Ä‚¢‚é‚©ŒŸõ‚·‚é
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
            Range("a17") = "’Ç‰Á‚³‚ê‚Ü‚µ‚½"
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
                Range("a17") = "•ÏX€–Ú‚ª‚ ‚è‚Ü‚¹‚ñ"
            Else
                If MsgBox("•ÏX€–Ú‚Í" & Wsw & "ƒ•Š‚Å‚·AXV‚µ‚Ä‚æ‚ë‚µ‚¢‚Å‚·‚©", vbOKCancel) = vbOK Then
                    .Update
                    Range("a17") = "XV‚³‚ê‚Ü‚µ‚½"
                End If
            End If
        End If
'        .Close
    End With
    Set Rst_Food = Nothing      'ƒIƒuƒWƒFƒNƒg‚Ì‰ğ•ú
    Call Eiyo920DB_Close        'DB Close
End Function
'--------------------------------------------------------------------------------
'   03_300  æÁ
'--------------------------------------------------------------------------------
Function Eiyo03_300æÁClick()
    Range("a17") = Empty
    Call Eiyo91DB_Open      'DB Open
    '€”õ‚±‚±‚Ü‚Å
    With Rst_Food
        'ƒCƒ“ƒfƒbƒNƒX‚Ìİ’è
        .Index = "PrimaryKey"
        'ƒŒƒR[ƒhƒZƒbƒg‚ğŠJ‚­
        Rst_Food.Open Source:=Tbl_Food, ActiveConnection:=myCon, _
            CursorType:=adOpenKeyset, LockType:=adLockOptimistic, _
            Options:=adCmdTableDirect
        '”Ô†‚ª“o˜^‚³‚ê‚Ä‚¢‚é‚©ŒŸõ‚·‚é
        If Not .EOF Then .Seek Range("b3")
        If .EOF Then
            Range("a17") = "ƒL[‚ª‘¶İ‚µ‚Ü‚¹‚ñ"
        Else
            If MsgBox("íœ‚µ‚Ä‚æ‚ë‚µ‚¢‚Å‚·‚©", vbOKCancel) = vbOK Then
                .Delete
                Range("a17") = "æÁ‚³‚ê‚Ü‚µ‚½"
            End If
        End If
        .Close
    End With
    Set Rst_Food = Nothing      'ƒIƒuƒWƒFƒNƒg‚Ì‰ğ•ú
    Call Eiyo920DB_Close        'DB Close
End Function
'--------------------------------------------------------------------------------
'   03_400  H•iƒR[ƒhƒf[ƒ^‚Ìì¬
'--------------------------------------------------------------------------------
Function Eiyo03_400ƒR[ƒhClick()
Dim i1      As Long
Dim Wsql    As String

    Call Eiyo930Screen_Hold
    Call Eiyo91DB_Open      'DB Open
    
    Wsql = "SELECT Foodc,Fname FROM " & Tbl_Food & " Order by Foodc"
    Set Rst_Food = myCon.Execute(Wsql)
    If Rst_Food.EOF Then
        Range("a17") = "ŠY“–ƒf[ƒ^‚Í‚ ‚è‚Ü‚¹‚ñ"
    Else
        With Rst_Food
            Range("AA1").CopyFromRecordset Rst_Food
            .Close
        End With
    End If
    Set Rst_Food = Nothing      'ƒIƒuƒWƒFƒNƒg‚Ì‰ğ•ú
    Call Eiyo920DB_Close        'DB Close
    
    Open ThisWorkbook.Path & "H•iƒR[ƒh.txt" For Output As #22
    For i1 = 1 To Range("aa60000").End(xlUp).Row
        Print #2, Cells(i1, 27) & vbTab & Cells(i1, 28)
    Next i1
    Close
    Columns("Aa:BZ").Delete Shift:=xlToLeft
    Call Eiyo940Screen_Start
End Function
'--------------------------------------------------------------------------------
'   03_810  ƒV[ƒg‚Ìì¬
'--------------------------------------------------------------------------------
Function Eiyo03_810H•isheet_make()
    Call Eiyo930Screen_Hold     '‰æ–Ê—}~‚Ù‚©
    Call Eiyo03_811_init        'ƒV[ƒg‚Ì‰Šú‰»
    Call Eiyo03_812_zokusei     'ƒL[A–¼ÌA‘®«
    Call Eiyo03_813_eiyoso      '‰h—{‘f
    Call Eiyo03_814_shokugun    'H•iŒQ
    Call Eiyo03_815_sisitu      '‰¿
    Call Eiyo03_816_keisen      'ŒrüA—ñ•
    Call Eiyo03_817_button      'ƒRƒ}ƒ“ƒhEƒ{ƒ^ƒ“
    Call Eiyo940Screen_Start    '‰æ–Ê•`‰æ‚Ù‚©
End Function
'--------------------------------------------------------------------------------
'   03_811  ƒV[ƒg‚Ì‰Šú‰»
'--------------------------------------------------------------------------------
Function Eiyo03_811_init()
    Sheets("H•iƒ}ƒXƒ^").Select
    While (ActiveSheet.Shapes.Count > 0)    'ƒRƒ}ƒ“ƒhƒ{ƒ^ƒ“æÁ
        ActiveSheet.Shapes(1).Cut
    Wend
    Cells.Delete Shift:=xlUp                '‘SÁ‹
    Cells.NumberFormatLocal = "@"           '‘S‰æ–Ê•¶š—ñ‘®«
    Range("e1") = "¦@H•iƒ}ƒXƒ^@Æ‰ïEXV@¦"
    Range("e1").Font.Size = 16
End Function
'--------------------------------------------------------------------------------
'   03_812  ƒL[A–¼ÌA‘®«
'--------------------------------------------------------------------------------
Function Eiyo03_812_zokusei()
Dim i1      As Long
Dim Wtext   As String
Dim Warray  As Variant
'   ‘®«‚Ù‚©
    Wtext = Empty
    Wtext = Wtext & vbLf & "H•iƒR[ƒh"
    Wtext = Wtext & vbLf & "H•i–¼"
    Wtext = Wtext & vbLf & "“Ç‚İ(•ª—Şj"
    Wtext = Wtext & vbLf & "“ü—Í’PˆÊ"
    Wtext = Wtext & vbLf & "“E—v"
    Wtext = Wtext & vbLf & "“o˜^’PˆÊ"
    Wtext = Wtext & vbLf & "’PˆÊŠ·Z’l"
    Wtext = Wtext & vbLf & "ÒÆ­°ˆÊ’u‚P"
    Wtext = Wtext & vbLf & "ÒÆ­°ˆÊ’u‚Q"
    Wtext = Wtext & vbLf & "H•i•ª—Ş"
    Wtext = Wtext & vbLf & "ÛH—Ê‰ºŒÀ"
    Wtext = Wtext & vbLf & "ÛH—ÊãŒÀ"
    Warray = Split(Wtext, vbLf)
    If UBound(Warray) <> 12 Then
        MsgBox "Program Error No.01 " & UBound(Warray)
        End
    End If
    For i1 = 1 To UBound(Warray)
        Cells(i1 + 2, 1) = Warray(i1)
    Next i1
    Range("c12") = "0:H•i 1:ğ—Ş 2:»ÌßØÒİÄ"
'   ƒZƒ‹Œ‹‡
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
        With Selection.Interior                         '”wŒiF
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        Selection.Locked = False                        '•ÛŒì‰ğœ
        Selection.FormulaHidden = False
    Next i1
    Range("a17").Locked = False                        '•ÛŒì‰ğœ
End Function
'--------------------------------------------------------------------------------
'   03_813  ‰h—{‘f
'--------------------------------------------------------------------------------
Function Eiyo03_813_eiyoso()
Dim i1      As Long
Dim Wtext   As String
Dim Warray  As Variant
    Wtext = Empty
    Wtext = Wtext & vbLf & "01:ƒGƒlƒ‹ƒM["
    Wtext = Wtext & vbLf & "02:‚½‚ñ‚Ï‚­¿"
    Wtext = Wtext & vbLf & "03:“®•¨«‚½‚ñ‚Ï‚­¿"
    Wtext = Wtext & vbLf & "04:A•¨«‚½‚ñ‚Ï‚­¿"
    Wtext = Wtext & vbLf & "05:‰¿"
    Wtext = Wtext & vbLf & "06:“œ¿"
    Wtext = Wtext & vbLf & "07:H•¨‚¹‚ñ‚¢"
    Wtext = Wtext & vbLf & "08:ƒJƒ‹ƒVƒEƒ€"
    Wtext = Wtext & vbLf & "09:ƒŠƒ“"
    Wtext = Wtext & vbLf & "10:“S"
    Wtext = Wtext & vbLf & "11:ƒiƒgƒŠƒEƒ€"
    Wtext = Wtext & vbLf & "12:ƒrƒ^ƒ~ƒ“‚`"
    Wtext = Wtext & vbLf & "13:ƒrƒ^ƒ~ƒ“‚a‚P"
    Wtext = Wtext & vbLf & "14:ƒrƒ^ƒ~ƒ“‚a‚Q"
    Wtext = Wtext & vbLf & "15:ƒiƒCƒAƒVƒ“"
    Wtext = Wtext & vbLf & "16:ƒrƒ^ƒ~ƒ“‚b"
    Wtext = Wtext & vbLf & "17:ƒrƒ^ƒ~ƒ“‚a‚U"
    Wtext = Wtext & vbLf & "18:ƒpƒ“ƒgƒeƒ“_"
    Wtext = Wtext & vbLf & "19:—t_"
    Wtext = Wtext & vbLf & "20:ƒrƒ^ƒ~ƒ“‚d"
    Wtext = Wtext & vbLf & "21:ƒJƒŠƒEƒ€"
    Wtext = Wtext & vbLf & "22:ƒ}ƒOƒlƒVƒEƒ€"
    Wtext = Wtext & vbLf & "23:H‰–"
    Wtext = Wtext & vbLf & "24:ƒRƒŒƒXƒeƒ[ƒ‹"
    Wtext = Wtext & vbLf & "25:•s–O˜a‰–b_"
    Wtext = Wtext & vbLf & "26:–O˜a‰–b_"
    Wtext = Wtext & vbLf & "27:»“œ"
    Warray = Split(Wtext, vbLf)
    If UBound(Warray) <> 27 Then
        MsgBox "Program Error No.02 " & UBound(Warray)
        End
    End If
    For i1 = 1 To UBound(Warray)
        Cells(i1 + 3, 6) = Warray(i1)
    Next i1
    Cells(3, 6) = "‰h—{‘f"
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
    Wtext = Wtext & vbLf & "ƒÊg"
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
'   03_814  H•iŒQ
'--------------------------------------------------------------------------------
Function Eiyo03_814_shokugun()
Dim i1      As Long
Dim Wtext   As String
Dim Warray  As Variant
    Wtext = Empty
    Wtext = Wtext & vbLf & "H•iŒQ"
    Wtext = Wtext & vbLf & "01:‘å“¤»•i"
    Wtext = Wtext & vbLf & "02:‹›‰î—Ş"
    Wtext = Wtext & vbLf & "03:“÷@—Ş"
    Wtext = Wtext & vbLf & "04:—‘"
    Wtext = Wtext & vbLf & "05:ŠC@‘"
    Wtext = Wtext & vbLf & "06:“û»•i"
    Wtext = Wtext & vbLf & "07:¬@‹›"
    Wtext = Wtext & vbLf & "08:—Î‰©F–ìØ"
    Wtext = Wtext & vbLf & "09:’WF–ìØ"
    Wtext = Wtext & vbLf & "10:‰Ê@•¨"
    Wtext = Wtext & vbLf & "11:’@—Ş"
    Wtext = Wtext & vbLf & "12:‚¢‚à—Ş"
    Wtext = Wtext & vbLf & "13:»@“œ"
    Wtext = Wtext & vbLf & "14:A•¨«–û‰"
    Wtext = Wtext & vbLf & "15:“®•¨«–û‰"
    Warray = Split(Wtext, vbLf)
    If UBound(Warray) <> 16 Then
        MsgBox "Program Error No.04 " & UBound(Warray)
        End
    End If
    For i1 = 1 To UBound(Warray)
        Cells(i1 + 2, 9) = Warray(i1)
    Next i1
    Cells(3, 10) = "¶ÛØ°"
    Cells(3, 11) = "d—Ê"
    Cells(3, 12) = "¶Ù¼³Ñ"
    Cells(19, 9) = "Total"
    Cells(20, 10) = "(kcal)"
    Cells(20, 11) = "(g)"
    Cells(20, 12) = "(g)"
    Range("j3:l3,i19,j20:l20").HorizontalAlignment = xlRight
End Function
'--------------------------------------------------------------------------------
'   03_815  ‰¿
'--------------------------------------------------------------------------------
Function Eiyo03_815_sisitu()
Dim i1      As Long
Dim Wtext   As String
Dim Warray  As Variant
    Wtext = Empty
    Wtext = Wtext & vbLf & "‰¿(“®•¨)"
    Wtext = Wtext & vbLf & "‰¿(‹›‰î)"
    Wtext = Wtext & vbLf & "‰¿(A•¨)"
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
'   03_816  ŒrüA—ñ•
'--------------------------------------------------------------------------------
Function Eiyo03_816_keisen()
Dim i1      As Long
Dim Wtext   As String
Dim Warray  As Variant
'   Œrü
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
    With Selection.Interior                         '”wŒiF
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Locked = False                        '•ÛŒì‰ğœ
    Selection.FormulaHidden = False
'   —ñ•
    Columns("F:F").ShrinkToFit = True
    Cells.EntireColumn.AutoFit
    Warray = Array(, 0, 1.75, 5, 18, 2, 15, 0, 7)
    For i1 = 1 To UBound(Warray)
        If Warray(i1) > 0 Then: Columns(i1).ColumnWidth = Warray(i1)
    Next i1
'    Range("B4:C4,B6:C6,B11:C11,B12:C12").NumberFormatLocal = "#,##0.00;[Ô]-#,##0.00"
    Range("B9,B13,B14").NumberFormatLocal = "#,##0.00;[Ô]-#,##0.00"
    Range("g4:g30,j4:l19,j22:j25").NumberFormatLocal = "#,##0.00;[Ô]-#,##0.00"
'   ƒJ[ƒ\ƒ‹ˆÊ’u‚ğ–¾Šm‰»‚·‚é
    Cells.FormatConditions.Delete               'ƒV[ƒg‘S‘Ì‚©‚çğŒ•t‚«‘®‚ğíœ‚·‚é
    Cells.FormatConditions.Add Type:=xlExpression, Formula1:="=AND(CELL(""row"")=ROW(),CELL(""col"")=COLUMN())"
    Cells.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Cells.FormatConditions(1).Interior.Color = 255
'   ‡Œv‚Ì®‚ÆğŒ•t‚«‘®
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
    
    Range("a17").Font.Bold = True           'ƒƒbƒZ[ƒWƒGƒŠƒAiÔ‘¾•¶šj
    Range("a17").Font.ColorIndex = 3
End Function
'--------------------------------------------------------------------------------
'   03_817 ƒRƒ}ƒ“ƒhEƒ{ƒ^ƒ“
'--------------------------------------------------------------------------------
Function Eiyo03_817_button()
    While (ActiveSheet.Shapes.Count > 0)    'ƒRƒ}ƒ“ƒhƒ{ƒ^ƒ“æÁ
        ActiveSheet.Shapes(1).Cut
    Wend
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=10, Top:=250, Width:=50, Height:=30)
        .Object.Caption = "ƒNƒŠƒA"
        .Name = "ƒNƒŠƒA"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=70, Top:=250, Width:=50, Height:=30)
        .Object.Caption = "ŒŸõ"
        .Name = "ŒŸõ"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=130, Top:=250, Width:=50, Height:=30)
        .Object.Caption = "XV"
        .Name = "XV"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=10, Top:=300, Width:=50, Height:=30)
        .Object.Caption = "‘OŒŸõ"
        .Name = "‘OŒŸõ"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=70, Top:=300, Width:=50, Height:=30)
        .Object.Caption = "ŸŒŸõ"
        .Name = "ŸŒŸõ"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=130, Top:=300, Width:=50, Height:=30)
        .Object.Caption = "æÁ"
        .Name = "æÁ"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=70, Top:=350, Width:=50, Height:=30)
        .Object.Caption = "I—¹"
        .Name = "I—¹"
    End With
End Function
'--------------------------------------------------------------------------------
'   ‹¤’Êˆ—@Eiyo.mdb ‚ÌƒI[ƒvƒ“
'--------------------------------------------------------------------------------
Function Eiyo91DB_Open()
    myCon.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
               "Data Source=" & ThisWorkbook.Path & "Eiyo.mdb;"""
End Function
'--------------------------------------------------------------------------------
'   ‹¤’Êˆ—@Eiyo.mdb ‚ÌƒNƒ[ƒY
'--------------------------------------------------------------------------------
Function Eiyo920DB_Close()
    myCon.Close
    Set myCon = Nothing
End Function
'--------------------------------------------------------------------------------
'   ‹¤’Êˆ—@‰æ–Ê—}~‚Ù‚©
'--------------------------------------------------------------------------------
Function Eiyo930Screen_Hold()
    Application.ScreenUpdating = False      '‰æ–Ê•`‰æ—}~
    Application.EnableEvents = False        'ƒCƒxƒ“ƒg”­¶—}~
    ActiveSheet.Unprotect                   'ƒV[ƒg‚Ì•ÛŒì‚ğ‰ğœ
End Function
'--------------------------------------------------------------------------------
'   ‹¤’Êˆ—@‰æ–Ê•`‰æ‚Ù‚©
'--------------------------------------------------------------------------------
Function Eiyo940Screen_Start()
    Application.ScreenUpdating = True           '‰æ–Ê•`‰æ‚Ì•œŠˆ
    Application.EnableEvents = True             'ƒCƒxƒ“ƒg”­¶ÄŠJ
    ActiveSheet.Protect UserInterfaceOnly:=True '•ÛŒì‚ğ—LŒø‚É‚·‚é
End Function
'--------------------------------------------------------------------------------
'   ‹¤’Êˆ—@ƒ{ƒ^ƒ“ì¬
'--------------------------------------------------------------------------------
Function Eiyo950Button_Add(in_L As Long, in_t As Long, in_W As Long, in_H As Long, in_text As String)
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=in_L, Top:=in_t, Width:=in_W, Height:=in_H)
        .Object.Caption = in_text
        .Name = in_text
    End With
End Function
'--------------------------------------------------------------------------------
'   ‹¤’Êˆ—@w’èƒV[ƒgíœ
'--------------------------------------------------------------------------------
Function Eiyo99_w’èƒV[ƒgíœ(Sname As String)
    Application.DisplayAlerts = False                                   'Šm”F—}~
    If Not IsError(Evaluate(Sname & "!a1")) Then: Sheets(Sname).Delete  'ƒV[ƒgíœ
    Application.DisplayAlerts = True                                    'Šm”F•œŠˆ
End Function

