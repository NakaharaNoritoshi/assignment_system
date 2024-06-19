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
'   Eiyo01_000 ��ʍ��ڒ�`
'   0:Field-Name
'   1:�Z���͈�
'   2:i/o ���͉ۂ𖾊m��
'   3:Field (D:DB G:Guid)
'   4:Type
'   5:Sample
'--------------------------------------------------------------------------------
Function Eiyo01_000init()
Dim Wtext   As String
    Wtext = "Gmesg,a025:a025,o,00,X ,Message"
    Wtext = Wtext & vbLf & "Fcode,g003:k003,i,D,X,1234567890"    'Fcode
    Wtext = Wtext & vbLf & "Fsave,l003:p003,o,G,90,1234567890"   'Fcode save
    Wtext = Wtext & vbLf & "Date1,g004:k004,i,D,Ds,2008/10/10"   '�������Ԏ�
    Wtext = Wtext & vbLf & "Nissu,l004:l004,i,D,90,1"            '����
    Wtext = Wtext & vbLf & "Namej,g005:o005,i,D,J ,�����P�����Q�����R��"
    Wtext = Wtext & vbLf & "Sex  ,g006:g006,i,D,X ,1"            '����"
    Wtext = Wtext & vbLf & "Birth,g007:k007,i,D,Ds,2001/1/1"     '���N����"
    Wtext = Wtext & vbLf & "Gyyyy,l007:m007,o,G,X ,"             '�a��N
    Wtext = Wtext & vbLf & "Age  ,n007:o007,o,D,90,"             '�N��"
    Wtext = Wtext & vbLf & "Hight,g008:i008,i,D,91,123.4"        '�g��
    Wtext = Wtext & vbLf & "Weght,g009:i009,i,D,91,123.4"        '�̏d
    Wtext = Wtext & vbLf & "Sibou,g010:i010,i,D,91,123.4"        '�牺���b
    Wtext = Wtext & vbLf & "Adrno,g011:j011,i,D,X ,123-4567"     '�X�֔ԍ�
    Wtext = Wtext & vbLf & "Adrs1,g012:v012,i,D,J ,�Z���[�P���Z���[�P���Z���[�P���Z���["
    Wtext = Wtext & vbLf & "Adrs2,g013:v013,i,D,J ,�Z���[�Q���Z���[�Q���Z���[�Q���Z���["
    Wtext = Wtext & vbLf & "Area1,g014:h014,o,D,X ,12"           '�n��
    Wtext = Wtext & vbLf & "Gare1,i014:i014,o,G,X ,�n�於"       '�n��
    Wtext = Wtext & vbLf & "Area2,g015:h015,o,D,X ,12"           '�n��
    Wtext = Wtext & vbLf & "Gare2,i015:i015,o,G,X ,�s�{��"       '�n��
    Wtext = Wtext & vbLf & "Q3rec,g016:k016,i,D,X ,1234567890"   'Q3.�H�K
    Wtext = Wtext & vbLf & "Q4rec,g017:i017,i,D,X ,12345"        'Q4.�x�{
    Wtext = Wtext & vbLf & "Q5rec,g018:h018,i,D,X ,123"          'Q5.�^��
    Wtext = Wtext & vbLf & "Q6r_a,g019:k019,i,D,X ,1234567890"   'Q6.���N-1
    Wtext = Wtext & vbLf & "Q6r_b,g020:k020,i,D,X ,1234567890"   'Q6.���N-2
    Wtext = Wtext & vbLf & "Q6r_c,g021:k021,i,D,X ,1234567890"   'Q6.���N-3
    Wtext = Wtext & vbLf & "Q6r_d,g022:k022,i,D,X ,1234567890"   'Q6.���N-4
    Wtext = Wtext & vbLf & "Q6r_e,g023:k023,i,D,X ,1234567890"   'Q6.���N-5
    Wtext = Wtext & vbLf & "Qjob1,q008:r008,i,D,X ,1234"         'Q7.�E��-1
    Wtext = Wtext & vbLf & "Qjob5,s008:s008,i,D,X ,1"            'Q7.�E��-2
    Wtext = Wtext & vbLf & "Qsyuf,q009:q009,i,D,X ,1"            'Qa.��w
    Wtext = Wtext & vbLf & "Qcnd1,q010:q010,i,D,X ,1"            'Qb.�D�P
    Wtext = Wtext & vbLf & "Qtony,q011:q011,i,G,X ,1"            'Qc.���A
    Wtext = Wtext & vbLf & "Qill1,r011:r011,o,D,X ,123456"       'Qc.���A
    Wtext = Wtext & vbLf & "Qkoke,q014:q014,i,G,X ,1"            'Qd.������
    Wtext = Wtext & vbLf & "Qill2,r014:r014,o,D,X ,123456"       'Qd.������
    Wtext = Wtext & vbLf & "Qsrmr,q015:r015,i,D,90,123"          'Qe.Spot-1
    Wtext = Wtext & vbLf & "Qsmin,s015:t015,i,D,90,123"          'Qe.Spot-2
    Wtext = Wtext & vbLf & "Qclab,q016:q016,i,D,90,1"            'Qf.�^����
    Wtext = Wtext & vbLf & "Qtobc,q017:q017,i,D,90,1"            'Qt.�i��
    Wtext = Wtext & vbLf & "Qsyog,q018:r018,i,D,90,12"           'Qg.�g�̏�Q
    Wtext = Wtext & vbLf & "Qwcnt,q019:q019,i,D,90,1"            'Qh.����CT
    Wtext = Wtext & vbLf & "Tenes,q020:q020,i,D,90,1"            'Qi.��ٷް�w��-1
    Wtext = Wtext & vbLf & "Tenee,r020:u020,i,D,92,12345.67"     'Qi.��ٷް�w��-2
    Wtext = Wtext & vbLf & "Tanps,q021:q021,i,D,90,1"            'Qj.���߸ �w��-1
    Wtext = Wtext & vbLf & "Tanpe,r021:u021,i,D,92,12345.67"     'Qj.���߸ �w��-2
    Wtext = Wtext & vbLf & "��ݾ�1,q023:af23,i,D,J ,��ݾ�1"      '
    Wtext = Wtext & vbLf & "��ݾ�2,q024:af24,i,D,J ,��ݾ�2"      '
    Wtext = Wtext & vbLf & "��ݾ�3,q025:af25,i,D,J ,��ݾ�3"      '
    Wtext = Wtext & vbLf & "Blood,ab03:ac03,i,D,X ,12"           'B1.���t�^
    Wtext = Wtext & vbLf & "Bscd1,ab04:ac04,i,D,X ,123"          'B2.�x�Е�-1
    Wtext = Wtext & vbLf & "Bscd2,ad04:ae04,i,D,X ,12"           'B2.�x�Е�-2
    Wtext = Wtext & vbLf & "Bhok1,ab05:ae05,i,D,X ,12345678"     'B3.�ی��L��
    Wtext = Wtext & vbLf & "Bhok2,ab06:ae06,i,D,X ,12345678"     'B4.�ی��ԍ�
    Wtext = Wtext & vbLf & "Bhant,ab07:ac07,i,D,X ,12"           'B5.������f����
    Wtext = Wtext & vbLf & "Barm ,ab08:ab08,i,D,X ,1"            'B6.�����r
    Wtext = Wtext & vbLf & "Bdate,ab09:af09,i,D,Ds,2008/10/10"   'B7.������
    Wtext = Wtext & vbLf & "Bbl01,ab10:ad10,i,D,91,123.41"       'B8.�Ԍ�����
    Wtext = Wtext & vbLf & "Bbl02,ab11:ad11,i,D,91,123.41"       'B8.���F�f��
    Wtext = Wtext & vbLf & "Bbl03,ab12:ad12,i,D,91,123.41"       'B8.��ĸد�
    Wtext = Wtext & vbLf & "Bbl04,ab13:ad13,i,D,91,123.41"       'B8.�ڽ�۰�
    Wtext = Wtext & vbLf & "Bbl05,ab14:ad14,i,D,91,123.41"       'B8.HDL
    Wtext = Wtext & vbLf & "Bbl06,ab15:ad15,i,D,91,123.41"       'B8.�������b
    Wtext = Wtext & vbLf & "Bbl07,ab16:ad16,i,D,91,123.41"       'B8.G.O.T.
    Wtext = Wtext & vbLf & "Bbl08,ab17:ad17,i,D,91,123.41"       'B8.G.P.T.
    Wtext = Wtext & vbLf & "Bbl09,ab18:ad18,i,D,91,123.41"       'B8.�A�_
    Wtext = Wtext & vbLf & "Bbl10,ab19:ad19,i,D,91,123.41"       'B8.����
    Wtext = Wtext & vbLf & "Bbl11,ab20:ad20,i,D,91,123.41"       'B8.�����ō�
    Wtext = Wtext & vbLf & "Bbl12,ab21:ad21,i,D,91,123.41"       'B8.�����Œ�
    Fld_Adrs1 = Split(Wtext, vbLf)

    Wtext = "100,�֓���,100" & vbLf & "101,�֓���,101" & vbLf & "102,�k�@��,102"
    Wtext = Wtext & vbLf & "103,���@�C,103" & vbLf & "104,�ߋE��,104" & vbLf & "105,�ߋE��,105"
    Wtext = Wtext & vbLf & "106,���@��,106" & vbLf & "107,�l�@��,107" & vbLf & "108,�k��B,108"
    Wtext = Wtext & vbLf & "109,���B,109" & vbLf & "110,�k�C��,110" & vbLf & "111,���@�k,111"
    Wtext = Wtext & vbLf & "201,�k�C��,110" & vbLf & "202,�X�@,111" & vbLf & "203,���@,111"
    Wtext = Wtext & vbLf & "204,�{��@,111" & vbLf & "205,�H�c�@,111" & vbLf & "206,�R�`�@,111"
    Wtext = Wtext & vbLf & "207,�����@,111" & vbLf & "208,���@,101" & vbLf & "209,�Ȗ؁@,101"
    Wtext = Wtext & vbLf & "210,�Q�n�@,101" & vbLf & "211,��ʁ@,100" & vbLf & "212,��t�@,100"
    Wtext = Wtext & vbLf & "213,�����@,100" & vbLf & "214,�_�ސ�,100" & vbLf & "215,�V���@,102"
    Wtext = Wtext & vbLf & "216,�x�R�@,102" & vbLf & "217,�ΐ�@,102" & vbLf & "218,����@,102"
    Wtext = Wtext & vbLf & "219,�R���@,101" & vbLf & "220,����@,101" & vbLf & "221,�򕌁@,103"
    Wtext = Wtext & vbLf & "222,�É��@,103" & vbLf & "223,���m�@,103" & vbLf & "224,�O�d�@,103"
    Wtext = Wtext & vbLf & "225,����@,105" & vbLf & "226,���s�@,104" & vbLf & "227,���@,104"
    Wtext = Wtext & vbLf & "228,���Ɂ@,104" & vbLf & "229,�ޗǁ@,105" & vbLf & "230,�a�̎R,105"
    Wtext = Wtext & vbLf & "231,����@,106" & vbLf & "232,�����@,106" & vbLf & "233,���R�@,106"
    Wtext = Wtext & vbLf & "234,�L���@,106" & vbLf & "235,�R���@,106" & vbLf & "236,�����@,107"
    Wtext = Wtext & vbLf & "237,����@,107" & vbLf & "238,���Q�@,107" & vbLf & "239,���m�@,107"
    Wtext = Wtext & vbLf & "240,�����@,108" & vbLf & "241,����@,108" & vbLf & "242,����@,108"
    Wtext = Wtext & vbLf & "243,�F�{�@,109" & vbLf & "244,�啪�@,108" & vbLf & "245,�{��@,109"
    Wtext = Wtext & vbLf & "246,������,109" & vbLf & "247,����@,109"
    Fld_Area = Split(Wtext, vbLf)
End Function
'--------------------------------------------------------------------------------
'   01_010 �ېH��ʂ̃��[�N�V�[�g���A�N�e�B�u�ɂȂ���
'--------------------------------------------------------------------------------
Function Eiyo01_010�ېH_Activate()
    ActiveSheet.Unprotect                           '�V�[�g�̕ی������
'    ActiveSheet.Protect UserInterfaceOnly:=True     '�ی��L���ɂ���
End Function
'--------------------------------------------------------------------------------
'   01_020 ��b��ʂ̃_�u���N���b�N
'   AA��i�����Y���������j�̃_�u���N���b�N�͊Y���ԍ����Z��[G3]�ɐݒ�
'--------------------------------------------------------------------------------
Function Eiyo01_020��b_BeforedoubleClick()
Dim Wadrs   As String
Dim Wcoul   As String
Dim Wtext   As String
Dim i1      As Long     'Fld_Adrs Index
Dim i3      As Long     '�_�u���N���b�N�̍s�ԍ�

End Function
'--------------------------------------------------------------------------------
'   01_030 �N���A_Click
'   ���͍��ڂ̏����A���[�E���؃V�[�g�̍폜
'--------------------------------------------------------------------------------
Function Eiyo01_030�N���AClick()
Dim i1      As Long
Dim FldItem As Variant
Dim Lmax    As Long

    Call Eiyo01_000init
    Call Eiyo930Screen_Hold     '��ʗ}�~�ق�
    
    For i1 = 0 To UBound(Fld_Adrs1)
        FldItem = Split(Fld_Adrs1(i1), ",")
        If FldItem(0) = "Gyyyy" Or _
           FldItem(0) = "Age  " Or _
           IsEmpty(Range(Trim(FldItem(0)))) Then
        Else
           Range(Trim(FldItem(0))) = Empty
        End If
    Next i1
    Call Eiyo01_820����K�C�h
    Lmax = Sheets("�ېH").UsedRange.Rows.Count
    If Lmax > 4 Then: Sheets("�ېH").Rows("5:" & Lmax).Delete Shift:=xlUp
    Call Eiyo99_�w��V�[�g�폜("����")
    Call Eiyo99_�w��V�[�g�폜("����2")
    Call Eiyo99_�w��V�[�g�폜("DBmirror")
    Call Eiyo99_�w��V�[�g�폜("��ݾ�ݸ޼��")
    Range("Fcode").Select
    Call Eiyo940Screen_Start    '��ʕ`��ق�
End Function
'--------------------------------------------------------------------------------
'   01_100 ����_Click
'       ��b���̌����A���肳�ꂽ�ꍇ�ɐېH�����擾����
'--------------------------------------------------------------------------------
Function Eiyo01_100����Click()
Dim FldItem     As Variant
Dim i1          As Long

    Call Eiyo930Screen_Hold     '��ʗ}�~�ق�
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
        Range("Gmesg") = "�����L�[������܂���"
    Else
        Call Eiyo01_110����(i1)
    End If
    
    If IsEmpty(Range("Fcode")) = False And _
       Range("Fcode") = Range("Fsave") Then         '���肳�ꂽ�ꍇ�͐ېH���
        Application.ScreenUpdating = False          '��ʕ`��}�~
        Call Eiyo01_130MealGet
        Sheets("��b").Select
    End If
    Range("Fcode").Select
    Call Eiyo940Screen_Start                        '��ʕ`��ق�
End Function
'--------------------------------------------------------------------------------
'   01_110 �c�a��������     F-024
'--------------------------------------------------------------------------------
Function Eiyo01_110����(i1 As Long)
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
        
    'SQL�œǂݍ��ރf�[�^���w�肷��
    in_key = Range(Trim(FldItem(0))).Text
    If Left(in_key, 1) = "��" Then: in_key = "%" & Right(in_key, Len(in_key) - 1)
    Call Eiyo91DB_Open      'DB Open
    If FldItem(0) = "Fcode" Then
        mySqlStr = "SELECT * FROM " & Tbl_Kiso & " Where Fcode = """ & in_key & """"
    Else
        mySqlStr = "SELECT * FROM " & Tbl_Kiso & " Where " & _
                   Trim(FldItem(0)) & " like """ & in_key & "%"""
    End If
    Set Rst_Kiso = myCon.Execute(mySqlStr)
    If Rst_Kiso.EOF Then
        Range("Gmesg") = "�Y���f�[�^�͂���܂���"
        Range("Fcode").Select
    Else
        With Rst_Kiso
            Range("Ah2").CopyFromRecordset Rst_Kiso           '���R�[�h
            If Range("Ah3") = Empty Then                        '�Y�����P���̂Ƃ�
                For i1 = 1 To UBound(Fld_Adrs1)                 '��ʍ��ڂ̏�������
                    FldItem = Split(Fld_Adrs1(i1), ",")
                    If FldItem(3) = "D" Then
                        For i2 = 0 To .Fields.Count - 1             '�t�B�[���h��
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
                Call Eiyo01_820����K�C�h
            Else
                For i1 = 1 To .Fields.Count                     '�t�B�[���h��
                    Cells(1, i1 + 33).Value = .Fields(i1 - 1).Name
                Next
                Columns("ah:hz").EntireColumn.AutoFit           '��
                i1 = Range("ah1").End(xlDown).Row
                Range("Ah2:ah" & i1).Locked = False             '���͉�
                Range("Ah2:ah" & i1).Interior.ColorIndex = 34
            End If
            .Close
        End With
    End If
    Set Rst_Kiso = Nothing                        '�I�u�W�F�N�g�̉��
    Call Eiyo920DB_Close    'DB Close
End Function
'--------------------------------------------------------------------------------
'   01_120 �n��E�s���{���\��
'--------------------------------------------------------------------------------
Function Eiyo01_120�n��(in_code As String) As String
Dim i1      As Long
Dim Witem   As Variant

    Eiyo01_120�n�� = Empty
    For i1 = 0 To UBound(Fld_Area)
        Witem = Split(Fld_Area(i1), ",")
        If Witem(0) = in_code Then
            Eiyo01_120�n�� = Witem(1)
            Exit For
        End If
    Next i1
End Function
'--------------------------------------------------------------------------------
'   01_130�@�ېH�擾
'--------------------------------------------------------------------------------
Function Eiyo01_130MealGet()
Dim mySqlStr    As String
Dim Lmax        As Long
Dim i1          As Long

    Sheets("�ېH").Select
    Application.EnableEvents = False                '�C�x���g�����}�~
'    ActiveSheet.Unprotect                           '�V�[�g�̕ی������
    Range("b1") = Empty
    Range("a2") = Range("Fcode") & ":" & Range("Namej")
    Lmax = ActiveSheet.UsedRange.Rows.Count
    If Lmax > 4 Then: Rows("5:" & Lmax).Delete Shift:=xlUp
        
    'SQL�œǂݍ��ރf�[�^���w�肷��
    Call Eiyo91DB_Open      'DB Open
    mySqlStr = "SELECT Sdate,Ekubn,Foodc,Suryo FROM " & Tbl_Meal & " Where Fcode = """ & Range("Fcode") & """"
    Set Rst_Meal = myCon.Execute(mySqlStr)
    If Rst_Meal.EOF Then
        Lmax = 0
    Else
        Range("A5").CopyFromRecordset Rst_Meal           '���R�[�h
    End If
    
    Lmax = ActiveSheet.UsedRange.Rows.Count
    For i1 = 5 To Lmax
        Cells(i1, 6) = Cells(i1, 4)
        Cells(i1, 4) = Cells(i1, 3)
        Cells(i1, 3) = Eiyo01_401�H���敪(Cells(i1, 2))
        Call Eiyo01_402�H�i�}�X�^(i1)
    Next i1
    Set Rst_Meal = Nothing                    '�I�u�W�F�N�g�̉��
    Call Eiyo920DB_Close    'DB Close
End Function
'--------------------------------------------------------------------------------
'   01_200 �X�V_Click
'--------------------------------------------------------------------------------
Function Eiyo01_200�X�VClick()
Dim Rtn As Long
    Call Eiyo930Screen_Hold                     '��ʗ}�~�ق�
    Call Eiyo01_000init
    Rtn = Eiyo01_210KeyCheck                    '�L�[�`�F�b�N
    If Rtn = 0 Then: Rtn = Eiyo01_220����Check  '���ڃ`�F�b�N
    If Rtn = 0 Then: Rtn = Eiyo01_230DB�X�V     'DB�X�V
    Call Eiyo940Screen_Start                    '��ʕ`��ق�
End Function
'--------------------------------------------------------------------------------
'   01_210 �L�[�`�F�b�N
'--------------------------------------------------------------------------------
Function Eiyo01_210KeyCheck() As Long
Dim mySqlStr    As String

    Call Eiyo91DB_Open      'DB Open
    mySqlStr = "SELECT * FROM " & Tbl_Kiso & " Where Fcode = """ & Range("Fcode") & """"
    Set Rst_Kiso = myCon.Execute(mySqlStr)
    If Rst_Kiso.EOF Then
        If Range("Fcode") = Range("Fsave") Then
            Range("Gmesg") = "Program Error Non Key & Save Key Same"    '�~�F�L�[�Ȃ��ASave����
            Eiyo01_210KeyCheck = 1
        Else
            Eiyo01_210KeyCheck = 0                                      '���F�L�[�Ȃ��ASave�قȂ�(�V�K)
        End If
    Else
        If Range("Fcode") = Range("Fsave") Then
            Eiyo01_210KeyCheck = 0                                      '���F�L�[����ASave����(�X�V)
        Else
            Range("Gmesg") = "�R�[�h���d�����Ă��܂�"                   '�~�F�L�[����ASave�قȂ�
            Eiyo01_210KeyCheck = 1
        End If
    End If
    Set Rst_Kiso = Nothing                        '�I�u�W�F�N�g�̉��
    Call Eiyo920DB_Close    'DB Close
End Function
'--------------------------------------------------------------------------------
'   01_220 ���ڃ`�F�b�N
'--------------------------------------------------------------------------------
Function Eiyo01_220����Check() As Long
Dim Witem   As Variant
Dim Wlen    As Long
Dim i1      As Long
Dim Wtemp   As String

    Eiyo01_220����Check = 1
    Range("Gmesg") = Empty
'   �R�[�h
'    Witem = Range("Fcode")
'    If IsNumeric(Witem) = True And Len(Witem) <= 10 Then
'    Else
'        Range("Gmesg") = "�R�[�h�͂P�O���ȓ��̐����ɂ��Ă��������@" & Len(Witem)
'        Range("Fcode").Activate
'        Exit Function
'    End If
'   �������ԊJ�n��
    If IsDate(Range("Date1")) Then
    Else
        Range("Gmesg") = "�������ԊJ�n�������ݓ��ɂ��Ă�������"
        Range("Date1").Activate
        Exit Function
    End If
'   �������ԓ���
    Witem = Range("Nissu")
    If IsNumeric(Witem) = True And Len(Witem) = 1 Then
    Else
        Range("Gmesg") = "�������ԓ����͂P���̐����ɂ��Ă�������"
        Range("Nissu").Activate
        Exit Function
    End If
'   ����
    If Eiyo01_221����check("Namej", "����", 10) = 1 Then: Exit Function
'   ����
    Witem = Range("sex")
    If Witem = "" Or Witem = "0" Or Witem = "1" Then
    Else
        Range("Gmesg") = "���ʂ͂P���̐����ɂ��Ă�������"
        Range("sex").Activate
        Exit Function
    End If
'   ���N����
    If IsDate(Range("Birth")) Then
    Else
        Range("Gmesg") = "���N���������ݓ��ɂ��Ă�������"
        Range("Birth").Activate
        Exit Function
    End If
    
    If Eiyo01_223���lcheck("Hight", "�g��", 3, 1, 300) = 1 Then: Exit Function
    If Eiyo01_223���lcheck("Weght", "�̏d", 3, 1, 300) = 1 Then: Exit Function
    If Eiyo01_223���lcheck("Sibou", "�牺���b", 2, 1, 50) = 1 Then: Exit Function
    
    If Eiyo01_221����check("Adrno", "�X�֔ԍ�", 18) = 1 Then: Exit Function
    If Eiyo01_221����check("Adrs1", "�Z���[�P", 18) = 1 Then: Exit Function
    If Eiyo01_221����check("Adrs2", "�Z���[�Q", 18) = 1 Then: Exit Function
'   �n��E�n��
    Wtemp = Left(Range("adrs1"), 2)
    For i1 = 0 To UBound(Fld_Area)
        Witem = Split(Fld_Area(i1), ",")
        If Left(Witem(0), 1) = "2" And _
           Left(Witem(1), 2) = Wtemp Then
            Range("Area1") = Right(Witem(2), 2)
            Range("Gare1") = Eiyo01_120�n��("1" & Range("Area1"))
            Range("Area2") = Right(Witem(0), 2)
            Range("Gare2") = Witem(1)
            Exit For
        End If
    Next i1
    If Eiyo01_222����check("Q3rec", "Q3.�H�K��", 10) = 1 Then: Exit Function
    If Eiyo01_222����check("Q4rec", "Q4.�x�{", 5) = 1 Then: Exit Function
    If Eiyo01_222����check("Q5rec", "Q5.�^��", 3) = 1 Then: Exit Function
    If Eiyo01_222����check("Q6r_a", "Q6.���N�P", 10) = 1 Then: Exit Function
    If Eiyo01_222����check("Q6r_b", "Q6.���N�Q", 10) = 1 Then: Exit Function
    If Eiyo01_222����check("Q6r_c", "Q6.���N�R", 10) = 1 Then: Exit Function
    If Eiyo01_222����check("Q6r_d", "Q6.���N�S", 10) = 1 Then: Exit Function
    If Eiyo01_222����check("Q6r_e", "Q6.���N�T", 10) = 1 Then: Exit Function
'   �E��
    Range("Qjob1") = UCase(Range("Qjob1"))
    If Len(Range("Qjob1")) = 4 Then
    Else
        Range("Gmesg") = "�E�Ƃ͂S���Ƃ��Ă��������@" & Len(Range("Qjob1"))
        Range("Qjob1").Activate
    End If

    If Eiyo01_222����check("Qsyuf", "QA.��w", 1) = 1 Then: Exit Function
    If Eiyo01_222����check("Qcnd1", "QB.�D�P", 1) = 1 Then: Exit Function
    If Eiyo01_222����check("Qtony", "QC.���A", 1) = 1 Then: Exit Function
    If Range("Qtony") = "0" Then
        Range("Qill1") = "000000"
    Else
        Range("Qill1") = "000321"
    End If
    If Eiyo01_222����check("Qkoke", "QC.���A", 1) = 1 Then: Exit Function
    If Range("Qkoke") = "0" Then
        Range("Qill2") = "000000"
    Else
        Range("Qill2") = "000313"
    End If
    If Eiyo01_223���lcheck("Qsrmr", "QE.��߰�1", 3, 0, 1000) = 1 Then: Exit Function
    If Eiyo01_223���lcheck("Qsmin", "QE.��߰�2", 3, 0, 1000) = 1 Then: Exit Function
    If Eiyo01_222����check("Qclab", "QF.�^����", 1) = 1 Then: Exit Function
    If Eiyo01_222����check("Qclab", "Q .�i��", 1) = 1 Then: Exit Function
    If Eiyo01_223���lcheck("Qsyog", "QG.�g��Q", 2, 0, 100) = 1 Then: Exit Function
    If Eiyo01_222����check("Qwcnt", "QG.����CT", 1) = 1 Then: Exit Function
    If Eiyo01_222����check("Tenes", "��ٷގw��", 1) = 1 Then: Exit Function
    If Eiyo01_222����check("Tanps", "���߸�w��", 1) = 1 Then: Exit Function
    If Eiyo01_223���lcheck("Tenee", "��ٷގw��", 5, 2, 100000) = 1 Then: Exit Function
    If Eiyo01_223���lcheck("Tanpe", "���߸�w��", 5, 2, 100000) = 1 Then: Exit Function
'
    Range("Blood") = UCase(Range("Blood"))
    Wtemp = Range("Blood")
    If Wtemp = "" Or Wtemp = "A" Or Wtemp = "B" Or Wtemp = "O" Or Wtemp = "AB" Then
    Else
        Range("Gmesg") = "���t�^���s���ł�"
        Range("Blood").Activate
        Exit Function
    End If

    If Eiyo01_221����check("Bscd1", "�x�X", 3) = 1 Then: Exit Function
    If Eiyo01_221����check("Bscd2", "�x��", 2) = 1 Then: Exit Function
    If Eiyo01_221����check("Bhok1", "�ی��؋L��", 8) = 1 Then: Exit Function
    If Eiyo01_221����check("Bhok2", "�ی���No", 8) = 1 Then: Exit Function
    If Eiyo01_221����check("Bhant", "������f", 2) = 1 Then: Exit Function
'
    Range("Barm") = UCase(Range("Barm"))
    Wtemp = Range("Barm")
    If Wtemp = "" Or Wtemp = "L" Or Wtemp = "R" Then
    Else
        Range("Gmesg") = "�����r���s���ł�"
        Range("Barm").Activate
        Exit Function
    End If
'   ���t������
    If IsEmpty(Range("Bdate")) Or IsDate(Range("Bdate")) Then
    Else
        Range("Gmesg") = "���t�����������ݓ��ɂ��Ă�������"
        Range("Bdate").Activate
        Exit Function
    End If
    If Eiyo01_223���lcheck("Bbl01", "�Ԍ�����", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223���lcheck("Bbl02", "���F�f��", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223���lcheck("Bbl03", "��ĸد�", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223���lcheck("Bbl04", "�ڽ�۰�", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223���lcheck("Bbl05", "HDL", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223���lcheck("Bbl06", "�������b", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223���lcheck("Bbl07", "G.O.T.", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223���lcheck("Bbl08", "G.P.T.", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223���lcheck("Bbl09", "�A�_", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223���lcheck("Bbl10", "����", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223���lcheck("Bbl11", "�����ō�", 3, 1, 10000) = 1 Then: Exit Function
    If Eiyo01_223���lcheck("Bbl12", "�����Œ�", 3, 1, 10000) = 1 Then: Exit Function
    Eiyo01_220����Check = 0
End Function
'--------------------------------------------------------------------------------
'   01_221 �����`�F�b�N
'--------------------------------------------------------------------------------
Function Eiyo01_221����check(Ifld As String, Iname As String, Ilen As Long) As Long
    If Len(Range(Ifld)) > Ilen Then
        Range("Gmesg") = Iname & "��" & Ilen & "���ȓ��ɂ��Ă��������@" & Len(Range(Ifld))
        Range(Ifld).Activate
        Eiyo01_221����check = 1
    Else
        Eiyo01_221����check = 0
    End If
End Function
'--------------------------------------------------------------------------------
'   01_222 �Œ茅�������ڃ`�F�b�N
'--------------------------------------------------------------------------------
Function Eiyo01_222����check(Ifld As String, Iname As String, Ilen As Long) As Long
Dim Witem   As Variant
Dim Wlen    As Long

    If Range(Ifld) = Empty Then: Range(Ifld) = String(Ilen, "0")
    Witem = Range(Ifld)
    Wlen = Len(Witem)
    If IsNumeric(Witem) And Wlen = Ilen Then
        Eiyo01_222����check = 0
    Else
        Range("Gmesg") = Iname & "��" & Ilen & "���̐����ɂ��Ă��������@" & Wlen
        Range(Ifld).Activate
        Eiyo01_222����check = 1
    End If
End Function
'--------------------------------------------------------------------------------
'   01_223 ���l���ڃ`�F�b�N
'--------------------------------------------------------------------------------
Function Eiyo01_223���lcheck(Ifld As String, Iname As String, _
                              Ilen1 As Long, Ilen2 As Long, Imax As Long) As Long
Dim Witem   As Variant
    
    Witem = Range(Ifld)
    If IsNumeric(Witem) And Witem < Imax Then
        Eiyo01_223���lcheck = 0
    Else
        Range("Gmesg") = Iname & "�͏�" & Ilen1 & "����" & Ilen2 & "���ȓ��̐��l�ɂ��Ă�������"
        Range(Ifld).Activate
        Eiyo01_223���lcheck = 1
    End If
End Function
'--------------------------------------------------------------------------------
'   01_230 �c�a�X�V                                     F-026
'   Microsoft ActiveX Data Objects 2.X Library �Q�Ɛݒ�
'--------------------------------------------------------------------------------
Function Eiyo01_230DB�X�V() As Long
Dim FldItem     As Variant
Dim FldName     As String
Dim i1          As Long

    Call Eiyo91DB_Open      'DB Open
    '���������܂�
    With Rst_Kiso
        '�C���f�b�N�X�̐ݒ�
        .Index = "PrimaryKey"
        '���R�[�h�Z�b�g���J��
        Rst_Kiso.Open Source:=Tbl_Kiso, ActiveConnection:=myCon, _
            CursorType:=adOpenKeyset, LockType:=adLockOptimistic, _
            Options:=adCmdTableDirect
        '�ԍ����o�^����Ă��邩��������
        If Not .EOF Then .Seek Range("Fcode")
        If .EOF Then
            .AddNew
            Range("Gmesg") = "�ǉ��o�^����܂����B"
            Range("Fsave") = Range("Fcode")
        Else
            Range("Gmesg") = "�X�V����܂����B"
        End If
        For i1 = 1 To UBound(Fld_Adrs1)                 '��ʍ��ڂ̏�������
            FldItem = Split(Fld_Adrs1(i1), ",")         '
            If FldItem(3) = "D" Then
                FldName = Trim(FldItem(0))
                .Fields(FldName).Value = Range(FldName).Value
            End If
        Next i1
        .Update
        .Close
    End With
    Set Rst_Kiso = Nothing      '�I�u�W�F�N�g�̉��
    Call Eiyo920DB_Close        'DB Close
    Eiyo01_230DB�X�V = 0
End Function
'--------------------------------------------------------------------------------
'   01_300�@���_Click
'--------------------------------------------------------------------------------
Function Eiyo01_300���Click()
    If Range("Fcode") = Range("Fsave") And _
        IsEmpty(Range("Fcode")) = False Then
        Call Eiyo91DB_Open      'DB Open
        myCon.Execute "DELETE FROM " & Tbl_Kiso & " Where Fcode = """ & Range("Fcode") & """"
        myCon.Execute "DELETE FROM " & Tbl_Meal & " Where Fcode = """ & Range("Fcode") & """"
        Range("Gmesg") = "����폜����܂����B"
        Range("Fsave") = Empty
        Call Eiyo920DB_Close    'DB Close
    Else
        Range("Gmesg") = "��������Ă��܂���B"
    End If
End Function
'--------------------------------------------------------------------------------
'   01_400�@�ېH�\��
'--------------------------------------------------------------------------------
Function Eiyo01_400MealDisp()
Dim Rtn     As Long
Dim Wmsg    As String
    
    Range("a2") = Range("Fcode") & ":" & Range("Namej")
    Wmsg = "��b���̌������s���Ă��܂���"
    If IsEmpty(Range("Fcode")) Or Range("Fcode") <> Range("Fsave") Then
        Rtn = CreateObject("WScript.Shell").Popup(Wmsg, 3, "Microsoft Excel", 0)
        Sheets("��b").Select
    End If
End Function
'--------------------------------------------------------------------------------
'   01_401�@�H���敪
'--------------------------------------------------------------------------------
Function Eiyo01_401�H���敪(kbn As Long) As String
    Select Case kbn
        Case 1: Eiyo01_401�H���敪 = "��"
        Case 2: Eiyo01_401�H���敪 = "��"
        Case 3: Eiyo01_401�H���敪 = "�["
        Case 4: Eiyo01_401�H���敪 = "��"
        Case 5: Eiyo01_401�H���敪 = "��"
        Case Else: Eiyo01_401�H���敪 = Empty
    End Select
End Function
'--------------------------------------------------------------------------------
'   01_402�@�H�i�}�X�^�擾
'--------------------------------------------------------------------------------
Function Eiyo01_402�H�i�}�X�^(in_line As Long)
Dim mySqlStr    As String
    If IsEmpty(Cells(in_line, 4)) Then
        Cells(in_line, 5) = Empty
        Range("g" & in_line & ":z" & in_line) = Empty
        Exit Function
    End If
        
    mySqlStr = "SELECT * FROM " & Tbl_Food & " Where Foodc = " & Cells(in_line, 4)
    Set Rst_Food = myCon.Execute(mySqlStr)
    If Rst_Food.EOF Then
        Cells(in_line, 5) = "�L�[�Ȃ�"
        Range("g" & in_line & ":z" & in_line) = Empty
    Else
        Cells(in_line, 7).CopyFromRecordset Rst_Food
        Cells(in_line, 5) = Cells(in_line, 8)
    End If
    Rst_Food.Close
    Set Rst_Food = Nothing
End Function
'--------------------------------------------------------------------------------
'   01_410�@�ېH��ʂ��ύX���ꂽ
'--------------------------------------------------------------------------------
Function Eiyo01_410MealChange(ChangeCell As String)
Dim Wl      As Long
Dim Wc      As Long

    Wl = Range(ChangeCell).Row
    Wc = Range(ChangeCell).Column
'    ActiveSheet.Unprotect                           '�V�[�g�̕ی������
    Select Case Wc
        Case 1: Cells(Wl, 2).Select
        Case 2
            Cells(Wl, 3) = Eiyo01_401�H���敪(Cells(Wl, 2))
            Cells(Wl, 4).Select
        Case 4
            'SQL�œǂݍ��ރf�[�^���w�肷��
            Call Eiyo91DB_Open      'DB Open
            Call Eiyo01_402�H�i�}�X�^(Wl)
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
'    ActiveSheet.Protect UserInterfaceOnly:=True     '�ی��L���ɂ���
End Function
'--------------------------------------------------------------------------------
'   01_420�@���j���[���I�����ꂽ
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
    If IsEmpty(Sheets("�ېH").Range("b1")) Then
        Wmsg = Wcode & ":" & Wname & "���I������܂����B"
        Rtn = CreateObject("WScript.Shell").Popup(Wmsg, 3, "Microsoft Excel", 0)
    Else
        Application.EnableEvents = False            '�C�x���g�����}�~
        Sheets("�ېH").Select
        Wcell = Range("b1")
        Range("b1") = Empty
        Application.EnableEvents = True             '�C�x���g�����ĊJ
        Range(Wcell) = Wcode
    End If
End Function
'--------------------------------------------------------------------------------
'   01_500�@�o�^ Click
'--------------------------------------------------------------------------------
Function Eiyo01_500MealCalc(in_Func As Long)
    Call Eiyo930Screen_Hold                                 '��ʗ}�~�ق�
    Call Eiyo91DB_Open                                      'DB Open
    If Eiyo01_501MealEntry = 1 Then: GoTo Eiyo01_503Exit    '�����̓`�F�b�N
    If Eiyo01_502Mealscope = 1 Then: GoTo Eiyo01_503Exit    '�ېH�ʂ͈̔̓`�F�b�N
    If Eiyo01_503Mealzerod = 1 Then: GoTo Eiyo01_503Exit    '�ېH�ʃ[���̍폜
    If Eiyo01_504MealDoubl = 1 Then: GoTo Eiyo01_503Exit    '�ېH�̏d������
    If Eiyo01_510MealUdate = 1 Then: GoTo Eiyo01_503Exit    '�c�a�X�V
    If Eiyo01_511MealFldgt = 1 Then: GoTo Eiyo01_501Exit    '���ڗv�f�擾
    If Eiyo01_512MealSheet = 1 Then: GoTo Eiyo01_501Exit    '�ېH�v�Z�V�[�g
    If Eiyo01_513kenso2sht = 1 Then: GoTo Eiyo01_501Exit    '���؂Q�V�[�g�쐬
    If Eiyo01_514MealCalc1 = 1 Then: GoTo Eiyo01_501Exit    '�ېH�v�Z
    If Eiyo01_515MealTotal = 1 Then: GoTo Eiyo01_501Exit    '�ېH�ʍ��v
    If Eiyo01_521CalcDbGet(1) = 1 Then: GoTo Eiyo01_501Exit '�ېH�ʍ��v
    If Eiyo01_522Mealcalc2 = 1 Then: GoTo Eiyo01_501Exit    '�W���̏d�ق�
    If Eiyo01_525MealDiffe = 1 Then: GoTo Eiyo01_501Exit    '�ߕs���A�h�o�C�X
    If Eiyo01_528Eiyohirit = 1 Then: GoTo Eiyo01_501Exit    '�h�{�䗦
    If in_Func = 2 Then
        If Eiyo01_540Old_Check = 1 Then: GoTo Eiyo01_501Exit    '���v�Z�l
'    Else
'        Call Eiyo99_�w��V�[�g�폜("DBmirror")
    End If
Eiyo01_501Exit:
    Call Eiyo01_550RstClose
Eiyo01_503Exit:
    Call Eiyo920DB_Close                'DB Close
    Sheets("��b").Select
    Call Eiyo940Screen_Start            '��ʕ`��ق�
End Function
'--------------------------------------------------------------------------------
'   01_501�@�ېH��񖢓��̓`�F�b�N
'--------------------------------------------------------------------------------
Function Eiyo01_501MealEntry() As Long
Dim Lmax    As Long
Dim i1      As Long
Dim Wnon    As Long

    Eiyo01_501MealEntry = 1
    Lmax = ActiveSheet.UsedRange.Rows.Count
    If Lmax < 5 Then
        MsgBox "�f�[�^������܂���"
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
        MsgBox ("���̍��ڂ��C�����Ă��������B")
    Else
        Eiyo01_501MealEntry = 0
    End If
End Function
'--------------------------------------------------------------------------------
'   01_502�@�ېH�ʂ͈̔̓`�F�b�N
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
        Wmsg = "�ېH�ʂُ̈�l��" & Wover & "��������܂�"
        i1 = CreateObject("WScript.Shell").Popup(Wmsg, 1, "Microsoft Excel", 0)
        Eiyo01_502Mealscope = 0
    End If
    Eiyo01_502Mealscope = 0
End Function
'--------------------------------------------------------------------------------
'   01_503�@�ېH�ʃ[���̍폜
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
        Wmsg = "�ېH�ʃ[����" & Wzero / 2 & "�s���폜���܂����B"
        i1 = CreateObject("WScript.Shell").Popup(Wmsg, 1, "Microsoft Excel", 0)
    End If
    Eiyo01_503Mealzerod = 0
End Function
'--------------------------------------------------------------------------------
'   01_504�@�ېH�̏d���`�F�b�N
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
'   01_510�@�ېHDB�o�^
'--------------------------------------------------------------------------------
Function Eiyo01_510MealUdate() As Long
Dim Lmax    As Long
Dim i1      As Long
Dim Wkey    As Variant

    Lmax = Range("a4").End(xlDown).Row
    myCon.Execute "DELETE FROM " & Tbl_Meal & " Where Fcode = """ & Range("Fcode") & """"
    '���������܂�
    With Rst_Meal
        '�C���f�b�N�X�̐ݒ�
        .Index = "PrimaryKey"
        '���R�[�h�Z�b�g���J��
        Rst_Meal.Open Source:=Tbl_Meal, ActiveConnection:=myCon, _
            CursorType:=adOpenKeyset, LockType:=adLockOptimistic, _
            Options:=adCmdTableDirect
        For i1 = 5 To Lmax
        '�ԍ����o�^����Ă��邩��������
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
    Set Rst_Meal = Nothing                    '�I�u�W�F�N�g�̉��
    Eiyo01_510MealUdate = 0
End Function
'--------------------------------------------------------------------------------
'   01_511�@�h�{�f���ڂ̊e����擾    F-018
'--------------------------------------------------------------------------------
Function Eiyo01_511MealFldgt() As Long
    
    Sheets.Add After:=Sheets(Sheets.Count)      '�V�[�g�ǉ�
    '���R�[�h�Z�b�g���J��
    Rst_Field.Open Source:=Tbl_Field, _
                ActiveConnection:=myCon, _
                CursorType:=adOpenForwardOnly, _
                LockType:=adLockReadOnly, _
                Options:=adCmdTableDirect
    '���R�[�h
    Range("a1").CopyFromRecordset Rst_Field
    Fld_Field = ActiveSheet.UsedRange
    Rst_Field.Close
    Set Rst_Field = Nothing    '�I�u�W�F�N�g�̉��
    Application.DisplayAlerts = False               '�m�F�}�~
    ActiveSheet.Delete
    Application.DisplayAlerts = True                '�m�F����
    Eiyo01_511MealFldgt = 0
End Function
'--------------------------------------------------------------------------------
'   01_512�@�ېH�v�Z�V�[�g�쐬
'--------------------------------------------------------------------------------
Function Eiyo01_512MealSheet() As Long
Dim i1      As Long     '�sIndex
Dim i2      As Long     '��Index
Dim Wno     As String
Dim Wtext   As String

    Call Eiyo99_�w��V�[�g�폜("����")
    Sheets.Add After:=Sheets(Sheets.Count)      '�V�[�g�ǉ�
    ActiveSheet.Name = "����"
    Range("d1") = "�h�{�v�Z�@�ېH����"
    Range("a2") = Sheets("�ېH").Range("a2")
    Wtext = Empty
    For i1 = 1 To 27                            '�h�{�f
        Wno = Format(i1, "00")
        Wtext = Wtext & "�ێ��" & Wno & vbTab
        Wtext = Wtext & "�M����" & Wno & vbTab
    Next i1
    For i1 = 1 To 15
        Wtext = Wtext & "��ٷ�C" & Format(i1, "00") & vbTab
    Next i1
    For i1 = 1 To 15
        Wtext = Wtext & "��ٷ�W" & Format(i1, "00") & vbTab
    Next i1
    For i1 = 1 To 15
        Wtext = Wtext & "�ټ��1" & Format(i1, "00") & vbTab
    Next i1
    For i1 = 1 To 15
        Wtext = Wtext & "�ټ��2" & Format(i1, "00") & vbTab
    Next i1
    Wtext = Wtext & "��������" & vbTab
    Wtext = Wtext & "��������" & vbTab
    Wtext = Wtext & "�����A��" & vbTab
    Wtext = Wtext & "�M������" & vbTab
    Wtext = Wtext & "�M������" & vbTab
    Wtext = Wtext & "�M���A��"
    Range("a4:dp4") = Split(Wtext, vbTab)
    ActiveWindow.FreezePanes = False        '�E�C���h�g�Œ�̉���
    Range("a5").Select
    ActiveWindow.FreezePanes = True         '�E�C���h�g�Œ�̐ݒ�
    Cells.NumberFormatLocal = "#,##0.00;[��]-#,##0.00"
    Eiyo01_512MealSheet = 0
End Function
'--------------------------------------------------------------------------------
'   01_513�@���؂Q�V�[�g�쐬
'--------------------------------------------------------------------------------
Function Eiyo01_513kenso2sht() As Long
Dim i1      As Long
    Call Eiyo99_�w��V�[�g�폜("����2")
    Sheets.Add After:=Sheets(Sheets.Count)      '�V�[�g�ǉ�
    ActiveSheet.Name = "����2"
    Cells.Interior.ColorIndex = 36              '�S��ʔw�i�F
'   �\��
    Range("C1:F1").Select
    Selection.MergeCells = True                 '�\��Z���A��
    Selection.HorizontalAlignment = xlCenter    '�\��Z���^�����O
    Selection.Interior.ColorIndex = 37          '�\��F�i�y�[���u���[�j
    With Selection.Font                         '�t�H���g
        .FontStyle = "����"
        .Size = 16
    End With
    Range("C1") = "�h�{�v�Z�@���؎����Q"
    
'   �h�{�f��
    Range("a4") = "No.�h�{�f��"
    Range("b4") = "�P��"
    For i1 = 1 To 27
        Cells(4 + i1, 1) = Format(i1, "00") & "." & Fld_Field(i1, 4)
        Cells(4 + i1, 2) = Fld_Field(i1, 5)
    Next i1
'   �ێ��
    Range("c3") = "<========= �M����ێ�� ==========>"
    Range("c4") = "����"
    Range("d4") = "�^��"
    Range("e4") = "��"
    Range("f4") = "�␳��"
    Range("c4:p4").HorizontalAlignment = xlCenter
    Range("c5:d31,f5:f31").Interior.ColorIndex = xlNone      '��������
    With Range("c5:d31,f5:f31").Borders                      '�g�r��
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .Weight = xlThin
    End With
    Range("c5:d31,f5:f31").NumberFormatLocal = "#,##0.00;[��]-#,##0.00"
    Range("c5").Name = "ks2_eiyoso"
'   �ێ�ʂ̕␳����
    Range("e:e,o:o").HorizontalAlignment = xlCenter '������
    Range("e15,e20,e24,e27").Interior.ColorIndex = xlNone   '��������
    With Range("e15,e20,e24,e27").Borders                   '�g�r��
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .Weight = xlThin
    End With
    Range("e15").Name = "ks2_hosei11"
    Range("e20").Name = "ks2_hosei16"
    Range("e24").Name = "ks2_hosei20"
    Range("e27").Name = "ks2_hosei23"
'   ��b���
    Range("i4") = "���̏d"
    Range("j4") = "�W���̏d"
    Range("h5") = "a.�̏d"
    Range("H6") = "b.�̕\�ʐ�"
    Range("H8") = "��b���"
    Range("h9") = "c.�����w��"
    Range("H10") = "d.�R�[�h"
    Range("H11") = "e.�ʐϓ���"
    Range("H12") = "f.�^��"
    Range("H13") = "g.�^��"
    Range("H15") = "�G�l���M�["
    Range("H16") = "h.�W����"
    Range("H17") = "i.�K�p����"
    Range("H18") = "j.��ٷް1"
    Range("H19") = "k.��ٷް2"
    Range("i21") = "f = b * e * 24"
    Range("i22") = "h = f * (1+c) * 1.1"
    Range("F05").Copy Range("I5:j6,i9:i11,i12:j13,i16:j16,i18:i19")
    Range("E15").Copy Range("I10,i17")
    Range("i05").Name = "ks2_weght"     '�̏d
    Range("i06").Name = "ks2_Aansa"     '�̕\�ʐ�
    Range("i09").Name = "ks2_Aansx"     '�����w��
    Range("i10").Name = "ks2_kisocd"    '��b�R�[�h
    Range("i11").Name = "ks2_kisot"     '�ʐϓ����b���
    Range("i12").Name = "ks2_Aansb"     '��b��Ӂ^��
    Range("i13").Name = "ks2_Aansc"     '��b��Ӂ^��
    Range("i16").Name = "ks2_Aansd"     '��ٷް�W����
    Range("i17").Name = "ks2_energ"     '���v�ʴ�ٷް����
'   ���v��(��ٷް�ȊO)
    Range("l3") = "���vTBL"
    Range("m3") = "�������x"
    Range("n3") = "�D�P�␳"
    Range("o3") = "��"
    Range("p3") = "���v��"
    Range("q3") = "�E�v"
    Range("E15").Copy Range("l4:n4")
    Range("F05").Copy Range("l6:n31")
    Range("l4").Name = "ks2_syoyo"
    Range("F05").Copy Range("p5:p31")

    With ActiveSheet.PageSetup
        .Orientation = xlLandscape                      '����
'        .PrintHeadings = True                           '�s��ԍ�
        .LeftMargin = Application.InchesToPoints(0.4)   '���]��
        .RightMargin = Application.InchesToPoints(0.2)  '�E�]��
        .Zoom = False
        .FitToPagesWide = 1                             '���P��
        .FitToPagesTall = 1                             '�c�P��
    End With
    Cells.EntireColumn.AutoFit                          '��
    Range("C:D,F:F,I:J,L:N,P:P").ColumnWidth = 10
    Range("G:G,K:K").ColumnWidth = 4
    Range("a01") = Range("Namej")
    Range("a02") = "[" & Range("Fcode") & "]"
    Range("q05") = "��ٷް1"
    Range("q06") = "�N�ߐ��ʂق�"
    Range("q07") = "����ς��� 1/2"
    Range("q08") = "����ς��� 1/2"
    Range("q09") = "��ٷް1 & �������x"
    Range("q10") = "��ٷް1����v�Z"
    Range("q11") = "��ٷް1 * 0.0099"
    Range("q12") = "���̏d�ƌW��"
    Range("q13") = "�ټ�ѓ��l"
    Range("q14") = "�N��"
    Range("q15") = "TBL�l(�������͎w��l)"
    Range("q16") = "TBL�l"
    Range("q17") = "��ٷް2 * 0.0004"
    Range("q18") = "��ٷް2 * 0.00055"
    Range("q19") = "��ٷް2 * 0.0066"
    Range("q20") = "TBL�l"
    Range("q21") = "(TBL�l)"
    Range("q22") = "(TBL�l)"
    Range("q23") = "(TBL�l)"
    Range("q24") = "�s�O�a���b�_ * 0.6"
    Range("q25") = "��سѓ��l"
    Range("q26") = "�ټ�� 1/2"
    Range("q27") = "�������͎w��l"
    Range("q28") = "�ꗥ"
    Range("q29") = "���� 66%"
    Range("q30") = "���� 34%"
    Range("q31") = "���A/���ĥ���۰�/���"
    Sheets("����").Select
    Eiyo01_513kenso2sht = 0
End Function
'--------------------------------------------------------------------------------
'   01_514�@�ېH�v�Z
'--------------------------------------------------------------------------------
Function Eiyo01_514MealCalc1() As Long
Dim aa      As Variant
Dim bb      As Worksheet
Dim i1      As Long     '�sIndex
Dim i2      As Long     '��Index
Dim Lmax    As Long     '�sMax
Dim Wtemp1  As Double
Dim wtemp2  As Double
    
    aa = Sheets("�ېH").UsedRange
    Set bb = Sheets("����")
    Lmax = UBound(aa, 1)
    For i1 = 5 To Lmax
        For i2 = 1 To 27
            If i1 = 10 And i2 = 2 Then
                Wtemp1 = 1
            End If
            
'           �h�{�f�v�Z =   �ێ��(F) * �h�{�f(S)       * ���Z�l(M)
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
'           ��ٷްC              =      �ێ��(F) * ��ٷްC(AT)     * ���Z�l(M)
            Cells(i1, i2 + 54) = WorksheetFunction.Round(aa(i1, 6) * aa(i1, i2 + 45) * aa(i1, 13), 2)
'           ��ٷްW              =      �ێ��(F) * ��ٷްW(BI)     * ���Z�l(M)
            Cells(i1, i2 + 69) = WorksheetFunction.Round(aa(i1, 6) * aa(i1, i2 + 60) * aa(i1, 13), 2)
'           �ټ��  =       �ێ��(F) * �ټ��           * ���Z�l
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
'           ����    =      �ێ��(F) * ����(CM)        * ���Z�l(M)
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
'   01_515�@�ېH�ʍ��v
'--------------------------------------------------------------------------------
Function Eiyo01_515MealTotal() As Long
Dim Lmax    As Long     '�sMax
Dim i1      As Long     '�sIndex
Dim i2      As Long     '��Index
Dim Wnissu  As Long

    Wnissu = Range("Nissu")
    Lmax = Sheets("�ېH").UsedRange.Rows.Count
    i1 = Lmax + 2
    For i2 = 1 To 120
        Cells(i1, i2) = "=SUM(R5C:R[-2]C)"
        Cells(i1 + 1, i2) = "=round(R[-1]C/" & Wnissu & ",2)"
    Next i2
    For i2 = 2 To 54 Step 2
        Range("ks2_eiyoso").Offset(i2 / 2 - 1, 0) = Cells(i1, i2)
        Range("ks2_eiyoso").Offset(i2 / 2 - 1, 1) = Cells(i1 + 1, i2)
    Next i2
    
    i1 = i1 + 1                             '��������荇�v�sIndex
    If Mid(Range("Q3rec"), 7, 1) = "3" Then '�����ɂ��␳
        If Cells(i1, 45) > 17 Then                           '��16G
            Cells(i1, 45) = Cells(i1, 45) - 8                 ' -8G
            Cells(i1, 21) = Cells(i1, 21) - 3149              '��س�
        ElseIf Cells(i1, 45) > 9 Then                        ' 9G��
            Cells(i1, 21) = Cells(i1, 21) - _
                          WorksheetFunction.RoundDown( _
                          (Cells(i1, 45) - 9) / 0.00254, 2)
            Cells(i1, 45) = 9                                 '����
        End If
        If Cells(i1, 46) > 17 Then                           '��16G
            Cells(i1, 46) = Cells(i1, 46) - 8                ' -8G
            Cells(i1, 22) = Cells(i1, 22) - 3149             '��س�
            Range("ks2_hosei23") = 1
        ElseIf Cells(i1, 46) > 9 Then                        ' 9G��
            Cells(i1, 22) = Cells(i1, 22) - _
                          WorksheetFunction.RoundDown( _
                          (Cells(i1, 46) - 9) / 0.00254, 2)
            Cells(i1, 46) = 9                                 '����
            Range("ks2_hosei23") = 2
        Else
            Range("ks2_hosei23") = 3
        End If
    Else
        Range("ks2_hosei23") = 4
    End If
    Range("ks2_hosei11") = Range("ks2_hosei23")
    
    
    If Cells(i1, 32) < 40 Then              'VC�␳(16)
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
    If Cells(i1, 40) < 3 Then               'VE�␳(20)
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
'   01_521�@��b���ق��擾
'           ���� Func 1:�X�V���� 2:�X�V�Ȃ�
'--------------------------------------------------------------------------------
Function Eiyo01_521CalcDbGet(Func As Long) As Long
Dim mySqlStr    As String
Dim i1          As Long
Dim i2          As Long

    Call Eiyo99_�w��V�[�g�폜("DBmirror")
    Sheets.Add After:=Sheets(Sheets.Count)      '�V�[�g�ǉ�
    ActiveSheet.Name = "DBmirror"
    Sheets("DBmirror").Range("m2:m3").NumberFormatLocal = "@"     '�Z���[�Q
'   ��b���擾
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
'   ���v�ʎ擾
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
'   �G�l���M�[�^�J�����[
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
        i2 = Sheets("����").UsedRange.Rows.Count
'       �����̖߂�
        Rst_Kiso.Fields("Sfods1").Value = Sheets("����").Cells(i2, 115)
        Rst_Kiso.Fields("Sfods2").Value = Sheets("����").Cells(i2, 116)
        Rst_Kiso.Fields("Sfods3").Value = Sheets("����").Cells(i2, 117)
        Rst_Kiso.Fields("Sfodh1").Value = Sheets("����").Cells(i2, 118)
        Rst_Kiso.Fields("Sfodh2").Value = Sheets("����").Cells(i2, 119)
        Rst_Kiso.Fields("Sfodh3").Value = Sheets("����").Cells(i2, 120)
'       ���v�ʂ̖߂�
        For i1 = 1 To 27
            Rst_Syoyo.Fields(i1 * 5 - 3).Value = Sheets("����").Cells(i2, i1 * 2 - 1)
            Rst_Syoyo.Fields(i1 * 5 - 2).Value = Sheets("����").Cells(i2, i1 * 2)
        Next i1
'       �G�l���M�[�^�J�����[�̖߂�
        For i1 = 1 To 15
            Rst_Energ.Fields(i1 + 1).Value = Sheets("����").Cells(i2, i1 + 54)
            Rst_Energ.Fields(i1 + 16).Value = Sheets("����").Cells(i2, i1 + 69)
            Rst_Energ.Fields(i1 + 31).Value = WorksheetFunction.Round(Sheets("����").Cells(i2, i1 + 54) / 80, 2)
            Rst_Energ.Fields(i1 + 58).Value = Sheets("����").Cells(i2, i1 + 84)
            Rst_Energ.Fields(i1 + 73).Value = Sheets("����").Cells(i2, i1 + 99)
        Next i1
    End If
    
    Eiyo01_521CalcDbGet = 0
End Function
'--------------------------------------------------------------------------------
'   01_522 �W���̏d�ق�
'--------------------------------------------------------------------------------
Function Eiyo01_522Mealcalc2() As Long
Dim mySqlStr    As String
Dim ���̏d      As Double
Dim �W���̏d    As Double
Dim �J��        As String   '�J�십�x�@Q7.�E�Ƃ̂S����
Dim �D�P        As Long
Dim Wtemp       As Double
Dim i1          As Long
Dim Wcondition  As Long     '���v�ʃG�l���M�[�K�p����
Dim Wenerg1     As Double   '�w�肠��̒�����ٷ�
Dim Wenerg2     As Double   '�w�菜�O�̒�����ٷ�
Dim KisoCd1     As Long     '�K�v�ʃ}�X�^�̊�b�R�[�h  ��b��ӁA�������
Dim KisoCd2     As Long     '�K�v�ʃ}�X�^�̊�b�R�[�h�@���v��
Dim KisoCd3     As Long     '�K�v�ʃ}�X�^�̊�b�R�[�h�@�D�P����
Dim �N��        As Long
Dim Warray      As Variant
Dim Wtext       As String

    ���̏d = Range("Weght")
    �N�� = Range("Age")
    �J�� = Mid(Range("Qjob1"), 4, 1)
    �D�P = Range("Qcnd1")
'   �W���̏d
    If �N�� <= 12 Then
        �W���̏d = ���̏d
    ElseIf Range("Hight") <= 150 Then
        �W���̏d = Range("Hight") - 100
    ElseIf Range("Hight") <= 165 Then
        �W���̏d = WorksheetFunction.Round((Range("Hight") - 100) * 0.9, 1)
    Else
        �W���̏d = Range("Hight") - 110
    End If
    Rst_Kiso.Fields("Aans1").Value = �W���̏d
    Range("ks2_weght") = ���̏d
    Range("ks2_weght").Offset(0, 1) = �W���̏d
'   �얞�x
    Rst_Kiso.Fields("Himanp").Value = WorksheetFunction.Round( _
                                 (���̏d - �W���̏d) / ���̏d * 100, 0)
'   �̊i�w��
    If �N�� <= 2 Then
        Wtemp = ���̏d / (Range("Hight") ^ 2) * 10 ^ 4              '���ߎw��
    ElseIf �N�� <= 12 Then
        Wtemp = ���̏d / (Range("Hight") ^ 3) * 10 ^ 7              '۰�َw��
    Else
        Wtemp = WorksheetFunction.Round(���̏d / �W���̏d * 100, 0) '��۰���w��
    End If
    Rst_Kiso.Fields("Taiis").Value = Wtemp
'   �̕\�ʐ�
    Rst_Kiso.Fields("Aansa").Value = Eiyo01_523_taihyou(���̏d)
    Rst_Kiso.Fields("Bansa").Value = Eiyo01_523_taihyou(�W���̏d)
    Range("ks2_Aansa") = Rst_Kiso.Fields("Aansa").Value
    Range("ks2_Aansa").Offset(0, 1) = Rst_Kiso.Fields("Bansa").Value
'   �����w��
    Select Case �J��
        Case "A":  Wtemp = 0.35
        Case "B":  Wtemp = 0.5
        Case "C":  Wtemp = 0.75
        Case Else: Wtemp = 1#
    End Select
    Rst_Kiso.Fields("Aansx").Value = Wtemp
    Range("ks2_Aansx") = Wtemp
'   �P�ʕ\�ʐς�����̊�b���
    If �N�� < 20 Then
        KisoCd1 = �N��
    ElseIf �N�� < 80 Then
        KisoCd1 = Int(�N�� / 10) * 10
    Else
        KisoCd1 = 80
    End If
    If Range("Sex") = 1 Then: KisoCd1 = KisoCd1 + 100
    mySqlStr = "SELECT ��b���,������� FROM " & Tbl_Need & " Where Ncode = " & KisoCd1
    Set Rst_Need = myCon.Execute(mySqlStr)
    If Rst_Need.EOF Then
        MsgBox "�K�v�ʃ}�X�^�̃L�[�Ȃ�:" & KisoCd1
    End If
    Rst_Kiso.Fields("Aans3").Value = Rst_Need.Fields("��b���").Value
    Range("ks2_kisocd") = KisoCd1
    Range("ks2_kisot") = Rst_Kiso.Fields("Aans3").Value
'   ��b���
    Rst_Kiso.Fields("Aansb").Value = WorksheetFunction.Round(Rst_Kiso.Fields("Aans3").Value _
                                                           * Rst_Kiso.Fields("Aansa").Value * 24, 2)    '���̏d�̊�b��Ӂ^��
    Rst_Kiso.Fields("Aansc").Value = WorksheetFunction.Round(Rst_Kiso.Fields("Aansb").Value / 1440, 2)  '���̏d�̊�b��Ӂ^��
    Rst_Kiso.Fields("Aansd").Value = Eiyo01_524ansd(Rst_Kiso.Fields("Aansb").Value)                     '���̏d��E�W����
    
    Rst_Kiso.Fields("Bansb").Value = WorksheetFunction.Round(Rst_Kiso.Fields("Aans3").Value _
                                                           * Rst_Kiso.Fields("Bansa").Value * 24, 2)    '�W���̏d�̊�b��Ӂ^��
    Rst_Kiso.Fields("Bansc").Value = WorksheetFunction.Round(Rst_Kiso.Fields("Bansb").Value / 1440, 2)  '�W���̏d�̊�b��Ӂ^��
    Rst_Kiso.Fields("Bansd").Value = Eiyo01_524ansd(Rst_Kiso.Fields("Bansb").Value)                     '�W���̏d��E�W����
    Range("ks2_Aansb") = Rst_Kiso.Fields("Aansb").Value                 '���̏d�̊�b��Ӂ^��
    Range("ks2_Aansc") = Rst_Kiso.Fields("Aansc").Value                 '���̏d�̊�b��Ӂ^��
    Range("ks2_Aansd") = Rst_Kiso.Fields("Aansd").Value                 '���̏d��E�W����
    Range("ks2_Aansb").Offset(0, 1) = Rst_Kiso.Fields("Bansb").Value    '�W���̏d�̊�b��Ӂ^��
    Range("ks2_Aansc").Offset(0, 1) = Rst_Kiso.Fields("Bansc").Value    '�W���̏d�̊�b��Ӂ^��
    Range("ks2_Aansd").Offset(0, 1) = Rst_Kiso.Fields("Bansd").Value    '�W���̏d��E�W����

'   ���v�ʁ@�G�l���M�[  -------------------------------------------------------------------------------------
    If Rst_Kiso.Fields("Tenes").Value = 1 Then              '�G�l���M�[�w��E�l
        Wenerg1 = Rst_Kiso.Fields("Tenee").Value
        Wenerg2 = Rst_Kiso.Fields("Aansd").Value
        Wcondition = 1
    ElseIf Rst_Kiso.Fields("Tenes").Value = 2 Then          '�G�l���M�[�w��E���̏d
        Wenerg1 = WorksheetFunction.RoundDown(Rst_Kiso.Fields("Tenee").Value * ���̏d, 2)
        Wenerg2 = Rst_Kiso.Fields("Aansd").Value
        Wcondition = 2
    ElseIf Rst_Kiso.Fields("Tenes").Value = 3 Then          '�G�l���M�[�w��E�W���̏d
        Wenerg1 = WorksheetFunction.RoundDown(Rst_Kiso.Fields("Tenee").Value * �W���̏d, 2)
        Wenerg2 = Rst_Kiso.Fields("Aansd").Value
        Wcondition = 3
    ElseIf �D�P = 1 Then                                    '�D�P�O��
        Wenerg1 = Rst_Kiso.Fields("Bansd").Value + 150
        Wenerg2 = Rst_Kiso.Fields("Aansd").Value + 150
        Wcondition = 4
    ElseIf �D�P = 2 Then                                    '�D�P���
        Wenerg1 = Rst_Kiso.Fields("Bansd").Value + 350
        Wenerg2 = Rst_Kiso.Fields("Aansd").Value + 350
        Wcondition = 5
    ElseIf �D�P = 3 Then                                    '������
        Wenerg1 = Rst_Kiso.Fields("Bansd").Value + 700
        Wenerg2 = Rst_Kiso.Fields("Aansd").Value + 700
        Wcondition = 6
    ElseIf Rst_Kiso.Fields("Qill1").Value <> 0 Then         '���A�a
        Wenerg1 = Rst_Kiso.Fields("Bansd").Value - 200
        Wenerg2 = Rst_Kiso.Fields("Aansd").Value
        Wcondition = 7
    ElseIf Rst_Kiso.Fields("Qsrmr").Value <> 0 Then         '�X�|�[�c
        Wenerg1 = Rst_Kiso.Fields("Aansb").Value _
                + WorksheetFunction.RoundDown((Rst_Kiso.Fields("Qsrmr").Value + 1.2) _
                                             * Rst_Kiso.Fields("Qsmin").Value _
                                             * Rst_Need.Fields("�������").Value _
                                             * Rst_Kiso.Fields("Aansc").Value, 2)
        Wenerg2 = Wenerg1
        Wcondition = 8
    ElseIf Rst_Kiso.Fields("Taiis").Value <= 90 Or _
           Rst_Kiso.Fields("Taiis").Value >= 120 Then      '�얞
        Wenerg1 = Rst_Kiso.Fields("Bansd").Value
        Wenerg2 = Rst_Kiso.Fields("Aansd").Value
        Wcondition = 9
    Else                                                    '���̑����
        Wenerg1 = Rst_Kiso.Fields("Aansd").Value
        Wenerg2 = Rst_Kiso.Fields("Aansd").Value
        Wcondition = 10
    End If
    Rst_Syoyo.Fields("Syoyo01").Value = Wenerg1
    Range("ks2_energ") = Wcondition
    Range("ks2_energ").Offset(1, 0) = Wenerg1
    Range("ks2_energ").Offset(2, 0) = Wenerg2
    
'   ���v�ʁ@���̑�  ------------------------------------------------------------------------------------------
    KisoCd2 = KisoCd1 + 1000
    mySqlStr = "SELECT * FROM " & Tbl_Need & " Where Ncode = " & KisoCd2
    Set Rst_Need = myCon.Execute(mySqlStr)
    If Rst_Need.EOF Then
        MsgBox "�K�v�ʃ}�X�^�̃L�[�Ȃ�:" & KisoCd2
    End If
    Range("ks2_syoyo") = KisoCd2
    For i1 = 2 To 27
        Rst_Syoyo.Fields(i1 * 5 - 1).Value = Rst_Need.Fields(i1 + 1).Value
        Range("ks2_syoyo").Offset(i1, 0) = Rst_Need.Fields(i1 + 1).Value
    Next i1
    If �J�� = "B" Or �N�� < 15 Then
    Else
        Select Case �J��
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
            MsgBox "�K�v�ʃ}�X�^�̃L�[�Ȃ�:" & KisoCd2
        End If
        Range("ks2_syoyo").Offset(0, 1) = KisoCd2
        For i1 = 2 To 27
            Rst_Syoyo.Fields(i1 * 5 - 1).Value = _
            Rst_Syoyo.Fields(i1 * 5 - 1).Value + Rst_Need.Fields(i1 + 1).Value
            Range("ks2_syoyo").Offset(i1, 1) = Rst_Need.Fields(i1 + 1).Value
        Next i1
    End If
    Select Case �D�P
        Case 1:    KisoCd3 = 1401   '�D�P�O��
        Case 2:    KisoCd3 = 1402   '�D�P���
        Case 3:    KisoCd3 = 1403   '������
        Case Else: KisoCd3 = 0
    End Select
    If KisoCd3 > 0 Then
        mySqlStr = "SELECT * FROM " & Tbl_Need & " Where Ncode = " & KisoCd3
        Set Rst_Need = myCon.Execute(mySqlStr)
        If Rst_Need.EOF Then
            MsgBox "�K�v�ʃ}�X�^�̃L�[�Ȃ�:" & KisoCd3
        End If
        Range("ks2_syoyo").Offset(0, 2) = KisoCd3
        For i1 = 2 To 27
            Rst_Syoyo.Fields(i1 * 5 - 1).Value = _
            Rst_Syoyo.Fields(i1 * 5 - 1).Value + Rst_Need.Fields(i1 + 1).Value
            Range("ks2_syoyo").Offset(i1, 2) = Rst_Need.Fields(i1 + 1).Value
        Next i1
    End If
'   ���v�ʁ@����ς���  --------------------------------------------------------------------------------------
    Wcondition = 0
    If Rst_Kiso.Fields("Tanps").Value = 1 Then              '����ς��w��E�l
        Wtemp = Rst_Kiso.Fields("Tanpe").Value
        Wcondition = 1
    ElseIf Rst_Kiso.Fields("Tanps").Value = 2 Then          '����ς��w��E���̏d
        Wtemp = WorksheetFunction.RoundDown(Rst_Kiso.Fields("Tanpe").Value * ���̏d, 2)
        Wcondition = 2
    ElseIf Rst_Kiso.Fields("Tanps").Value = 3 Then          '����ς��w��E�W���̏d
        Wtemp = WorksheetFunction.RoundDown(Rst_Kiso.Fields("Tanpe").Value * �W���̏d, 2)
        Wcondition = 3
    ElseIf Rst_Kiso.Fields("Tanps").Value = 4 Then          '����ς��w��E�G�l���M�[��
        Wtemp = WorksheetFunction.RoundDown(Rst_Kiso.Fields("Tanpe").Value _
                                          * Wenerg1 / 400, 2)
        Wcondition = 4
    ElseIf Rst_Kiso.Fields("Qsrmr").Value <> 0 Then         '�X�|�[�c
        Wtemp = ���̏d * 1.4
        Wcondition = 5
    ElseIf Rst_Kiso.Fields("Qwcnt").Value <> 0 Then         '���ĥ���۰�
        If Wenerg1 < 1412 Then
            Wtemp = 60
            Wcondition = 6
        Else
            Wtemp = Wenerg1 * 17 / 400
            Wcondition = 7
        End If
    Else
        If �N�� < 21 Then
            If Range("Sex") = 0 Then
'                          3   4   5   6   7   8   9  10  11  12  13  14  15  16  17  18  19  20��
                Wtext = "117,116,122,124,128,135,138,144,140,138,136,129,125,121,117,117,113,113"
            Else
                Wtext = "117,120,126,132,137,144,149,148,144,146,139,133,131,129,129,124,121,121"
            End If
            Warray = Split(Wtext, ",")
            Wtemp = Warray(�N�� - 3)
            If Rst_Kiso.Fields("Qsyog").Value = 1 Then: Wtemp = Wtemp + 20  '��Q��
            Wcondition = 8
        Else
            i1 = Int(�N�� / 10) - 2
            If i1 > 6 Then: i1 = 6
            Select Case �J��       ' 20  30  40  50  60  70  80      �ˑ�
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
        Select Case �D�P
            Case 1: Wtemp = Wtemp + 10  '�D�P�O��
            Case 2: Wtemp = Wtemp + 20  '�D�P���
            Case 3: Wtemp = Wtemp + 20  '������
        End Select
    End If
    Wtemp = WorksheetFunction.RoundDown(Wtemp, 2)
    Rst_Syoyo.Fields("Syoyo02").Value = WorksheetFunction.RoundDown(Wtemp, 2)       '����ς���  (02)
    Rst_Syoyo.Fields("Syoyo03").Value = WorksheetFunction.RoundDown(Wtemp / 2, 2)   '��������ς�(03)
    Rst_Syoyo.Fields("Syoyo04").Value = WorksheetFunction.RoundDown(Wtemp / 2, 2)   '�A������ς�(04)
    Range("ks2_syoyo").Offset(2, 3) = Wcondition
'   ���v�ʁ@����  --------------------------------------------------------------------------
    If �D�P > 0 Then
        Wtemp = 275
    ElseIf �N�� < 21 Then
        If �J�� = "A" Then
            Wtemp = 225
        Else
            Wtemp = 275
        End If
    Else
        Select Case �J��
            Case "A", "B": Wtemp = 225
            Case Else:     Wtemp = 275
        End Select
    End If
    Wtemp = WorksheetFunction.RoundDown(Wenerg1 * Wtemp / 9000, 2)
    Rst_Syoyo.Fields("Syoyo05").Value = Wtemp                                       '����   (05)
    Rst_Syoyo.Fields("Syoyo26").Value = WorksheetFunction.Round(Wtemp * 0.34, 2)    'S      (24)
    Rst_Syoyo.Fields("Syoyo25").Value = WorksheetFunction.Round(Wtemp * 0.66, 2)    'P      (25)
    Rst_Syoyo.Fields("Syoyo24").Value = 300                                         '�ڽ�۰�(24)
    
    Rst_Syoyo.Fields("Syoyo06").Value = WorksheetFunction.RoundDown((Wenerg1 _
                                      - Rst_Syoyo.Fields("Syoyo02").Value * 4 _
                                      - Rst_Syoyo.Fields("Syoyo05").Value * 9) / 4, 2)  '����(06)
    If Rst_Kiso.Fields("Qill1").Value <> 0 Then                                         '����(27)
        Rst_Syoyo.Fields("Syoyo27").Value = 10          '���A�a
    ElseIf Rst_Kiso.Fields("Qwcnt").Value <> 0 Then
        Rst_Syoyo.Fields("Syoyo27").Value = 10          '���ĥ���۰�
    Else
        Rst_Syoyo.Fields("Syoyo27").Value = 30          '���̑�(���)
    End If
    Rst_Syoyo.Fields("Syoyo07").Value = WorksheetFunction.RoundDown(Wenerg1 * 0.0099, 2)    '�H������(07)
'   �ټ��(08) ------------------------------------------------------------------------------------------------
    Wcondition = 0
    If �N�� < 21 Then
        If Range("Sex") = 0 Then
'                      3   4   5   6   7   8   9  10  11  12  13  14  15  16  17  18  19  20��
            Wtext = "171,168,169,173,176,165,169,177,187,188,177,156,134,124,115,109,103,103"
        Else
            Wtext = "173,169,169,174,182,177,184,184,175,158,142,133,119,108,100,100,100,100"
        End If
        Warray = Split(Wtext, ",")
        Wtemp = WorksheetFunction.RoundDown(���̏d * Warray(�N�� - 3) / 10, 2)
        Wcondition = 1
    ElseIf �N�� < 60 Then
        Wtemp = ���̏d * 10
        Wcondition = 2
    Else
        Wtemp = 600
        Wcondition = 3
    End If
    Select Case �D�P
        Case 1, 2
            Wtemp = Wtemp + 400 '�D�P�O���
            Wcondition = 4
        Case 3
            Wtemp = Wtemp + 500 '������
            Wcondition = 5
    End Select
    Rst_Syoyo.Fields("Syoyo08").Value = Wtemp   '�ټ��(08)
    Rst_Syoyo.Fields("Syoyo09").Value = Wtemp   '���� (09)
    Range("ks2_syoyo").Offset(8, 3) = Wcondition
'   �S  ------------------------------------------------------------------------------------------------------
    Select Case �D�P
        Case 1:    Wtemp = 15               '�D�P�O��
        Case 2, 3: Wtemp = 20               '�D�P����E������
        Case Else
            Select Case �N��
                Case 1 To 5:   Wtemp = 8    '    �T�ˈȉ�
                Case 6 To 8:   Wtemp = 9    ' 6�` 8��
                Case 9 To 11:  Wtemp = 10   ' 9�`11��
                Case 12 To 19: Wtemp = 12   '12�`19��
                Case 20 To 49
                    Select Case Range("Sex")
                        Case 0:    Wtemp = 10   '20�`49�˂̒j
                        Case Else: Wtemp = 12   '20�`49�˂̏�
                    End Select
                Case Else: Wtemp = 10           '50�ˈȏ�
            End Select
    End Select
    Rst_Syoyo.Fields("Syoyo10").Value = Wtemp
'   VB1/VB2/Ų���  -------------------------------------------------------------------------------------------
    Rst_Syoyo.Fields("Syoyo13").Value = WorksheetFunction.RoundDown(Wenerg2 * 0.0004, 2)
    Rst_Syoyo.Fields("Syoyo14").Value = WorksheetFunction.RoundDown(Wenerg2 * 0.00055, 2)
    Rst_Syoyo.Fields("Syoyo15").Value = WorksheetFunction.RoundDown(Wenerg2 * 0.0066, 2)
    Select Case �D�P
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
        Rst_Syoyo.Fields("Syoyo23").Value = 6       '��
        Rst_Syoyo.Fields("Syoyo11").Value = 2800    '��س�
        Range("ks2_syoyo").Offset(11, 3) = 1
        Range("ks2_syoyo").Offset(23, 3) = 1
    Else
        Rst_Syoyo.Fields("Syoyo23").Value = 10      '��
        Range("ks2_syoyo").Offset(23, 3) = 2
    End If
    Rst_Syoyo.Fields("Syoyo20").Value = WorksheetFunction.Round(Rst_Syoyo.Fields("Syoyo25").Value * 0.6, 2)     'VE
    Rst_Syoyo.Fields("Syoyo21").Value = Rst_Syoyo.Fields("Syoyo11").Value                                       '�س� <= ��س�
    Rst_Syoyo.Fields("Syoyo22").Value = WorksheetFunction.RoundDown(Rst_Syoyo.Fields("Syoyo08").Value / 2, 2)   'Mg=Ca/2
'   �h�{�f�䗦
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

'   �X�V���ʕ\��
    For i1 = 1 To Rst_Kiso.Fields.Count
        Cells(3, i1).Value = Rst_Kiso.Fields(i1 - 1).Value
    Next
    For i1 = 1 To 27
        Range("ks2_syoyo").Offset(i1, 4) = Rst_Syoyo.Fields(i1 * 5 - 1).Value
    Next i1
End Function
'--------------------------------------------------------------------------------
'   01_523 �̕\�ʐ�
'       �T�Έȉ�    �̏d^0.423 * �g��^0.362 * 382.89 / 10000
'       �U�Έȏ�    �̏d^0.444 * �g��^0.663 *  88.83 / 10000
'--------------------------------------------------------------------------------
Function Eiyo01_523_taihyou(�̏d As Double) As Double
Dim Wtemp   As Double

    If Range("Age") < 6 Then
        Wtemp = WorksheetFunction.Round(�̏d ^ 0.423 * Range("hight") ^ 0.362 * 382.89 / 10000, 2)
    Else
        Wtemp = WorksheetFunction.Round(�̏d ^ 0.444 * Range("hight") ^ 0.663 * 88.83 / 10000, 2)
    End If
    Eiyo01_523_taihyou = Wtemp
End Function
'--------------------------------------------------------------------------------
'   01_524 �G�l���M�[�W���ʁ@�����������x�␳
'       ��Q��      50%
'       �U�O�Α�    90%
'       �V�O�Α�    80%
'       �W�O�Έȏ�  70%
'       ��b��Ӂ^�� * (�␳�����������x+1) * 1.1
'--------------------------------------------------------------------------------
Function Eiyo01_524ansd(��b��� As Double) As Double
Dim Wtemp   As Double
    
    If Rst_Kiso.Fields("Qsyog").Value = 1 Then                      '��Q��
        Wtemp = WorksheetFunction.Round(Rst_Kiso.Fields("Aansx").Value * 0.5, 2)
    ElseIf Range("Age") < 60 Then                                   '�U�O�Ζ���
        Wtemp = Rst_Kiso.Fields("Aansx").Value
    ElseIf Range("Age") >= 60 And Range("Age") <= 69 Then           '�U�O�Α�
        Wtemp = WorksheetFunction.Round(Rst_Kiso.Fields("Aansx").Value * 0.9, 2)
    ElseIf Range("Age") >= 70 And Range("Age") <= 79 Then           '�V�O�Α�
        Wtemp = WorksheetFunction.Round(Rst_Kiso.Fields("Aansx").Value * 0.8, 2)
    Else                                                            '�W�O�Έȏ�
        Wtemp = WorksheetFunction.Round(Rst_Kiso.Fields("Aansx").Value * 0.7, 2)
    End If
    Wtemp = WorksheetFunction.Round(��b��� * (1 + Wtemp) * 1.1, 2)
    Eiyo01_524ansd = Wtemp
End Function
'--------------------------------------------------------------------------------
'   01_525 �ߕs���v�Z�A�A�h�o�C�X
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
                                          / Rst_Syoyo.Fields(i1 * 5 - 1).Value * 100, 0) - 100   '�ߕs����
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
'   �A�h�o�C�X
    Wtext = Rst_Kiso.Fields("Q3rec").Value      '�H�K��
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
        Wans1 = 98  '2008/4/25 3050��0098�ɕύX
    End If
    Rst_Kiso.Fields("Badv1").Value = Wans1
    
    Wtext = Rst_Kiso.Fields("Q4rec").Value      '�x�{
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
        Wans1 = 98  '2008/4/25 3150��0098�ɕύX
    End If
    Rst_Kiso.Fields("Badv2").Value = Wans1
    
    Wtext = Rst_Kiso.Fields("Q5rec").Value      '�^��
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
        Wans1 = 98  '2008/4/25 3250��0098�ɕύX
    End If
    Rst_Kiso.Fields("Badv3").Value = Wans1
    
    Wtext = Rst_Kiso.Fields("Q6r_a").Value & Rst_Kiso.Fields("Q6r_b").Value _
          & Rst_Kiso.Fields("Q6r_c").Value & Rst_Kiso.Fields("Q6r_d").Value _
          & Rst_Kiso.Fields("Q6r_e").Value                              '���N����
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
    
    If Rst_Kiso.Fields("Qsrmr").Value <> 0 Then            '  ���� ����޲�
       Wans1 = 2801
    ElseIf Rst_Kiso.Fields("age").Value <= 12 Then
       Wans1 = 2806
    ElseIf Rst_Kiso.Fields("Taiis").Value < 120 Or _
           Rst_Kiso.Fields("Qcnd1").Value = 1 Or _
           Rst_Kiso.Fields("Qcnd1").Value = 2 Then      ' 120%��� OR �ݼ�
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
    If Rst_Syoyo.Fields("Syort20").Value < -37 Then     'VE ̿�
       If Rst_Kiso.Fields("age").Value < 40 And _
          Rst_Kiso.Fields("Sex").Value = 1 Then
           Call Eiyo01_526Cadvs(3630)                   '39��  ���
       Else
           Call Eiyo01_526Cadvs(3610)                   '�ĺ & 40��ޮ� ���
       End If
    End If
    If Rst_Syoyo.Fields("Syort12") < -37 Then: Call Eiyo01_526Cadvs(3640)   'VA
    If Rst_Syoyo.Fields("Syort13") < -37 Then: Call Eiyo01_526Cadvs(3620)   'VB1
    If Rst_Syoyo.Fields("Syort08") < -37 Then: Call Eiyo01_526Cadvs(3650)   'CA
    If Rst_Syoyo.Fields("Syort07") < -37 Then: Call Eiyo01_526Cadvs(3660)   '�ݲ
    
    If Rst_Syoyo.Fields("Syort23") > 12 Then: Call Eiyo01_527Cadvs(3730)    '��
    If Rst_Energ.Fields("Enet08") < 85 Then: Call Eiyo01_527Cadvs(3720)     '3��� ��ֳ ����
    If Rst_Syoyo.Fields("Syort08") < -37 Then: Call Eiyo01_527Cadvs(3760)   'CA
    If Rst_Energ.Fields("Enet08") < 85 Or _
       Rst_Energ.Fields("Enet09") < 85 Then: Call Eiyo01_527Cadvs(3740)     '4���
    If Rst_Kiso.Fields("PER04") > 50 Then: Call Eiyo01_527Cadvs(3710)      '�޳��� ���߸�� �
    
    
'                     ---- 0.35 ----  ---- 0.5 -----
    Wtext = Empty   '<=-110=><=111-=><=-110=><=111-=>
    Wtext = Wtext & "00012023000120270001021500010211"  '   -20 �ĺ
    Wtext = Wtext & "00012028000120210001021600010214"  ' 21-30
    Wtext = Wtext & "00013335000133340001202800012029"  ' 31-40
    Wtext = Wtext & "00013337000133360001202100012025"  ' 41-50
    Wtext = Wtext & "00012025000120290001021100010214"  ' 51-60
    Wtext = Wtext & "00012030000120260001021700010218"  ' 61-70
    Wtext = Wtext & "00012032000120310001020900010219"  ' 71-
    Wtext = Wtext & "00010204000102030001021000010212"  '   -20 ���
    Wtext = Wtext & "00012021000120220001020300010211"  ' 21-30
    Wtext = Wtext & "00012024000120230001021300010212"  ' 31-40
    Wtext = Wtext & "00012026000120250001020700010214"  ' 41-50
    Wtext = Wtext & "00010205000120260001020500010208"  ' 51-60
    Wtext = Wtext & "00010206000102050001020600010209"  ' 61-70
    Wtext = Wtext & "00010209000102080001020900010208"  ' 71-
    If Rst_Kiso.Fields("Aansx") > 50 Then
        Wtext2 = "38394041"
    Else
        If Rst_Kiso.Fields("Taiis").Value < 111 Then   '���� ���
            Wans1 = 0
        Else
            Wans1 = 1
        End If
        If Rst_Kiso.Fields("Aansx") = 0.5 Then: Wans1 = Wans1 + 2        '���� ���
        Select Case Rst_Kiso.Fields("Age")
            Case 0 To 20:
            Case 21 To 30: Wans1 = Wans1 + 4
            Case 31 To 40: Wans1 = Wans1 + 8
            Case 41 To 50: Wans1 = Wans1 + 12
            Case 51 To 60: Wans1 = Wans1 + 16
            Case 61 To 70: Wans1 = Wans1 + 20
            Case Else:     Wans1 = Wans1 + 24
        End Select
        If Rst_Kiso.Fields("Sex") = 1 Then: Wans1 = Wans1 + 28        '���
        Wtext2 = Mid(Wtext, Wans1 * 8 + 1, 8)
    End If
    Rst_Kiso.Fields("Dadv1").Value = Left(Wtext2, 2)
    Rst_Kiso.Fields("Dadv2").Value = Mid(Wtext2, 3, 2)
    Rst_Kiso.Fields("Dadv3").Value = Mid(Wtext2, 5, 2)
    Rst_Kiso.Fields("Dadv4").Value = Mid(Wtext2, 7, 2)
    
    Eiyo01_525MealDiffe = 0
End Function
'--------------------------------------------------------------------------------
'   01_526�@C�A�h�o�C�X�P
'--------------------------------------------------------------------------------
Function Eiyo01_526Cadvs(advc As Long)
    If Rst_Kiso.Fields("Cadv1").Value = 98 Then
        Rst_Kiso.Fields("Cadv1").Value = advc
    ElseIf Rst_Kiso.Fields("Cadv2").Value = 98 Then
        Rst_Kiso.Fields("Cadv2").Value = advc
    End If
End Function
'--------------------------------------------------------------------------------
'   01_527�@C�A�h�o�C�X�Q
'--------------------------------------------------------------------------------
Function Eiyo01_527Cadvs(advc As Long)
    If Rst_Kiso.Fields("Cadv3").Value = 98 Then
        Rst_Kiso.Fields("Cadv3").Value = advc
    ElseIf Rst_Kiso.Fields("Cadv4").Value = 98 Then
        Rst_Kiso.Fields("Cadv4").Value = advc
    End If
End Function
'--------------------------------------------------------------------------------
'   01_528�@�h�{�䗦
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
'   01_540�@���ېH�v�Z�l�̔�r�p
'--------------------------------------------------------------------------------
Function Eiyo01_540Old_Check() As Long
Dim mySqlStr    As String
Dim i1          As Long
Dim i2          As Long
Dim Lmax1       As Long
Dim Lmax2       As Long
Dim Lmax3       As Long
Dim Errcnt      As Long

'   �X�V���ʕ\��
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
    If Errcnt > 0 Then: MsgBox "�s��v " & Errcnt
    
End Function
'--------------------------------------------------------------------------------
'   01_541�@��r
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
'   01_550�@��b���ق�Close
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
'   01_700�@�J�E���Z�����O�V�[�g��\
'--------------------------------------------------------------------------------
Function Eiyo01_700��\Click()
    
    If IsEmpty(Range("Fcode")) Or _
       Range("Fcode") <> Range("Fsave") Then
        MsgBox "��b���̌������s���Ă��܂���"
        Exit Function
    End If
    Application.ScreenUpdating = False  '��ʕ`��}�~
    Call Eiyo91DB_Open                  'DB Open
    Call Eiyo01_511MealFldgt            '���ڗv�f�擾
    Call Eiyo01_701Sheet                '��ݾ�ݸ޼�Ēǉ�
    Call Eiyo01_702DbGet                'DB Get(521)
    Call Eiyo01_703Pset                 '������ڂ̐ݒ�
    Call Eiyo01_704Advic                '�A�h�o�C�X
    Call Eiyo01_705Footer               '�R�[�h�����t�A�J���E�Z���[
    Call Eiyo920DB_Close                'DB Close
'    Call Eiyo99_�w��V�[�g�폜("DBmirror")
    Sheets("��ݾ�ݸ޼��").Select
End Function
'--------------------------------------------------------------------------------
'   01_701�@�V�[�g�ǉ�
'--------------------------------------------------------------------------------
Function Eiyo01_701Sheet()
Const ShtName = "��ݾ�ݸ޼��"
Const Eiyo01Bk = "Eiyo01_��b�ېH����.xls"
Const Eiyo02Bk = "Eiyo02_��ݾ�ݸ޼��.xls"
    
    Call Eiyo99_�w��V�[�g�폜(ShtName)
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
'   01_703�@������ڂ̐ݒ�
'--------------------------------------------------------------------------------
Function Eiyo01_703Pset()
Dim aa  As Worksheet
Dim bb  As Worksheet
Dim i1  As Long
Dim i2  As Long
Dim i3  As Long
Dim i4  As Long

    Set aa = Sheets("DBmirror")
    Set bb = Sheets("��ݾ�ݸ޼��")
'   �������E����
    bb.Range("p_date1") = Format(aa.Range("b2"), " yyyy""�N"" mm""��"" dd""������") & _
                     Format(aa.Range("b2") + aa.Range("c2") - 1, " mm""��"" dd""���܂�(") & _
                     aa.Range("c2") & "����)"
'   ����
    If aa.Range("e2") = 0 Then
        bb.Range("P_sex") = "�j"
    Else
        bb.Range("P_sex") = "��"
    End If
'
    bb.Range("P_age") = aa.Range("g2")              '�N��
    bb.Range("P_adrno") = aa.Range("k2")            '�X�֔ԍ�
    bb.Range("P_adrs1") = aa.Range("l2")            '�Z���[�P
    bb.Range("P_adrs2") = "'" & aa.Range("m2")      '�Z���[�Q
    bb.Range("P_namej") = aa.Range("d2") & "�@�l"   '����
    bb.Range("P_fcode") = aa.Range("a2")            'Fcode
    bb.Range("P_hok1") = aa.Range("at2")            '�ی��؋L��
    bb.Range("P_hok2") = aa.Range("au2")            '�ی��؂m�n
'   �̈�
    bb.Range("P_hight") = aa.Range("h2")            '�g��
    bb.Range("P_weght") = aa.Range("i2")            '�̏d
    If aa.Range("g2") > 12 Or _
       aa.Range("ad2") = 0 Or _
       aa.Range("ag2") = 0 Then
        bb.Range("P_aans1") = aa.Range("bl2")       '�W���̏d
    Else
        bb.Range("P_aans1") = Empty
    End If
    If aa.Range("j2") = 0 Then                      '�牺���b
        bb.Range("P_sibou") = Empty
    Else
        bb.Range("P_sibou") = aa.Range("j2")
    End If
    If aa.Range("ad2") = 0 And aa.Range("ag2") = 0 Then '�D�P/��߰�
        bb.Range("P_taii") = aa.Range("bx2")        '�̈ʎw��
        If aa.Range("g2") < 3 Then
            bb.Range("P_tsisu") = "(�J�E�v�w��)"    '�Q�ˈȉ�
        ElseIf aa.Range("g2") < 13 Then
            bb.Range("P_tsisu") = "(���[�����w��)"  '�R�`�P�Q��
        ElseIf aa.Range("i2") < 150 Then
            bb.Range("P_tsisu") = "(�u���[�J�[�w���ϖ@)"  '�g��150cm����
        Else
            bb.Range("P_tsisu") = "(�u���[�J�[�w���ϖ@)"  '�g��150cm�ȏ�
        End If
    Else
        bb.Range("P_taii") = Empty
        bb.Range("P_tsisu") = Empty
    End If
'   �H�i�ێ�o�����X
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
'   ���t����
    bb.Range("P_bdate") = aa.Cells(2, 50)
    For i1 = 1 To 12
        bb.Range("P_bbl01").Offset(i1 - 1, 0) = aa.Cells(2, i1 + 50)
    Next i1
'   �h�{�f�ێ�o�����X  i1:�h�{�fIndex 1�`27  i2:�sIndex 1�`24
    For i1 = 1 To 27
        i2 = Fld_Field(i1, 24)
        If i2 > 0 Then
            bb.Range("i44").Offset(i2, 0) = aa.Range("e7").Offset(0, (i1 * 5 - 5))
            bb.Range("l44").Offset(i2, 0) = aa.Range("d7").Offset(0, (i1 * 5 - 5))
            bb.Range("o44").Offset(i2, 0) = bb.Range("l44").Offset(i2, 0) - bb.Range("i44").Offset(i2, 0)
            i3 = aa.Range("f7").Offset(0, (i1 * 5 - 5))
            If i3 > -62.5 Then
'                bb.Range("r44").Offset(i2, 0) = String((i3 + 62.5) * 52 / 125, "*")
                bb.Range("r44").Offset(i2, 0) = String((i3 + 62.5) * 52 / 250, "��")
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
'   �h�{�f�䗦
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
'   01_704�@�A�h�o�C�X���ڂ̐ݒ�
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
'       �E�G�C�g�E�A�h�o�C�X
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
    
'   ����
    If Sheets("DBmirror").Range("ck2") = 3330 Then
        Wadvic2(4) = Empty
        For i1 = 3 To 10
            If Mid(Range("q2"), i1, 1) = "1" Then
                Select Case i1
                    Case 3: Wadvic2(4) = Wadvic2(4) & "�����w���@�ċz�@�\�@"
                    Case 4: Wadvic2(4) = Wadvic2(4) & "�S�d�}�@�����@�����@"
                    Case 5: Wadvic2(4) = Wadvic2(4) & "������n�@"
                    Case 6: Wadvic2(4) = Wadvic2(4) & "���t�@"
                    Case 7: Wadvic2(4) = Wadvic2(4) & "���@"
                    Case 8: Wadvic2(4) = Wadvic2(4) & "�̋@�\�@"
                    Case 9: Wadvic2(4) = Wadvic2(4) & "���@��A�@"
                    Case 10: Wadvic2(4) = Wadvic2(4) & "��ȁ@"
                End Select
            End If
        Next i1
    End If
    
    Sheets("��ݾ�ݸ޼��").Range("an13") = Left(Wadvic1(1), 18)      '�H�����^�K��
    Sheets("��ݾ�ݸ޼��").Range("an14") = Mid(Wadvic1(1), 19, 18)
    Sheets("��ݾ�ݸ޼��").Range("an15") = Mid(Wadvic1(1), 37, 18)
    Sheets("��ݾ�ݸ޼��").Range("aj16") = Left(Wadvic2(1), 22)
    Sheets("��ݾ�ݸ޼��").Range("aj17") = Mid(Wadvic2(1), 23, 22)
    Sheets("��ݾ�ݸ޼��").Range("an20") = Left(Wadvic1(2), 18)      '�����Ƌx�{
    Sheets("��ݾ�ݸ޼��").Range("an21") = Mid(Wadvic1(2), 19, 18)
    Sheets("��ݾ�ݸ޼��").Range("an22") = Mid(Wadvic1(2), 37, 18)
    Sheets("��ݾ�ݸ޼��").Range("aj23") = Left(Wadvic2(2), 22)
    Sheets("��ݾ�ݸ޼��").Range("aj24") = Mid(Wadvic2(2), 23, 22)
    Sheets("��ݾ�ݸ޼��").Range("an27") = Left(Wadvic1(3), 18)      '�^��
    Sheets("��ݾ�ݸ޼��").Range("an28") = Mid(Wadvic1(3), 19, 18)
    Sheets("��ݾ�ݸ޼��").Range("an29") = Mid(Wadvic1(3), 37, 18)
    Sheets("��ݾ�ݸ޼��").Range("aj30") = Left(Wadvic2(3), 22)
    Sheets("��ݾ�ݸ޼��").Range("aj31") = Mid(Wadvic2(3), 23, 22)
    Sheets("��ݾ�ݸ޼��").Range("an34") = Left(Wadvic1(4), 18)      '���N���
    Sheets("��ݾ�ݸ޼��").Range("an35") = Mid(Wadvic1(4), 19, 18)
    Sheets("��ݾ�ݸ޼��").Range("an36") = Mid(Wadvic1(4), 37, 18)
    Sheets("��ݾ�ݸ޼��").Range("aj37") = Left(Wadvic1(5), 22)
    Sheets("��ݾ�ݸ޼��").Range("aj38") = Mid(Wadvic1(5), 23, 22)
    
    Sheets("��ݾ�ݸ޼��").Range("bc17") = Wadvic3(1)
    Sheets("��ݾ�ݸ޼��").Range("bc18") = Wadvic3(2)
    Sheets("��ݾ�ݸ޼��").Range("bc19") = Wadvic3(3)
    Sheets("��ݾ�ݸ޼��").Range("bc20") = Wadvic3(4)
    Sheets("��ݾ�ݸ޼��").Range("bc21") = Wadvic3(5)
    
    Sheets("��ݾ�ݸ޼��").Range("bc63") = Left(Wadvic1(6), 18)
    Sheets("��ݾ�ݸ޼��").Range("bc64") = Mid(Wadvic1(6), 19, 18)
    Sheets("��ݾ�ݸ޼��").Range("bc65") = Left(Wadvic2(6), 18)
    Sheets("��ݾ�ݸ޼��").Range("bc66") = Mid(Wadvic2(6), 19, 18)
    
    Sheets("��ݾ�ݸ޼��").Range("u73") = Wadvic1(12)
    Sheets("��ݾ�ݸ޼��").Range("u74") = Wadvic2(12)
    Sheets("��ݾ�ݸ޼��").Range("u75") = Wadvic1(13)
    Sheets("��ݾ�ݸ޼��").Range("u76") = Wadvic2(13)
End Function
'--------------------------------------------------------------------------------
'   01_705�@�R�[�h�����t
'--------------------------------------------------------------------------------
Function Eiyo01_705Footer()
Dim Wtext   As String
    Wtext = "(" & Sheets("��b").Range("g3") & ":" & Format(Date, "yymmdd") & ")"
    Sheets("��ݾ�ݸ޼��").Range("b80") = Wtext
    Sheets("��ݾ�ݸ޼��").Range("bd75") = Sheets("DBmirror").Range("db2")
    Sheets("��ݾ�ݸ޼��").Range("bd76") = Sheets("DBmirror").Range("dc2")
    Sheets("��ݾ�ݸ޼��").Range("bd77") = Sheets("DBmirror").Range("dd2")
End Function
'--------------------------------------------------------------------------------
'   01_810�@��b��ʍ쐬
'--------------------------------------------------------------------------------
Function Eiyo01_810��b��ʍ쐬()
Const PgmName = "Eiyo01_��b�ېH����.xls"
Const ShtName = "��b"
Dim i1      As Long
Dim i2      As Long
Dim FldItem As Variant

    If ActiveWorkbook.Name <> PgmName Then
        MsgBox PgmName & " �ł͂���܂���"
        End
    End If
    If ActiveSheet.Name <> ShtName Then
        MsgBox ShtName & " �ł͂���܂���"
        End
    End If
    Call Eiyo01_000init
'   ��ʂ̍쐬
    Call Eiyo930Screen_Hold                 '��ʗ}�~�ق�
    While (ActiveSheet.Shapes.Count > 0)    '�R�}���h�{�^�����
        ActiveSheet.Shapes(1).Cut
    Wend
    Cells.Delete Shift:=xlUp                '�S����
    Cells.NumberFormatLocal = "@"           '�S��ʕ����񑮐�
    Cells.Select
    With Selection.Font                     '�����t�H���g
        .Name = "�l�r �S�V�b�N"
        .Size = 11
    End With
    Selection.ColumnWidth = 1.75            '��
    Selection.Interior.ColorIndex = 40      '�S��ʔw�i�F�i�W���j
    
'   �\��
    Range("G1:AA1").Select
    Selection.MergeCells = True                 '�\��Z���A��
    Selection.HorizontalAlignment = xlCenter    '�\��Z���^�����O
    Selection.Interior.ColorIndex = 37          '�\��F�i�y�[���u���[�j
    With Selection.Font                         '�t�H���g
        .FontStyle = "����"
        .Size = 16
    End With
    Range("G01") = "�h�{�v�Z�i��b�j�Q�V�h�{�f��"
    Range("A01").VerticalAlignment = xlTop
    Range("A01") = "v-01"
    Range("A03") = "�R�[�h"
    Range("A04") = "��������"
    Range("A05") = "����"
    Range("A06") = "����"
    Range("i06") = "(0:�j 1:��)"
    Range("A07") = "���N����"
    Range("a08") = "�g��"
    Range("j08") = "cm"
    Range("a09") = "�̏d"
    Range("j09") = "Kg"
    Range("a10") = "�牺���b"
    Range("j10") = "cm"
    Range("A11") = "�X�֔ԍ�"
    Range("A12") = "�Z���[�P"
    Range("A13") = "�Z���[�Q"
    Range("A14") = "�n��"
    Range("a15") = "�s���{��"
    Range("a16") = "3.�H�K"
    Range("a17") = "4.�x�{"
    Range("a18") = "5.�^��"
    Range("a19") = "6.���N"
    Range("m08") = "7.�E��"
    Range("m09") = "A.��w"
    Range("m10") = "B.�D�P"
    Range("m11") = "C.���A"
    Range("m14") = "D.������"
    Range("m15") = "E.��߰�"
    Range("m16") = "F.�^����"
    Range("m17") = "*.�i��"
    Range("m18") = "G.�g��Q"
    Range("m19") = "H.����CT"
    Range("m20") = "��ٷް�w��"
    Range("M20").Characters(Start:=7, Length:=2).Font.Size = 9
    Range("m21") = "���߸ �w��"
    Range("M21").Characters(Start:=7, Length:=2).Font.Size = 9
    Range("m22") = "(0:�� 1:�w�� 2:���̏d 3:�W���̏d)"
    Range("m23") = "��ݾװ1"
    Range("m24") = "��ݾװ2"
    Range("m25") = "��ݾװ3"
    
    Range("x03") = "���t�^"
    Range("x04") = "�x�Е�CD"
    Range("x05") = "�ی��L��"
    Range("x06") = "      No"
    Range("x07") = "������f"
    Range("x08") = "�r(L,R)"
    Range("x09") = "������"
    Range("x10") = "�Ԍ�����"
    Range("x11") = "���F�f��"
    Range("x12") = "��ĸد�"
    Range("x13") = "�ڽ�۰�"
    Range("x14") = "HDL"
    Range("x15") = "�������b"
    Range("x16") = "G.O.T."
    Range("x17") = "G.P.T."
    Range("x18") = "�A�_"
    Range("x19") = "����"
    Range("x20") = "�����ō�"
    Range("x21") = "    �Œ�"
    
    Cells.Locked = True                             '�S�Z�������b�N
    For i1 = 0 To UBound(Fld_Adrs1)
        FldItem = Split(Fld_Adrs1(i1), ",")
        Range(Trim(FldItem(1))).Select
        Selection.MergeCells = True                 '�Z������
        Range(Left(FldItem(1), 4)).Name = Trim(FldItem(0))
        If FldItem(2) = "i" Then
            With Selection.Borders                      '���͍��ڂ̘g�r��
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .Weight = xlThin
            End With
            Selection.Interior.ColorIndex = xlNone      '���͍��ڂ̔�������
            Selection.Locked = False                    '���͍��ڂ̕ی����
        End If
        Select Case FldItem(4)
            Case "90": Selection.NumberFormatLocal = "G/�W��"
            Case "91": Selection.NumberFormatLocal = "#0.0"
            Case "92": Selection.NumberFormatLocal = "#0.00"
            Case "Ds"
                Selection.NumberFormatLocal = "yyyy/mm/dd"
                Selection.HorizontalAlignment = xlLeft
            Case "Dw"
                Selection.NumberFormatLocal = "gy.m.d"
                Selection.HorizontalAlignment = xlLeft
            Case "J "
                With Selection.Validation           '��������
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
    Range("Gyyyy").NumberFormatLocal = "gy"     '���N�����̘a��N�\��
    Range("Gyyyy") = "=RC[-5]"
    Range("Age").NumberFormatLocal = "G/�W��"
    Range("Age") = "=DATEDIF(RC[-7],R[-3]C[-7],""y"")"
    Range("p07") = "��"

    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=10, Top:=350, Width:=60, Height:=30)
        .Object.Caption = "��ʏ���"
        .Name = "�N���A"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=100, Top:=350, Width:=60, Height:=30)
        .Object.Caption = "�ް��ďo"
        .Name = "����"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=170, Top:=350, Width:=60, Height:=30)
        .Object.Caption = "�ް��o�^"
        .Name = "�X�V"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=240, Top:=350, Width:=60, Height:=30)
        .Object.Caption = "��ݾ�ݸ�" & vbLf & "��č�\"
        .Name = "��\"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=330, Top:=450, Width:=60, Height:=30)
        .Object.Caption = "�ް����"
        .Name = "���"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=400, Top:=350, Width:=60, Height:=30)
        .Object.Caption = "�I��"
        .Name = "�I��"
    End With
    
    Range("Gmesg").Font.Bold = True                            '���b�Z�[�W�G���A
    Range("Gmesg").Font.ColorIndex = 3
    Cells.FormatConditions.Delete               '�V�[�g�S�̂�������t���������폜����
    Cells.FormatConditions.Add Type:=xlExpression, Formula1:="=AND(CELL(""row"")=ROW(),CELL(""col"")=COLUMN())"
    Cells.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Cells.FormatConditions(1).Interior.Color = 255
    
    Call Eiyo01_820����K�C�h
    Range("g4").Select
    Call Eiyo940Screen_Start    '��ʕ`��ق�
End Function
'--------------------------------------------------------------------------------
'   01_820 ����K�C�h
'--------------------------------------------------------------------------------
Function Eiyo01_820����K�C�h()
    Call Eiyo930Screen_Hold     '��ʗ}�~�ق�
    Columns("ah:hz").Delete Shift:=xlToLeft
    Range("ah01") = "1.�l����������"
    Range("ah02") = "�@��ʂ̂��Âꂩ�̍��ڂ�"
    Range("ah03") = "�@���͌�A�u�����v���������Ă��������B"
    Range("ah04") = "�@�����ȂǕ����Y���҂̏ꍇ�́A�E���̈ꗗ����"
    Range("ah05") = "�@�R�[�h���_�u���N���b�N���đI�����܂��B"
    Range("ah07") = "�@�u�����v�͌����w�O����v�x�ł��A"
    Range("ah08") = "�@�擪��[%]��t����Ɓw�܂ށx�ɂȂ�܂��B"
    Range("ah10") = "2.�l��o�^����"
    Range("ah11") = "�@��ʂ̊e���ڂ���͂�"
    Range("ah12") = "�@�u�X�V�v���������Ă��������B"
    Range("ah14") = "3.�l�̕ύX�E���"
    Range("ah15") = "�@�l���������A"
    Range("ah16") = "�@�C����Ɂu�X�V�v�܂��́u����v���������Ă��������B"
    Range("ah18") = "4.�ېH�̓o�^"
    Range("ah19") = "�@�l�̓o�^�܂��͏Ɖ��u�ېH�v�V�[�g�ɐ؂�ւ��Ă�������"
    Call Eiyo940Screen_Start    '��ʕ`��ق�
End Function
'--------------------------------------------------------------------------------
'   01_830�@�ېH��ʍ쐬
'--------------------------------------------------------------------------------
Function Eiyo01_830�ېH��ʍ쐬()
Const PgmName = "Eiyo01_��b�ېH����.xls"
Const ShtName = "�ېH"
Dim i1      As Long
Dim i2      As Long
Dim FldItem As Variant

    If ActiveWorkbook.Name <> PgmName Then
        MsgBox PgmName & " �ł͂���܂���"
        End
    End If
    If ActiveSheet.Name <> ShtName Then
        MsgBox ShtName & " �ł͂���܂���"
        End
    End If
    Call Eiyo01_000init
'   ��ʂ̍쐬
    Call Eiyo930Screen_Hold     '��ʗ}�~�ق�
    ActiveWindow.FreezePanes = False        '�E�C���h�g�Œ�̉���
    
    While (ActiveSheet.Shapes.Count > 0)    '�R�}���h�{�^�����
        ActiveSheet.Shapes(1).Cut
    Wend
    Cells.Delete Shift:=xlUp                '�S����
    Cells.Select
    With Selection.Font                     '�����t�H���g
        .Name = "�l�r �S�V�b�N"
        .Size = 11
    End With
    Cells.Interior.ColorIndex = 34          '�S��ʔw�i�F�i�W�΁j
    Columns("A:B").Interior.ColorIndex = xlNone
    Columns("d").Interior.ColorIndex = xlNone
    Columns("f").Interior.ColorIndex = xlNone
    Rows("1:4").Interior.ColorIndex = 34       '�S��ʔw�i�F�i�W�΁j
    Cells.Select                               '�r��
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

'   �\��
    With Range("d1").Font                         '�t�H���g
        .FontStyle = "����"
        .Size = 16
    End With
    Range("D01") = "�h�{�v�Z�i�ېH�j"
    Range("A01").VerticalAlignment = xlTop
    Range("A01") = "v-01"
'    Range("a2") = "=Fcode & ":" & Namej
    Columns("A:A").NumberFormatLocal = "yyyy/mm/dd"
    Columns("F:F").NumberFormatLocal = "0.0 "
    Range("A4") = "�ېH��"
    Range("B4") = "�H�敪"
    Range("D4") = "�H�iCD"
    Range("E4") = "�i���E�ޗ��@�������"
    Range("F4") = "�ێ��"
    Range("k3") = "�E�H�iCD�����_�u���N���b�N����ƃ��j���[�ɕς��܂�"
    Range("k2") = "�E�ێ�ʂ��[���̍s�͍폜����܂�"
    Range("k3") = "�E�ǉ��͍ŏI�s�̌��ɓ��͂��Ă�������"
    Range("k1:k2").Font.Size = 9
    Call Eiyo01_840�H�i�}�X�^(4, 6)
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
    
    Cells.Locked = True                             '�S�Z�������b�N
    Range("A:B,D:D,F:F").Locked = False             '���͗�̉���
    Rows("1:3").Locked = True                       '�\��s�̃��b�N
    
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=300, Top:=5, Width:=50, Height:=25)
        .Object.Caption = "�o�^"
        .Name = "�o�^"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=650, Top:=5, Width:=50, Height:=25)
        .Object.Caption = "����"
        .Name = "����"
    End With
    
    Range("a3").Font.Bold = True                    '���b�Z�[�W�G���A
    Range("a3").Font.ColorIndex = 3
    Range("E5").Select
    ActiveWindow.FreezePanes = True                 '�E�C���h�g�Œ�̐ݒ�
    
'    ActiveSheet.Protect UserInterfaceOnly:=True     '�ی��L���ɂ���
    Range("g4").Select
    Call Eiyo940Screen_Start                        '��ʕ`��ق�
End Function
'--------------------------------------------------------------------------------
'   01_840�@�H�i�}�X�^����
'--------------------------------------------------------------------------------
Function Eiyo01_840�H�i�}�X�^(il As Long, ic As Long)
Dim Wtext   As String
Dim Warray  As Variant
Dim i1      As Long

    Wtext = Empty
    Wtext = Wtext & vbLf & "�R�[�h"
    Wtext = Wtext & vbLf & "�H�i��"
    Wtext = Wtext & vbLf & "�ǂ݁i���ށj"
    Wtext = Wtext & vbLf & "�P��"           '���͒P��
    Wtext = Wtext & vbLf & "�R�����g"
    Wtext = Wtext & vbLf & "�o�^�P��"
    Wtext = Wtext & vbLf & "���Z�W��"
    Wtext = Wtext & vbLf & "�ƭ��ʒu�P"
    Wtext = Wtext & vbLf & "�ƭ��ʒu�Q"
    Wtext = Wtext & vbLf & "�H��"           '0:�H 1:�� 2:�������
    Wtext = Wtext & vbLf & "�ېH�͈͉���"
    Wtext = Wtext & vbLf & "�ېH�͈͏��"
    Wtext = Wtext & vbLf & "�h�{�f-01"
    Wtext = Wtext & vbLf & "�h�{�f-02"
    Wtext = Wtext & vbLf & "�h�{�f-03"
    Wtext = Wtext & vbLf & "�h�{�f-04"
    Wtext = Wtext & vbLf & "�h�{�f-05"
    Wtext = Wtext & vbLf & "�h�{�f-06"
    Wtext = Wtext & vbLf & "�h�{�f-07"
    Wtext = Wtext & vbLf & "�h�{�f-08"
    Wtext = Wtext & vbLf & "�h�{�f-09"
    Wtext = Wtext & vbLf & "�h�{�f-10"
    Wtext = Wtext & vbLf & "�h�{�f-11"
    Wtext = Wtext & vbLf & "�h�{�f-12"
    Wtext = Wtext & vbLf & "�h�{�f-13"
    Wtext = Wtext & vbLf & "�h�{�f-14"
    Wtext = Wtext & vbLf & "�h�{�f-15"
    Wtext = Wtext & vbLf & "�h�{�f-16"
    Wtext = Wtext & vbLf & "�h�{�f-17"
    Wtext = Wtext & vbLf & "�h�{�f-18"
    Wtext = Wtext & vbLf & "�h�{�f-19"
    Wtext = Wtext & vbLf & "�h�{�f-20"
    Wtext = Wtext & vbLf & "�h�{�f-21"
    Wtext = Wtext & vbLf & "�h�{�f-22"
    Wtext = Wtext & vbLf & "�h�{�f-23"
    Wtext = Wtext & vbLf & "�h�{�f-24"
    Wtext = Wtext & vbLf & "�h�{�f-25"
    Wtext = Wtext & vbLf & "�h�{�f-26"
    Wtext = Wtext & vbLf & "�h�{�f-27"
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
    Wtext = Wtext & vbLf & "�����E����"
    Wtext = Wtext & vbLf & "�����E����"
    Wtext = Wtext & vbLf & "�����E�A��"
    Warray = Split(Wtext, vbLf)
    For i1 = 1 To UBound(Warray)
        Cells(il, ic + i1) = Warray(i1)
    Next i1
End Function
'--------------------------------------------------------------------------------
'   01_900  ������Ԃɂc�a�̃R�s�[���Ƃ�B
'           �^�C���X�^���v�͑O���̍ŏI�X�V�����ƂȂ�
'--------------------------------------------------------------------------------
Function Eiyo01_900WorkbookOpen()
Dim F_name          As String   '���������t�@�C����
Dim F_dbname_today  As String   'DB+�{��
Dim F_dbname_min    As String   'DB+00000000
Dim F_dbname_max    As String   'DB+2�T�ԑO
Dim W_path          As String

    W_path = ThisWorkbook.Path & "BackUp"""
    F_dbname_today = W_path & "Eiyo_" & Format(Date, "yyyymmdd") & ".mdb"""
    F_dbname_min = "Eiyo_00000000.mdb"
    F_dbname_max = "Eiyo_" & Format(Date - 14, "yyyymmdd") & ".mdb"
    
    SetCurrentDirectory (W_path)            'Dir�ύX
    If Dir(F_dbname_today) = "" Then        '�����̕ۑ��t�@�C�������݂��Ȃ�
        F_name = Dir("*", vbNormal)
        Do While F_name <> ""
            If (F_name > F_dbname_min And F_name < F_dbname_max) Then
               Kill F_name
            End If
            F_name = Dir                    ' ���̃t�H���_����Ԃ��܂��B
        Loop
        FileCopy ThisWorkbook.Path & myFileName, F_dbname_today
    End If
End Function
'--------------------------------------------------------------------------------
'   03_030  �N���A�̃{�^���E�N���b�N
'--------------------------------------------------------------------------------
Function Eiyo03_030�N���AClick()
    Call Eiyo930Screen_Hold     '��ʗ}�~�ق�
    Range("b3:b11") = Empty
    Range("b12") = Empty
    Range("b13:b14") = Empty
    Range("g4:g30").ClearContents
    Range("j4:l18").ClearContents
    Range("j22:j24").ClearContents
    Range("a17") = Empty
    Columns("n:hz").Delete Shift:=xlToLeft
    
    Range("b3").Select
    Call Eiyo940Screen_Start    '��ʕ`��ق�
End Function
'--------------------------------------------------------------------------------
'   03_100  ����_Click
'--------------------------------------------------------------------------------
Function Eiyo03_100����Click()
Dim Wsql    As String
Dim i1      As Long

    Range("a17") = Empty
    For i1 = 3 To 14
        If Not IsEmpty(Cells(i1, 2)) Then: Exit For
    Next i1
    If i1 > 14 Then
        Range("a17") = "�L�[������܂���"
        Exit Function
    End If
    Call Eiyo930Screen_Hold     '��ʗ}�~�ق�
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
        Range("a17") = "�Y���f�[�^�͂���܂���"
    Else
        With Rst_Food
            Range("n2").CopyFromRecordset Rst_Food  '���R�[�h
            If IsEmpty(Range("n3")) Then            '�Y�����P���̂Ƃ�
                For i1 = 1 To 12                    '��ʍ��ڂ̏�������
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
                For i1 = 1 To .Fields.Count                     '�t�B�[���h��
                    Cells(1, i1 + 13).Value = .Fields(i1 - 1).Name
                Next
                Columns("n:hz").EntireColumn.AutoFit           '��
                i1 = Range("n1").End(xlDown).Row
                Range("N:N").Locked = False                     '���͗�̉���
            End If
            .Close
        End With
    End If
    Set Rst_Food = Nothing              '�I�u�W�F�N�g�̉��
    Call Eiyo920DB_Close                'DB Close
    Columns("J:L").EntireColumn.AutoFit
    Call Eiyo940Screen_Start            '��ʕ`��ق�
End Function
'--------------------------------------------------------------------------------
'   03_110  �O����_Click
'--------------------------------------------------------------------------------
Function Eiyo03_110�O����Click()
Dim Wsql    As String
Dim Wkey    As Long

    Range("a17") = Empty
    Wkey = Range("b03")
    Call Eiyo930Screen_Hold             '��ʗ}�~�ق�
    Call Eiyo91DB_Open                  'DB Open
    Wsql = "SELECT Foodc FROM " & Tbl_Food & " Where Foodc < " & Wkey & " Order by Foodc DESC"
    Set Rst_Food = myCon.Execute(Wsql)
    If Rst_Food.EOF Then
        Range("a17") = "�Y���f�[�^�͂���܂���"
    Else
        With Rst_Food
            Range("b03") = .Fields(0).Value
            .Close
        End With
    End If
    Set Rst_Food = Nothing      '�I�u�W�F�N�g�̉��
    Call Eiyo920DB_Close        'DB Close
    Call Eiyo03_100����Click
End Function
'--------------------------------------------------------------------------------
'   03_120  ������_Click
'--------------------------------------------------------------------------------
Function Eiyo03_120������Click()
Dim Wsql    As String
Dim Wkey    As Long

    Range("a17") = Empty
    Wkey = Range("b03")
    Call Eiyo930Screen_Hold             '��ʗ}�~�ق�
    Call Eiyo91DB_Open                  'DB Open
    Wsql = "SELECT Foodc FROM " & Tbl_Food & " Where Foodc > " & Wkey & " Order by Foodc"
    Set Rst_Food = myCon.Execute(Wsql)
    If Rst_Food.EOF Then
        Range("a17") = "�Y���f�[�^�͂���܂���"
    Else
        With Rst_Food
            Range("b03") = .Fields(0).Value
            .Close
        End With
    End If
    Set Rst_Food = Nothing      '�I�u�W�F�N�g�̉��
    Call Eiyo920DB_Close        'DB Close
    Call Eiyo03_100����Click
End Function
'--------------------------------------------------------------------------------
'   03_200  �X�V
'--------------------------------------------------------------------------------
Function Eiyo03_200�X�VClick()
Dim i1      As Long
Dim Wsw     As Long
Dim Wtemp   As Variant

    Wsw = 0
    Range("a17") = Empty
    If Range("b3") < 1 Then: Exit Function
    Call Eiyo91DB_Open                      'DB Open
    '���������܂�
    With Rst_Food
        '�C���f�b�N�X�̐ݒ�
        .Index = "PrimaryKey"
        '���R�[�h�Z�b�g���J��
        Rst_Food.Open Source:=Tbl_Food, ActiveConnection:=myCon, _
            CursorType:=adOpenKeyset, LockType:=adLockOptimistic, _
            Options:=adCmdTableDirect
        '�ԍ����o�^����Ă��邩��������
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
            Range("a17") = "�ǉ�����܂���"
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
                Range("a17") = "�ύX���ڂ�����܂���"
            Else
                If MsgBox("�ύX���ڂ�" & Wsw & "�����ł��A�X�V���Ă�낵���ł���", vbOKCancel) = vbOK Then
                    .Update
                    Range("a17") = "�X�V����܂���"
                End If
            End If
        End If
'        .Close
    End With
    Set Rst_Food = Nothing      '�I�u�W�F�N�g�̉��
    Call Eiyo920DB_Close        'DB Close
End Function
'--------------------------------------------------------------------------------
'   03_300  ���
'--------------------------------------------------------------------------------
Function Eiyo03_300���Click()
    Range("a17") = Empty
    Call Eiyo91DB_Open      'DB Open
    '���������܂�
    With Rst_Food
        '�C���f�b�N�X�̐ݒ�
        .Index = "PrimaryKey"
        '���R�[�h�Z�b�g���J��
        Rst_Food.Open Source:=Tbl_Food, ActiveConnection:=myCon, _
            CursorType:=adOpenKeyset, LockType:=adLockOptimistic, _
            Options:=adCmdTableDirect
        '�ԍ����o�^����Ă��邩��������
        If Not .EOF Then .Seek Range("b3")
        If .EOF Then
            Range("a17") = "�L�[�����݂��܂���"
        Else
            If MsgBox("�폜���Ă�낵���ł���", vbOKCancel) = vbOK Then
                .Delete
                Range("a17") = "�������܂���"
            End If
        End If
        .Close
    End With
    Set Rst_Food = Nothing      '�I�u�W�F�N�g�̉��
    Call Eiyo920DB_Close        'DB Close
End Function
'--------------------------------------------------------------------------------
'   03_400  �H�i�R�[�h�f�[�^�̍쐬
'--------------------------------------------------------------------------------
Function Eiyo03_400�R�[�hClick()
Dim i1      As Long
Dim Wsql    As String

    Call Eiyo930Screen_Hold
    Call Eiyo91DB_Open      'DB Open
    
    Wsql = "SELECT Foodc,Fname FROM " & Tbl_Food & " Order by Foodc"
    Set Rst_Food = myCon.Execute(Wsql)
    If Rst_Food.EOF Then
        Range("a17") = "�Y���f�[�^�͂���܂���"
    Else
        With Rst_Food
            Range("AA1").CopyFromRecordset Rst_Food
            .Close
        End With
    End If
    Set Rst_Food = Nothing      '�I�u�W�F�N�g�̉��
    Call Eiyo920DB_Close        'DB Close
    
    Open ThisWorkbook.Path & "�H�i�R�[�h.txt" For Output As #22
    For i1 = 1 To Range("aa60000").End(xlUp).Row
        Print #2, Cells(i1, 27) & vbTab & Cells(i1, 28)
    Next i1
    Close
    Columns("Aa:BZ").Delete Shift:=xlToLeft
    Call Eiyo940Screen_Start
End Function
'--------------------------------------------------------------------------------
'   03_810  �V�[�g�̍쐬
'--------------------------------------------------------------------------------
Function Eiyo03_810�H�isheet_make()
    Call Eiyo930Screen_Hold     '��ʗ}�~�ق�
    Call Eiyo03_811_init        '�V�[�g�̏�����
    Call Eiyo03_812_zokusei     '�L�[�A���́A����
    Call Eiyo03_813_eiyoso      '�h�{�f
    Call Eiyo03_814_shokugun    '�H�i�Q
    Call Eiyo03_815_sisitu      '����
    Call Eiyo03_816_keisen      '�r���A��
    Call Eiyo03_817_button      '�R�}���h�E�{�^��
    Call Eiyo940Screen_Start    '��ʕ`��ق�
End Function
'--------------------------------------------------------------------------------
'   03_811  �V�[�g�̏�����
'--------------------------------------------------------------------------------
Function Eiyo03_811_init()
    Sheets("�H�i�}�X�^").Select
    While (ActiveSheet.Shapes.Count > 0)    '�R�}���h�{�^�����
        ActiveSheet.Shapes(1).Cut
    Wend
    Cells.Delete Shift:=xlUp                '�S����
    Cells.NumberFormatLocal = "@"           '�S��ʕ����񑮐�
    Range("e1") = "���@�H�i�}�X�^�@�Ɖ�E�X�V�@��"
    Range("e1").Font.Size = 16
End Function
'--------------------------------------------------------------------------------
'   03_812  �L�[�A���́A����
'--------------------------------------------------------------------------------
Function Eiyo03_812_zokusei()
Dim i1      As Long
Dim Wtext   As String
Dim Warray  As Variant
'   �����ق�
    Wtext = Empty
    Wtext = Wtext & vbLf & "�H�i�R�[�h"
    Wtext = Wtext & vbLf & "�H�i��"
    Wtext = Wtext & vbLf & "�ǂ�(���ށj"
    Wtext = Wtext & vbLf & "���͒P��"
    Wtext = Wtext & vbLf & "�E�v"
    Wtext = Wtext & vbLf & "�o�^�P��"
    Wtext = Wtext & vbLf & "�P�ʊ��Z�l"
    Wtext = Wtext & vbLf & "�ƭ��ʒu�P"
    Wtext = Wtext & vbLf & "�ƭ��ʒu�Q"
    Wtext = Wtext & vbLf & "�H�i����"
    Wtext = Wtext & vbLf & "�ېH�ʉ���"
    Wtext = Wtext & vbLf & "�ېH�ʏ��"
    Warray = Split(Wtext, vbLf)
    If UBound(Warray) <> 12 Then
        MsgBox "Program Error No.01 " & UBound(Warray)
        End
    End If
    For i1 = 1 To UBound(Warray)
        Cells(i1 + 2, 1) = Warray(i1)
    Next i1
    Range("c12") = "0:�H�i 1:��� 2:�������"
'   �Z������
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
        With Selection.Interior                         '�w�i�F
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        Selection.Locked = False                        '�ی����
        Selection.FormulaHidden = False
    Next i1
    Range("a17").Locked = False                        '�ی����
End Function
'--------------------------------------------------------------------------------
'   03_813  �h�{�f
'--------------------------------------------------------------------------------
Function Eiyo03_813_eiyoso()
Dim i1      As Long
Dim Wtext   As String
Dim Warray  As Variant
    Wtext = Empty
    Wtext = Wtext & vbLf & "01:�G�l���M�["
    Wtext = Wtext & vbLf & "02:����ς���"
    Wtext = Wtext & vbLf & "03:����������ς���"
    Wtext = Wtext & vbLf & "04:�A��������ς���"
    Wtext = Wtext & vbLf & "05:����"
    Wtext = Wtext & vbLf & "06:����"
    Wtext = Wtext & vbLf & "07:�H������"
    Wtext = Wtext & vbLf & "08:�J���V�E��"
    Wtext = Wtext & vbLf & "09:����"
    Wtext = Wtext & vbLf & "10:�S"
    Wtext = Wtext & vbLf & "11:�i�g���E��"
    Wtext = Wtext & vbLf & "12:�r�^�~���`"
    Wtext = Wtext & vbLf & "13:�r�^�~���a�P"
    Wtext = Wtext & vbLf & "14:�r�^�~���a�Q"
    Wtext = Wtext & vbLf & "15:�i�C�A�V��"
    Wtext = Wtext & vbLf & "16:�r�^�~���b"
    Wtext = Wtext & vbLf & "17:�r�^�~���a�U"
    Wtext = Wtext & vbLf & "18:�p���g�e���_"
    Wtext = Wtext & vbLf & "19:�t�_"
    Wtext = Wtext & vbLf & "20:�r�^�~���d"
    Wtext = Wtext & vbLf & "21:�J���E��"
    Wtext = Wtext & vbLf & "22:�}�O�l�V�E��"
    Wtext = Wtext & vbLf & "23:�H��"
    Wtext = Wtext & vbLf & "24:�R���X�e���[��"
    Wtext = Wtext & vbLf & "25:�s�O�a���b�_"
    Wtext = Wtext & vbLf & "26:�O�a���b�_"
    Wtext = Wtext & vbLf & "27:����"
    Warray = Split(Wtext, vbLf)
    If UBound(Warray) <> 27 Then
        MsgBox "Program Error No.02 " & UBound(Warray)
        End
    End If
    For i1 = 1 To UBound(Warray)
        Cells(i1 + 3, 6) = Warray(i1)
    Next i1
    Cells(3, 6) = "�h�{�f"
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
    Wtext = Wtext & vbLf & "��g"
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
'   03_814  �H�i�Q
'--------------------------------------------------------------------------------
Function Eiyo03_814_shokugun()
Dim i1      As Long
Dim Wtext   As String
Dim Warray  As Variant
    Wtext = Empty
    Wtext = Wtext & vbLf & "�H�i�Q"
    Wtext = Wtext & vbLf & "01:�哤���i"
    Wtext = Wtext & vbLf & "02:�����"
    Wtext = Wtext & vbLf & "03:���@��"
    Wtext = Wtext & vbLf & "04:��"
    Wtext = Wtext & vbLf & "05:�C�@��"
    Wtext = Wtext & vbLf & "06:�����i"
    Wtext = Wtext & vbLf & "07:���@��"
    Wtext = Wtext & vbLf & "08:�Ή��F���"
    Wtext = Wtext & vbLf & "09:�W�F���"
    Wtext = Wtext & vbLf & "10:�ʁ@��"
    Wtext = Wtext & vbLf & "11:���@��"
    Wtext = Wtext & vbLf & "12:������"
    Wtext = Wtext & vbLf & "13:���@��"
    Wtext = Wtext & vbLf & "14:�A��������"
    Wtext = Wtext & vbLf & "15:����������"
    Warray = Split(Wtext, vbLf)
    If UBound(Warray) <> 16 Then
        MsgBox "Program Error No.04 " & UBound(Warray)
        End
    End If
    For i1 = 1 To UBound(Warray)
        Cells(i1 + 2, 9) = Warray(i1)
    Next i1
    Cells(3, 10) = "��ذ"
    Cells(3, 11) = "�d��"
    Cells(3, 12) = "�ټ��"
    Cells(19, 9) = "Total"
    Cells(20, 10) = "(kcal)"
    Cells(20, 11) = "(g)"
    Cells(20, 12) = "(g)"
    Range("j3:l3,i19,j20:l20").HorizontalAlignment = xlRight
End Function
'--------------------------------------------------------------------------------
'   03_815  ����
'--------------------------------------------------------------------------------
Function Eiyo03_815_sisitu()
Dim i1      As Long
Dim Wtext   As String
Dim Warray  As Variant
    Wtext = Empty
    Wtext = Wtext & vbLf & "����(����)"
    Wtext = Wtext & vbLf & "����(����)"
    Wtext = Wtext & vbLf & "����(�A��)"
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
'   03_816  �r���A��
'--------------------------------------------------------------------------------
Function Eiyo03_816_keisen()
Dim i1      As Long
Dim Wtext   As String
Dim Warray  As Variant
'   �r��
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
    With Selection.Interior                         '�w�i�F
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Locked = False                        '�ی����
    Selection.FormulaHidden = False
'   ��
    Columns("F:F").ShrinkToFit = True
    Cells.EntireColumn.AutoFit
    Warray = Array(, 0, 1.75, 5, 18, 2, 15, 0, 7)
    For i1 = 1 To UBound(Warray)
        If Warray(i1) > 0 Then: Columns(i1).ColumnWidth = Warray(i1)
    Next i1
'    Range("B4:C4,B6:C6,B11:C11,B12:C12").NumberFormatLocal = "#,##0.00;[��]-#,##0.00"
    Range("B9,B13,B14").NumberFormatLocal = "#,##0.00;[��]-#,##0.00"
    Range("g4:g30,j4:l19,j22:j25").NumberFormatLocal = "#,##0.00;[��]-#,##0.00"
'   �J�[�\���ʒu�𖾊m������
    Cells.FormatConditions.Delete               '�V�[�g�S�̂�������t���������폜����
    Cells.FormatConditions.Add Type:=xlExpression, Formula1:="=AND(CELL(""row"")=ROW(),CELL(""col"")=COLUMN())"
    Cells.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Cells.FormatConditions(1).Interior.Color = 255
'   ���v�̎��Ə����t������
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
    
    Range("a17").Font.Bold = True           '���b�Z�[�W�G���A�i�ԑ������j
    Range("a17").Font.ColorIndex = 3
End Function
'--------------------------------------------------------------------------------
'   03_817 �R�}���h�E�{�^��
'--------------------------------------------------------------------------------
Function Eiyo03_817_button()
    While (ActiveSheet.Shapes.Count > 0)    '�R�}���h�{�^�����
        ActiveSheet.Shapes(1).Cut
    Wend
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=10, Top:=250, Width:=50, Height:=30)
        .Object.Caption = "�N���A"
        .Name = "�N���A"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=70, Top:=250, Width:=50, Height:=30)
        .Object.Caption = "����"
        .Name = "����"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=130, Top:=250, Width:=50, Height:=30)
        .Object.Caption = "�X�V"
        .Name = "�X�V"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=10, Top:=300, Width:=50, Height:=30)
        .Object.Caption = "�O����"
        .Name = "�O����"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=70, Top:=300, Width:=50, Height:=30)
        .Object.Caption = "������"
        .Name = "������"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=130, Top:=300, Width:=50, Height:=30)
        .Object.Caption = "���"
        .Name = "���"
    End With
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=70, Top:=350, Width:=50, Height:=30)
        .Object.Caption = "�I��"
        .Name = "�I��"
    End With
End Function
'--------------------------------------------------------------------------------
'   ���ʏ����@Eiyo.mdb �̃I�[�v��
'--------------------------------------------------------------------------------
Function Eiyo91DB_Open()
    myCon.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
               "Data Source=" & ThisWorkbook.Path & "Eiyo.mdb;"""
End Function
'--------------------------------------------------------------------------------
'   ���ʏ����@Eiyo.mdb �̃N���[�Y
'--------------------------------------------------------------------------------
Function Eiyo920DB_Close()
    myCon.Close
    Set myCon = Nothing
End Function
'--------------------------------------------------------------------------------
'   ���ʏ����@��ʗ}�~�ق�
'--------------------------------------------------------------------------------
Function Eiyo930Screen_Hold()
    Application.ScreenUpdating = False      '��ʕ`��}�~
    Application.EnableEvents = False        '�C�x���g�����}�~
    ActiveSheet.Unprotect                   '�V�[�g�̕ی������
End Function
'--------------------------------------------------------------------------------
'   ���ʏ����@��ʕ`��ق�
'--------------------------------------------------------------------------------
Function Eiyo940Screen_Start()
    Application.ScreenUpdating = True           '��ʕ`��̕���
    Application.EnableEvents = True             '�C�x���g�����ĊJ
    ActiveSheet.Protect UserInterfaceOnly:=True '�ی��L���ɂ���
End Function
'--------------------------------------------------------------------------------
'   ���ʏ����@�{�^���쐬
'--------------------------------------------------------------------------------
Function Eiyo950Button_Add(in_L As Long, in_t As Long, in_W As Long, in_H As Long, in_text As String)
    With ActiveSheet.OLEObjects.Add("Forms.CommandButton.1", Left:=in_L, Top:=in_t, Width:=in_W, Height:=in_H)
        .Object.Caption = in_text
        .Name = in_text
    End With
End Function
'--------------------------------------------------------------------------------
'   ���ʏ����@�w��V�[�g�폜
'--------------------------------------------------------------------------------
Function Eiyo99_�w��V�[�g�폜(Sname As String)
    Application.DisplayAlerts = False                                   '�m�F�}�~
    If Not IsError(Evaluate(Sname & "!a1")) Then: Sheets(Sname).Delete  '�V�[�g�폜
    Application.DisplayAlerts = True                                    '�m�F����
End Function

