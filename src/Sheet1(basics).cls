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

Private Sub �N���A_Click()
    Application.Run macro:="Eiyo_Macro.xla!Eiyo01_030�N���AClick"
End Sub

Private Sub ����_Click()
    Application.Run macro:="Eiyo_Macro.xla!Eiyo01_100����Click"
End Sub

Private Sub �X�V_Click()
    Application.Run macro:="Eiyo_Macro.xla!Eiyo01_200�X�VClick"
End Sub

Private Sub ���_Click()
    If MsgBox("��������s���܂���", vbOKCancel) = vbOK Then
        Application.Run macro:="Eiyo_Macro.xla!Eiyo01_300���Click"
    End If
End Sub

Private Sub ��\_Click()
    Application.Run macro:="Eiyo_Macro.xla!Eiyo01_700��\Click"
End Sub

Private Sub �I��_Click()
    ActiveWorkbook.Close savechanges:=False     '�ۑ������ɏI��
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Left(Target.Address, 4) = "$AH$" Then        '�t�@�~���[�R�[�h�ꗗ����̑I��
        Range("Fcode") = Target.Value
        Cancel = True
        ����_Click
    ElseIf Left(Target.Address, 4) = "$Q$8" Then    '�E�ƂƘJ��
        Cancel = True
        Sheets("�E�ƂƘJ��").Select
    End If
End Sub
'--------------------------------------------------------------------------------
'   ���[�N�V�[�g�̒l���ύX���ꂽ
'--------------------------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address = Range("Adrno").Address Then         '�X�֔ԍ�
        If Len(Target.Text) > 3 And Range("adrs1") = Empty Then
            Application.EnableEvents = False    '�C�x���g�����}�~
            ' �X�֔ԍ���S�p�ϊ����Z���ɓ]�L
            Range("adrs1").Value = StrConv(Range("Adrno").Value, vbWide)
            ' �Z���̾ق�I��
            Range("adrs1").Select
            ' F2 �� Shift+Home �� F13 ���[�����͂���
            SendKeys "{F2}", True       ' �ҏWӰ��
            SendKeys "+{HOME}", True    ' ������S�̂�I��
            SendKeys "{F13}", True      ' �ĕϊ�(MS-IME)
            Application.EnableEvents = True     '�C�x���g�����ĊJ
        End If
    End If
End Sub
'--------------------------------------------------------------------------------
'   �A�N�e�B�u�ɂȂ����Z���������\�����邽�߃Z�����I���������Ƃ���ʍX�V����
'--------------------------------------------------------------------------------
Private Sub Worksheet_SelectionChange(ByVal Traget As Range)
    Application.ScreenUpdating = True
End Sub