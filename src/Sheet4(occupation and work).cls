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
'   �E�ƂƘJ����_�u���N���b�N�őI��
'--------------------------------------------------------------------------------
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim Wline As Long '�s
    Dim Wcolu As Long '��
    Dim ��b�V�[�g As Worksheet

    ' ��b�V�[�g���I�u�W�F�N�g�Ƃ��Ď擾
    Set ��b�V�[�g = ThisWorkbook.Sheets("��b")

    ' �_�u���N���b�N�̃C�x���g���L�����Z��
    Cancel = True

    ' �_�u���N���b�N���ꂽ�Z���̍s�Ɨ���擾
    Wline = Target.Row
    Wcolu = Target.Column

    ' ��̒l�𒲐�
    Select Case Wcolu
        Case 2: Wcolu = 3
        Case 6: Wcolu = 7
        Case 10: Wcolu = 11
    End Select

    ' ��b�V�[�g�̃Z���ɒl��ݒ�
    ��b�V�[�g.Range("Qjob1").Value = Target.Parent.Cells(Wline, Wcolu).Value

    ' ��b�V�[�g��I��
    ��b�V�[�g.Select
End Sub