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
'   �ېH�̃��[�N�V�[�g���I�����ꂽ
'--------------------------------------------------------------------------------
Private Sub Worksheet_Activate()
    Application.Run macro:="Eiyo_Macro.xla!Eiyo01_400MealDisp"
End Sub
'--------------------------------------------------------------------------------
'   ���[�N�V�[�g�l���ύX���ꂽ
'--------------------------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
    Application.EnableEvents = False            '�C�x���g�����}�~
    Application.Run "Eiyo_Macro.xla!Eiyo01_410MealChange", Target.Address
    Application.EnableEvents = True             '�C�x���g�����ĊJ
End Sub
'--------------------------------------------------------------------------------
'   �_�u���N���b�N���ꂽ
'--------------------------------------------------------------------------------
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Application.EnableEvents = False            '�C�x���g�����}�~
    Cancel = True
    Range("b1") = Target.Address
    Application.EnableEvents = True             '�C�x���g�����ĊJ
    Sheets("�ƭ�").Select
'    Application.Run "Eiyo_Macro.xla!Eiyo01_420Menu", Target.Address
End Sub
'--------------------------------------------------------------------------------
'   ����_Click
'--------------------------------------------------------------------------------
Private Sub ����_Click()
    Application.Run "Eiyo_Macro.xla!Eiyo01_500MealCalc", 2
End Sub
'--------------------------------------------------------------------------------
'   �o�^_Click
'--------------------------------------------------------------------------------
Private Sub �o�^_Click()
    Application.Run "Eiyo_Macro.xla!Eiyo01_500MealCalc", 1
End Sub