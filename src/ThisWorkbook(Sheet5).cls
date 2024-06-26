VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'--------------------------------------------------------------------------------
'   オープンされたらマクロもオープンする
'--------------------------------------------------------------------------------
Private Sub Workbook_Open()
    Dim macroPath As String
    On Error GoTo ErrorHandler
    macroPath = ThisWorkbook.Path & "Eiyo_macro.xla"""
    If Dir(macroPath) <> "" Then
        Workbooks.Open macroPath, ReadOnly:=True
    Else
        MsgBox "Eiyo_macro.xlaが見つかりません。", vbExclamation
    End If
    Application.Run "Eiyo01_900WorkbookOpen"
    Exit Sub

ErrorHandler:
    MsgBox "マクロの読み込み中にエラーが発生しました: " & Err.Description, vbCritical
End Sub
'--------------------------------------------------------------------------------
'   自ブックを閉じる前にマクロを閉じる
'--------------------------------------------------------------------------------
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Const Gname = "Eiyo"
    Dim i1 As Long
    Dim Cnt As Long
    Cnt = 0
    
    For i1 = 1 To Workbooks.Count
        If Left(Workbooks(i1).Name, Len(Gname)) = Gname Then
            Cnt = Cnt + 1
        End If
    Next i1
    
    If Cnt < 2 Then
        On Error Resume Next
        If Workbooks("Eiyo_macro.xla").ReadOnly Then
            Workbooks("Eiyo_macro.xla").Close savechanges:=False
        ElseIf MsgBox("Eiyo_macro.xla を保存しますか", vbYesNo) = vbYes Then
            Workbooks("Eiyo_macro.xla").Close savechanges:=True
        Else
            Workbooks("Eiyo_macro.xla").Close savechanges:=False
        End If
        On Error GoTo 0
    End If
End Sub


