Attribute VB_Name = "Module1"
Sub �Z���̌���()
Attribute �Z���̌���.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' �Z���̌��� Macro
'
' Keyboard Shortcut: Ctrl+q
'
    Selection.Merge
End Sub


Sub �擪�̃Z����I��()
Attribute �擪�̃Z����I��.VB_ProcData.VB_Invoke_Func = " \n14"
'
' �擪�̃Z����I�� Macro
'
    For i = Worksheets.Count To 1 Step -1
        
        If Sheets(i).Visible = False Then
        Else
            ActiveWindow.ScrollColumn = 1
            ActiveWindow.ScrollRow = 1
            Sheets(i).Select
            ActiveWindow.ActiveSheet.Range("A1").Select
        End If
    Next i
End Sub


Sub Zoom()
'
' �{���ݒ� Macro
'
    Dim per As Variant
re:
    per = Application.InputBox(Type:=1, prompt:="���l�Ŕ{������͂��Ă�������", Title:="�{���ݒ�")
    
    If per = False Then
        GoTo Endsub
    End If
    
    If per >= 10 And per <= 400 Then
        For i = Worksheets.Count To 1 Step -1
           Sheets(i).Select
           ActiveWindow.Zoom = per
        Next i
    Else
        MsgBox ("10����400�͈̔͂œ��͂��Ă�������")
        GoTo re
    End If
    
Endsub:
End Sub

