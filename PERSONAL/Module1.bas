Attribute VB_Name = "Module1"
Sub セルの結合()
Attribute セルの結合.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' セルの結合 Macro
'
' Keyboard Shortcut: Ctrl+q
'
    Selection.Merge
End Sub


Sub 先頭のセルを選択()
Attribute 先頭のセルを選択.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 先頭のセルを選択 Macro
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
' 倍率設定 Macro
'
    Dim per As Variant
re:
    per = Application.InputBox(Type:=1, prompt:="数値で倍率を入力してください", Title:="倍率設定")
    
    If per = False Then
        GoTo Endsub
    End If
    
    If per >= 10 And per <= 400 Then
        For i = Worksheets.Count To 1 Step -1
           Sheets(i).Select
           ActiveWindow.Zoom = per
        Next i
    Else
        MsgBox ("10から400の範囲で入力してください")
        GoTo re
    End If
    
Endsub:
End Sub

