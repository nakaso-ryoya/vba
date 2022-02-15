Attribute VB_Name = "Module1"
Sub ボタン1_Click()
    Dim path As String
    Dim replaceFrom As String
    Dim replaceTo As String
    
    path = Range("B1").Text
    replaceFrom = Range("B2").Text
    replaceTo = Range("B3").Text
    
    Call replaceFileNameAll(path, replaceFrom, replaceTo)
    MsgBox ("完了")
End Sub

Sub replaceFileNameAll(path As String, replaceFrom As String, replaceTo As String)
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    ' フォルダを取得
    Dim fl As Folder
    Set fl = fso.GetFolder(path)
    
    Dim fileName As String
    Dim f As File
    On Error Resume Next
    For Each sf In fl.SubFolders
        Call replaceFileNameAll(sf.path, replaceFrom, replaceTo)
    Next sf
    For Each f In fl.Files ' フォルダ内のファイルを取得
        fileName = f.Name             ' ファイル名の取得
       
       If InStr(f.Name, replaceFrom) > 0 Then
        f.Name = Replace(fileName, replaceFrom, replaceTo)
        cnt = cnt + 1
        End If
    Next f
End Sub
