Attribute VB_Name = "Module1"
Sub �{�^��1_Click()
    Dim path As String
    Dim replaceFrom As String
    Dim replaceTo As String
    
    path = Range("B1").Text
    replaceFrom = Range("B2").Text
    replaceTo = Range("B3").Text
    
    Call replaceFileNameAll(path, replaceFrom, replaceTo)
    MsgBox ("����")
End Sub

Sub replaceFileNameAll(path As String, replaceFrom As String, replaceTo As String)
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    ' �t�H���_���擾
    Dim fl As Folder
    Set fl = fso.GetFolder(path)
    
    Dim fileName As String
    Dim f As File
    On Error Resume Next
    For Each sf In fl.SubFolders
        Call replaceFileNameAll(sf.path, replaceFrom, replaceTo)
    Next sf
    For Each f In fl.Files ' �t�H���_���̃t�@�C�����擾
        fileName = f.Name             ' �t�@�C�����̎擾
       
       If InStr(f.Name, replaceFrom) > 0 Then
        f.Name = Replace(fileName, replaceFrom, replaceTo)
        cnt = cnt + 1
        End If
    Next f
End Sub
