Attribute VB_Name = "Module1"
Public Function Call_EncodeBase64(ByRef text As String) As String
    '�I�u�W�F�N�g����
    Dim node As Object, obj As Object
    Set node = CreateObject("Msxml2.DOMDocument.3.0").createElement("base64")
    Set obj = CreateObject("ADODB.Stream")
  
    '�G���R�[�h(text��BASE64�֕ϊ�)
    node.DataType = "bin.base64"
    With obj
        .Type = 2
        .Charset = "us-ascii"
        .Open
        .WriteText text
        .Position = 0
        .Type = 1
        .Position = 0
    End With
    node.nodeTypedValue = obj.Read
  
    '���s���폜���ĕԋp(��L�Ŏ�菜���Ȃ���)
    Call_EncodeBase64 = Replace(node.text, vbLf, "")
End Function
