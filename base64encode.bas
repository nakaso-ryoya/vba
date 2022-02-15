Attribute VB_Name = "Module1"
Public Function Call_EncodeBase64(ByRef text As String) As String
    'オブジェクト準備
    Dim node As Object, obj As Object
    Set node = CreateObject("Msxml2.DOMDocument.3.0").createElement("base64")
    Set obj = CreateObject("ADODB.Stream")
  
    'エンコード(textをBASE64へ変換)
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
  
    '改行を削除して返却(上記で取り除けない為)
    Call_EncodeBase64 = Replace(node.text, vbLf, "")
End Function
