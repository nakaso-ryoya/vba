Attribute VB_Name = "Common"
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

' HTTPリクエスト送信ファンクション
Public Function SendRequest(ByRef URL As String, ByRef ID As String, ByRef PW As String, ByRef body As String) As Object
    ' HTTP通信のオブジェクトを取得
    Dim objHTTP As Object
    Set objHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
    
    ' URL設定
    objHTTP.Open "POST", URL
    ' IDとパスワードをBase64でエンコードし、認証情報に設定
    objHTTP.setRequestHeader "Authorization", "Basic " + Call_EncodeBase64(ID + ":" + PW)
    objHTTP.setRequestHeader "Content-Type", "application/json; charset=UTF-8"
    ' 送信
    objHTTP.send body
    
    ' レスポンスが帰ってくるまで待機
    Do While objHTTP.readyState < 4
        DoEvents
    Loop
    
    ' レスポンスボディを返却
    Set SendRequest = objHTTP
End Function


