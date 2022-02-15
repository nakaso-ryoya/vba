Attribute VB_Name = "Module4"
' HTTPリクエスト送信ファンクション
Public Function SendRequest(ByRef URL As String, ByRef ID As String, ByRef PW As String, ByRef body As String) As String
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
    SendRequest = objHTTP.responseText
End Function

