Attribute VB_Name = "Module4"
Public Function SendRequest(ByRef URL As String, ByRef id As String, ByRef pw As String, ByRef body As String) As Object
    Dim objHTTP As Object
    Set objHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
    objHTTP.Open "POST", URL
    objHTTP.setRequestHeader "Authorization", "Basic " + Call_EncodeBase64(id + ":" + pw)
    objHTTP.setRequestHeader "Content-Type", "application/json; charset=UTF-8"
    objHTTP.send body
    
    Do While objHTTP.readyState < 4
        DoEvents
    Loop
    
    Set SendRequest = objHTTP
End Function

