Attribute VB_Name = "Common"
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

' HTTP���N�G�X�g���M�t�@���N�V����
Public Function SendRequest(ByRef URL As String, ByRef ID As String, ByRef PW As String, ByRef body As String) As Object
    ' HTTP�ʐM�̃I�u�W�F�N�g���擾
    Dim objHTTP As Object
    Set objHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
    
    ' URL�ݒ�
    objHTTP.Open "POST", URL
    ' ID�ƃp�X���[�h��Base64�ŃG���R�[�h���A�F�؏��ɐݒ�
    objHTTP.setRequestHeader "Authorization", "Basic " + Call_EncodeBase64(ID + ":" + PW)
    objHTTP.setRequestHeader "Content-Type", "application/json; charset=UTF-8"
    ' ���M
    objHTTP.send body
    
    ' ���X�|���X���A���Ă���܂őҋ@
    Do While objHTTP.readyState < 4
        DoEvents
    Loop
    
    ' ���X�|���X�{�f�B��ԋp
    Set SendRequest = objHTTP
End Function


