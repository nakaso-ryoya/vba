Attribute VB_Name = "Module4"
' HTTP���N�G�X�g���M�t�@���N�V����
Public Function SendRequest(ByRef URL As String, ByRef ID As String, ByRef PW As String, ByRef body As String) As String
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
    SendRequest = objHTTP.responseText
End Function

