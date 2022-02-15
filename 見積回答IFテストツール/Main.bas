Attribute VB_Name = "Main"
Const URL As String = "http://172.23.146.143:8000/lindo/logic/api/poifad021" ' �Ώ�IF��URL
Const DETAIL_ROW As Long = 8 ' ���׍s�X�^�[�g�ʒu

Sub send()
    Application.ScreenUpdating = False
    Dim ID As String  ' ���[�UID
    Dim PW As String  ' �p�X���[�h
    Dim requestNo As String  ' ���N�G�X�g�ԍ�
    Dim mode As String ' ���샂�[�h
    Dim body As String '���N�G�X�g�{�f�B
    
    Range("A8:X1000000").ClearContents
    
    ID = Range("B1").text
    PW = Range("B2").text
    requestNo = Range("B3").text
    
    If Range("B4").text = "�ʏ�" Then
        mode = "normal"
    End If
    If Range("B4").text = "�Ď擾" Then
        mode = "recovery"
    End If
    
    Dim Common As New Dictionary
    Common.Add "SYSTEM_ID", "01"
    Common.Add "REQUEST_NO", requestNo
    Common.Add "MODE", mode
    
    Dim jsonObject As New Dictionary
    jsonObject.Add "COMMON", Common
    
    body = JsonConverter.ConvertToJson(jsonObject)
    
    Dim response As Object
    Dim responseTxt As String
    Dim responseJson As Object
    Set response = SendRequest(URL, ID, PW, body)
    responseTxt = response.responseText
    
    
    If Not response.Status = 200 Then
        MsgBox "HTTP���M�G���[�ł��I" & Chr(13) & "�X�e�[�^�X�R�[�h�F" & response.Status
    End If
    
    Set responseJson = JsonConverter.ParseJson(responseTxt)
    
    
    If responseJson("RESULT") = "00" Then
        Dim n As Long
        n = 0
        Dim size As Long
        size = responseJson("DATA").Count
        Dim cellValue()
        ReDim cellValue(size, 23)
        For Each i In responseJson("DATA")
            cellValue(n, 0) = i("ESTIMREP_SHORI_NO")
            cellValue(n, 1) = i("ESTIMREP_INF_CD")
            cellValue(n, 2) = i("ESTIMREP_DATA_CRE_YMD")
            cellValue(n, 3) = i("ESTIMREP_MA_CUSTOMERCD")
            cellValue(n, 4) = i("ESTIMREP_SHORI_KBN")
            cellValue(n, 5) = i("ESTIMREP_HINMCD")
            cellValue(n, 6) = i("SEISAN_KOJOH_CD")
            cellValue(n, 7) = i("CHUMON_NO")
            cellValue(n, 8) = i("ESTIMREP_QTY")
            cellValue(n, 9) = i("ESTIMREP_AMT")
            cellValue(n, 10) = i("ESTIMREP_NOUKI_YMD")
            cellValue(n, 11) = i("ESTIMREP_YMD")
            cellValue(n, 12) = i("ESTIMREP_STYMD")
            cellValue(n, 13) = i("ESTIMREP_MINORD_QTY")
            cellValue(n, 14) = i("ESTIMREP_PURCHASE_TARGET_AMT")
            cellValue(n, 15) = i("CHUMON_OCCURYM")
            cellValue(n, 16) = i("SINSEI_BUMON_CD")
            cellValue(n, 17) = i("APPLICATION_NO")
            cellValue(n, 18) = i("MA_HINNM")
            cellValue(n, 19) = i("MA_SPEC")
            cellValue(n, 20) = i("MRUME_QTY")
            cellValue(n, 21) = i("UNITSIG")
            cellValue(n, 22) = i("SUPPLY_LEAD_TIME")
            cellValue(n, 23) = i("ESTIMREQ_TYPE")
            
            n = n + 1
        Next
        
        Range(Cells(DETAIL_ROW, 1), Cells(size + DETAIL_ROW - 1, 24)).Value = cellValue
    ElseIf responseJson("RESULT") = "01" Then
        MsgBox "���ʂ�0���ł�"
    Else
        MsgBox "���̓G���[�ł��I" & Chr(13) & "RESULT�F" & responseJson("RESULT")
    End If
    
    Application.ScreenUpdating = True
    
End Sub
