Attribute VB_Name = "Module2"
Public Function outputTxtFile(ByRef text As String, ByRef path As String) As Integer
    On Error GoTo error
    Open path For Output As #1
    Print #1, text
    Close #1
    outputTxtFile = 0
    Exit Function
error:
    outputTxtFile = 1
    Exit Function
End Function
