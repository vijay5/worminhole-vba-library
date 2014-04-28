''' ����������� ������� ���������� (encode = -1 ��� ����������)
Function encrypt(ByVal inString As String, Optional encode As Integer = 1) As String
    Dim length As String
    Dim key As String
    Dim i As Integer
    Dim isOdd As Integer
    Dim index As Integer
    Dim activechar As String, codedChar As String
    
    key = "314159265358979" ' pi ��� ����� ������ ����
    length = IIf(Len(inString) = 0, 1, Len(inString))
    key = Mid(key, length) + Left(key, length - 1) ' �������� ������
    
    encrypt = ""
    
    For i = 1 To Len(inString)
        activechar = Mid(inString, i, 1)
        isOdd = IIf((i / 2) = (i \ 2), 1, -1) ' ��������/����������
        index = (i - 1) Mod Len(key) + 1      ' ����� ������� ������ �����
        codedChar = Chr(Asc(activechar) + encode * isOdd * CInt(Mid(key, index, 1)))
        encrypt = encrypt + codedChar
    Next i
End Function