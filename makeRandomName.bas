Function MakeRandomName(Optional num As Integer = 15) As String
    ' Aggregate - ��� ��������� ���� ������
    Dim st As String
    Dim strArray As Variant
    Dim i As Long
    
    Randomize Timer
    st = ""
    strArray = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
    For i = 1 To WorksheetFunction.Min(num, 20) ' ��� ������ ��� ������������ �����
        st = st + Mid(strArray, CInt(Rnd() * Len(strArray) + 1), 1)
    Next i
    MakeRandomName = st
End Function