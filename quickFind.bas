' ��� �������� ������ ������� ��������
Function quickFind(ByVal text1 As Variant, firstSymbol As String, lastSymbol As String, Optional startFromSymbol As Integer = 1) As Variant
    Dim st As Integer, en As Integer, stAlternative As Integer
    quickFind = ""
    st = InStr(startFromSymbol, text1, firstSymbol)
    If st > 0 Then ' ����� ������
        en = InStr(st + 1, text1, lastSymbol)
        If en > 0 Then ' ����� �����
            stAlternative = InStrRev(text1, firstSymbol, en)
            quickFind = Mid(text1, stAlternative, en - stAlternative + 1)
        End If
    End If
End Function