''' ������� ��� �������� ��������� �������� � ��������� ������, �������� ������������ ������� �� "str1" + Chr(xx) + "str2"
' REQUIRES: addToText
Function isolateSymbol(inString As String, symbol As String) As String
    Dim pos As Long
    Dim strToProcess As String
    Dim border As String, lBorder As String, rBorder As String
    Dim tmpStr As String
    Dim tmpArray As Variant
    Dim i As Long
    
    strToProcess = inString
    
    border = Chr(1) + Chr(3) + Chr(2) ' ������ �� ������������, ������� �� ����� ���� � ������� ������
    lBorder = border + "+" ' ����� "������"
    rBorder = "+" + border ' ����� "������"
    
    pos = InStr(strToProcess, symbol)
    If pos > 0 Then ' ���� ���� ���-�� �������
        tmpStr = ""
        For i = 1 To Len(symbol) ' �� ������, ���� � symbol ������ ��������� ��������
            addToText tmpStr, "Chr(" + CStr(Asc(Mid(symbol, i, 1))) + ")", "+"
        Next i
    End If
    
    tmpArray = Split(strToProcess, symbol)
    strToProcess = Join(tmpArray, lBorder + tmpStr + rBorder)
    If InStr(strToProcess, lBorder) = 1 Then ' ������� ��������� ������
        strToProcess = Mid(strToProcess, Len(lBorder) + 1)
    End If
    If InStrRev(strToProcess, rBorder) > 0 And _
       (InStrRev(strToProcess, rBorder) = Len(strToProcess) - Len(rBorder) + 1) Then ' ������� �������� ������
        strToProcess = Left(strToProcess, Len(strToProcess) - Len(rBorder))
    End If
    
    isolateSymbol = Replace(strToProcess, border, """")
    
End Function
