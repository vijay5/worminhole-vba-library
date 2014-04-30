''' �-�� MAP ��� hexString: ����������� hex ������ � ������ ��������, �������� � ������ ��������� chr(cdbl(hex(x)))
'REQUIRES: addJournal, addToText
Function hex2str(hexString_in As Variant, Optional divisor As String = ",") As Variant
    Dim funcName As String
    Dim isUnicode As Boolean
    Dim symbol As Variant
    Dim resultString As Variant
    Dim hexString As Variant
    Dim hexArray As Variant
    Dim hexAtom As Variant
    Dim i As Long
    
    funcName = "hex2str"
    hexString = Replace(hexString_in, "&H", "")
    hexString = Replace(hexString, "&h", "")
    hexArray = Split(hexString, divisor)
    hex2str = "" ' �� ��������� �� ������ ������ ������
    
    For i = LBound(hexArray) To UBound(hexArray)
        hexAtom = hexArray(i)
        isUnicode = False
        
        Select Case Len(hexAtom) ' ������� �� ����� ���������
        Case 0
            hexAtom = "00"
        Case 1
            hexAtom = "0" + hexAtom
        Case 2
            ' �� �����
        Case 3
            isUnicode = True
            hexAtom = "0" + hexAtom ' 4-� ������� ��� - ChrW
        Case 4
            isUnicode = True
        Case Else
            addJournal funcName, "[Warning]", "�������� ���������� �������� � �����. ��������� �����������. ��� ������� ����� ���� ����� 1-4 ���������. ������: " + CStr(hexString)
            Exit Function
        End Select
        
        ' ����������� ������� &H
        hexAtom = "&H" + hexAtom
        
        ' ��������� ��������� �� ��������
        On Error Resume Next
            symbol = Array(0)
            If isUnicode Then
                symbol = ChrW(CLng(hexAtom)) ' ������ Unicode
            Else
                symbol = Chr(CLng(hexAtom))  ' ������ ASCII
            End If
        On Error GoTo 0
        If IsArray(symbol) Then
            addJournal funcName, "[Warning]", "�� ������� ������������� ��������� """ + CStr(hexAtom) + """ � hex-������"
            Exit Function
        End If
        
        ' �������� ������ �� ������ (��������)
        addToText resultString, symbol, ""
    Next i
    hex2str = resultString
End Function
