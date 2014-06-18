''' ����� ���������� ���������
''' ����� ������������ "array" � �������� returnValue
Function reFind(ByVal text1 As Variant, ByVal patt As Variant, Optional ByVal returnValue As Variant = "boolean", Optional ByVal index As Variant = -1) As Variant
    Set re = CreateObject("vbscript.regexp") ' ������� ���������� - ���������� RegExp
    Dim Matches, Match, tmp, m
    
    With re
          .MultiLine = False
          .Global = True
          .IgnoreCase = False
          .Pattern = patt
    End With
    Set Matches = re.Execute(text1) ' �������� ����� (����� ����� ������� ���)
    
    reFind = ""
    index = CInt(index)
    Select Case LCase(returnValue)
    Case "boolean" ' True/False
        If Matches.count > 0 Then reFind = True Else reFind = False
    Case "value" ' ������� ������
        If Matches.count > 0 Then
            reFind = Matches(0).value
        Else
            reFind = ""
        End If
    Case "array" ' ������ (������ ���������, ����������� � ������, � ��������� ������)
        If Matches.count > 0 Then
            ReDim tmp(0 To Matches(0).SubMatches.count)
            tmp(0) = Matches(0).value
            For m = 1 To Matches(0).SubMatches.count
                tmp(m) = Matches(0).SubMatches(m - 1)
            Next m
            
            If index >= 0 And index <= Matches.count Then ' ���� ����� ������, ��� ������ � �����
                reFind = tmp(index)
            Else
                reFind = tmp
            End If
        Else
            reFind = ""
        End If
    End Select
    Set re = Nothing
End Function