''' ��� ������� � �������� �����
''' ��� ������� ���������� ������ 2003 ������
Function switcher(ParamArray VarExp() As Variant) As Variant
    If (UBound(VarExp) - LBound(VarExp) + 1) Mod 2 = 0 And (UBound(VarExp) - LBound(VarExp) + 1) <= 30 Then
        switcher = Switch(VarExp)
    Else
        switcher = "#N/A"
    End If
End Function