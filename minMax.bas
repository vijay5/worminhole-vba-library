Function MinMax(value As Variant, minValue As Variant, maxValue As Variant) As Variant
    ' ���������� �������� �� ������ minValue � �� ������ maxValue
    If value < minValue Then
        MinMax = minValue
    ElseIf value > maxValue Then
        MinMax = maxValue
    Else
        MinMax = value
    End If
End Function