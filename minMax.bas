''' возвращает значение не меньше minValue и не больше maxValue
Function MinMax(value As Variant, minValue As Variant, maxValue As Variant) As Variant
    If value < minValue Then
        MinMax = minValue
    ElseIf value > maxValue Then
        MinMax = maxValue
    Else
        MinMax = value
    End If
End Function