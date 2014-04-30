''' Получаем значения разделителей, действующих в системе
Public dateSeparator as String
Public timeSeparator as String
Public decimalSeparator as String
Public antiDecimalSeparator as String

Sub getSeparators()
    dateSeparator = Mid(Format(Date, "General Date"), 3, 1)
    timeSeparator = Mid(Format(0.5, "Long Time"), 3, 1)
    decimalSeparator = Mid(Format(1.1, "General Number"), 2, 1)
    If decimalSeparator = "." Then antiDecimalSeparator = ","
    If decimalSeparator = "," Then antiDecimalSeparator = "."
End Sub
