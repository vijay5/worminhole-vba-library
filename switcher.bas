' для запуска с рабочего листа
' при большом количестве формул 2003 глючит
Function switcher(ParamArray VarExp() As Variant) As Variant
    Select Case (UBound(VarExp) - LBound(VarExp) + 1)
    Case 2: switcher = Switch(VarExp(0), VarExp(1))
    Case 4: switcher = Switch(VarExp(0), VarExp(1), VarExp(2), VarExp(3))
    Case 6: switcher = Switch(VarExp(0), VarExp(1), VarExp(2), VarExp(3), VarExp(4), VarExp(5))
    Case 8: switcher = Switch(VarExp(0), VarExp(1), VarExp(2), VarExp(3), VarExp(4), VarExp(5), VarExp(6), VarExp(7))
    Case 10: switcher = Switch(VarExp(0), VarExp(1), VarExp(2), VarExp(3), VarExp(4), VarExp(5), VarExp(6), VarExp(7), VarExp(8), VarExp(9))
    Case 12: switcher = Switch(VarExp(0), VarExp(1), VarExp(2), VarExp(3), VarExp(4), VarExp(5), VarExp(6), VarExp(7), VarExp(8), VarExp(9), VarExp(10), VarExp(11))
    Case 14: switcher = Switch(VarExp(0), VarExp(1), VarExp(2), VarExp(3), VarExp(4), VarExp(5), VarExp(6), VarExp(7), VarExp(8), VarExp(9), VarExp(10), VarExp(11), VarExp(12), VarExp(13))
    Case Else
        switcher = "#N/A"
    End Select
End Function