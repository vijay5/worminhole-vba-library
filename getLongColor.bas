''' узнаёт цвет заливки или шрифта ячейки
Function getLongColor(cl As Variant, Optional getColorForFill As Boolean = True) As Variant
    Dim tmpValue As Variant
    getLongColor = "!Error"
    If TypeName(cl) = "Range" Then
        If getColorForFill Then
            tmpValue = CLng(cl.Cells(1, 1).Interior.Color)
        Else
            tmpValue = CLng(cl.Cells(1, 1).Font.Color)
        End If
    ElseIf IsNumeric(cl) Then
        If CLng(cl) >= 0 And CLng(cl) <= CLng("&HFFFFFF") Then
            tmpValue = CLng(cl)
        Else
            Exit Function
        End If
    Else
        Exit Function
    End If
    
    getLongColor = tmpValue
End Function