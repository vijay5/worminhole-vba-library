''' преобразует массив в строку формата "12/20.40"
Function Collapse(inArray As Variant, Optional dotDivisor As String = ".", Optional slashDivisor As String = "/", Optional interval As Boolean = False) As Variant
    Dim i As Long, j As Long, array1 As Variant, st As Variant, en As Variant
    Dim OutString As String
    st = ""
    en = ""
    Collapse = "" ' по умолчанию - пусто
    If arrayDepth(inArray) = 1 Then ' если у нас одномерный массив
        ' в случае если нам скормили текстовый массив - перегоняем в цифры
        ReDim array1(LBound(inArray) To UBound(inArray))
        For i = LBound(inArray) To UBound(inArray)
            If IsNumeric(inArray(i)) Then
                array1(i) = CLng(inArray(i))
            Else
                array1(i) = inArray(i)
            End If
        Next i
        array1 = ShellSort(array1) ' сортируем по возрастанию
        
        If InStr(TypeName(array1), "()") = 0 Then Exit Function ' тоже ошибка
        For i = LBound(array1) To UBound(array1)
            If st = "" Then st = array1(i)
            If i < UBound(array1) Then
                If IsNumeric(array1(i)) And IsNumeric(array1(i + 1)) Then
                    If array1(i + 1) - array1(i) > 1 Then en = array1(i)
                Else
                    en = array1(i)
                End If
            Else
                en = array1(UBound(array1))
            End If
            If st <> "" And en <> "" Then
                If st = en Then
                    If interval Then ' если описываем интервал
                        addToText OutString, "=" + CStr(st), dotDivisor
                    Else
                        addToText OutString, CStr(st), dotDivisor
                    End If
                    
                ElseIf st <> en Then
                    If interval Then ' если описываем интервал
                        addToText OutString, ">=" + CStr(st) + slashDivisor + "<=" + CStr(en), dotDivisor
                    Else
                        addToText OutString, CStr(st) + slashDivisor + CStr(en), dotDivisor
                    End If
                End If
                st = ""
                en = ""
            End If
        Next i
    End If
    Collapse = OutString
End Function
