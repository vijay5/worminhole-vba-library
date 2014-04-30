''' преобразует строку с множеством кодов в массив, где перечисляется каждый код
''' на выходе - 1) 0-based Array из цифр или 2) 0-based Array из строк или 3) Dictionary
''' 175/190.207/209 -> (175, 176, 177, ..., 189, 190, 207, 208, 209)
Function Expand(sourceString As String, Optional stringOrNumeric As String = "numeric", Optional useSort As Boolean = True)
    Expand = False
    
' Stage 1 - очистка исходной строки, проверка правильности символов
    Dim Step1, Step2 As Variant
    Dim Step3 As Variant
    Dim Step4 As Variant
    Dim outDic As Variant
    Dim k As Long, j As Long, i As Long
    Dim tmp As Variant, brk As Variant
    Dim chk1 As Boolean, chk2 As Boolean
    Dim changedString As Variant
    Dim el As Variant
    
    Set outDic = CreateObject("Scripting.Dictionary")
    outDic.RemoveAll
    
    changedString = sourceString
    
    changedString = Replace(changedString, "   ", " ")
    changedString = Replace(changedString, "  ", " ")
    changedString = Replace(changedString, " , ", ".")
    changedString = Replace(changedString, ", ", ".")
    changedString = Replace(changedString, " ,", ".")
    changedString = Replace(changedString, " thru ", "/")
    changedString = Replace(changedString, " thr ", "/")
    changedString = Replace(changedString, " to ", "/")
    changedString = Replace(changedString, ",", ".")
    changedString = Replace(changedString, ";", ".")
    changedString = Replace(changedString, "`", "")
    changedString = Replace(changedString, ":", "/")
    changedString = Replace(changedString, "-", "/")
    changedString = Replace(changedString, "\", "/")
    changedString = Replace(changedString, "..", ".")
    changedString = Replace(changedString, "//", "/")
    changedString = Replace(changedString, " ", ".")
    
    ' проверяем на наличие левых символов
    brk = reFind(LCase(changedString), "[^0-9a-z.\/]")
    
    If brk Or changedString = "" Then
        Expand = "" ' возвращаем "пусто"
    Else
        k = 0 ' длина конечного массива
        Step1 = Split(changedString, ".") ' бьем по "точкам"
        For i = LBound(Step1) To UBound(Step1)
            If Step1(i) <> "" Then
                Step2 = Split(Step1(i), "/")  ' бьем по "слешам"
                ' может разбиться (тогда UBound > 0) или не разбиться (тогда UBound = 0)
                If UBound(Step2) = 1 Then ' может быть 1 элемент - 122.134.156
                    chk1 = IsNumeric(Step2(0))
                    chk2 = IsNumeric(Step2(1))
                    If chk1 And chk2 Then ' все числовые
                        If CDbl(Step2(0)) > CDbl(Step2(1)) Then ' меняем местами начало и конец
                            tmp = Step2(0)
                            Step2(0) = Step2(1)
                            Step2(1) = tmp
                        End If
                    Else ' нечисловые (трудно сказать, что тут происходит)
                        If Step2(0) > Step2(1) Then ' меняем местами начало и конец
                            tmp = Step2(0)
                            Step2(0) = Step2(1)
                            Step2(1) = tmp
                        End If
                    End If
                ElseIf UBound(Step2) > 1 Then ' запись типа: 122/126/140 - ошибка
                    Expand = ""
                    Exit Function
                Else ' UBound(Step2) =0
                    ' случай когда у нас только один элемент после разбиения не обрабатываем
                    ' этот элемент перепишется в Step4 "как есть"
                End If
                ' пишем в конец Step3 все числа попорядку
                If IsNumeric(Step2(LBound(Step2))) And IsNumeric(Step2(UBound(Step2))) Then
                    For j = CLng(Step2(LBound(Step2))) To CLng(Step2(UBound(Step2)))
                        If Not outDic.Exists(j) Then outDic.Add j, 1
                    Next j
                Else ' а если не числа - пишем все что есть, не раскрывая
                    For Each el In Step2
                        If Not outDic.Exists(el) Then outDic.Add el, 1
                    Next el
                End If
            End If
        Next i
        
        If outDic.count = 0 Or Not (LCase(stringOrNumeric) = "numeric" Or LCase(stringOrNumeric) = "string" Or LCase(stringOrNumeric) = "dic") Then
            Expand = ""
        ElseIf LCase(stringOrNumeric) = "numeric" Then
            If useSort Then
                Expand = ShellSort(outDic.keys)    ' numericArray
            Else
                Expand = outDic.keys    ' numericArray
            End If
        ElseIf LCase(stringOrNumeric) = "string" Then
            If useSort Then
                Step4 = ShellSort(outDic.keys)
            Else
                Step4 = outDic.keys
            End If
            ReDim Step3(LBound(Step4) To UBound(Step4))
            For i = LBound(Step4) To UBound(Step4)
                Step3(i) = CStr(Step4(i))
            Next i
            Expand = Step3               ' stringArray
        ElseIf LCase(stringOrNumeric) = "dic" Then
            Set Expand = outDic          ' dictionary
        End If
    End If
End Function
