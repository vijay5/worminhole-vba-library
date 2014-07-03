''' Сортировка Шелла для двумерных массивов
Function ShellSort2D(inArr As Variant, colNum As Long, Optional isAscending As Boolean = True) As Variant
    Dim arr As Variant
    Dim chk As Boolean
    Dim step As Long
    Dim i As Long
    Dim j As Long
    Dim n As Long
    Dim l As Long
    Dim tmp As Variant
    
    arr = inArr ' копируем исходный массив (можно было через ByVal сделать)
    n = UBound(arr, 1) - LBound(arr, 1) + 1 ' длина массива
    If n = 0 Then ' проверка на косячность
        ShellSort2D = ""
        Exit Function
    End If
    
    step = n \ 2 ' шаг (первоначально предложенный Шеллом, есть модификации - http://ru.wikipedia.org/wiki/Сортировка_Шелла)
    Do ' изменение шага
        i = step  ' база для сравнения (от нее вычисляется левый и правый элемент)
        Do ' первый цикл (внешний)
            j = i - step + LBound(arr, 1) ' левый элемент
            chk = True
            Do ' второй цикл (внутренний)
                If isAscending Then ' по возрастанию
                    If arr(j, colNum) <= arr(j + step, colNum) Then ' элементы упорядочены - ничего не делаем
                        chk = False
                    Else ' элементы не упорядочены - меняем порядок
                        For l = LBound(arr, 2) To UBound(arr, 2) ' обмен записей по всем столбцам
                            tmp = arr(j, l)
                            arr(j, l) = arr(j + step, l)
                            arr(j + step, l) = tmp
                        Next l
                    End If
                Else ' по убыванию
                    If arr(j, colNum) >= arr(j + step, colNum) Then ' элементы упорядочены - ничего не делаем
                        chk = False
                    Else ' элементы не упорядочены - меняем порядок
                        For l = LBound(arr, 2) To UBound(arr, 2) ' обмен записей по всем столбцам
                            tmp = arr(j, l)
                            arr(j, l) = arr(j + step, l)
                            arr(j + step, l) = tmp
                        Next l
                    End If
                End If
                j = j - 1
            Loop Until (chk = False) Or (j < LBound(arr, 1))
            i = i + 1
        Loop Until i = n
        step = step \ 2 ' уменьшаем шаг вдвое
    Loop Until step = 0
    ShellSort2D = arr
End Function
