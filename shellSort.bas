'Процедура для сортировки массива методом Шелла
Function ShellSort(inArr As Variant) As Variant
    Dim arr As Variant
    Dim chk As Boolean
    Dim step As Long
    Dim i As Long
    Dim j As Long
    Dim n As Long
    Dim tmp As Variant
    
    arr = inArr ' копируем исходный массив (можно было через ByVal сделать)
    n = arrayLength(arr) ' длина массива
    If n = 0 Then ' проверка на косячность
        ShellSort = ""
        Exit Function
    End If
    
    step = n \ 2 ' шаг (первоначально предложенный Шеллом, есть модификации - http://ru.wikipedia.org/wiki/Сортировка_Шелла)
    Do ' изменение шага
        i = step  ' база для сравнения (от нее вычисляется левый и правый элемент)
        Do ' первый цикл (внешний)
            j = i - step + LBound(arr) ' левый элемент
            chk = True
            Do ' второй цикл (внутренний)
                If arr(j) <= arr(j + step) Then ' элементы упорядочены - ничего не делаем
                    chk = False
                Else ' элементы не упорядочены - меняем порядок
                    tmp = arr(j)
                    arr(j) = arr(j + step)
                    arr(j + step) = tmp
                End If
                j = j - 1
            Loop Until (chk = False) Or (j < LBound(arr))
            i = i + 1
        Loop Until i = n
        step = step \ 2 ' уменьшаем шаг вдвое
    Loop Until step = 0
    ShellSort = arr
End Function