''' ���������� ����� ��� ��������� ��������
Function ShellSort2D(inArr As Variant, colNum As Long, Optional isAscending As Boolean = True) As Variant
    Dim arr As Variant
    Dim chk As Boolean
    Dim step As Long
    Dim i As Long
    Dim j As Long
    Dim n As Long
    Dim l As Long
    Dim tmp As Variant
    
    arr = inArr ' �������� �������� ������ (����� ���� ����� ByVal �������)
    n = UBound(arr, 1) - LBound(arr, 1) + 1 ' ����� �������
    If n = 0 Then ' �������� �� ����������
        ShellSort2D = ""
        Exit Function
    End If
    
    step = n \ 2 ' ��� (������������� ������������ ������, ���� ����������� - http://ru.wikipedia.org/wiki/����������_�����)
    Do ' ��������� ����
        i = step  ' ���� ��� ��������� (�� ��� ����������� ����� � ������ �������)
        Do ' ������ ���� (�������)
            j = i - step + LBound(arr, 1) ' ����� �������
            chk = True
            Do ' ������ ���� (����������)
                If isAscending Then ' �� �����������
                    If arr(j, colNum) <= arr(j + step, colNum) Then ' �������� ����������� - ������ �� ������
                        chk = False
                    Else ' �������� �� ����������� - ������ �������
                        For l = LBound(arr, 2) To UBound(arr, 2) ' ����� ������� �� ���� ��������
                            tmp = arr(j, l)
                            arr(j, l) = arr(j + step, l)
                            arr(j + step, l) = tmp
                        Next l
                    End If
                Else ' �� ��������
                    If arr(j, colNum) >= arr(j + step, colNum) Then ' �������� ����������� - ������ �� ������
                        chk = False
                    Else ' �������� �� ����������� - ������ �������
                        For l = LBound(arr, 2) To UBound(arr, 2) ' ����� ������� �� ���� ��������
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
        step = step \ 2 ' ��������� ��� �����
    Loop Until step = 0
    ShellSort2D = arr
End Function
