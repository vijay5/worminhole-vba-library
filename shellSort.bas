'��������� ��� ���������� ������� ������� �����
Function ShellSort(inArr As Variant) As Variant
    Dim arr As Variant
    Dim chk As Boolean
    Dim step As Long
    Dim i As Long
    Dim j As Long
    Dim n As Long
    Dim tmp As Variant
    
    arr = inArr ' �������� �������� ������ (����� ���� ����� ByVal �������)
    n = arrayLength(arr) ' ����� �������
    If n = 0 Then ' �������� �� ����������
        ShellSort = ""
        Exit Function
    End If
    
    step = n \ 2 ' ��� (������������� ������������ ������, ���� ����������� - http://ru.wikipedia.org/wiki/����������_�����)
    Do ' ��������� ����
        i = step  ' ���� ��� ��������� (�� ��� ����������� ����� � ������ �������)
        Do ' ������ ���� (�������)
            j = i - step + LBound(arr) ' ����� �������
            chk = True
            Do ' ������ ���� (����������)
                If arr(j) <= arr(j + step) Then ' �������� ����������� - ������ �� ������
                    chk = False
                Else ' �������� �� ����������� - ������ �������
                    tmp = arr(j)
                    arr(j) = arr(j + step)
                    arr(j + step) = tmp
                End If
                j = j - 1
            Loop Until (chk = False) Or (j < LBound(arr))
            i = i + 1
        Loop Until i = n
        step = step \ 2 ' ��������� ��� �����
    Loop Until step = 0
    ShellSort = arr
End Function