''' ������������ n-������ ������ � 1-D ������
'REQUIRES: arrayDepth, arrayLength
Function getFlatArray(arr As Variant, Optional origin As Integer = 0) As Variant
    Dim outArr As Variant
    Dim axisNum As Long
    Dim numOfEls As Long
    Dim i As Long
    Dim el As Variant
    Dim cnt As Long
    
    '������� ������� ������� � ���������� ��������� � �������
    axisNum = arrayDepth(arr)
    numOfEls = 1
    For i = 1 To axisNum
        numOfEls = numOfEls * arrayLength(arr, CByte(i))
    Next i
    
    ' ������ ������ � ��������� ������
    ReDim outArr(origin To origin + numOfEls - 1)
    cnt = origin - 1
    For Each el In arr
        cnt = cnt + 1
        outArr(cnt) = el
    Next el
    
    getFlatArray = outArr
    
End Function