Function InvertSelection(inputArea As Variant) As Variant
    Dim firstRow As Long, lastRow As Long, firstColumn As Long, lastColumn As Long
    Dim mergedRange As Range, finalRange As Range, ar As Range
    Dim elementsOfRange() As Range
    Dim i As Integer, cnt As Integer
    Dim processRange As Range
    
    ' /////// ���������� ///////
    maxRowGlobal = 65536
    maxColGlobal = 256
    If CDbl(Application.Version) >= 12 Then ' ����������
        If Not ActiveWorkbook Is Nothing Then
            If Not ActiveWorkbook.Excel8CompatibilityMode Then
                maxRowGlobal = 1048576
                maxColGlobal = 16384
            End If
        End If
    End If
    ' \\\\\\\ ���������� \\\\\\\

    
    ' �������� ���� - ������ �������� �������� ������� ���������� ������� � �� ��� ��������� ������ ���� "�����"
    ' ����� ��� ����� ����� � ���-���� ��������
    If inputArea Is Nothing Then ' ���� �� ���� ���� ������ ��������
        Set InvertSelection = Cells
        Exit Function
    End If
    
    InvertSelection = inputArea ' �� ������ �� ��� �� �����
    If TypeName(inputArea) = "Range" Then
        Set processRange = inputArea
    Else
        processRange = Range(inputArea)
    End If
    
    Set finalRange = Cells                 ' ���� ��������
    For Each ar In processRange.Areas ' ���� �� ��������
        cnt = 0
        With ar
            firstRow = .row
            lastRow = .row + .Rows.count - 1
            firstColumn = .Column
            lastColumn = .Column + .Columns.count - 1

            Set mergedRange = Nothing
            If firstRow > 1 Then ' ������ ������ 1 - �������� "���"
                cnt = cnt + 1
                ReDim Preserve elementsOfRange(1 To cnt)
                Set elementsOfRange(cnt) = Range(Cells(1, 1), Cells(firstRow - 1, maxColGlobal))
            End If
            If lastRow < maxRowGlobal Then ' ������ ������ ����������� - �������� "���"
                cnt = cnt + 1
                ReDim Preserve elementsOfRange(1 To cnt)
                Set elementsOfRange(cnt) = Range(Cells(lastRow + 1, 1), Cells(maxRowGlobal, maxColGlobal))
            End If
            If firstColumn > 1 Then ' ������� ������ 1 - ������� "�����"
                cnt = cnt + 1
                ReDim Preserve elementsOfRange(1 To cnt)
                Set elementsOfRange(cnt) = Range(Cells(1, 1), Cells(maxRowGlobal, firstColumn - 1))
            End If
            If lastColumn < maxColGlobal Then ' ������� ������ ����������� - ������� "������"
                cnt = cnt + 1
                ReDim Preserve elementsOfRange(1 To cnt)
                Set elementsOfRange(cnt) = Range(Cells(1, lastColumn + 1), Cells(maxRowGlobal, maxColGlobal))
            End If
            Set mergedRange = elementsOfRange(1)
            For i = 2 To cnt
                Set mergedRange = Union(mergedRange, elementsOfRange(i))
            Next i
            ' ���������� ���� �������� � ���������� ������������
            Set finalRange = Intersect(finalRange, mergedRange) ' �������� ��� ������� � ����
        End With
    Next ar
    
    If TypeName(inputArea) = "Range" Then
        Set InvertSelection = finalRange
    Else
        InvertSelection = finalRange.Address
    End If

End Function