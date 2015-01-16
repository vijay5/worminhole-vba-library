'' �������� ���������� ������ ListBox ����� ��� ����
' REQUIRES: getListOfSelected, col2Array, arrayLength, Collapse, MatrixPart
Private Sub cmb_DownArrow_Click() ' �������� ������ ����
    Dim listOfSelected As Variant, listLength As Variant
    Dim tmp As Variant
    Dim i As Long
    Dim j As Long
    Dim lb As MSForms.listBox
    Dim direction As String
    Dim minPos As Long, maxPos As Long
    Dim stepAmount As Long
    
    ' ///��������� �������� �������� � ����������� ������
    Set lb = Me.lb_destList ' �������� ��������
    direction = "Down"      ' ����������� ������
    ' ///
    
    listOfSelected = col2Array(getListOfSelected(lb))
    listLength = arrayLength(listOfSelected)
    If direction = "Down" Then
        minPos = listLength
        maxPos = 1
        stepAmount = -1
    Else
        minPos = 1
        maxPos = listLength
        stepAmount = 1
    End If
    
    If listLength > 0 Then
        tmp = Collapse(listOfSelected)
        If InStr(tmp, ".") = 0 Then ' ������, �������� ��������
            If direction = "Down" And (listOfSelected(listLength) < lb.ListCount - 1) Or _
               direction = "Up" And (listOfSelected(1) > 0) Then ' ����� �������� �� ������� ����
                
                ' ������ ���/��� ����������
                tmpRow = MatrixPart(lb.List, listOfSelected(minPos) - stepAmount, listOfSelected(minPos) - stepAmount, 0, 9, , False)
                
                ' ������������ ���������� ������ �� 1 ����
                For i = minPos To maxPos Step stepAmount
                    For j = 0 To 9
                        If IsNull(lb.List(listOfSelected(i), j)) Then
                            lb.List(listOfSelected(i) - stepAmount, j) = ""
                        Else
                            lb.List(listOfSelected(i) - stepAmount, j) = lb.List(listOfSelected(i), j)
                        End If
                    Next j
                    lb.Selected(listOfSelected(i) - stepAmount) = True ' ��������� ���������
                Next i
                
                ' ������/��������� ������� ����������� ������ - ��������� ������� ������ (��������� � tmpRow)
                For j = 0 To 9
                    If IsNull(tmpRow(1, j + 1)) Then
                        lb.List(listOfSelected(maxPos), j) = ""
                    Else
                        lb.List(listOfSelected(maxPos), j) = tmpRow(1, j + 1)
                    End If
                Next j
                lb.Selected(listOfSelected(maxPos)) = False  ' ������� ���������
                
            Else ' �������� �� �����
            End If
        End If
    Else ' ������ ���������� ����
    End If
End Sub