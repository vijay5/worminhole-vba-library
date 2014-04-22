' ����������� ������� � �������� ������� ������/������ ������
Function pix2Point(pixels As Integer, Optional forColumn As Boolean = True) As Single
    If pixels >= 1 Then
        If forColumn Then  ' ������� ��� ��������
            If pixels >= 12 Then
                pix2Point = Round((pixels - 5) / 7, 2)
            Else
                pix2Point = Round((pixels) / 12, 2)
            End If
        Else               ' ������� ��� �����
            pix2Point = Round(pixels * 0.75, 2)
        End If
    Else
        pix2Point = 0 ' ������� ������� / ������
    End If
End Function