' ��� �������� ��������
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Dim i As Integer               '������� ������
Dim fX As Single, fY As Single '������� �����
Dim fW As Single, fH As Single '������/������ ��������
 
' ������ ��� �������� i, fX � fY, ��������� ��� ��� ���������� fW � fH ��� ���������� ��� �������� ������� �������� ��������. ������� ���������� ���������� ����� ��� � � ���������� �������:
 
Private Sub Form_Resize()
 
  '�������� ������ � fX � fY, ����� �������� _
  ���� �� ������ �����
  fX = Round((Me.ScaleWidth - imgPic(0).Width) / 2)
  fY = Round((Me.ScaleHeight - imgPic(0).Height) / 2)
 
End Sub
 
������ �������� ��������� � ������� Form_Load:
 
Private Sub Form_Load()
 
  '������������� ������ �������� �������� _
  � ����������� �� ������� ������ ��������
  fW = imgPic(0).Width
  fH = imgPic(0).Height
 
End Sub
 
�����, ��������� ��� ���������� � ������:
 
Private Sub Timer1_Timer()
 
  Me.Cls
 
  picMask.Picture = imgMask(i).Picture
  picPic.Picture = imgPic(i).Picture
 
  '������� ������ �����
  BitBlt Me.hDC, fX, fY, fW, fH, _
  picMask.hDC, 0, 0, vbMergePaint
 
  '����� ��������
  BitBlt Me.hDC, fX, fY, fW, fH, _
  picPic.hDC, 0, 0, vbSrcAnd
 
  '��������� �����, � �� ������ �� ����� �����
  Me.Refresh
 
  '���������� �������
  i = i + 1
  '���� �������� �������� ������ ���������� ������ _
  �� ���������� ���
  If i > (imgPic.Count - 1) Then i = 0
 
End Sub