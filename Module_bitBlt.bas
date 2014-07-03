' дл€ анимации картинок
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Dim i As Integer               'счетчик кадров
Dim fX As Single, fY As Single 'позици€ ввода
Dim fW As Single, fH As Single 'ширина/высота картинки
 
' ѕомимо уже знакомых i, fX и fY, по€вились еще две переменные fW и fH они необходимы дл€ указани€ размера рисуемой картинки. ѕозицию прорисовки определ€ем также как и в предыдущем способе:
 
Private Sub Form_Resize()
 
  'забиваем данные в fX и fY, чтобы картинка _
  была по центру формы
  fX = Round((Me.ScaleWidth - imgPic(0).Width) / 2)
  fY = Round((Me.ScaleHeight - imgPic(0).Height) / 2)
 
End Sub
 
–азмер картинок установим в событие Form_Load:
 
Private Sub Form_Load()
 
  'устанавливаем размер рисуемой картинки _
  в зависимости от размера первой картинки
  fW = imgPic(0).Width
  fH = imgPic(0).Height
 
End Sub
 
«атем, добавл€ем код прорисовки в таймер:
 
Private Sub Timer1_Timer()
 
  Me.Cls
 
  picMask.Picture = imgMask(i).Picture
  picPic.Picture = imgPic(i).Picture
 
  '—начала рисуем маску
  BitBlt Me.hDC, fX, fY, fW, fH, _
  picMask.hDC, 0, 0, vbMergePaint
 
  '«атем картинку
  BitBlt Me.hDC, fX, fY, fW, fH, _
  picPic.hDC, 0, 0, vbSrcAnd
 
  'ќбновл€ем форму, а то ничего не будет видно
  Me.Refresh
 
  'ѕрибавл€ем счетчик
  i = i + 1
  '≈сли значение счетчика больше количества кадров _
  то аннулируем его
  If i > (imgPic.Count - 1) Then i = 0
 
End Sub