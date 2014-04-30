''' проверяем нажатие клавиши Ctrl в момент запуска функции
Public Function isCtrlPressed() As Boolean
    isCtrlPressed = (GetKeyState(&H11) <= -127)
End Function
