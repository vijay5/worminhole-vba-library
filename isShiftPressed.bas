' ��������� ������� ������� Shift � ������ ������� �������
Public Function isShiftPressed() As Boolean
    isShiftPressed = (GetKeyState(&H10) <= -127)
End Function
