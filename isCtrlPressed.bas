''' ��������� ������� ������� Ctrl � ������ ������� �������
Public Function isCtrlPressed() As Boolean
    isCtrlPressed = (GetKeyState(&H11) <= -127)
End Function
