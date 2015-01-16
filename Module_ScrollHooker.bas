Option Explicit

'''''''' end normal module code
''''' Userform code
'Private Sub xxx_MouseMove(???)
'    TurnHookOn Me, Me.ComboBox1
'End Sub
'
' ��� �������� ����� ����������� ������� hook - ����� � ������ �� ��������
' � �� �����, �����-������ ������� � ��� �� hwnd ��������� � ����� ������� ��� ��������� ������ :)
'Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'    TurnHookOff
'End Sub
''''''' end Userform code


Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Type MOUSEHOOKSTRUCT
        pt As POINTAPI
        hwnd As Long
        wHitTestCode As Long
        dwExtraInfo As Long
End Type

Public Const WH_MOUSE_LL As Long = 14
Public Const WM_MOUSEWHEEL As Long = &H20A
Public Const HC_ACTION As Long = 0
Public Const GWL_HINSTANCE As Long = (-6)

Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const VK_UP As Long = &H26
Private Const VK_DOWN As Long = &H28
Private Const WM_LBUTTONDOWN As Long = &H201

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
' ��� ���������, � ������� ��� ���������� ���������
Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private mHookNumber As Long          ' ����� Hook'� � ������ Hook'��
Private mControlHwnd As Long         ' hwnd ��������
Private mHookIsActive As Boolean     ' ������� Hook'�
Private mControl As MSForms.Control  ' ������ �� ������� �����
Dim n As Long

' ������ Hook
Sub TurnHookOn(frm As Object, ctl As MSForms.Control)
    Dim lngAppInst As Long
    Dim hwndUnderCursor As Long
    Dim tPoint As POINTAPI
    
    ' � ������� ������ ������ ���� ��� hook
    ' ���� �� ��������� �� ����� ��������, ��� �������� ��������� hook
    ' � ����� ������ �� ������ �������, �� ������� ������ ���� hook
    ' �� ������� ������ hook � ������ ����� hook
    
    GetCursorPos tPoint ' ������� ������� �������
    hwndUnderCursor = WindowFromPoint(tPoint.X, tPoint.Y) ' hwnd ��� ��������
    If Not frm.ActiveControl Is ctl Then ' ������ ����� �� �������
        ctl.SetFocus
    End If
    
    If mControlHwnd <> hwndUnderCursor Then ' ���� hwnd ��� �������� <> hwnd �������������� hook'�
        TurnHookOff                         ' ������� ������� (������) hook
        Set mControl = ctl
        mControlHwnd = hwndUnderCursor      ' ����� hwnd ����
        lngAppInst = GetWindowLong(mControlHwnd, GWL_HINSTANCE)
        If Not mHookIsActive Then
            mHookNumber = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseProc, lngAppInst, 0) ' ������ ����� hook
            mHookIsActive = mHookNumber <> 0 ' ��������, ��� hook ����������
        End If
    End If
End Sub

' ������� Hook
Sub TurnHookOff()
    If mHookIsActive Then
        Set mControl = Nothing
        UnhookWindowsHookEx mHookNumber ' ������� ������ hook �� ������
        mHookNumber = 0
        mControlHwnd = 0
        mHookIsActive = False
    End If
End Sub

' ���������, ������� ����������� Hook'��
' �� ���� ����������, ����������� �� �������, ��� hook'��
' ������ ������� breakpoint'� � ��������� ������.
' ����� ��������� � ���� ��������� � 99% ������� ������� ������ � ��������� ���� ������ �������������� (!!!)
' ����� ����������� ������ ����������� :)
Private Function MouseProc(ByVal nCode As Long, ByVal wParam As Long, ByRef lParam As MOUSEHOOKSTRUCT) As Long
    Dim direction As Long
    Dim newScrollPos As Single
    Dim maxScrollHeight As Single
    
    On Error GoTo errH
    If (nCode = HC_ACTION) Then
        If WindowFromPoint(lParam.pt.X, lParam.pt.Y) = mControlHwnd Then ' �������� hwnd ���� �� �����������
            If wParam = WM_MOUSEWHEEL Then ' ���� ��������/������� ����������� ������� ����� (?)
                MouseProc = True
                
                direction = -Sgn(lParam.hwnd) ' ���������� ����������� ������ ������
                ' � ����������� �� ���� ������� �������� ������ ��������
                Select Case TypeName(mControl)
                Case "MultiPage" ' ������ ������������
                    ' ��� ��������������� ���������� �� .ScrollLeft � .ScrollWidth
                    newScrollPos = mControl.Item(mControl.Value).ScrollTop + direction * 24 ' ����� �� 24 ������
                    maxScrollHeight = mControl.Item(mControl.Value).ScrollHeight - mControl.Height + 17.25 ' 17.25 - ��� ������ ����������, �������� �
                    
                    If newScrollPos < 0 Then newScrollPos = 0
                    If newScrollPos > maxScrollHeight Then newScrollPos = maxScrollHeight
                    
                    mControl.Item(mControl.Value).ScrollTop = newScrollPos
                    
                Case "ListBox" ' ������ �����
                    newScrollPos = mControl.ListIndex + direction * 3
                    
                    If newScrollPos >= 0 And newScrollPos <= mControl.ListCount - 1 Then
                        mControl.ListIndex = newScrollPos
                    ElseIf newScrollPos < 0 Then
                        mControl.ListIndex = 0
                    ElseIf newScrollPos > mControl.ListCount - 1 Then
                        mControl.ListIndex = mControl.ListCount - 1
                    End If
                    
'                Case "MultiPage" ' ������ ����� ����������
'                    newScrollPos = mControl.value + direction * 1
'
'                    If newScrollPos >= 0 And newScrollPos < mControl.Pages.Count Then
'                        mControl.value = newScrollPos
'                    End If

                Case "TextBox" ' ��� multiline'� �������
                    If direction < 0 Then
                        PostMessage mControlHwnd, WM_KEYDOWN, VK_UP, 0   ' �������� ������ �����
                    Else
                        PostMessage mControlHwnd, WM_KEYDOWN, VK_DOWN, 0 ' �������� ������ ����
                    End If
                    
                    PostMessage mControlHwnd, WM_KEYUP, VK_UP, 0 ' ��������� ������
                
                End Select
                
                Exit Function
            End If
        Else ' �� ������� �������� hwnd ���� �� �����
            TurnHookOff
        End If
    End If
    ' ������� ���������� ���������� hook'� � �������
    MouseProc = CallNextHookEx(mHookNumber, nCode, wParam, ByVal lParam)
    Exit Function
errH:
    TurnHookOff
End Function


