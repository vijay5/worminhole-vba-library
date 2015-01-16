Option Explicit

'''''''' end normal module code
''''' Userform code
'Private Sub xxx_MouseMove(???)
'    TurnHookOn Me, Me.ComboBox1
'End Sub
'
' при закрытии формы обязательно убиваем hook - чтобы в памяти не болтался
' а то вдруг, какой-нибудь контрол с тем же hwnd откроется и будет сюрприз при прокрутке мышкой :)
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
' для контролов, у которых нет встроенной прокрутки
Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private mHookNumber As Long          ' номер Hook'а в списке Hook'ов
Private mControlHwnd As Long         ' hwnd контрола
Private mHookIsActive As Boolean     ' наличие Hook'а
Private mControl As MSForms.Control  ' ссылка на элемент формы
Dim n As Long

' ставим Hook
Sub TurnHookOn(frm As Object, ctl As MSForms.Control)
    Dim lngAppInst As Long
    Dim hwndUnderCursor As Long
    Dim tPoint As POINTAPI
    
    ' в очереди всегда только один наш hook
    ' если мы находимся на одном контроле, для которого поставили hook
    ' а потом уходим на другой контрол, на котором должен быть hook
    ' мы убиваем старый hook и ставим новый hook
    
    GetCursorPos tPoint ' текущая позиция курсора
    hwndUnderCursor = WindowFromPoint(tPoint.X, tPoint.Y) ' hwnd под курсором
    If Not frm.ActiveControl Is ctl Then ' ставим фокус на контрол
        ctl.SetFocus
    End If
    
    If mControlHwnd <> hwndUnderCursor Then ' если hwnd под курсором <> hwnd установленного hook'а
        TurnHookOff                         ' снимаем текущий (старый) hook
        Set mControl = ctl
        mControlHwnd = hwndUnderCursor      ' узнаём hwnd окна
        lngAppInst = GetWindowLong(mControlHwnd, GWL_HINSTANCE)
        If Not mHookIsActive Then
            mHookNumber = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseProc, lngAppInst, 0) ' ставим новый hook
            mHookIsActive = mHookNumber <> 0 ' проверям, что hook поставился
        End If
    End If
End Sub

' снимаем Hook
Sub TurnHookOff()
    If mHookIsActive Then
        Set mControl = Nothing
        UnhookWindowsHookEx mHookNumber ' снимаем старый hook по номеру
        mHookNumber = 0
        mControlHwnd = 0
        mHookIsActive = False
    End If
End Sub

' процедура, которая запускается Hook'ом
' Во всех процедурах, запускаемых по таймеру, или hook'ом
' нельзя ставить breakpoint'ы и допускать ошибок.
' Любая остановка в этой процедуре в 99% случаев убивает Эксель с потрохами безо всяких предупреждений (!!!)
' Перед изменениями всегда сохраняемся :)
Private Function MouseProc(ByVal nCode As Long, ByVal wParam As Long, ByRef lParam As MOUSEHOOKSTRUCT) As Long
    Dim direction As Long
    Dim newScrollPos As Single
    Dim maxScrollHeight As Single
    
    On Error GoTo errH
    If (nCode = HC_ACTION) Then
        If WindowFromPoint(lParam.pt.X, lParam.pt.Y) = mControlHwnd Then ' получили hwnd окна по координатам
            If wParam = WM_MOUSEWHEEL Then ' если действие/событие произведено колесом мышки (?)
                MouseProc = True
                
                direction = -Sgn(lParam.hwnd) ' определяем направление сдвига колеса
                ' в зависимости от типа объекта изменяем разные свойства
                Select Case TypeName(mControl)
                Case "MultiPage" ' скролл вертикальный
                    ' для горизонтального переделать на .ScrollLeft и .ScrollWidth
                    newScrollPos = mControl.Item(mControl.Value).ScrollTop + direction * 24 ' сдвиг на 24 поинта
                    maxScrollHeight = mControl.Item(mControl.Value).ScrollHeight - mControl.Height + 17.25 ' 17.25 - это высота закладочки, вычитаем её
                    
                    If newScrollPos < 0 Then newScrollPos = 0
                    If newScrollPos > maxScrollHeight Then newScrollPos = maxScrollHeight
                    
                    mControl.Item(mControl.Value).ScrollTop = newScrollPos
                    
                Case "ListBox" ' скролл строк
                    newScrollPos = mControl.ListIndex + direction * 3
                    
                    If newScrollPos >= 0 And newScrollPos <= mControl.ListCount - 1 Then
                        mControl.ListIndex = newScrollPos
                    ElseIf newScrollPos < 0 Then
                        mControl.ListIndex = 0
                    ElseIf newScrollPos > mControl.ListCount - 1 Then
                        mControl.ListIndex = mControl.ListCount - 1
                    End If
                    
'                Case "MultiPage" ' скролл между страницами
'                    newScrollPos = mControl.value + direction * 1
'
'                    If newScrollPos >= 0 And newScrollPos < mControl.Pages.Count Then
'                        mControl.value = newScrollPos
'                    End If

                Case "TextBox" ' для multiline'а подойдёт
                    If direction < 0 Then
                        PostMessage mControlHwnd, WM_KEYDOWN, VK_UP, 0   ' нажимаем кнопку вверх
                    Else
                        PostMessage mControlHwnd, WM_KEYDOWN, VK_DOWN, 0 ' нажимаем кнопку вниз
                    End If
                    
                    PostMessage mControlHwnd, WM_KEYUP, VK_UP, 0 ' отпускаем кнопку
                
                End Select
                
                Exit Function
            End If
        Else ' не удалось получить hwnd окна по точке
            TurnHookOff
        End If
    End If
    ' передаём управление следующему hook'у в очереди
    MouseProc = CallNextHookEx(mHookNumber, nCode, wParam, ByVal lParam)
    Exit Function
errH:
    TurnHookOff
End Function


