''' включатель для объектов
''' ссылки на TextBox, yesLabel и noLabel
''' объекты yesLabel и noLabel делаются видимыми в зависимости от результата проверки
''' keySwitch - ключ, если keySwitch = False, то yesLabel и noLabel - оба выключены
Sub yesNoSwitcher(value_in As Variant, yesLabel As Object, noLabel As Object, Optional globalSwitch As Boolean = True)
    Dim value As Integer
    Dim tmp As Variant
    
    value = 0 ' по-умолчанию, не показываем, ни то, ни другое
    
    If globalSwitch Then ' если label'ы вообще надо показывать
        ' пытаемся понять, что подано на вход
        On Error Resume Next
            If IsObject(value_in) Then
                value = IIf(CBool(value_in.value), 1, -1) ' пытаемся получить значение объекта
            ElseIf TypeName(value_in) = "Boolean" Then
                value = IIf(value_in, 1, -1)
            ElseIf TypeName(value_in) = "Integer" Then
                value = value_in
            End If
        On Error GoTo 0
    End If
    
    Select Case value
    Case 1   ' включаем
        yesLabel.Visible = True
        noLabel.Visible = False
    Case -1  ' выключаем
        yesLabel.Visible = False
        noLabel.Visible = True
    Case Else '0   ' скрываем
        yesLabel.Visible = False
        noLabel.Visible = False
    End Select
End Sub
