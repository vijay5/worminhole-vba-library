''' ���������� ��� ��������
''' ������ �� TextBox, yesLabel � noLabel
''' ������� yesLabel � noLabel �������� �������� � ����������� �� ���������� ��������
''' keySwitch - ����, ���� keySwitch = False, �� yesLabel � noLabel - ��� ���������
Sub yesNoSwitcher(value_in As Variant, yesLabel As Object, noLabel As Object, Optional globalSwitch As Boolean = True)
    Dim value As Integer
    Dim tmp As Variant
    
    value = 0 ' ��-���������, �� ����������, �� ��, �� ������
    
    If globalSwitch Then ' ���� label'� ������ ���� ����������
        ' �������� ������, ��� ������ �� ����
        On Error Resume Next
            If IsObject(value_in) Then
                value = IIf(CBool(value_in.value), 1, -1) ' �������� �������� �������� �������
            ElseIf TypeName(value_in) = "Boolean" Then
                value = IIf(value_in, 1, -1)
            ElseIf TypeName(value_in) = "Integer" Then
                value = value_in
            End If
        On Error GoTo 0
    End If
    
    Select Case value
    Case 1   ' ��������
        yesLabel.Visible = True
        noLabel.Visible = False
    Case -1  ' ���������
        yesLabel.Visible = False
        noLabel.Visible = True
    Case Else '0   ' ��������
        yesLabel.Visible = False
        noLabel.Visible = False
    End Select
End Sub
