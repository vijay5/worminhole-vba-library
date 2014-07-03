''' ������ ��������� ��������� ��������� Enabled/Disabled �� �������� � chk (CheckBox)
Sub chkBoxReactor(chk As Variant, changeVisibility As Boolean, ParamArray manyObjecs() As Variant)
    Dim chkValue As Variant
    Dim obj As Variant
    
    If IsObject(chk) Then
        chkValue = chk.value
    Else
        chkValue = chk
    End If
    
    For Each obj In manyObjecs ' ������� ������ ��������
        If changeVisibility Then ' ��������� ��������/���������� ������
            obj.Visible = chkValue
        Else                     ' ������ ��� Enabled/Disabled
            obj.Enabled = chkValue
            Select Case TypeName(obj)
            Case "CheckBox"
                ' pass
            Case "TextBox", "ComboBox", "ListBox" ' ��������� - "����������" ���� �����, if any
                obj.BackColor = IIf(chkValue, &H80000005, &H80000004)
            Case "Label"
                ' pass
            Case "CommandButton"
                ' pass
            Case Else
                ' pass
            End Select
        End If
        
    Next obj
End Sub
