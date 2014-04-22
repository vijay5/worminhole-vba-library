Sub makeDropDownList(targetRng As Range, sourceDataRng As Range, Optional ignoreBlank As Boolean = True, Optional showError As Boolean = True)
    ''' ��������� ���������� ������ � ��������� ������ (2003 compatible)
    ''' ������ �������� ������� �� sourceDataRng (��������, ���� ����������� �� ���-�� �����/��������, �.�. ������ ���� ���� ���� ������, ���� ���� �������)
    Dim shName As String
    Dim firstCellRow As String, firstCellCol As String
    Dim lastCellRow As String, lastCellCol As String
    
    If targetRng Is Nothing Or sourceDataRng Is Nothing Then Exit Sub ' �������
    
    shName = sourceDataRng.Parent.Name
    firstCellRow = CStr(sourceDataRng.Cells(1, 1).Row)
    firstCellCol = CStr(sourceDataRng.Cells(1, 1).Column)
    lastCellRow = CStr(sourceDataRng.Cells(1, 1).Row + sourceDataRng.Rows.Count - 1)
    lastCellCol = CStr(sourceDataRng.Cells(1, 1).Column + sourceDataRng.Columns.Count - 1)
    
    If CDbl(Application.Version) < 12 Then
        targetRng.Cells(1, 1).Select ' ��� ������������� � 2003 �������, ���������� �������� ������ ����� ���������� ���������, ���...
    Else
        targetRng.Select
    End If
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
         Formula1:="=INDIRECT(ADDRESS(" + firstCellRow + "," + firstCellCol + ",,,""" + shName + """)&"":""&ADDRESS(" + lastCellRow + "," + lastCellCol + "))"
        .ignoreBlank = ignoreBlank ' ������������ ��������
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "������!"
        .InputMessage = ""
        .ErrorMessage = "������� �������� ��������. �������� �������� �� ����������� ������!"
        .ShowInput = True
        .showError = showError
    End With
End Sub