''' ��������� ���������� ������ � ��������� ������ (2003 compatible)
''' ������ �������� ������� �� sourceDataRng (��������, ���� ����������� �� ���-�� �����/��������, �.�. ������ ���� ���� ���� ������, ���� ���� �������)
' REQUIRES: col2Array, array2col
Sub makeDropDownList(targetRng As Range, sourceData As Variant, Optional ignoreBlank As Boolean = True, Optional showError As Boolean = True)
    Dim shName As String
    Dim sourceDataRng As Range
    Dim firstCellRow As String, firstCellCol As String
    Dim lastCellRow As String, lastCellCol As String
    Dim isRange As Boolean
    Dim validObj As Validation
    Dim oldSelectionAddr As String
    Dim sourceDataCol As Collection
    Dim sourceDataArr As Variant
    Dim formulaStr As String
    Dim i As Long
    
    If targetRng Is Nothing Then Exit Sub ' �������
    
    isRange = False
    Select Case TypeName(sourceData)
    Case "Range"
        Set sourceDataRng = sourceData
        isRange = True
    Case "Collection"
        Set sourceDataCol = sourceData
        sourceDataArr = col2Array(sourceDataCol)
    Case "Variant()", "String()", "Integer()", "Single()", "Long()", "Double()"
        sourceDataArr = col2Array(array2col(sourceData))
    Case Else
        Exit Sub
    End Select
    
    
    If isRange Then ' ����� ��������
        shName = sourceDataRng.Parent.Name
        firstCellRow = CStr(sourceDataRng.Cells(1, 1).row)
        firstCellCol = CStr(sourceDataRng.Cells(1, 1).Column)
        lastCellRow = CStr(sourceDataRng.Cells(1, 1).row + sourceDataRng.Rows.Count - 1)
        lastCellCol = CStr(sourceDataRng.Cells(1, 1).Column + sourceDataRng.Columns.Count - 1)
        
        formulaStr = "=INDIRECT(ADDRESS(" + firstCellRow + "," + firstCellCol + ",,,""" + shName + """)&"":""&ADDRESS(" + lastCellRow + "," + lastCellCol + "))"
    Else ' ����� ������
        
        ReDim sourceDataArrStr(LBound(sourceDataArr) To UBound(sourceDataArr))
        For i = LBound(sourceDataArr) To UBound(sourceDataArr)
            sourceDataArrStr(i) = CStr(sourceDataArr(i))
        Next i
        
        formulaStr = Join(sourceDataArrStr, ",")
        
    End If
    
    ' ������������� �� ������ �������
    If CDbl(Application.Version) < 12 Then
        oldSelectionAddr = Selection.Address
        targetRng.Cells(1, 1).Select ' ��� ������������� � 2003 �������, ���������� �������� ������ ����� ���������� ���������, ���...
        Set validObj = Selection.Validation
    Else
        oldSelectionAddr = ""
        Set validObj = targetRng.Cells(1, 1).Validation
    End If
    
    
    With validObj
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
         Formula1:=formulaStr
        .ignoreBlank = ignoreBlank ' ������������ ��������
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "������!"
        .InputMessage = ""
        .ErrorMessage = "������� �������� ��������. �������� �������� �� ����������� ������!"
        .ShowInput = True
        .showError = showError
    End With
    
    
    
    If oldSelectionAddr <> "" Then ' ��������������� ����������� ������
        Range(oldSelectionAddr).Select
    End If
End Sub
