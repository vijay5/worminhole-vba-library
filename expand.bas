''' ����������� ������ � ���������� ����� � ������, ��� ������������� ������ ���
''' �� ������ - 1) 0-based Array �� ���� ��� 2) 0-based Array �� ����� ��� 3) Dictionary
''' 175/190.207/209 -> (175, 176, 177, ..., 189, 190, 207, 208, 209)
Function Expand(sourceString As String, Optional stringOrNumeric As String = "numeric", Optional useSort As Boolean = True)
    Expand = False
    
' Stage 1 - ������� �������� ������, �������� ������������ ��������
    Dim Step1, Step2 As Variant
    Dim Step3 As Variant
    Dim Step4 As Variant
    Dim outDic As Variant
    Dim k As Long, j As Long, i As Long
    Dim tmp As Variant, brk As Variant
    Dim chk1 As Boolean, chk2 As Boolean
    Dim changedString As Variant
    Dim el As Variant
    
    Set outDic = CreateObject("Scripting.Dictionary")
    outDic.RemoveAll
    
    changedString = sourceString
    
    changedString = Replace(changedString, "   ", " ")
    changedString = Replace(changedString, "  ", " ")
    changedString = Replace(changedString, " , ", ".")
    changedString = Replace(changedString, ", ", ".")
    changedString = Replace(changedString, " ,", ".")
    changedString = Replace(changedString, " thru ", "/")
    changedString = Replace(changedString, " thr ", "/")
    changedString = Replace(changedString, " to ", "/")
    changedString = Replace(changedString, ",", ".")
    changedString = Replace(changedString, ";", ".")
    changedString = Replace(changedString, "`", "")
    changedString = Replace(changedString, ":", "/")
    changedString = Replace(changedString, "-", "/")
    changedString = Replace(changedString, "\", "/")
    changedString = Replace(changedString, "..", ".")
    changedString = Replace(changedString, "//", "/")
    changedString = Replace(changedString, " ", ".")
    
    ' ��������� �� ������� ����� ��������
    brk = reFind(LCase(changedString), "[^0-9a-z.\/]")
    
    If brk Or changedString = "" Then
        Expand = "" ' ���������� "�����"
    Else
        k = 0 ' ����� ��������� �������
        Step1 = Split(changedString, ".") ' ���� �� "������"
        For i = LBound(Step1) To UBound(Step1)
            If Step1(i) <> "" Then
                Step2 = Split(Step1(i), "/")  ' ���� �� "������"
                ' ����� ��������� (����� UBound > 0) ��� �� ��������� (����� UBound = 0)
                If UBound(Step2) = 1 Then ' ����� ���� 1 ������� - 122.134.156
                    chk1 = IsNumeric(Step2(0))
                    chk2 = IsNumeric(Step2(1))
                    If chk1 And chk2 Then ' ��� ��������
                        If CDbl(Step2(0)) > CDbl(Step2(1)) Then ' ������ ������� ������ � �����
                            tmp = Step2(0)
                            Step2(0) = Step2(1)
                            Step2(1) = tmp
                        End If
                    Else ' ���������� (������ �������, ��� ��� ����������)
                        If Step2(0) > Step2(1) Then ' ������ ������� ������ � �����
                            tmp = Step2(0)
                            Step2(0) = Step2(1)
                            Step2(1) = tmp
                        End If
                    End If
                ElseIf UBound(Step2) > 1 Then ' ������ ����: 122/126/140 - ������
                    Expand = ""
                    Exit Function
                Else ' UBound(Step2) =0
                    ' ������ ����� � ��� ������ ���� ������� ����� ��������� �� ������������
                    ' ���� ������� ����������� � Step4 "��� ����"
                End If
                ' ����� � ����� Step3 ��� ����� ���������
                If IsNumeric(Step2(LBound(Step2))) And IsNumeric(Step2(UBound(Step2))) Then
                    For j = CLng(Step2(LBound(Step2))) To CLng(Step2(UBound(Step2)))
                        If Not outDic.Exists(j) Then outDic.Add j, 1
                    Next j
                Else ' � ���� �� ����� - ����� ��� ��� ����, �� ���������
                    For Each el In Step2
                        If Not outDic.Exists(el) Then outDic.Add el, 1
                    Next el
                End If
            End If
        Next i
        
        If outDic.count = 0 Or Not (LCase(stringOrNumeric) = "numeric" Or LCase(stringOrNumeric) = "string" Or LCase(stringOrNumeric) = "dic") Then
            Expand = ""
        ElseIf LCase(stringOrNumeric) = "numeric" Then
            If useSort Then
                Expand = ShellSort(outDic.keys)    ' numericArray
            Else
                Expand = outDic.keys    ' numericArray
            End If
        ElseIf LCase(stringOrNumeric) = "string" Then
            If useSort Then
                Step4 = ShellSort(outDic.keys)
            Else
                Step4 = outDic.keys
            End If
            ReDim Step3(LBound(Step4) To UBound(Step4))
            For i = LBound(Step4) To UBound(Step4)
                Step3(i) = CStr(Step4(i))
            Next i
            Expand = Step3               ' stringArray
        ElseIf LCase(stringOrNumeric) = "dic" Then
            Set Expand = outDic          ' dictionary
        End If
    End If
End Function
