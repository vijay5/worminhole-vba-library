''' удал€ютс€/пр€чутс€ все строки, если в левых €чейках выделени€ есть нули или пустые €чейки
''' работает с одной областью
' REQUIRES: economyModeOn, economyModeOff
Sub DeleteByZeroLColumn(Optional hide As Boolean = False, Optional askForMultiSheets As Boolean = True)
    Dim rep As String
    Dim a, b, cnt
    Dim rowsToDel() As Variant
    Dim numOfSelections As Integer
    
    If ActiveWorkbook.Windows(1).SelectedSheets.count > 1 And askForMultiSheets Then
        Select Case MsgBox("”далить строки с нулевыми €чейками на всех выделенных листах (Yes) или только на текущем (No)?", vbYesNoCancel)
        Case vbYes
            Set tmp = ActiveWorkbook.Windows(1).SelectedSheets
        Case vbNo
            ActiveSheet.Select
            Set tmp = ActiveWorkbook.Windows(1).SelectedSheets
        Case vbCancel
            Exit Sub
        End Select
    Else
        ActiveSheet.Select
        Set tmp = ActiveWorkbook.Windows(1).SelectedSheets
    End If
    
    economyModeOn
    
    For Each sh In tmp
        sh.Select
        sh.Activate
        If Selection.Areas.count = 1 Then ' не говорим про несколько выделений, просто не работаем...
            col = Selection.Column
            a = ActiveCell.row
            b = ActiveCell.Column
            rep = ""
            cnt = 0
            numOfSelections = 0
            PrevIsEmpty = False ' предыдуща€ €чейка не пуста€
            For i = Selection.row To Selection.row + Selection.Rows.count - 1
                ' условие, при котором текуща€ €чейка считаетс€ пустой
                CurIsEmpty = (Cells(i, col).value = 0) Or (Cells(i, col).value = "") Or (Cells(i, col).value = ".")
                    If PrevIsEmpty And CurIsEmpty Then
                    ' ничего не делаем
                ElseIf PrevIsEmpty And Not CurIsEmpty Then
                    ' закрываем множество
                    PrevIsEmpty = False
                    rep = rep + ":" + CStr(i - 1) + ","
                    numOfSelections = numOfSelections + 1
                    If numOfSelections = 20 Then
                    ' чтобы не было переполнени€ когда удал€ем слишком много выделений
                    ' режем список всех выделенных строк на кусочки по 20 выделений в каждом
                    ' потом удал€ем с конца по этим кусочкам
                        cnt = cnt + 1
                        If Right(rep, 1) = "," Then rep = Left(rep, Len(rep) - 1)
                        ReDim Preserve rowsToDel(1 To cnt)
                        rowsToDel(cnt) = rep
                        rep = ""
                        numOfSelections = 0
                    End If
                ElseIf Not PrevIsEmpty And CurIsEmpty Then
                    ' новое множество
                    PrevIsEmpty = True
                    rep = rep + CStr(i)
                ElseIf Not PrevIsEmpty And Not CurIsEmpty Then
                    ' ничего не делаем
                    PrevIsEmpty = False
                End If
    
            Next i
            CurIsEmpty = False
            If PrevIsEmpty And Not CurIsEmpty Then
                ' закрываем длинное множество
                rep = rep + ":" + CStr(i - 1)
            End If
            
            cnt = cnt + 1
            If Right(rep, 1) = "," Then rep = Left(rep, Len(rep) - 1)
            ReDim Preserve rowsToDel(1 To cnt)
            rowsToDel(cnt) = rep
            
            For i = UBound(rowsToDel) To LBound(rowsToDel) Step -1
                If rowsToDel(i) <> "" Then
                    Range(rowsToDel(i)).Select
                    ' пр€чем или удал€ем строку
                    If hide Then Selection.EntireRow.Hidden = True Else Selection.EntireRow.Delete
                End If
            Next i
            Cells(a, b).Select
            Cells(a, b).Activate

        End If
    Next sh
    tmp.Select
    economyModeOff
    
End Sub