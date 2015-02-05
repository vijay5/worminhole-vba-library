''' возвращает рабочий день, считая от текущей даты назад/вперёд
''' если текущий день рабочий - возвращает его
''' если текущий день - не рабочий, вовзращает предыдущий рабочий день
''' нерабочие дни должны быть заданы в одной ячейке, каждый день в новой строке через Chr(10)
''' dt - дата
''' direction - направление поиска >=0 - вперёд, <0 - назад
''' nonWorkingDays - праздничные дни
''' workingDays - рабочие выходные дни
''' REQUIRES: isInCollection
Function getNearestWorkingDay(dt As Date, Optional direction As Integer = 1, Optional nonWorkingDays As Range = Nothing, Optional workingDays As Range = Nothing) As Date
    Dim dtArr As Variant
    Dim nonWorkingDaysArr As Variant
    Dim nonWorkDaysCol As Collection
    Dim workingDaysArr As Variant
    Dim workDaysCol As Collection
    Dim i As Long
    Dim itm As Date
    Dim key As String
    Dim returnDate As Date
    Dim dtShift As Integer
    Dim chk As Boolean
    Dim chk1 As Boolean
    Dim chk2 As Boolean
    Dim chk3 As Boolean
    Dim curDate As Date
    
    
    ' нерабочие дни
    Set nonWorkDaysCol = New Collection
    If Not nonWorkingDays Is Nothing Then
        nonWorkingDaysArr = Split(nonWorkingDays.Cells(1, 1).Value, Chr(10))
        For i = LBound(nonWorkingDaysArr) To UBound(nonWorkingDaysArr)
            key = Format(CDate(nonWorkingDaysArr(i)), "DD.MM.YYYY")
            itm = CDate(nonWorkingDaysArr(i))
            If Not isInCollection(key, nonWorkDaysCol) Then
                nonWorkDaysCol.Add itm, key
            Else
                ' pass
            End If
        Next i
    Else
        ' pass
    End If
    
    ' рабочие дни - если заданы
    Set workDaysCol = New Collection
    If workingDays Is Nothing Then
        workingDaysArr = Split(workingDays.Cells(1, 1).Value, Chr(10))
        For i = LBound(workingDaysArr) To UBound(workingDaysArr)
            key = Format(CDate(workingDaysArr(i)), "DD.MM.YYYY")
            itm = CDate(workingDaysArr(i))
            If Not isInCollection(key, workDaysCol) Then
                workDaysCol.Add itm, key
            Else
                ' pass
            End If
        Next i
    Else
        ' pass
    End If
    
    ' ищем последний рабочий день
    returnDate = DateSerial(1990, 1, 1)
    dtShift = 0
    chk = True
    Do While chk And ((direction < 0 And dtShift >= -14) Or (direction >= 0 And dtShift <= 14))
        curDate = dt + dtShift ' текущая дата
        chk1 = (Weekday(curDate, vbMonday) >= 1 And Weekday(curDate, vbMonday) <= 5)
        chk2 = isInCollection(Format(curDate, "DD.MM.YYYY"), workDaysCol)
        chk3 = isInCollection(Format(curDate, "DD.MM.YYYY"), nonWorkDaysCol)
        
        If (chk1 And Not chk3) Or (Not chk1 And chk2) Then
            returnDate = curDate
            chk = False ' выходим из цикла
        Else
            ' pass
        End If
        
        dtShift = dtShift + IIf(direction >= 0, 1, -1)
    Loop
    getNearestWorkingDay = returnDate
End Function