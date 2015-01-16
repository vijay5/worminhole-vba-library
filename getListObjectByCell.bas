''' по адресу на листе возвращает таблицу
Function getListObjectByCell(rng As Range) As ListObject
    Dim sh As Worksheet
    Dim lstObject As ListObject
    
    
    Set getListObjectByCell = Nothing ' значение по умолчанию
    Set sh = rng.Parent ' лист, на котором указан адрес
    
    For Each lstObject In sh.ListObjects
        If Not Intersect(rng.Cells(1, 1), lstObject.Range) Is Nothing Then ' если есть пересечение
            Set getListObjectByCell = lstObject
            Exit For
        Else
            ' идём дальше
        End If
    Next lstObject
End Function