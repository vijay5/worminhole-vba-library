Function FindCell(textToFind As String, rangeToSearch As Range, Optional lookIn As XlFindLookIn = xlValues, Optional lookAt As XlLookAt = xlWhole, Optional matchCase As Boolean = False) As Range
    ''' Ищет ячейку в заданном диапазоне
    ''' В качестве rangeToSearch можно передавать Sheets("somesheet").Range(...), тогда поиск
    ''' будет происходить по указанному диапазону, причём лист может быть неактивным (!)
    ''' Если значение не найдено - возвращает Nothing
    Dim cl As Range
    
    
    Set cl = rangeToSearch.Find( _
    What:=textToFind, _
    after:=rangeToSearch.Cells(1, 1), _
    lookIn:=lookIn, _
    lookAt:=lookAt, _
    SearchOrder:=xlByRows, _
    SearchDirection:=xlNext, _
    matchCase:=matchCase, _
    SearchFormat:=False)
    
    Set FindCell = cl
End Function