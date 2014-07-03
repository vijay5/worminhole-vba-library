''' стырено с http://www.cpearson.com/excel/FindAll.aspx
''' чтобы искать €чейки по формату перед поиском определ€ем Application.SearchFormat
''' чуть перекорЄжил чтобы нормально искала следующую €чейку с заданным форматом
Function FindAll(SearchRange As Range, _
                FindWhat As Variant, _
               Optional LookIn As XlFindLookIn = xlValues, _
                Optional LookAt As XlLookAt = xlWhole, _
                Optional SearchOrder As XlSearchOrder = xlByRows, _
                Optional MatchCase As Boolean = False, _
                Optional BeginsWith As String = vbNullString, _
                Optional EndsWith As String = vbNullString, _
                Optional BeginEndCompare As VbCompareMethod = vbTextCompare, _
                Optional SearchFormat As Boolean = False) As Range
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' FindAll
    ' This searches the range specified by SearchRange and returns a Range object
    ' that contains all the cells in which FindWhat was found. The search parameters to
    ' this function have the same meaning and effect as they do with the
    ' Range.Find method. If the value was not found, the function return Nothing. If
    ' BeginsWith is not an empty string, only those cells that begin with BeginWith
    ' are included in the result. If EndsWith is not an empty string, only those cells
    ' that end with EndsWith are included in the result. Note that if a cell contains
    ' a single word that matches either BeginsWith or EndsWith, it is included in the
    ' result.  If BeginsWith or EndsWith is not an empty string, the LookAt parameter
    ' is automatically changed to xlPart. The tests for BeginsWith and EndsWith may be
    ' case-sensitive by setting BeginEndCompare to vbBinaryCompare. For case-insensitive
    ' comparisons, set BeginEndCompare to vbTextCompare. If this parameter is omitted,
    ' it defaults to vbTextCompare. The comparisons for BeginsWith and EndsWith are
    ' in an OR relationship. That is, if both BeginsWith and EndsWith are provided,
    ' a match if found if the text begins with BeginsWith OR the text ends with EndsWith.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim FoundCell As Range
    Dim ResultRange As Range
    Dim XLookAt As XlLookAt
    Dim Include As Boolean
    Dim needToExit As Boolean
    
    
    Set ResultRange = Nothing
    
    If BeginsWith <> vbNullString Or EndsWith <> vbNullString Then
        XLookAt = xlPart
    Else
        XLookAt = LookAt
    End If
    
    Set FoundCell = SearchRange.Find(what:=FindWhat, _
            LookIn:=LookIn, _
            LookAt:=XLookAt, _
            SearchOrder:=SearchOrder, _
            MatchCase:=MatchCase, _
            SearchFormat:=SearchFormat)
    
    If Not FoundCell Is Nothing Then ' нашли €чейку
        needToExit = False
        
        Do Until needToExit ' Loop forever. We'll "Exit Do" when necessary.
            Include = False ' флаг: найденна€ €чейка удовлетвор€ет услови€м
            If BeginsWith = vbNullString And EndsWith = vbNullString Then
                Include = True
            Else
                If BeginsWith <> vbNullString Then
                    If StrComp(Left(FoundCell.Text, Len(BeginsWith)), BeginsWith, BeginEndCompare) = 0 Then
                        Include = True
                    End If
                End If
                If EndsWith <> vbNullString Then
                    If StrComp(Right(FoundCell.Text, Len(EndsWith)), EndsWith, BeginEndCompare) = 0 Then
                        Include = True
                    End If
                End If
            End If
            
            If Include = True Then
                If ResultRange Is Nothing Then ' дл€ самого первого поиска
                    Set ResultRange = FoundCell
                Else ' дл€ всех остальных поисков
                    needToExit = Not (Intersect(ResultRange, FoundCell) Is Nothing) ' условие выхода - повторное добавление €чейки в Range
                    Set ResultRange = Application.Union(ResultRange, FoundCell)
                End If
            End If
            ' пытаемс€ искать дальше
            Set FoundCell = SearchRange.Find(what:=FindWhat, _
                    after:=FoundCell, _
                    LookIn:=LookIn, _
                    LookAt:=XLookAt, _
                    SearchOrder:=SearchOrder, _
                    MatchCase:=MatchCase, _
                    SearchFormat:=SearchFormat)
            
            If (FoundCell Is Nothing) Then ' если в диапазоне только одна €чейка (after:=FoundCell не найдЄт больше ничего)
                needToExit = True
            End If
        Loop
    Else
        ' pass
    End If
        
    Set FindAll = ResultRange

End Function
