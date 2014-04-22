' стырено с http://www.cpearson.com/excel/FindAll.aspx
' чтобы искать €чейки по формату перед поиском определ€ем Application.SearchFormat
' чуть перекорЄжил чтобы нормально искала следующую €чейку с заданным форматом
Function FindAll(ByVal SearchRange As Range, _
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
    Dim FirstFound As Range
    Dim LastCell As Range
    Dim ResultRange As Range
    Dim XLookAt As XlLookAt
    Dim Include As Boolean
    Dim CompMode As VbCompareMethod
    Dim Area As Range
    Dim maxRow As Long
    Dim MaxCol As Long
    Dim BeginB As Boolean
    Dim EndB As Boolean
    
    
    CompMode = BeginEndCompare
    If BeginsWith <> vbNullString Or EndsWith <> vbNullString Then
        XLookAt = xlPart
    Else
        XLookAt = LookAt
    End If
    
    ' this loop in Areas is to find the last cell
    ' of all the areas. That is, the cell whose row
    ' and column are greater than or equal to any cell
    ' in any Area.
    For Each Area In SearchRange.Areas
        With Area
            If .Cells(.Cells.count).row > maxRow Then
                maxRow = .Cells(.Cells.count).row
            End If
            If .Cells(.Cells.count).Column > MaxCol Then
                MaxCol = .Cells(.Cells.count).Column
            End If
        End With
    Next Area
    Set LastCell = SearchRange.Worksheet.Cells(maxRow, MaxCol)
    
    On Error Resume Next
    Set FoundCell = SearchRange.Find(What:=FindWhat, _
            after:=LastCell, _
            LookIn:=LookIn, _
            LookAt:=XLookAt, _
            SearchOrder:=SearchOrder, _
            MatchCase:=MatchCase, _
            SearchFormat:=SearchFormat)
    On Error GoTo 0
    
    Set FirstFound = FoundCell
    
    If Not FoundCell Is Nothing Then
        Do Until False ' Loop forever. We'll "Exit Do" when necessary.
            Include = False
            ' если условие на начало или конец строки не задано - берЄм все €чейки
            If BeginsWith = vbNullString And EndsWith = vbNullString Then
                Include = True
            Else ' если условие на начало и/или конец строки задано - провер€ем начало и конец строки соотв. €чейки
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
            If Include = True Then ' текущую €чейку надо брать - она прошла проверку (или проверки не было)
                If ResultRange Is Nothing Then ' дл€ самой первой найденной €чейки - приравниваем первый найденный диапазон
                    Set ResultRange = FoundCell
                Else ' дл€ всех последующих найденныхй €чеек - склеиванем
                    Set ResultRange = Application.Union(ResultRange, FoundCell)
                End If
            End If
            
            'Set FoundCell = SearchRange.FindNext(after:=FoundCell)
            ' ищем следующую €чейку
            Set FoundCell = SearchRange.Find(What:=FindWhat, _
                    after:=FoundCell, _
                    LookIn:=LookIn, _
                    LookAt:=XLookAt, _
                    SearchOrder:=SearchOrder, _
                    MatchCase:=MatchCase, _
                    SearchFormat:=SearchFormat)
            
            If (FoundCell Is Nothing) Then ' ничего не нашли
                Exit Do
            End If
            If (FoundCell.Address = FirstFound.Address) Then ' нашли, но это сама€ перва€ €чейка
                Exit Do
            End If
        Loop
    End If
        
    Set FindAll = ResultRange
End Function