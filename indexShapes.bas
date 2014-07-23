' индексирует все shape'ы, на выходе коллекци€: ключ = адрес €чйки , значение: Collection(shape1.Name, shape2.Name, ...)
' REQUIRES: addUniqToCol, isInCollection, buildIndex, BinarySearch
Sub indexShapes(sh As Worksheet, outCol As Collection, Optional shapeTypes As Variant = "")
    Dim shp As Shape
    Dim shapeTypesCol As New Collection
    Dim checkShpType As Boolean
    Dim shpType As String
    Dim key As String
    Dim item As Variant
    
    
    ' формируем список допустимых типов Shape'ов
    checkShpType = True
    If IsArray(shapeTypes) Then ' массив на входе
        For Each el In shapeTypes
            addUniqToCol shapeTypesCol, CStr(el), CStr(el)
        Next el
    ElseIf TypeName(shapeTypes) = "Collection" Then ' на входе коллекци€
        Set shapeTypesCol = shapeTypes
    ElseIf TypeName(shapeTypes) = "String" Then
        If shapeTypes = "" Then
            checkShpType = False
        Else
            addUniqToCol shapeTypesCol, CStr(shapeTypes), CStr(shapeTypes)
        End If
    Else
        ' pass
        Exit Sub
    End If
    ' список типов Shape'ов сформирован
    
    Dim maxRow As Long
    Dim maxCol As Long
    Dim leftIndexes As New Collection
    Dim topIndexes As New Collection
    Dim i As Long
    Dim shpTop As Single
    Dim shpLeft As Single
    Dim shpBottom As Single
    Dim shpRight As Single
    
    Dim rowCoord As Variant
    Dim colCoord As Variant
    Dim cl As Range
    
    maxRow = getMaxRow(sh)
    maxCol = getMaxCol(sh)
    
    ' индексы Left'ов столбцов
    For i = 1 To maxCol
        item = Array(sh.Cells(1, i).Left) ' в Array, чтобы использовать уже написаный BinarySearch
        key = CStr(i)
        leftIndexes.Add item, key
    Next i
    
    ' индексы Top'ов строк
    For j = 1 To maxRow
        item = Array(sh.Cells(j, 1).Top)
        key = CStr(j)
        topIndexes.Add item, key
    Next j
    
        
    ' индексируем объекты на листе
    Set outCol = New Collection
    For Each shp In sh.Shapes
        chk = False
        ' проврека типа Shape'а
        shpType = shp.Type
        If checkShpType Then ' список типов задан
            If isInCollection(shpType, shapeTypesCol) Then
                chk = True
            Else ' не подходит - не индексируем
                ' pass
            End If
        Else ' список типов не задан - берЄм всЄ
            chk = True
        End If
        
        If chk Then ' Shape подходит по типу
            
            shpTop = shp.Top
            shpLeft = shp.Left
            shpBottom = shp.Top + shp.Height
            shpRight = shp.Left + shp.Width
            
            ' top left corner
            rowCoord = BinarySearch(topIndexes, shpTop, 0) ' получили позицию координаты Shape'а внутри списка Top'ов строк
            colCoord = BinarySearch(leftIndexes, shpLeft, 0)
            
            colBegNum = colCoord(1) + IIf(colCoord(0) = -1, -1, 0)
            rowBegNum = rowCoord(1) + IIf(rowCoord(0) = -1, -1, 0)
            
            ' bottom right corner
            rowCoord = BinarySearch(topIndexes, shpBottom, 0)
            colCoord = BinarySearch(leftIndexes, shpRight, 0)
            
            colEndNum = colCoord(1) + IIf(colCoord(0) = -1, -1, 0)
            rowEndNum = rowCoord(1) + IIf(rowCoord(0) = -1, -1, 0)
            
            Set begCell = sh.Cells(rowBegNum, colBegNum)
            Set endCell = sh.Cells(rowEndNum, colEndNum)
            Set shpRange = Range(begCell, endCell)
            
            For Each cl In shpRange
                buildIndex outCol, cl.Address, shp.Name
            Next cl
        End If
    Next shp
End Sub