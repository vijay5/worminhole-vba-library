' REQUIRES: arrayDepth
Sub insert2DArrToSh(arr As Variant, target As Range)
    Dim sh As Worksheet
    Dim minRow As Long
    Dim minCol As Long
    Dim maxRow As Long
    Dim maxCol As Long
    
    Set sh = target.Parent
    
    If arrayDepth(arr) = 2 Then
        minRow = target.Cells(1, 1).Row
        minCol = target.Cells(1, 1).Column
        maxRow = minRow + (UBound(arr, 1) - LBound(arr, 1) + 1) - 1
        maxCol = minCol + (UBound(arr, 2) - LBound(arr, 2) + 1) - 1
        
        ' вставляем массив
        Range(sh.Cells(minRow, minCol), sh.Cells(maxRow, maxCol)).Value = arr
    Else
        ' pass - ничего не вставляем
    End If
    
End Sub