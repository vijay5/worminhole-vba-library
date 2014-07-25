' возвращает смещение rngToOffset относительно pivot.Cells(1,1)
Sub getRngOffsets(rngToOffset As Range, pivot As Range, rowOffset As Long, colOffset As Long)
    Set rng1 = rngToOffset.Cells(1, 1)
    Set rng2 = pivot.Cells(1, 1)
    
    rowOffset = rng1.Row - rng2.Row
    colOffset = rng1.Column - rng2.Column
End Sub
