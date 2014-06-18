''' для приклеивания элементов в зад большой строки
Function addToText(source As String, appndx As String, Optional divisor As String = ",") As String
    addToText = IIf(Len(source) = 0, appndx, source & divisor & appndx)
End Function