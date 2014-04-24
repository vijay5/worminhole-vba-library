Function getColumnLetter(cl As Range) As String
    Dim tmp As String
    Dim tmpArray As Variant

    tmp = cl.EntireColumn.Address(False, False)
    tmpArray = Split(tmp, ":")
    getColumnLetter = tmpArray(0)
End Function