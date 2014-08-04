''' убирает нули (актуально при получении данных из БД)
Sub isolateNull(destVar As Variant, value As Variant)
    If IsNull(value) Then
        Select Case TypeName(destVar)
        Case "String"
            destVar = ""
        Case "Integer", "Long", "Single", "Double"
            destVar = 0
        Case "Date"
            destVar = DateSerial(1990, 1, 1)
        Case Else
            destVar = ""
        End Select
    Else
        destVar = value
    End If
End Sub
