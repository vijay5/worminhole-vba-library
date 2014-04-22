' отображает статус задачи в строке состояния (чтобы юзер не пугался)
Sub showStatus(text As String, Optional curValue As Variant = "", Optional maxValue As Variant = "")
    If text <> "" Then
        If curValue <> "" And maxValue <> "" Then
            If IsNumeric(curValue) And IsNumeric(maxValue) And curValue <= maxValue Then
                Application.StatusBar = text + " " + Format(curValue / maxValue, "0%")
            End If
        ElseIf curValue <> "" And maxValue = "" Then
            If IsNumeric(curValue) Then
                Application.StatusBar = text + " " + Format(curValue, "0")
            End If
        Else ' curValue = ""
            Application.StatusBar = text
        End If
        
    Else
        Application.StatusBar = ""
    End If
    DoEvents
End Sub