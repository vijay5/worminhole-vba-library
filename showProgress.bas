''' показывает прогресс выполнения
Sub showProgress(Optional msg As String = "")
    Dim patt As String
    Dim tmp As String
    Dim begPos As Integer
    Dim endPos As Integer
    Dim pointPos As Integer
    Dim pnt As Integer
    Dim resStr As String
    Dim rightPart As String
    
    patt = "[||||||||||]"
    
    tmp = Application.StatusBar
    begPos = InStr(tmp, "[")
    pointPos = InStr(WorksheetFunction.Max(1, begPos), tmp, ".")
    endPos = InStr(tmp, "]")
    rightPart = ""
    
    If begPos > 0 And pointPos > begPos And endPos > begPos Then
        rightPart = Mid(tmp, endPos + 1)
        pnt = pointPos - begPos
        pnt = pnt + 1
        If pnt > 10 Then pnt = 1
        resStr = Left(patt, pnt) + "." + Mid(patt, pnt + 2)
    Else
        resStr = "[.|||||||||]"
    End If
    
    Application.StatusBar = resStr + IIf(msg = "", rightPart, " " & msg)
End Sub
