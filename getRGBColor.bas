' узнаёт цвет заливки или шрифта ячейки
Function getRGBColor(cl As Variant, Optional getRGBforFill As Boolean = True) As Variant
    getRGBColor = "!Error"
    If TypeName(cl) = "Range" Then
        If getRGBforFill Then
            tmpValue = Hex(cl.Cells(1, 1).Interior.color)
        Else
            tmpValue = Hex(cl.Cells(1, 1).Font.color)
        End If
    ElseIf IsNumeric(cl) Then
        If CDbl(cl) > 0 And CDbl(cl) < CDbl("&HFFFFFF") Then
            tmpValue = Hex(CDbl(cl))
        Else
            Exit Function
        End If
    Else
        Exit Function
    End If
    
    tmp = IIf(Len(tmpValue) < 6, String(6 - Len(tmpValue), "0"), "") + tmpValue
    redClr = CInt("&H" & Right(tmp, 2))
    greenClr = CInt("&H" & Mid(tmp, 3, 2))
    blueClr = CInt("&H" & Left(tmp, 2))
    
    getRGBColor = "RGB(" & CStr(redClr) & "," & CStr(greenClr) & "," & CStr(blueClr) & ")"
End Function