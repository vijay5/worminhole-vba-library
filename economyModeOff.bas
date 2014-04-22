Sub economyModeOff()
    ' эконом-мода
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub
