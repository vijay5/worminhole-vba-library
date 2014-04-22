Sub economyModeOn()
    ' эконом-мода
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With
End Sub