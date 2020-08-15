Public origcalc

Private Sub Class_Initialize()
    Me.origcalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
End Sub

Private Sub Class_Terminate()
    Application.Calculation = Me.origcalc
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = ""
End Sub