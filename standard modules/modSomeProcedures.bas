Attribute VB_Name = "modSomeProcedures"
Option Explicit

Sub calculation_events_screen(bolEnabling As Boolean)
    Application.Calculation = IIf(bolEnabling, xlCalculationAutomatic, xlCalculationManual)
    Application.EnableEvents = bolEnabling
    Application.ScreenUpdating = bolEnabling
End Sub

