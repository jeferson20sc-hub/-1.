Attribute VB_Name = "modRuntime"
'==============================================================================
' modRuntime
'   Salvar / restaurar estado do Excel (ScreenUpdating, EnableEvents,
'   Calculation, DisplayAlerts, StatusBar). Usado por toda macro pesada
'   para garantir que a aplicacao volta ao estado original mesmo em erro.
'==============================================================================
Option Explicit

Public Type AppState
    Calc       As XlCalculation
    Screen     As Boolean
    Events     As Boolean
    Alerts     As Boolean
    StatusBar  As Variant
End Type

Public Function SalvarEstado() As AppState
    Dim s As AppState
    s.Calc = Application.Calculation
    s.Screen = Application.ScreenUpdating
    s.Events = Application.EnableEvents
    s.Alerts = Application.DisplayAlerts
    s.StatusBar = Application.StatusBar
    SalvarEstado = s
End Function

Public Sub PrepararExecucao()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
End Sub

Public Sub RestaurarEstado(ByRef s As AppState)
    On Error Resume Next
    Application.Calculation = s.Calc
    Application.ScreenUpdating = s.Screen
    Application.EnableEvents = s.Events
    Application.DisplayAlerts = s.Alerts
    Application.StatusBar = s.StatusBar
    On Error GoTo 0
End Sub

Public Sub LimparRuntime()
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.StatusBar = False
    On Error GoTo 0
End Sub
