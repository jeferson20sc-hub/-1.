Attribute VB_Name = "modDashboard"
'==============================================================================
' modDashboard
'   Atualizar KPIs, formulas e graficos do DASHBOARD. Tambem atualiza
'   tabelas dependentes em ANALISE.
'==============================================================================
Option Explicit

Public Sub AtualizarDashboard()
    On Error GoTo TrataErro
    Dim st As AppState: st = SalvarEstado()
    PrepararExecucao

    Dim ws As Worksheet
    Set ws = GetSheet(SH_DASHBOARD)
    If ws Is Nothing Then GoTo Finalizar

    ' Atualiza data/periodo no cabecalho
    Dim dtIni As Date, dtFim As Date
    GetPeriodo dtIni, dtFim
    ws.Range("B3").Value = "Periodo: " & Format(dtIni, "dd/mm/yyyy") & " - " & Format(dtFim, "dd/mm/yyyy")
    ws.Range("H3").Value = "Ultima atualizacao: " & Format(Now, "dd/mm/yyyy hh:mm")

    ' Forca recalculo
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFull

    ' Atualiza graficos
    Dim ch As ChartObject
    For Each ch In ws.ChartObjects
        ch.Chart.Refresh
    Next ch

    RegistrarLog "Atualizar Dashboard", SH_DASHBOARD, "", "Sucesso"

Finalizar:
    RestaurarEstado st
    Exit Sub
TrataErro:
    MsgBox "Erro em AtualizarDashboard: " & Err.Description, vbCritical, AppTitle
    Resume Finalizar
End Sub

Public Sub AtualizarTudo360()
    On Error GoTo TrataErro
    Dim st As AppState: st = SalvarEstado()
    PrepararExecucao

    AtualizarDashboardInterno

    ' Atualiza graficos em todas as planilhas
    Dim ws As Worksheet
    Dim ch As ChartObject
    For Each ws In ThisWorkbook.Worksheets
        For Each ch In ws.ChartObjects
            ch.Chart.Refresh
        Next ch
    Next ws

    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFull

    RegistrarLog "Atualizar Tudo", "TODAS", "", "Sucesso"

    MsgBox "Atualizacao geral concluida.", vbInformation, AppTitle

Finalizar:
    RestaurarEstado st
    Exit Sub
TrataErro:
    MsgBox "Erro em AtualizarTudo360: " & Err.Description, vbCritical, AppTitle
    Resume Finalizar
End Sub

Private Sub AtualizarDashboardInterno()
    Dim ws As Worksheet
    Set ws = GetSheet(SH_DASHBOARD)
    If ws Is Nothing Then Exit Sub
    Dim dtIni As Date, dtFim As Date
    GetPeriodo dtIni, dtFim
    ws.Range("B3").Value = "Periodo: " & Format(dtIni, "dd/mm/yyyy") & " - " & Format(dtFim, "dd/mm/yyyy")
    ws.Range("H3").Value = "Ultima atualizacao: " & Format(Now, "dd/mm/yyyy hh:mm")
End Sub

Private Sub GetPeriodo(ByRef dtIni As Date, ByRef dtFim As Date)
    Dim lo As ListObject
    Dim r As Range
    Dim minDt As Date, maxDt As Date
    Dim primeira As Boolean

    Set lo = GetTable(TB_EVENTOS)
    primeira = True
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            For Each r In lo.ListColumns("Data").DataBodyRange.Cells
                If IsDate(r.Value) Then
                    If primeira Then
                        minDt = CDate(r.Value): maxDt = CDate(r.Value)
                        primeira = False
                    Else
                        If CDate(r.Value) < minDt Then minDt = CDate(r.Value)
                        If CDate(r.Value) > maxDt Then maxDt = CDate(r.Value)
                    End If
                End If
            Next r
        End If
    End If

    If primeira Then
        dtIni = DateSerial(Year(Date), Month(Date) - 3, 1)
        dtFim = Date
    Else
        dtIni = minDt
        dtFim = maxDt
    End If
End Sub
