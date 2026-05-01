Attribute VB_Name = "modSimulador"
'==============================================================================
' modSimulador
'   Cenarios financeiros e simulacao de Monte Carlo simples para o
'   horizonte de 3 anos. Os resultados sao gravados na aba SIMULADOR
'   (linha 14, abaixo do bloco de cenarios) e em uma faixa de saida
'   livre para inspecao.
'==============================================================================
Option Explicit

Public Sub SimularCenarios()
    On Error GoTo TrataErro
    Dim st As AppState: st = SalvarEstado()
    PrepararExecucao

    Dim ws As Worksheet
    Set ws = GetSheet(SH_SIMULADOR)
    If ws Is Nothing Then GoTo Finalizar

    ' Forca recalculo das formulas dos cenarios (que dependem de INPUT_*)
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate

    RegistrarLog "Simular Cenarios", SH_SIMULADOR, "", "Sucesso"
    MsgBox "Cenarios atualizados com base nas premissas atuais.", _
           vbInformation, AppTitle

Finalizar:
    RestaurarEstado st
    Exit Sub
TrataErro:
    MsgBox "Erro em SimularCenarios: " & Err.Description, vbCritical, AppTitle
    Resume Finalizar
End Sub

Public Sub SimularMonteCarlo360()
    On Error GoTo TrataErro
    Dim st As AppState: st = SalvarEstado()
    PrepararExecucao

    Dim ws As Worksheet
    Set ws = GetSheet(SH_SIMULADOR)
    If ws Is Nothing Then GoTo Finalizar

    Dim perda As Double, invest As Double
    Dim n As Long, i As Long
    Dim total As Double, melhor As Double, pior As Double
    Dim media As Double, desvio As Double
    Dim economia As Double, pct As Double
    Dim valores() As Double

    perda = CDbl(Range("INPUT_PerdaAnual").Value)
    invest = CDbl(Range("INPUT_Investimento").Value)
    n = 5000

    ReDim valores(1 To n)
    Randomize

    melhor = -1E+15
    pior = 1E+15
    total = 0

    For i = 1 To n
        ' Reducao % entre 3% e 25%, distribuicao triangular aproximada
        pct = (Rnd() + Rnd() + Rnd()) / 3 * 0.22 + 0.03
        economia = perda * pct - invest
        valores(i) = economia
        total = total + economia
        If economia > melhor Then melhor = economia
        If economia < pior Then pior = economia
    Next i

    media = total / n

    Dim soma2 As Double
    For i = 1 To n
        soma2 = soma2 + (valores(i) - media) ^ 2
    Next i
    desvio = Sqr(soma2 / n)

    ' Escreve resultados a partir da linha 26 (area livre)
    With ws
        .Range("B26:K33").ClearContents
        .Range("B26").Value = "MONTE CARLO - " & n & " simulacoes"
        .Range("B26").Font.Bold = True
        .Range("B27").Value = "Cenario"
        .Range("C27").Value = "Resultado liquido (R$)"
        .Range("B28").Value = "Pior caso"
        .Range("B29").Value = "Media"
        .Range("B30").Value = "Melhor caso"
        .Range("B31").Value = "Desvio padrao"
        .Range("C28").Value = pior
        .Range("C29").Value = media
        .Range("C30").Value = melhor
        .Range("C31").Value = desvio
        .Range("C28:C31").NumberFormat = "R$ #,##0.00"

        ' Probabilidade de payback positivo
        Dim positivos As Long: positivos = 0
        For i = 1 To n
            If valores(i) > 0 Then positivos = positivos + 1
        Next i
        .Range("B33").Value = "P(retorno positivo no 1 ano)"
        .Range("C33").Value = positivos / n
        .Range("C33").NumberFormat = "0.0%"
    End With

    RegistrarLog "Monte Carlo", SH_SIMULADOR, n & " sims", "Sucesso"
    MsgBox "Monte Carlo concluido. " & n & " simulacoes." & vbCrLf & _
           "Media de retorno: " & FormatBRL(media), _
           vbInformation, AppTitle

Finalizar:
    RestaurarEstado st
    Exit Sub
TrataErro:
    MsgBox "Erro em SimularMonteCarlo360: " & Err.Description, vbCritical, AppTitle
    Resume Finalizar
End Sub
