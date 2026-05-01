Attribute VB_Name = "modSetup"
'==============================================================================
' modSetup
'   Inicializacao do workbook: define estado seguro ao abrir, vincula
'   shapes/botoes e prepara o ambiente de uso.
'==============================================================================
Option Explicit

Public Sub Inicializar360()
    On Error GoTo TrataErro
    Dim st As AppState: st = SalvarEstado()
    PrepararExecucao

    ' Por padrao iniciamos em Modo Usuario
    ModoSupervisor = False
    AtivarModoUsuario

    ' Posiciona em DASHBOARD e atualiza
    Dim ws As Worksheet
    Set ws = GetSheet(SH_DASHBOARD)
    If Not ws Is Nothing Then
        ws.Activate
        Application.Goto ws.Range("A1"), Scroll:=True
    End If

    AtualizarDashboardSilencioso
    RegistrarLog "Inicializar", "WORKBOOK", APP_VERSION, "Sucesso"

Finalizar:
    RestaurarEstado st
    Exit Sub
TrataErro:
    ' Erro durante inicializacao nao deve travar abertura do arquivo
    LimparRuntime
    Exit Sub
End Sub

Private Sub AtualizarDashboardSilencioso()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = GetSheet(SH_DASHBOARD)
    If ws Is Nothing Then Exit Sub
    ws.Range("H3").Value = "Ultima atualizacao: " & Format(Now, "dd/mm/yyyy hh:mm")
    Application.Calculate
    On Error GoTo 0
End Sub
