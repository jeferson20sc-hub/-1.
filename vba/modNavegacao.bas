Attribute VB_Name = "modNavegacao"
'==============================================================================
' modNavegacao
'   Navegacao rapida entre abas. Cada Sub leva o usuario a uma aba e
'   posiciona em A1.
'==============================================================================
Option Explicit

Public Sub NavegarDashboard()
    NavegarPara SH_DASHBOARD
End Sub

Public Sub NavegarRegistro()
    NavegarPara SH_REGISTRO
End Sub

Public Sub NavegarBase()
    NavegarPara SH_BASE
End Sub

Public Sub NavegarAnalise()
    NavegarPara SH_ANALISE
End Sub

Public Sub NavegarSimulador()
    NavegarPara SH_SIMULADOR
End Sub

Public Sub NavegarComercial()
    NavegarPara SH_COMERCIAL
End Sub

Public Sub NavegarConfig()
    If Not ModoSupervisor Then
        MsgBox "Aba CONFIG so pode ser acessada em Modo Supervisor.", _
               vbExclamation, AppTitle
        Exit Sub
    End If
    NavegarPara SH_CONFIG
End Sub

Public Sub NavegarPlanoAcao()
    On Error GoTo TrataErro
    Dim ws As Worksheet
    Set ws = GetSheet(SH_ANALISE)
    If ws Is Nothing Then Exit Sub
    ws.Activate
    ' Posiciona perto da tabela tbPlanoAcao
    Dim lo As ListObject
    Set lo = GetTable(TB_PLANO)
    If Not lo Is Nothing Then
        Application.Goto lo.Range.Cells(1, 1), Scroll:=True
    End If
    Exit Sub
TrataErro:
    MsgBox "Erro em NavegarPlanoAcao: " & Err.Description, vbCritical, AppTitle
End Sub

Private Sub NavegarPara(ByVal nomeAba As String)
    On Error GoTo TrataErro
    Dim ws As Worksheet
    Set ws = GetSheet(nomeAba)
    If ws Is Nothing Then
        MsgBox "Aba '" & nomeAba & "' nao encontrada.", vbExclamation, AppTitle
        Exit Sub
    End If
    If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
    ws.Activate
    Application.Goto ws.Range("A1"), Scroll:=True
    Exit Sub
TrataErro:
    MsgBox "Erro ao navegar para " & nomeAba & ": " & Err.Description, _
           vbCritical, AppTitle
End Sub
