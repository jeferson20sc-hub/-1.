Attribute VB_Name = "modValidacao"
'==============================================================================
' modValidacao
'   Regras de validacao do formulario REGISTRO. Devolve uma string vazia se
'   tudo estiver OK ou uma mensagem detalhada com os campos pendentes.
'==============================================================================
Option Explicit

Public Function ValidarFormularioRegistro() As String
    Dim ws As Worksheet
    Dim msgs As String

    Set ws = GetSheet(SH_REGISTRO)
    If ws Is Nothing Then
        ValidarFormularioRegistro = "Aba REGISTRO nao encontrada."
        Exit Function
    End If

    msgs = ""
    If EhCelulaVazia(ws.Range("C5").Value) Then msgs = msgs & vbCrLf & " - Data e obrigatoria"
    If EhCelulaVazia(ws.Range("C6").Value) Then msgs = msgs & vbCrLf & " - Hora e obrigatoria"
    If EhCelulaVazia(ws.Range("C7").Value) Then msgs = msgs & vbCrLf & " - Operador e obrigatorio"
    If EhCelulaVazia(ws.Range("C9").Value) Then msgs = msgs & vbCrLf & " - Forno e obrigatorio"
    If EhCelulaVazia(ws.Range("C10").Value) Then msgs = msgs & vbCrLf & " - Tipo de Evento e obrigatorio"
    If EhCelulaVazia(ws.Range("F6").Value) Then msgs = msgs & vbCrLf & " - Criticidade e obrigatoria"

    ' Numericos: duracao e custo devem ser >= 0 quando preenchidos
    If Not EhCelulaVazia(ws.Range("F7").Value) Then
        If Not IsNumeric(ws.Range("F7").Value) Then
            msgs = msgs & vbCrLf & " - Duracao deve ser numerica"
        ElseIf CDbl(ws.Range("F7").Value) < 0 Then
            msgs = msgs & vbCrLf & " - Duracao deve ser >= 0"
        End If
    End If
    If Not EhCelulaVazia(ws.Range("F8").Value) Then
        If Not IsNumeric(ws.Range("F8").Value) Then
            msgs = msgs & vbCrLf & " - Custo deve ser numerico"
        ElseIf CDbl(ws.Range("F8").Value) < 0 Then
            msgs = msgs & vbCrLf & " - Custo deve ser >= 0"
        End If
    End If

    ValidarFormularioRegistro = msgs
End Function

Public Sub ValidarRegistro()
    On Error GoTo TrataErro
    Dim st As AppState: st = SalvarEstado()
    PrepararExecucao

    Dim ws As Worksheet
    Set ws = GetSheet(SH_REGISTRO)
    If ws Is Nothing Then GoTo Finalizar

    Dim msgs As String
    msgs = ValidarFormularioRegistro()

    If Len(msgs) = 0 Then
        ws.Range("B18").Value = "OK: Dados validos. Pronto para registrar."
        ws.Range("B18").Font.Color = RGB(46, 125, 50)
        MsgBox "Formulario validado com sucesso. Clique em SALVAR para registrar.", _
               vbInformation, AppTitle
    Else
        ws.Range("B18").Value = "ATENCAO: corrigir campos antes de salvar." & msgs
        ws.Range("B18").Font.Color = RGB(198, 40, 40)
        MsgBox "Foram encontrados problemas:" & msgs, vbExclamation, AppTitle
    End If

Finalizar:
    RestaurarEstado st
    Exit Sub
TrataErro:
    MsgBox "Erro em ValidarRegistro: " & Err.Description, vbCritical, AppTitle
    Resume Finalizar
End Sub
