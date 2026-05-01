Attribute VB_Name = "modRegistro"
'==============================================================================
' modRegistro
'   Macros do formulario REGISTRO: registrar evento (grava em tbEventos),
'   limpar formulario, abrir formulario.
'==============================================================================
Option Explicit

Public Sub RegistrarEvento360()
    On Error GoTo TrataErro
    Dim st As AppState: st = SalvarEstado()
    PrepararExecucao

    Dim wsForm As Worksheet
    Dim lo As ListObject
    Dim novaLinha As ListRow
    Dim msgs As String

    Set wsForm = GetSheet(SH_REGISTRO)
    If wsForm Is Nothing Then
        MsgBox "Aba REGISTRO nao encontrada.", vbCritical, AppTitle
        GoTo Finalizar
    End If

    msgs = ValidarFormularioRegistro()
    If Len(msgs) > 0 Then
        wsForm.Range("B18").Value = "ATENCAO: corrigir antes de salvar." & msgs
        wsForm.Range("B18").Font.Color = RGB(198, 40, 40)
        MsgBox "Nao e possivel salvar:" & msgs, vbExclamation, AppTitle
        GoTo Finalizar
    End If

    Set lo = GetTable(TB_EVENTOS)
    If lo Is Nothing Then
        MsgBox "Tabela tbEventos nao encontrada na aba BASE.", vbCritical, AppTitle
        GoTo Finalizar
    End If

    Set novaLinha = lo.ListRows.Add
    With novaLinha.Range
        .Cells(1, 1).Value = ProximoIDEvento()                 ' ID
        .Cells(1, 2).Value = wsForm.Range("C5").Value          ' Data
        .Cells(1, 3).Value = wsForm.Range("C6").Value          ' Hora
        .Cells(1, 4).Value = wsForm.Range("C7").Value          ' Operador
        .Cells(1, 5).Value = wsForm.Range("C9").Value          ' Forno
        .Cells(1, 6).Value = wsForm.Range("C10").Value         ' TipoEvento
        .Cells(1, 7).Value = wsForm.Range("C11").Value         ' Categoria
        .Cells(1, 8).Value = wsForm.Range("F5").Value          ' Componente
        .Cells(1, 9).Value = wsForm.Range("F6").Value          ' Criticidade
        .Cells(1, 10).Value = wsForm.Range("F7").Value         ' DuracaoHoras
        .Cells(1, 11).Value = wsForm.Range("F8").Value         ' CustoEstimado
        .Cells(1, 12).Value = wsForm.Range("C13").Value        ' Observacao
        .Cells(1, 13).Value = wsForm.Range("F9").Value         ' Status
    End With

    RegistrarLog "Registrar Evento", SH_BASE, _
                 "ID=" & novaLinha.Range.Cells(1, 1).Value, "Sucesso"

    wsForm.Range("B18").Value = "OK: evento registrado com sucesso."
    wsForm.Range("B18").Font.Color = RGB(46, 125, 50)

    LimparRegistroInterno wsForm

    ' Atualiza dashboard de forma silenciosa (mas restaura calc no fim)
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate

    MsgBox "Evento registrado com sucesso.", vbInformation, AppTitle

Finalizar:
    RestaurarEstado st
    Exit Sub
TrataErro:
    RegistrarLog "Registrar Evento", SH_REGISTRO, "", "ERRO: " & Err.Description
    MsgBox "Erro em RegistrarEvento360: " & Err.Description, vbCritical, AppTitle
    Resume Finalizar
End Sub

' Alias para compatibilidade com botoes antigos
Public Sub RegistrarEvento()
    RegistrarEvento360
End Sub

Public Sub LimparRegistro()
    On Error GoTo TrataErro
    Dim st As AppState: st = SalvarEstado()
    PrepararExecucao

    Dim ws As Worksheet
    Set ws = GetSheet(SH_REGISTRO)
    If ws Is Nothing Then GoTo Finalizar

    LimparRegistroInterno ws
    ws.Range("B18").Value = "Pronto para novo registro."
    ws.Range("B18").Font.Color = RGB(107, 107, 107)

Finalizar:
    RestaurarEstado st
    Exit Sub
TrataErro:
    MsgBox "Erro em LimparRegistro: " & Err.Description, vbCritical, AppTitle
    Resume Finalizar
End Sub

' Alias para compatibilidade
Public Sub LimparFormularioCompleto()
    LimparRegistro
End Sub

Private Sub LimparRegistroInterno(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Range("C5").Value = Date
    ws.Range("C6").Value = Format(Now, "hh:mm")
    ws.Range("C7").Value = ""
    ws.Range("C8").Value = ""
    ws.Range("C9").Value = ""
    ws.Range("C10").Value = ""
    ws.Range("C11").Value = ""
    ws.Range("F5").Value = ""
    ws.Range("F6").Value = ""
    ws.Range("F7").Value = 0
    ws.Range("F8").Value = 0
    ws.Range("F9").Value = "Aberto"
    ws.Range("F10").Value = ""
    ws.Range("F11").Value = ""
    ws.Range("C13").Value = ""
    On Error GoTo 0
End Sub

Public Sub AbrirFormularioRegistro()
    On Error GoTo TrataErro
    Dim ws As Worksheet
    Set ws = GetSheet(SH_REGISTRO)
    If ws Is Nothing Then Exit Sub
    ws.Activate
    ws.Range("C7").Select
    Exit Sub
TrataErro:
    MsgBox "Erro em AbrirFormularioRegistro: " & Err.Description, vbCritical, AppTitle
End Sub
