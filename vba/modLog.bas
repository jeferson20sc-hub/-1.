Attribute VB_Name = "modLog"
'==============================================================================
' modLog
'   Wrapper publico para auditoria. Mantido como modulo separado para
'   permitir desativar/redirecionar logs sem mexer nas outras macros.
'==============================================================================
Option Explicit

Public Sub LogAuditoria(ByVal acao As String, ByVal aba As String, _
                        ByVal registro As String, ByVal resultado As String)
    On Error Resume Next
    RegistrarLog acao, aba, registro, resultado
    On Error GoTo 0
End Sub

Public Sub LimparLogAntigo(ByVal manterDias As Long)
    On Error GoTo TrataErro
    Dim st As AppState: st = SalvarEstado()
    PrepararExecucao

    Dim lo As ListObject
    Dim r As ListRow
    Dim limite As Date
    Dim removidos As Long

    Set lo = GetTable(TB_LOG)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then GoTo Finalizar

    limite = DateAdd("d", -manterDias, Now)
    Dim i As Long
    For i = lo.ListRows.Count To 1 Step -1
        Dim v As Variant
        v = lo.ListRows(i).Range.Cells(1, 1).Value
        If IsDate(v) Then
            If CDate(v) < limite Then
                lo.ListRows(i).Delete
                removidos = removidos + 1
            End If
        End If
    Next i

    MsgBox "Linhas de log removidas: " & removidos, vbInformation, AppTitle

Finalizar:
    RestaurarEstado st
    Exit Sub
TrataErro:
    MsgBox "Erro em LimparLogAntigo: " & Err.Description, vbCritical, AppTitle
    Resume Finalizar
End Sub
