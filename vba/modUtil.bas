Attribute VB_Name = "modUtil"
'==============================================================================
' modUtil
'   Utilitarios genericos: localizar planilhas/tabelas, gerar IDs,
'   escrever em LOG, formatar timestamps.
'==============================================================================
Option Explicit

Public Function GetSheet(ByVal nome As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, nome, vbTextCompare) = 0 Then
            Set GetSheet = ws
            Exit Function
        End If
    Next ws
    Set GetSheet = Nothing
End Function

Public Function GetTable(ByVal nomeTabela As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, nomeTabela, vbTextCompare) = 0 Then
                Set GetTable = lo
                Exit Function
            End If
        Next lo
    Next ws
    Set GetTable = Nothing
End Function

Public Function ProximoIDEvento() As Long
    Dim lo As ListObject
    Dim r As Range
    Dim maxId As Long
    Dim v As Variant

    Set lo = GetTable(TB_EVENTOS)
    If lo Is Nothing Then
        ProximoIDEvento = 1
        Exit Function
    End If
    maxId = 0
    If Not lo.DataBodyRange Is Nothing Then
        For Each r In lo.ListColumns("ID").DataBodyRange.Cells
            v = r.Value
            If IsNumeric(v) Then
                If CLng(v) > maxId Then maxId = CLng(v)
            End If
        Next r
    End If
    ProximoIDEvento = maxId + 1
End Function

Public Sub RegistrarLog(ByVal acao As String, ByVal aba As String, _
                        ByVal registro As String, ByVal resultado As String)
    On Error Resume Next
    Dim lo As ListObject
    Dim novaLinha As ListRow
    Set lo = GetTable(TB_LOG)
    If lo Is Nothing Then Exit Sub
    Set novaLinha = lo.ListRows.Add
    novaLinha.Range.Cells(1, 1).Value = Now
    novaLinha.Range.Cells(1, 2).Value = Environ("USERNAME")
    novaLinha.Range.Cells(1, 3).Value = acao
    novaLinha.Range.Cells(1, 4).Value = aba
    novaLinha.Range.Cells(1, 5).Value = registro
    novaLinha.Range.Cells(1, 6).Value = resultado
    On Error GoTo 0
End Sub

Public Function EhCelulaVazia(ByVal c As Variant) As Boolean
    EhCelulaVazia = (Trim(CStr(c & "")) = "")
End Function

Public Function FormatBRL(ByVal v As Double) As String
    FormatBRL = "R$ " & Format(v, "#,##0.00")
End Function

Public Function FormatMI(ByVal v As Double) As String
    FormatMI = "R$ " & Format(v / 1000000, "#,##0.00") & " mi"
End Function

Public Function ParametroValor(ByVal nomeParam As String) As Variant
    Dim lo As ListObject
    Dim r As ListRow
    Set lo = GetTable(TB_PARAMETROS)
    If lo Is Nothing Then
        ParametroValor = Empty
        Exit Function
    End If
    For Each r In lo.ListRows
        If StrComp(CStr(r.Range.Cells(1, 1).Value), nomeParam, vbTextCompare) = 0 Then
            ParametroValor = r.Range.Cells(1, 2).Value
            Exit Function
        End If
    Next r
    ParametroValor = Empty
End Function
