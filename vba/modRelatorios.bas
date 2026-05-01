Attribute VB_Name = "modRelatorios"
'==============================================================================
' modRelatorios
'   Geracao de relatorios e exportacao para PDF (DASHBOARD + ANALISE).
'==============================================================================
Option Explicit

Public Sub ExportarRelatorioPDF360()
    On Error GoTo TrataErro
    Dim st As AppState: st = SalvarEstado()
    PrepararExecucao

    Dim caminho As String
    Dim nomeBase As String
    Dim wsDash As Worksheet, wsAnal As Worksheet

    Set wsDash = GetSheet(SH_DASHBOARD)
    Set wsAnal = GetSheet(SH_ANALISE)
    If wsDash Is Nothing Then GoTo Finalizar

    nomeBase = "EXAUSTAO360_Relatorio_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"
    caminho = ThisWorkbook.Path
    If Len(caminho) = 0 Then caminho = Environ("USERPROFILE") & "\Desktop"
    caminho = caminho & Application.PathSeparator & nomeBase

    ' Garante area de impressao adequada
    wsDash.PageSetup.Orientation = xlLandscape
    wsDash.PageSetup.Zoom = False
    wsDash.PageSetup.FitToPagesWide = 1
    wsDash.PageSetup.FitToPagesTall = False
    If Not wsAnal Is Nothing Then
        wsAnal.PageSetup.Orientation = xlLandscape
        wsAnal.PageSetup.Zoom = False
        wsAnal.PageSetup.FitToPagesWide = 1
        wsAnal.PageSetup.FitToPagesTall = False
    End If

    Dim arrSheets() As String
    If wsAnal Is Nothing Then
        ReDim arrSheets(0): arrSheets(0) = wsDash.Name
    Else
        ReDim arrSheets(1): arrSheets(0) = wsDash.Name: arrSheets(1) = wsAnal.Name
    End If
    ThisWorkbook.Sheets(arrSheets).Select

    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, Filename:=caminho, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True

    wsDash.Select

    RegistrarLog "Exportar PDF", "DASHBOARD+ANALISE", caminho, "Sucesso"
    MsgBox "Relatorio exportado em:" & vbCrLf & caminho, vbInformation, AppTitle

Finalizar:
    RestaurarEstado st
    Exit Sub
TrataErro:
    RegistrarLog "Exportar PDF", "DASHBOARD+ANALISE", "", "ERRO: " & Err.Description
    MsgBox "Erro em ExportarRelatorioPDF360: " & Err.Description, vbCritical, AppTitle
    Resume Finalizar
End Sub

Public Sub ImprimirRelatorio()
    On Error GoTo TrataErro
    Dim ws As Worksheet
    Set ws = GetSheet(SH_DASHBOARD)
    If ws Is Nothing Then Exit Sub
    ws.PrintPreview
    Exit Sub
TrataErro:
    MsgBox "Erro em ImprimirRelatorio: " & Err.Description, vbCritical, AppTitle
End Sub

Public Sub ExportarConsulta()
    On Error GoTo TrataErro
    Dim st As AppState: st = SalvarEstado()
    PrepararExecucao

    Dim lo As ListObject
    Set lo = GetTable(TB_EVENTOS)
    If lo Is Nothing Then GoTo Finalizar

    Dim caminho As String
    caminho = ThisWorkbook.Path & Application.PathSeparator & _
              "EXAUSTAO360_Consulta_" & Format(Now, "yyyymmdd_hhnnss") & ".csv"

    Dim ff As Integer
    ff = FreeFile
    Open caminho For Output As #ff
    ' cabecalho
    Print #ff, JoinHeaders(lo)
    Dim r As ListRow
    For Each r In lo.ListRows
        Print #ff, JoinRow(r.Range)
    Next r
    Close #ff

    RegistrarLog "Exportar Consulta", SH_BASE, caminho, "Sucesso"
    MsgBox "Consulta exportada em:" & vbCrLf & caminho, vbInformation, AppTitle

Finalizar:
    RestaurarEstado st
    Exit Sub
TrataErro:
    MsgBox "Erro em ExportarConsulta: " & Err.Description, vbCritical, AppTitle
    Resume Finalizar
End Sub

Private Function JoinHeaders(ByVal lo As ListObject) As String
    Dim s As String
    Dim h As ListColumn
    Dim primeiro As Boolean: primeiro = True
    For Each h In lo.ListColumns
        If primeiro Then
            s = h.Name
            primeiro = False
        Else
            s = s & ";" & h.Name
        End If
    Next h
    JoinHeaders = s
End Function

Private Function JoinRow(ByVal rng As Range) As String
    Dim s As String
    Dim i As Long
    For i = 1 To rng.Columns.Count
        Dim v As String
        v = CStr(rng.Cells(1, i).Value)
        v = Replace(v, ";", ",")
        v = Replace(v, vbCrLf, " ")
        If i = 1 Then s = v Else s = s & ";" & v
    Next i
    JoinRow = s
End Function
