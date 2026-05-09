Attribute VB_Name = "modFornos"
Option Explicit

' =====================================================================
' Modulo de macros - Controle de Fornos
' Macros: RegistrarForno, LimparCampos, LimparFiltros
' =====================================================================

Public Sub RegistrarForno()
    Dim ws As Worksheet, wsD As Worksheet
    Dim lo As ListObject
    Dim novaLinha As ListRow
    Dim dt As Variant, fn As Variant, tp As Variant, ld As Variant
    Dim novoID As Long

    Set ws = ThisWorkbook.Worksheets("Registro")
    Set wsD = ThisWorkbook.Worksheets("_Dados")
    Set lo = wsD.ListObjects("tblFornos")

    dt = ws.Range("B5").Value
    fn = ws.Range("C5").Value
    tp = ws.Range("D5").Value
    ld = ws.Range("E5").Value

    If Not IsDate(dt) Then
        MsgBox "Informe uma data valida.", vbExclamation, "Registro"
        ws.Range("B5").Select: Exit Sub
    End If
    If Not IsNumeric(fn) Or fn = "" Then
        MsgBox "Informe o numero do forno.", vbExclamation, "Registro"
        ws.Range("C5").Select: Exit Sub
    End If
    If Trim(CStr(tp)) = "" Then
        MsgBox "Selecione o tipo.", vbExclamation, "Registro"
        ws.Range("D5").Select: Exit Sub
    End If
    If Trim(CStr(ld)) = "" Then
        MsgBox "Selecione o lado.", vbExclamation, "Registro"
        ws.Range("E5").Select: Exit Sub
    End If

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    If lo.ListRows.Count = 0 Then
        novoID = 1
    Else
        novoID = Application.WorksheetFunction.Max(lo.ListColumns("ID").DataBodyRange) + 1
    End If

    Set novaLinha = lo.ListRows.Add
    novaLinha.Range.Cells(1, 1).Value = novoID
    novaLinha.Range.Cells(1, 2).Value = CDate(dt)
    novaLinha.Range.Cells(1, 2).NumberFormat = "dd/mm/yyyy"
    novaLinha.Range.Cells(1, 3).Value = CLng(fn)
    novaLinha.Range.Cells(1, 4).Value = CStr(tp)
    novaLinha.Range.Cells(1, 5).Value = CStr(ld)

    ws.Range("C5").Value = ""
    ws.Range("C5").Select

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = "Forno " & fn & " (" & tp & ", lado " & ld & ") registrado em " & Format(dt, "dd/mm/yyyy") & "."
End Sub

Public Sub LimparCampos()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Registro")
    Application.EnableEvents = False
    ws.Range("B5").Value = Date
    ws.Range("C5").Value = ""
    ws.Range("D5").Value = "Operacao"
    ws.Range("E5").Value = "A"
    ws.Range("C5").Select
    Application.EnableEvents = True
End Sub

Public Sub LimparFiltros()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Filtros")
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ws.Range("B5").Value = ""
    ws.Range("C5").Value = ""
    ws.Range("D5").Value = ""
    ws.Range("E5").Value = ""
    ws.Range("F5").Value = ""
    ws.Range("G5").Value = ""
    ws.Range("H5").Value = ""
    ws.Range("I5").Value = ""
    ws.Range("C7").Value = ""
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Public Sub IrParaDashboard()
    ThisWorkbook.Worksheets("Dashboard").Activate
End Sub

Public Sub IrParaRegistro()
    ThisWorkbook.Worksheets("Registro").Activate
    ThisWorkbook.Worksheets("Registro").Range("C5").Select
End Sub

Public Sub IrParaFiltros()
    ThisWorkbook.Worksheets("Filtros").Activate
End Sub
