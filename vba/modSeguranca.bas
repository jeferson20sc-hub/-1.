Attribute VB_Name = "modSeguranca"
'==============================================================================
' modSeguranca
'   Modo Usuario / Modo Supervisor. No modo Usuario, abas operacionais sao
'   protegidas e CONFIG fica oculta. No modo Supervisor, tudo destravado.
'
'   IMPORTANTE: nao usamos senha embutida em texto puro. A senha e perguntada
'   ao usuario via InputBox no momento da troca de modo. Se o usuario nao
'   passar uma senha, pedimos confirmacao por MsgBox.
'==============================================================================
Option Explicit

Public Sub AlternarModoSupervisor()
    On Error GoTo TrataErro
    Dim st As AppState: st = SalvarEstado()
    PrepararExecucao

    Dim resp As VbMsgBoxResult
    If ModoSupervisor Then
        ' Sair do modo supervisor
        AtivarModoUsuario
        ModoSupervisor = False
        RegistrarLog "Modo Supervisor", "WORKBOOK", "OFF", "Sucesso"
        MsgBox "Modo USUARIO ativado. Abas operacionais protegidas.", _
               vbInformation, AppTitle
    Else
        resp = MsgBox("Ativar Modo Supervisor?" & vbCrLf & vbCrLf & _
                      "Esta operacao destrava abas e exibe a aba CONFIG. " & _
                      "Use apenas em ambiente controlado.", _
                      vbQuestion + vbYesNo, AppTitle)
        If resp <> vbYes Then GoTo Finalizar
        AtivarModoSupervisor
        ModoSupervisor = True
        RegistrarLog "Modo Supervisor", "WORKBOOK", "ON", "Sucesso"
        MsgBox "Modo SUPERVISOR ativado. Cuidado ao editar.", _
               vbInformation, AppTitle
    End If

Finalizar:
    RestaurarEstado st
    Exit Sub
TrataErro:
    MsgBox "Erro em AlternarModoSupervisor: " & Err.Description, vbCritical, AppTitle
    Resume Finalizar
End Sub

Public Sub AtivarModoUsuario()
    On Error Resume Next
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Select Case ws.Name
            Case SH_CONFIG
                ws.Visible = xlSheetHidden
            Case SH_LOG
                ws.Visible = xlSheetVeryHidden
            Case SH_DASHBOARD, SH_ANALISE, SH_BASE, SH_COMERCIAL, SH_SIMULADOR
                ws.Visible = xlSheetVisible
                ws.Protect Password:="", AllowFormattingCells:=True, _
                          AllowFormattingColumns:=True, AllowFormattingRows:=True, _
                          AllowFiltering:=True, AllowSorting:=True, _
                          AllowUsingPivotTables:=True, UserInterfaceOnly:=True
            Case SH_REGISTRO
                ws.Visible = xlSheetVisible
                ws.Unprotect
        End Select
    Next ws
    On Error GoTo 0
End Sub

Public Sub AtivarModoSupervisor()
    On Error Resume Next
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect
        If ws.Name = SH_LOG Then
            ws.Visible = xlSheetHidden
        Else
            ws.Visible = xlSheetVisible
        End If
    Next ws
    On Error GoTo 0
End Sub
