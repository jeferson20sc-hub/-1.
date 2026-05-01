Attribute VB_Name = "modBackup"
'==============================================================================
' modBackup
'   Backup do arquivo .xlsm em pasta local com timestamp.
'==============================================================================
Option Explicit

Public Sub CriarBackup360()
    On Error GoTo TrataErro
    Dim st As AppState: st = SalvarEstado()
    PrepararExecucao

    Dim pastaBackup As String
    Dim arquivoOrigem As String
    Dim arquivoDestino As String

    arquivoOrigem = ThisWorkbook.FullName
    If Len(arquivoOrigem) = 0 Then
        MsgBox "Salve o arquivo antes de criar o backup.", vbExclamation, AppTitle
        GoTo Finalizar
    End If

    pastaBackup = ThisWorkbook.Path & Application.PathSeparator & "backup"
    On Error Resume Next
    MkDir pastaBackup
    On Error GoTo TrataErro

    arquivoDestino = pastaBackup & Application.PathSeparator & _
                     "EXAUSTAO360_backup_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsm"

    ThisWorkbook.Save
    FileCopy arquivoOrigem, arquivoDestino

    RegistrarLog "Backup", "WORKBOOK", arquivoDestino, "Sucesso"
    MsgBox "Backup criado em:" & vbCrLf & arquivoDestino, vbInformation, AppTitle

Finalizar:
    RestaurarEstado st
    Exit Sub
TrataErro:
    RegistrarLog "Backup", "WORKBOOK", "", "ERRO: " & Err.Description
    MsgBox "Erro em CriarBackup360: " & Err.Description, vbCritical, AppTitle
    Resume Finalizar
End Sub
