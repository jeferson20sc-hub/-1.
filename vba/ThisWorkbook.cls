VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    Inicializar360
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    RegistrarLog "Fechar Arquivo", "WORKBOOK", "", "Sucesso"
    LimparRuntime
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error Resume Next
    RegistrarLog "Salvar Arquivo", "WORKBOOK", IIf(SaveAsUI, "SaveAs", "Save"), _
                 "Iniciando"
    On Error GoTo 0
End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    On Error Resume Next
    RegistrarLog "Salvar Arquivo", "WORKBOOK", "", IIf(Success, "Sucesso", "Falha")
    On Error GoTo 0
End Sub
