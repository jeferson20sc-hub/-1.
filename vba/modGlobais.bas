Attribute VB_Name = "modGlobais"
'==============================================================================
' modGlobais
'   Constantes globais, versionamento e nomes de objetos do EXAUSTAO 360
'   ENTERPRISE PRO. Tudo que e "configuravel em tempo de design" mora aqui.
'==============================================================================
Option Explicit

' --- Identidade do produto ----------------------------------------------------
Public Const APP_NAME      As String = "EXAUSTAO 360 ENTERPRISE PRO"
Public Const APP_VERSION   As String = "1.0.0"
Public Const APP_VENDOR    As String = "EXAUSTAO 360 PRO"

' --- Nomes de planilhas (manter sincronizado com build.py) --------------------
Public Const SH_DASHBOARD  As String = "DASHBOARD"
Public Const SH_REGISTRO   As String = "REGISTRO"
Public Const SH_BASE       As String = "BASE"
Public Const SH_ANALISE    As String = "ANALISE"
Public Const SH_SIMULADOR  As String = "SIMULADOR"
Public Const SH_COMERCIAL  As String = "MODELO_COMERCIAL"
Public Const SH_CONFIG     As String = "CONFIG"
Public Const SH_LOG        As String = "LOG"

' --- Nomes das tabelas estruturadas ------------------------------------------
Public Const TB_EVENTOS    As String = "tbEventos"
Public Const TB_PARAMETROS As String = "tbParametros"
Public Const TB_FORNOS     As String = "tbFornos"
Public Const TB_TIPOS      As String = "tbTiposEvento"
Public Const TB_COMPS      As String = "tbComponentes"
Public Const TB_OPERS      As String = "tbOperadores"
Public Const TB_FMEA       As String = "tbFMEA"
Public Const TB_PLANO      As String = "tbPlanoAcao"
Public Const TB_LOG        As String = "tbLogEdicao"

' --- Modo supervisor (em memoria) --------------------------------------------
Public ModoSupervisor As Boolean

' --- Helper para titulo de MsgBox --------------------------------------------
Public Function AppTitle() As String
    AppTitle = APP_NAME & " v" & APP_VERSION
End Function
