# =============================================================================
# assemble.ps1 - Windows-only VBA injector for EXAUSTAO 360 ENTERPRISE PRO
#
# Pre-requisitos:
#   - Excel instalado (32 ou 64 bits)
#   - Acesso ao modelo de objeto VBE habilitado:
#       Excel -> Arquivo -> Opcoes -> Central de Confiabilidade ->
#       Configuracoes -> Confiar no acesso ao modelo de objeto do
#       projeto VBA
#   - Ter executado antes:  python build.py
#
# Uso:
#   powershell -ExecutionPolicy Bypass -File assemble.ps1
# =============================================================================

[CmdletBinding()]
param(
    [string]$XlsxPath = "dist\EXAUSTAO_360_ENTERPRISE_PRO.xlsx",
    [string]$XlsmPath = "dist\EXAUSTAO_360_ENTERPRISE_PRO.xlsm",
    [string]$VbaDir   = "vba"
)

$ErrorActionPreference = "Stop"

function Write-Info($msg)  { Write-Host "[INFO] $msg" -ForegroundColor Cyan }
function Write-Ok($msg)    { Write-Host "[ OK ] $msg" -ForegroundColor Green }
function Write-Warn($msg)  { Write-Host "[WARN] $msg" -ForegroundColor Yellow }
function Write-Err($msg)   { Write-Host "[ERR ] $msg" -ForegroundColor Red }

# Resolve absolute paths
$XlsxFull = (Resolve-Path -LiteralPath $XlsxPath -ErrorAction SilentlyContinue).Path
if (-not $XlsxFull) {
    Write-Err "Arquivo $XlsxPath nao encontrado. Rode 'python build.py' antes."
    exit 1
}
$XlsmFull = Join-Path (Split-Path $XlsxFull -Parent) (Split-Path $XlsmPath -Leaf)
$VbaFull = (Resolve-Path -LiteralPath $VbaDir).Path

Write-Info "XLSX origem  : $XlsxFull"
Write-Info "XLSM destino : $XlsmFull"
Write-Info "VBA folder   : $VbaFull"

# Module type constants
# 1 = vbext_ct_StdModule, 2 = vbext_ct_ClassModule, 100 = vbext_ct_Document
$VBEXT_CT_STDMODULE  = 1
$VBEXT_CT_CLASSMODULE = 2
$VBEXT_CT_DOCUMENT   = 100

# xlOpenXMLWorkbookMacroEnabled = 52
$XL_FORMAT_XLSM = 52

Write-Info "Iniciando Excel..."
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Excel.DisplayAlerts = $false
$Excel.AutomationSecurity = 3   # msoAutomationSecurityForceDisable

try {
    Write-Info "Abrindo workbook..."
    $wb = $Excel.Workbooks.Open($XlsxFull)

    Write-Info "Acessando projeto VBA..."
    try {
        $vbProject = $wb.VBProject
    } catch {
        Write-Err "Nao foi possivel acessar o projeto VBA."
        Write-Err "Habilite: Arquivo > Opcoes > Central de Confiabilidade >"
        Write-Err "Configuracoes > Confiar no acesso ao modelo de objeto VBA."
        throw
    }

    # ----- Limpa modulos pre-existentes (exceto ThisWorkbook/Sheets) ----------
    $toRemove = @()
    foreach ($comp in $vbProject.VBComponents) {
        if ($comp.Type -eq $VBEXT_CT_STDMODULE -or $comp.Type -eq $VBEXT_CT_CLASSMODULE) {
            $toRemove += $comp.Name
        }
    }
    foreach ($name in $toRemove) {
        try {
            $c = $vbProject.VBComponents.Item($name)
            $vbProject.VBComponents.Remove($c) | Out-Null
            Write-Info "Removido modulo existente: $name"
        } catch { }
    }

    # ----- Importa modulos .bas e .cls ---------------------------------------
    $basFiles = Get-ChildItem -LiteralPath $VbaFull -Filter "*.bas" -File
    foreach ($f in $basFiles) {
        $vbProject.VBComponents.Import($f.FullName) | Out-Null
        Write-Ok "Importado: $($f.Name)"
    }

    # ----- ThisWorkbook.cls e tratado de forma especial ----------------------
    $twPath = Join-Path $VbaFull "ThisWorkbook.cls"
    if (Test-Path -LiteralPath $twPath) {
        $tw = $vbProject.VBComponents.Item("ThisWorkbook")
        $cm = $tw.CodeModule
        if ($cm.CountOfLines -gt 0) {
            $cm.DeleteLines(1, $cm.CountOfLines)
        }
        # Le o conteudo, descarta cabecalho VERSION/BEGIN/END/Attribute - VBA
        # rejeita esses cabecalhos quando colados em modulo de documento ja
        # existente.
        $rawLines = Get-Content -LiteralPath $twPath
        $skip = $true
        $body = New-Object System.Collections.Generic.List[string]
        foreach ($line in $rawLines) {
            if ($skip) {
                if ($line -match '^\s*Attribute\s+VB_Exposed') { $skip = $false; continue }
                if ($line -match '^\s*(VERSION|BEGIN|END|MultiUse|Attribute)') { continue }
                # First non-header line - stop skipping
                $skip = $false
            }
            $body.Add($line)
        }
        $cm.AddFromString(($body -join "`r`n"))
        Write-Ok "Importado ThisWorkbook.cls"
    }

    # ----- Renomeia projeto VBA ----------------------------------------------
    try {
        $vbProject.Name = "EXAUSTAO360"
    } catch {
        Write-Warn "Nao foi possivel renomear o projeto VBA: $($_.Exception.Message)"
    }

    # ----- Compila para detectar erros ---------------------------------------
    Write-Info "Compilando projeto VBA..."
    try {
        $Excel.Run("Inicializar360") | Out-Null
        Write-Ok "Macro de inicializacao executada (smoke test)."
    } catch {
        Write-Warn "Smoke test falhou: $($_.Exception.Message)"
    }

    # ----- Salva como .xlsm --------------------------------------------------
    Write-Info "Salvando como .xlsm..."
    $wb.SaveAs($XlsmFull, $XL_FORMAT_XLSM)
    Write-Ok "Arquivo final: $XlsmFull"

    $wb.Close($false)
}
finally {
    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Ok "EXAUSTAO 360 ENTERPRISE PRO assemble concluido."
