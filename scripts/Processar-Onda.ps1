<#
.SYNOPSIS
    Organiza coletas de netstat e consolida dados de uma onda de migração.
.DESCRIPTION
    Equivalente PowerShell de processar_onda.sh
    - Lê hostnames da planilha Excel pela onda informada
    - Cria estrutura de diretórios: ONDA{N}/COLETAS, ONDA{N}/CONSOLIDADO, PROCESSADOS
    - Copia arquivos netstat da pasta raw para COLETAS
    - Executa os scripts AWK de consolidação (dependencias_ocvs.awk e aglutinar_ocvs.awk)
      via WSL ou via awk.exe se disponível no PATH

    Pré-requisitos:
      - Módulo ImportExcel: Install-Module ImportExcel -Scope CurrentUser
      - Scripts AWK em: $HOME\dependencias_ocvs.awk e $HOME\aglutinar_ocvs.awk
        (ou equivalente em $env:USERPROFILE)
      - awk.exe no PATH, ou WSL disponível para executar os AWK scripts

.PARAMETER NumeroOnda
    Número da onda a processar
.PARAMETER ArquivoExcel
    Caminho para o arquivo .xlsx (padrão: planilha na pasta pai)
.PARAMETER DirBase
    Diretório base onde ficam as pastas ONDA*, raw, PROCESSADOS
    (padrão: pasta pai do script, equivalente ao DIR do bash)
.EXAMPLE
    .\Processar-Onda.ps1 -NumeroOnda 2
    .\Processar-Onda.ps1 -NumeroOnda 2 -DirBase "C:\ocvs\netstat"
#>
param(
    [Parameter(Mandatory)][string]$NumeroOnda,
    [string]$ArquivoExcel = "",
    [string]$DirBase      = ""
)

# Resolver DirBase aqui, depois do param, onde $PSScriptRoot ja esta disponivel
if (-not $DirBase) {
    $DirBase = Join-Path (Split-Path $PSScriptRoot -Parent) "dados"
}

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

# Resolver caminhos — procurar xlsx automaticamente se nao informado
if ($ArquivoExcel) {
    $ArquivoExcel = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ArquivoExcel)
} else {
    $candidatos = @(
        Get-ChildItem -Path (Split-Path $PSScriptRoot -Parent) -Filter "*.xlsx" -ErrorAction SilentlyContinue
    )
    if ($candidatos.Count -eq 0) {
        Write-Error "Nenhum arquivo .xlsx encontrado em '$(Split-Path $PSScriptRoot -Parent)'. Coloque a planilha na pasta V2/ ou use -ArquivoExcel para especificar o caminho."
        exit 1
    }
    if ($candidatos.Count -gt 1) {
        Write-Host "Multiplos .xlsx encontrados, usando o primeiro:" -ForegroundColor Yellow
        $candidatos | ForEach-Object { Write-Host "  $($_.FullName)" }
    }
    $ArquivoExcel = $candidatos[0].FullName
    Write-Host "Planilha: $ArquivoExcel" -ForegroundColor Cyan
}
$DirBase = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DirBase)

# ── Caminhos ─────────────────────────────────────────────────────────────────
$DirOnda        = Join-Path $DirBase "ONDA$NumeroOnda"
$DirColetas     = Join-Path $DirOnda "COLETAS"
$DirConsolidado = Join-Path $DirOnda "CONSOLIDADO"
$DirProcessados = Join-Path $DirBase "PROCESSADOS"
$DirRaw         = Join-Path $DirBase "raw"

$ArquivoConsolidado = Join-Path $DirConsolidado "ONDA${NumeroOnda}_consolidado.txt"
$ArquivoProcessado  = Join-Path $DirProcessados  "ONDA${NumeroOnda}_processado.txt"

# Scripts AWK — mesma pasta dos scripts (V2/scripts/)
$AwkDependencias = Join-Path $PSScriptRoot "dependencias_ocvs.awk"
$AwkAglutinar    = Join-Path $PSScriptRoot "aglutinar_ocvs.awk"

Write-Host ("=" * 40)
Write-Host "Lendo servidores com Onda $NumeroOnda do arquivo Excel..."
Write-Host ("=" * 40)

# ── Carregar servidores ───────────────────────────────────────────────────────
$scriptDir  = $PSScriptRoot
$servidores = @(& "$scriptDir\Extrair-Hostnames.ps1" -NumeroOnda $NumeroOnda -ArquivoExcel $ArquivoExcel)

if (-not $servidores -or $servidores.Count -eq 0) {
    Write-Error "Nenhum servidor com Onda $NumeroOnda encontrado!"
    exit 1
}

Write-Host "Servidores encontrados: $($servidores.Count)"
foreach ($s in $servidores) { Write-Host "  - $s" }
Write-Host ("=" * 40)

# ── Criar estrutura de diretórios ─────────────────────────────────────────────
foreach ($dir in @($DirColetas, $DirConsolidado, $DirProcessados)) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
}

# ── Copiar coletas da pasta raw ───────────────────────────────────────────────
Write-Host "Organizando coletas de netstat dos servidores da ONDA $NumeroOnda..."

# Limpar COLETAS antes de copiar — garante que só ficam os servidores atuais da onda
if (Test-Path $DirColetas) {
    Get-ChildItem -Path $DirColetas -Filter "netstat_*.txt" | Remove-Item -Force
    Write-Host "  Pasta COLETAS limpa para reprocessamento"
}

$copiados = 0
$naoEncontrados = @()

foreach ($servidor in $servidores) {
    $origem = Join-Path $DirRaw "netstat_${servidor}.txt"
    if (Test-Path $origem) {
        Copy-Item -Path $origem -Destination $DirColetas -Force
        $copiados++
    } else {
        $naoEncontrados += $servidor
    }
}

Write-Host "Arquivos organizados em $DirColetas ($copiados copiados)"

if ($naoEncontrados.Count -gt 0) {
    Write-Warning "Arquivos não encontrados em raw para: $($naoEncontrados -join ', ')"
}

# ── Normalizar hostname na coluna 2 de cada arquivo de coleta ─────────────────
# Garante que todas as linhas tenham o hostname correto (da planilha),
# independente do que o servidor reportou (ex: -, hostname.localdomain, typo)
Write-Host "Normalizando hostnames nos arquivos de coleta..."

$normalizados = 0
foreach ($servidor in $servidores) {
    $arquivo = Join-Path $DirColetas "netstat_${servidor}.txt"
    if (-not (Test-Path $arquivo)) { continue }

    $linhas = [System.IO.File]::ReadAllLines($arquivo, [System.Text.Encoding]::UTF8)
    $alterado = $false
    $linhasNormalizadas = for ($i = 0; $i -lt $linhas.Count; $i++) {
        $cols = $linhas[$i] -split ";"
        if ($cols.Count -ge 2 -and $cols[1] -ne $servidor) {
            $cols[1] = $servidor
            $alterado = $true
            $cols -join ";"
        } else {
            $linhas[$i]
        }
    }
    if ($alterado) {
        [System.IO.File]::WriteAllLines($arquivo, $linhasNormalizadas, [System.Text.Encoding]::UTF8)
        $normalizados++
    }
}

if ($normalizados -gt 0) {
    Write-Host "  $normalizados arquivo(s) com hostname normalizado" -ForegroundColor Yellow
}

# ── Executar AWK: dependencias_ocvs.awk ──────────────────────────────────────
Write-Host "Adicionando informacoes extras..."

$netstatFiles = Get-ChildItem -Path $DirColetas -Filter "netstat_*.txt" | Select-Object -ExpandProperty FullName

if ($netstatFiles.Count -eq 0) {
    Write-Error "Nenhum arquivo netstat encontrado em $DirColetas"
    exit 1
}

function Invoke-Awk {
    param([string]$ScriptAwk, [string[]]$Arquivos, [string]$Saida)

    # 1. awk.exe na mesma pasta dos scripts (empacotado com a solucao)
    $AwkLocal = Join-Path $PSScriptRoot "awk.exe"
    if (Test-Path $AwkLocal) {
        & $AwkLocal -f $ScriptAwk @Arquivos | Set-Content -Path $Saida -Encoding UTF8
        return $LASTEXITCODE -eq 0
    }

    # 2. awk no PATH do sistema
    if (Get-Command awk -ErrorAction SilentlyContinue) {
        $args = @($ScriptAwk) + $Arquivos
        awk -f @args > $Saida
        return $LASTEXITCODE -eq 0
    }

    # 3. Fallback: WSL
    if (Get-Command wsl -ErrorAction SilentlyContinue) {
        $awkWsl   = wsl wslpath ($ScriptAwk -replace '\\', '/')
        $saidaWsl = wsl wslpath ($Saida     -replace '\\', '/')
        $filesWsl = $Arquivos | ForEach-Object { wsl wslpath ($_ -replace '\\', '/') }
        wsl awk -f $awkWsl @filesWsl > $Saida
        return $LASTEXITCODE -eq 0
    }

    Write-Error "awk nao encontrado. Coloque awk.exe na pasta scripts/ ou instale Git for Windows."
    return $false
}

if (-not (Test-Path $AwkDependencias)) {
    Write-Error "Script AWK não encontrado: $AwkDependencias"
    exit 1
}

$ok = Invoke-Awk -ScriptAwk $AwkDependencias -Arquivos $netstatFiles -Saida $ArquivoConsolidado
if (-not $ok) {
    Write-Error "Falha ao executar dependencias_ocvs.awk"
    exit 1
}

# ── Executar AWK: aglutinar_ocvs.awk ─────────────────────────────────────────
Write-Host "Consolidando dados..."

if (-not (Test-Path $AwkAglutinar)) {
    Write-Error "Script AWK não encontrado: $AwkAglutinar"
    exit 1
}

$ok = Invoke-Awk -ScriptAwk $AwkAglutinar -Arquivos @($ArquivoConsolidado) -Saida $ArquivoProcessado
if (-not $ok) {
    Write-Error "Falha ao executar aglutinar_ocvs.awk"
    exit 1
}

Write-Host "Finalizado." -ForegroundColor Green
Write-Host "Arquivo gerado: $ArquivoProcessado"
