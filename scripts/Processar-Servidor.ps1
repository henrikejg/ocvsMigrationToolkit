<#
.SYNOPSIS
    Processa dados netstat de um servidor individual (enriquecimento + aglutinação incremental)
.PARAMETER Hostname
    Nome do servidor a processar
.PARAMETER ArquivoExcel
    Caminho para o arquivo .xlsx (opcional)
.PARAMETER IncluirPublicos
    Se informado, processa também comunicações com IPs públicos
.PARAMETER Forcar
    Se informado, ignora controle e reprocessa tudo do zero
.PARAMETER DirBase
    Diretório base (padrão: pasta pai do script)
.EXAMPLE
    .\Processar-Servidor.ps1 -Hostname "AE-B21-WEBH01"
    .\Processar-Servidor.ps1 -Hostname "AE-B21-WEBH01" -Forcar
#>
param(
    [Parameter(Mandatory)][string]$Hostname,
    [string]$ArquivoExcel     = "",
    [switch]$IncluirPublicos,
    [switch]$Forcar,
    [string]$DirBase          = ""
)

# ── Resolver caminhos ──────────────────────────────────────────────────────────
if (-not $DirBase) { $DirBase = Split-Path $PSScriptRoot -Parent }
$DirBase = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DirBase)

$DirDados       = Join-Path $DirBase "dados"
$DirRaw         = Join-Path $DirDados "raw"
$DirPrivado     = Join-Path $DirDados "privado"
$DirPublico     = Join-Path $DirDados "publico"
$DirProcessados = Join-Path $DirDados "processados_servidor"
$DirProcPub     = Join-Path $DirDados "processados_servidor_publico"
$DirControle    = Join-Path $DirDados "controle"

# Criar pastas se não existem
foreach ($d in @($DirPrivado, $DirPublico, $DirProcessados, $DirProcPub, $DirControle)) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}

# Scripts AWK
$AwkFiltrar      = Join-Path $PSScriptRoot "filtrar_ip.awk"
$AwkDependencias = Join-Path $PSScriptRoot "dependencias_ocvs.awk"
$AwkAglutinar    = Join-Path $PSScriptRoot "aglutinar_ocvs.awk"

# awk.exe na mesma pasta dos scripts
$AwkExe = Join-Path $PSScriptRoot "awk.exe"
if (-not (Test-Path $AwkExe)) {
    Write-Error "awk.exe nao encontrado em $PSScriptRoot"
    exit 1
}

# ── Helpers ────────────────────────────────────────────────────────────────────
function Write-Ok   { param($m) Write-Host "  v $m" -ForegroundColor Green  }
function Write-Fail { param($m) Write-Host "  x $m" -ForegroundColor Red    }
function Write-Info { param($m) Write-Host "  $m" -ForegroundColor Cyan     }

# Converter paths Windows para formato que o AWK aceita (barras normais)
function ConvertTo-AwkPath { param($p) $p -replace '\\', '/' }

# ── Localizar arquivo raw ──────────────────────────────────────────────────────
$HostnameUpper = $Hostname.ToUpper()
$HostnameLower = $Hostname.ToLower()

# Tentar variações de nome
$rawCandidatos = @(
    Join-Path $DirRaw "netstat_$HostnameUpper.txt"
    Join-Path $DirRaw "netstat_$HostnameLower.txt"
    Join-Path $DirRaw "netstat_$Hostname.txt"
)
$ArquivoRaw = $null
foreach ($c in $rawCandidatos) {
    if (Test-Path $c) { $ArquivoRaw = $c; break }
}
if (-not $ArquivoRaw) {
    Write-Fail "Arquivo raw nao encontrado para $Hostname"
    Write-Host "    Procurado em: $DirRaw\netstat_$Hostname.txt"
    exit 1
}

# ── Arquivos de saída ──────────────────────────────────────────────────────────
$ArqPrivado    = Join-Path $DirPrivado     "netstat_$HostnameUpper.txt"
$ArqPublico    = Join-Path $DirPublico     "netstat_$HostnameUpper.txt"
$ArqProcessado = Join-Path $DirProcessados "netstat_$HostnameUpper.txt"
$ArqProcPub    = Join-Path $DirProcPub     "netstat_$HostnameUpper.txt"
$ArqControle   = Join-Path $DirControle    "$HostnameUpper.json"

# ── Verificar se precisa processar ─────────────────────────────────────────────
$rawModificado = (Get-Item $ArquivoRaw).LastWriteTime
$controle = $null
if ((Test-Path $ArqControle) -and -not $Forcar) {
    try {
        $controle = Get-Content $ArqControle -Raw | ConvertFrom-Json
        $ultimoProcessamento = [DateTime]::ParseExact($controle.dataProcessamento, "o", [System.Globalization.CultureInfo]::InvariantCulture)
    } catch {
        # Controle corrompido ou formato inválido — ignorar e reprocessar
        $controle = $null
    }
    if ($controle -and $rawModificado -le $ultimoProcessamento) {
        Write-Ok "$Hostname — ja esta atualizado (raw nao mudou)"
        exit 0
    }
}

Write-Host "Processando $Hostname..." -ForegroundColor Yellow
$linhasRaw = (Get-Content $ArquivoRaw | Measure-Object -Line).Lines
Write-Info "$linhasRaw linhas no arquivo raw"

# ── Etapa 0: Normalizar hostname na coluna 2 do arquivo raw ───────────────────
# Garante que todas as linhas tenham o hostname correto (da planilha),
# independente do que o servidor reportou (ex: -, hostname.localdomain, typo)
$linhasArq = [System.IO.File]::ReadAllLines($ArquivoRaw, [System.Text.Encoding]::UTF8)
$alterado = $false
$linhasNorm = for ($li = 0; $li -lt $linhasArq.Count; $li++) {
    $cols = $linhasArq[$li] -split ";"
    if ($cols.Count -ge 2 -and $cols[1] -ne $HostnameUpper) {
        $cols[1] = $HostnameUpper
        $alterado = $true
        $cols -join ";"
    } else {
        $linhasArq[$li]
    }
}
if ($alterado) {
    [System.IO.File]::WriteAllLines($ArquivoRaw, $linhasNorm, [System.Text.Encoding]::UTF8)
    Write-Info "Hostname normalizado no arquivo raw"
}

# ── Etapa 1: Pré-filtro (separar privado/público) ─────────────────────────────
Write-Info "Separando IPs privados e publicos..."
& $AwkExe -f $AwkFiltrar -v "PRIVADO=$(ConvertTo-AwkPath $ArqPrivado)" -v "PUBLICO=$(ConvertTo-AwkPath $ArqPublico)" $ArquivoRaw

$linhasPrivado = if (Test-Path $ArqPrivado) { (Get-Content $ArqPrivado | Measure-Object -Line).Lines } else { 0 }
$linhasPublico = if (Test-Path $ArqPublico) { (Get-Content $ArqPublico | Measure-Object -Line).Lines } else { 0 }
Write-Info "Privado: $linhasPrivado | Publico: $linhasPublico"

# ── Etapa 2: Identificar linhas novas (incremental) ───────────────────────────
$arquivoParaEnriquecer = $ArqPrivado
$modoIncremental = $false

if ($controle -and -not $Forcar -and (Test-Path $ArqProcessado)) {
    $ultimoTimestamp = $controle.ultimoTimestamp
    if ($ultimoTimestamp) {
        Write-Info "Modo incremental — filtrando linhas apos $ultimoTimestamp"
        $tempNovas = Join-Path $env:TEMP "ocvs_novas_$HostnameUpper.txt"
        # Filtrar linhas com timestamp maior que o último processado
        & $AwkExe -F ";" -v "ULTIMO=$ultimoTimestamp" 'BEGIN{OFS=";"} $1 > ULTIMO {print}' $ArqPrivado | Set-Content -Path $tempNovas -Encoding UTF8
        $linhasNovas = (Get-Content $tempNovas | Measure-Object -Line).Lines
        if ($linhasNovas -eq 0) {
            Write-Ok "$Hostname — sem linhas novas no privado"
            Remove-Item $tempNovas -Force -ErrorAction SilentlyContinue
            # Atualizar controle mesmo sem linhas novas (raw pode ter mudado por outro motivo)
            $controle.dataProcessamento = (Get-Date).ToString("o")
            $controle | ConvertTo-Json | Set-Content -Path $ArqControle -Encoding UTF8
            exit 0
        }
        Write-Info "$linhasNovas linhas novas detectadas"
        $arquivoParaEnriquecer = $tempNovas
        $modoIncremental = $true
    }
}

# ── Etapa 3: Enriquecer com dependencias_ocvs.awk ─────────────────────────────
Write-Info "Enriquecendo dados..."
$tempEnriquecido = Join-Path $env:TEMP "ocvs_enriquecido_$HostnameUpper.txt"
& $AwkExe -f $AwkDependencias $arquivoParaEnriquecer | Set-Content -Path $tempEnriquecido -Encoding UTF8

# ── Etapa 4: Aglutinar ────────────────────────────────────────────────────────
Write-Info "Aglutinando conexoes..."
$tempAglutinado = Join-Path $env:TEMP "ocvs_aglutinado_$HostnameUpper.txt"
& $AwkExe -f $AwkAglutinar $tempEnriquecido | Set-Content -Path $tempAglutinado -Encoding UTF8

# ── Etapa 5: Merge com processado existente (se incremental) ──────────────────
if ($modoIncremental -and (Test-Path $ArqProcessado)) {
    Write-Info "Mesclando com dados anteriores e re-aglutinando..."
    $tempMerge = Join-Path $env:TEMP "ocvs_merge_$HostnameUpper.txt"
    # Concatenar anterior + novo
    Get-Content $ArqProcessado, $tempAglutinado | Set-Content -Path $tempMerge -Encoding UTF8
    # Re-aglutinar
    & $AwkExe -f $AwkAglutinar $tempMerge | Set-Content -Path $ArqProcessado -Encoding UTF8
    Remove-Item $tempMerge -Force -ErrorAction SilentlyContinue
} else {
    # Primeira vez ou forçado — usar aglutinado direto
    Copy-Item $tempAglutinado $ArqProcessado -Force
}

$linhasProcessadas = (Get-Content $ArqProcessado | Measure-Object -Line).Lines
Write-Ok "$Hostname — $linhasProcessadas linhas processadas (privado)"

# ── Etapa 6: Processar públicos (se solicitado) ───────────────────────────────
if ($IncluirPublicos -and (Test-Path $ArqPublico) -and $linhasPublico -gt 0) {
    Write-Info "Processando comunicacoes publicas..."
    $tempEnrPub = Join-Path $env:TEMP "ocvs_enr_pub_$HostnameUpper.txt"
    $tempAglPub = Join-Path $env:TEMP "ocvs_agl_pub_$HostnameUpper.txt"
    & $AwkExe -f $AwkDependencias $ArqPublico | Set-Content -Path $tempEnrPub -Encoding UTF8
    & $AwkExe -f $AwkAglutinar $tempEnrPub | Set-Content -Path $tempAglPub -Encoding UTF8
    Copy-Item $tempAglPub $ArqProcPub -Force
    $linhasProcPub = (Get-Content $ArqProcPub | Measure-Object -Line).Lines
    Write-Ok "$Hostname — $linhasProcPub linhas processadas (publico)"
    Remove-Item $tempEnrPub, $tempAglPub -Force -ErrorAction SilentlyContinue
}

# ── Etapa 7: Extrair último timestamp para controle ───────────────────────────
$ultimoTS = ""
if (Test-Path $ArqPrivado) {
    # Pegar o maior timestamp do arquivo privado (campo 1, formato "YYYY-MM-DD HH:MM:SS")
    $ultimoTS = & $AwkExe -F ";" 'BEGIN{max=""} {if($1>max)max=$1} END{print max}' $ArqPrivado
    $ultimoTS = $ultimoTS.Trim()
}

# ── Etapa 8: Salvar controle ──────────────────────────────────────────────────
$novoControle = @{
    hostname           = $HostnameUpper
    ultimoTimestamp     = $ultimoTS
    linhasRaw          = $linhasRaw
    linhasPrivado      = $linhasPrivado
    linhasPublico      = $linhasPublico
    linhasProcessadas  = $linhasProcessadas
    dataProcessamento  = (Get-Date).ToString("o")
    incremental        = $modoIncremental
}
$novoControle | ConvertTo-Json | Set-Content -Path $ArqControle -Encoding UTF8

# ── Limpeza ────────────────────────────────────────────────────────────────────
Remove-Item (Join-Path $env:TEMP "ocvs_novas_$HostnameUpper.txt") -Force -ErrorAction SilentlyContinue
Remove-Item $tempEnriquecido -Force -ErrorAction SilentlyContinue
Remove-Item $tempAglutinado -Force -ErrorAction SilentlyContinue

Write-Ok "$Hostname concluido"
