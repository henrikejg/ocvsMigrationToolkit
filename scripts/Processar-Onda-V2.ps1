<#
.SYNOPSIS
    Processa todos os servidores de uma onda (orquestrador do Processar-Servidor.ps1)
.PARAMETER NumeroOnda
    Número da onda a processar
.PARAMETER ArquivoExcel
    Caminho para o arquivo .xlsx (opcional)
.PARAMETER IncluirPublicos
    Se informado, processa também comunicações com IPs públicos
.PARAMETER Forcar
    Se informado, ignora controle e reprocessa tudo do zero
.EXAMPLE
    .\Processar-Onda-V2.ps1 -NumeroOnda 2
    .\Processar-Onda-V2.ps1 -NumeroOnda 2 -Forcar
#>
param(
    [Parameter(Mandatory)][string]$NumeroOnda,
    [string]$ArquivoExcel     = "",
    [switch]$IncluirPublicos,
    [switch]$Forcar
)

$scriptDir = $PSScriptRoot
$DirBase   = Split-Path $scriptDir -Parent

# ── Resolver Excel ─────────────────────────────────────────────────────────────
if ($ArquivoExcel) {
    $ArquivoExcel = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ArquivoExcel)
} else {
    $candidatos = @(Get-ChildItem -Path $DirBase -Filter "*.xlsx" -File -ErrorAction SilentlyContinue)
    if ($candidatos.Count -eq 0) {
        Write-Error "Nenhum arquivo .xlsx encontrado. Use -ArquivoExcel para especificar."
        exit 1
    }
    $ArquivoExcel = $candidatos[0].FullName
}
Write-Host "Planilha: $ArquivoExcel" -ForegroundColor Cyan

# ── Carregar hostnames da onda ─────────────────────────────────────────────────
$hostnames = @(& "$scriptDir\Extrair-Hostnames.ps1" -NumeroOnda $NumeroOnda -ArquivoExcel $ArquivoExcel)
if (-not $hostnames -or $hostnames.Count -eq 0) {
    Write-Error "Nenhum servidor encontrado para Onda $NumeroOnda"
    exit 1
}

Write-Host ""
Write-Host ("=" * 50)
Write-Host "Processamento por Servidor — Onda $NumeroOnda" -ForegroundColor Yellow
Write-Host ("=" * 50)
Write-Host "$($hostnames.Count) servidores na onda"
Write-Host ""

# ── Processar cada servidor ────────────────────────────────────────────────────
$ok = 0; $pulados = 0; $erros = 0

for ($i = 0; $i -lt $hostnames.Count; $i++) {
    $h = $hostnames[$i].Trim()
    if (-not $h) { continue }

    Write-Host "[$($i+1)/$($hostnames.Count)] " -NoNewline -ForegroundColor Gray

    $args_servidor = @{
        Hostname     = $h
        ArquivoExcel = $ArquivoExcel
        DirBase      = $DirBase
    }
    if ($IncluirPublicos) { $args_servidor["IncluirPublicos"] = $true }
    if ($Forcar) { $args_servidor["Forcar"] = $true }

    try {
        & "$scriptDir\Processar-Servidor.ps1" @args_servidor
        $exitCode = $LASTEXITCODE
        if ($exitCode -eq 0) {
            # Verificar se foi pulado (já atualizado) pelo output
            $ok++
        } else {
            $erros++
        }
    } catch {
        Write-Host "  x Erro: $_" -ForegroundColor Red
        $erros++
    }
}

# ── Resumo ─────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host ("=" * 50)
Write-Host "RESUMO — Onda $NumeroOnda" -ForegroundColor Yellow
Write-Host ("=" * 50)
Write-Host "Total de servidores: $($hostnames.Count)"
Write-Host "Processados/atualizados: $ok" -ForegroundColor Green
Write-Host "Erros: $erros" -ForegroundColor $(if ($erros -gt 0) { "Red" } else { "Gray" })
Write-Host ("=" * 50)
Write-Host ""
Write-Host "Processamento concluido!" -ForegroundColor Green
