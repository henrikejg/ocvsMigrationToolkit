<#
.SYNOPSIS
    Inicia o OCVS Migration Dashboard v0.2.2
#>

$scriptDir = $PSScriptRoot

# Verificar Node.js
if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
    Write-Host ""
    Write-Host "Node.js nao encontrado!" -ForegroundColor Red
    Write-Host "Baixe e instale em: https://nodejs.org (versao LTS)" -ForegroundColor Yellow
    Write-Host ""
    Start-Process "https://nodejs.org"
    exit 1
}

$nodeVersion = node --version
Write-Host "Node.js $nodeVersion detectado" -ForegroundColor Green

# Instalar/atualizar dependencias
$nodeModules = Join-Path $scriptDir "node_modules"
$sqlJsPath   = Join-Path $nodeModules "sql.js"
if ((-not (Test-Path $nodeModules)) -or (-not (Test-Path $sqlJsPath))) {
    Write-Host "Instalando dependencias..." -ForegroundColor Cyan
    Push-Location $scriptDir
    npm install
    Pop-Location
}

# Iniciar servidor — injetar o path do PowerShell atual para o Node usar
Write-Host ""
Write-Host "Iniciando dashboard em http://localhost:5000" -ForegroundColor Cyan
Write-Host "Pressione Ctrl+C para encerrar" -ForegroundColor Gray
Write-Host ""

$env:OCVS_PWSH = (Get-Process -Id $PID).Path

Push-Location $scriptDir
node server.js
Pop-Location
