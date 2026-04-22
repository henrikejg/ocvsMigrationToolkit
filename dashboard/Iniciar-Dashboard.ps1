<#
.SYNOPSIS
    Inicia o OCVS Migration Dashboard v0.3.8
#>

$scriptDir = $PSScriptRoot

# Garantir que a execution policy permite rodar scripts nesta sessao
$currentPolicy = Get-ExecutionPolicy -Scope Process
if ($currentPolicy -eq "Restricted" -or $currentPolicy -eq "Undefined") {
    try {
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned -Force
        Write-Host "Execution policy ajustada para RemoteSigned (sessao atual)" -ForegroundColor Yellow
    } catch {
        Write-Host ""
        Write-Host "Erro: politica de execucao do PowerShell impede a execucao de scripts." -ForegroundColor Red
        Write-Host "Execute o comando abaixo como Administrador e tente novamente:" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned" -ForegroundColor Cyan
        Write-Host ""
        exit 1
    }
}

# Garantir SSH config para compatibilidade com servidores legado
$sshDir    = Join-Path $env:USERPROFILE ".ssh"
$sshConfig = Join-Path $sshDir "config"
if (-not (Test-Path $sshConfig)) {
    if (-not (Test-Path $sshDir)) {
        New-Item -ItemType Directory -Path $sshDir -Force | Out-Null
    }
    $sshContent = @"
Host *
    KexAlgorithms +diffie-hellman-group1-sha1
    HostKeyAlgorithms +ssh-rsa
    Ciphers +aes128-cbc
    StrictHostKeyChecking no
"@
    Set-Content -Path $sshConfig -Value $sshContent -Encoding UTF8
    Write-Host "SSH config criado em $sshConfig (compatibilidade com servidores legado)" -ForegroundColor Yellow
}

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
