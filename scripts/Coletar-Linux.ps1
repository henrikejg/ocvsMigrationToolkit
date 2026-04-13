<#
.SYNOPSIS
    Coleta arquivos netstat de servidores Linux via SSH com compressão.
.DESCRIPTION
    Equivalente PowerShell de coletar_linux.sh
    - Solicita credenciais via prompt seguro do Windows (Get-Credential)
    - Lê servidores da planilha Excel pela onda informada (IPs)
    - Detecta SO via TTL do ping (Linux: 50-64 ou 240-255 | Windows: 110-128)
    - Conecta via SSH com opções de compatibilidade para servidores legados
    - Transfere /tmp/netstat_*.txt com compressão tar+gzip
    - Exibe relatório final com estatísticas e economia de banda

    Pre-requisitos:
      - OpenSSH instalado (ssh.exe disponivel no PATH)
      - Modulo ImportExcel: Install-Module ImportExcel -Scope CurrentUser

.PARAMETER NumeroOnda
    Numero da onda a processar
.PARAMETER ArquivoExcel
    Caminho para o arquivo .xlsx (padrao: planilha na pasta pai)
.PARAMETER Destino
    Pasta de destino para os arquivos coletados (padrao: ..\raw\)
.PARAMETER Usuario
    Usuario SSH (padrao: migracao)
.EXAMPLE
    .\Coletar-Linux.ps1 -NumeroOnda 2
    .\Coletar-Linux.ps1 -NumeroOnda 2 -ArquivoExcel "C:\dados\planilha.xlsx"
#>
param(
    [Parameter(Mandatory)][string]$NumeroOnda,
    [string]$ArquivoExcel = "",
    [string]$Destino      = "",
    [string]$Usuario      = "migracao",
    [string]$Senha        = ""
)

# Resolver Destino aqui, depois do param, onde $PSScriptRoot ja esta disponivel
if (-not $Destino) {
    $Destino = Join-Path (Split-Path $PSScriptRoot -Parent) "dados\raw\"
}

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

# Resolver caminhos relativos com base no diretorio de trabalho atual
if ($ArquivoExcel) {
    $ArquivoExcel = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ArquivoExcel)
} else {
    # Procurar automaticamente: primeiro na pasta do script, depois na pasta pai
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
$Destino = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Destino)

# ── Helpers de cor ────────────────────────────────────────────────────────────
function Write-Ok   { param($m) Write-Host "  v $m" -ForegroundColor Green  }
function Write-Fail { param($m) Write-Host "  x $m" -ForegroundColor Red    }
function Write-Warn { param($m) Write-Host "  ! $m" -ForegroundColor Yellow }
function Write-Sep  { Write-Host ("=" * 40) }
function Write-Dash { Write-Host ("-" * 40) }

function Format-Bytes {
    param([long]$Bytes)
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    return "$Bytes bytes"
}

# ── Solicitar credenciais via prompt seguro ───────────────────────────────────
Write-Host ""
Write-Host "Informe as credenciais SSH para os servidores da Onda ${NumeroOnda}:" -ForegroundColor Cyan

# Modo headless: senha via parametro ou variavel de ambiente (chamado pelo dashboard)
if (-not $Senha) { $Senha = $env:OCVS_SENHA }

if ($Senha) {
    Write-Host "Usando credenciais fornecidas para usuario '$Usuario'" -ForegroundColor Green
    $SenhaPlana = $Senha
} else {
    $credencial = Get-Credential -UserName $Usuario -Message "Credenciais SSH - Onda ${NumeroOnda}"
    if (-not $credencial) {
        Write-Error "Credenciais nao informadas. Abortando."
        exit 1
    }
    $Usuario    = $credencial.UserName
    $SenhaPlana = $credencial.GetNetworkCredential().Password
}

# ── Criar script askpass temporario ──────────────────────────────────────────
# Pasta unica por execucao para evitar conflitos
$askpassDir = Join-Path $env:TEMP ("ssh_askpass_" + [System.IO.Path]::GetRandomFileName().Replace('.',''))
New-Item -ItemType Directory -Path $askpassDir -Force | Out-Null
$askpassBat = Join-Path $askpassDir "askpass.bat"

# O .bat imprime a senha — chamado automaticamente pelo ssh.exe via SSH_ASKPASS
# Usar WriteAllBytes para evitar que caracteres especiais (!, ^, &, %) sejam
# interpretados pelo cmd.exe durante a interpolacao de string no Set-Content
$batConteudo = "@echo off`r`necho $SenhaPlana"
# Escapar caracteres especiais do cmd: !, ^, &, |, <, >, %
# O echo sem aspas no cmd interpreta esses caracteres — usar escape ^
$SenhaEscapada = $SenhaPlana `
    -replace '\^', '^^' `
    -replace '!',  '^!' `
    -replace '&',  '^&' `
    -replace '\|', '^|' `
    -replace '<',  '^<' `
    -replace '>',  '^>' `
    -replace '%',  '%%'
$batConteudo = "@echo off`r`necho $SenhaEscapada"
[System.IO.File]::WriteAllText($askpassBat, $batConteudo, [System.Text.Encoding]::ASCII)

# ── Opcoes SSH para servidores legados ────────────────────────────────────────
$SshOpts = @(
    "-o", "KexAlgorithms=+diffie-hellman-group1-sha1",
    "-o", "HostKeyAlgorithms=+ssh-rsa",
    "-o", "Ciphers=+aes128-cbc",
    "-o", "MACs=+hmac-sha1",
    "-o", "StrictHostKeyChecking=no",
    "-o", "PasswordAuthentication=yes",
    "-o", "BatchMode=no"
)

# ── Funcao auxiliar: executar SSH com senha via askpass ───────────────────────
function Invoke-SSH {
    param(
        [string[]]$SshArgs,
        [bool]$CapturarSaida = $false
    )

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo.FileName               = "ssh"
    $proc.StartInfo.Arguments              = ($SshArgs -join " ")
    $proc.StartInfo.UseShellExecute        = $false
    $proc.StartInfo.RedirectStandardOutput = $true
    $proc.StartInfo.RedirectStandardError  = $true
    $proc.StartInfo.CreateNoWindow         = $true
    $proc.StartInfo.EnvironmentVariables["SSH_ASKPASS"]         = $askpassBat
    $proc.StartInfo.EnvironmentVariables["SSH_ASKPASS_REQUIRE"] = "force"
    $proc.StartInfo.EnvironmentVariables["DISPLAY"]             = "localhost:0"

    $proc.Start() | Out-Null

    if ($CapturarSaida) {
        # Ler stderr em task paralela para nao bloquear
        $stderrTask = $proc.StandardError.ReadToEndAsync()
        $saida = $proc.StandardOutput.ReadToEnd()
        $proc.WaitForExit()
        $stderr = $stderrTask.Result
        return [PSCustomObject]@{ Saida = $saida; ExitCode = $proc.ExitCode; Stderr = $stderr }
    }

    # Modo binario: retornar processo para o chamador ler stdout como stream
    # Stderr sera lido apos WaitForExit pelo chamador via $proc.StandardError
    return [PSCustomObject]@{ Proc = $proc }
}

# ── Detectar SO via TTL ───────────────────────────────────────────────────────
function Get-ServidorTipo {
    param([string]$Servidor)

    # Test-Connection compativel com PS5 (sem -TimeoutSeconds)
    $ping = Test-Connection -ComputerName $Servidor -Count 1 -ErrorAction SilentlyContinue
    if (-not $ping) { return "indisponivel" }

    $ttl = $null
    # PS5: objeto retornado tem propriedade ReplySize, TTL via .ResponseTimeToLive ou ping nativo
    if ($ping.PSObject.Properties['ResponseTimeToLive']) {
        $ttl = $ping.ResponseTimeToLive
    }
    if (-not $ttl) {
        $raw = (ping -n 1 -w 3000 $Servidor 2>$null) -join "`n"
        if ($raw -match "TTL=(\d+)") { $ttl = [int]$Matches[1] }
    }

    if (-not $ttl) { return "desconhecido" }

    if    ($ttl -ge 50  -and $ttl -le 64)  { return "linux"        }
    elseif($ttl -ge 240 -and $ttl -le 255) { return "linux"        }
    elseif($ttl -ge 110 -and $ttl -le 128) { return "windows"      }
    else                                    { return "indeterminado" }
}

# ── Copiar com compressao tar+gzip ────────────────────────────────────────────
function Copiar-ComCompressao {
    param([string]$Servidor, [string]$DestinoLocal, [string]$NomeServidor)

    $tempDir  = Join-Path $env:TEMP "netstat_coleta"
    $tempFile = Join-Path $tempDir "${NomeServidor}_netstat.tar.gz"
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

    $sshArgs = $SshOpts + @("${Usuario}@${Servidor}", "cd /tmp && tar czf - netstat_*.txt 2>/dev/null")

    try {
        $r = Invoke-SSH -SshArgs $sshArgs -CapturarSaida $false
        $proc = $r.Proc

        $stderrTask = $proc.StandardError.ReadToEndAsync()

        # ── Transferencia com barra de progresso ──────────────────────────────
        $srcStream  = $proc.StandardOutput.BaseStream
        $fs         = [System.IO.File]::OpenWrite($tempFile)
        $buffer     = New-Object byte[] 65536  # 64 KB por leitura
        $totalBytes = 0
        $swProgress = [System.Diagnostics.Stopwatch]::StartNew()

        Write-Host "    Transferindo..." -NoNewline

        while ($true) {
            $lidos = $srcStream.Read($buffer, 0, $buffer.Length)
            if ($lidos -eq 0) { break }
            $fs.Write($buffer, 0, $lidos)
            $totalBytes += $lidos

            # Atualizar progresso a cada ~200ms para nao sobrecarregar o terminal
            if ($swProgress.ElapsedMilliseconds -ge 200) {
                $mbRecebidos = [math]::Round($totalBytes / 1MB, 2)
                Write-Progress -Activity "Coletando $Servidor" `
                    -Status ("Recebido: {0}" -f (Format-Bytes $totalBytes)) `
                    -PercentComplete -1  # indeterminado — nao sabemos o tamanho final
                $swProgress.Restart()
            }
        }

        $fs.Close()
        $proc.WaitForExit()
        Write-Progress -Activity "Coletando $Servidor" -Completed

        $stderrMsg         = $stderrTask.Result.Trim()
        $tamanhoCompactado = (Get-Item $tempFile -ErrorAction SilentlyContinue).Length

        if ($proc.ExitCode -ne 0 -or -not $tamanhoCompactado -or $tamanhoCompactado -eq 0) {
            Write-Host ""
            Write-Fail "Falha na compactacao/transferencia (exit $($proc.ExitCode))"
            if ($stderrMsg) { Write-Host "    SSH erro: $stderrMsg" -ForegroundColor DarkRed }
            Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
            return $false
        }

        # Extrair tar.gz — tar.exe nativo (Win10 1803+) ou fallback WSL
        Write-Host "`r    Extraindo...                    " -NoNewline
        $extraido = $false
        if (Get-Command tar -ErrorAction SilentlyContinue) {
            tar xzf $tempFile -C $DestinoLocal 2>$null
            $extraido = ($LASTEXITCODE -eq 0)
        }
        if (-not $extraido -and (Get-Command wsl -ErrorAction SilentlyContinue)) {
            $tmpWsl  = wsl wslpath ($tempFile     -replace '\\', '/')
            $destWsl = wsl wslpath ($DestinoLocal -replace '\\', '/')
            wsl tar xzf $tmpWsl -C $destWsl
            $extraido = ($LASTEXITCODE -eq 0)
        }

        if (-not $extraido) {
            Write-Host ""
            Write-Fail "Falha na extracao (tar nao disponivel)"
            Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
            return $false
        }

        # Calcular tamanho original e estatisticas
        $sshSizeArgs = $SshOpts + @(
            "${Usuario}@${Servidor}",
            'cd /tmp && du -sb netstat_*.txt 2>/dev/null | awk ''{sum+=$1} END {print sum}'''
        )
        $r2              = Invoke-SSH -SshArgs $sshSizeArgs -CapturarSaida $true
        $tamanhoOriginal = ($r2.Saida.Trim()) -as [long]

        Remove-Item $tempFile -Force -ErrorAction SilentlyContinue

        if ($tamanhoOriginal -and $tamanhoCompactado) {
            $economia = $tamanhoOriginal - $tamanhoCompactado
            $script:TotalEconomiaBanda    += $economia
            $script:TotalBaixado          += $tamanhoCompactado
            $script:TotalOriginal         += $tamanhoOriginal
            Write-Host ("`r    v {0} originais -> {1} transferidos (economia de {2})" -f `
                (Format-Bytes $tamanhoOriginal), (Format-Bytes $tamanhoCompactado), (Format-Bytes $economia)) `
                -ForegroundColor Green
        } else {
            $script:TotalBaixado  += $tamanhoCompactado
            Write-Host ("`r    v Transferido: {0}" -f (Format-Bytes $tamanhoCompactado)) -ForegroundColor Green
        }

        Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
        return $true

    } catch {
        Write-Fail "Erro inesperado: $_"
        Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
        return $false
    }
}

# ── Inicio ────────────────────────────────────────────────────────────────────
Write-Sep
Write-Host "Lendo servidores com Onda $NumeroOnda do arquivo Excel..."
Write-Sep

$scriptDir  = $PSScriptRoot
$servidores = @(& "$scriptDir\Extrair-IPs.ps1" -NumeroOnda $NumeroOnda -ArquivoExcel $ArquivoExcel)

if (-not $servidores -or $servidores.Count -eq 0) {
    Write-Error "Nenhum servidor com Onda $NumeroOnda encontrado!"
    Remove-Item $askpassDir -Recurse -Force -ErrorAction SilentlyContinue
    exit 1
}

Write-Host "Servidores encontrados: $($servidores.Count)"
foreach ($s in $servidores) { Write-Host "  - $s" }
Write-Sep

New-Item -ItemType Directory -Path $Destino -Force | Out-Null

$script:TotalEconomiaBanda = 0
$script:TotalBaixado       = 0
$script:TotalOriginal      = 0
$totalServidores    = 0
$totalLinux         = 0
$totalWindows       = 0
$totalIndisponivel  = 0
$totalIndeterminado = 0
$totalSucesso       = 0
$totalFalha         = 0

# ── Loop principal ────────────────────────────────────────────────────────────
foreach ($servidor in $servidores) {
    $totalServidores++
    $nomeServidor = $servidor -replace '\.', '_'

    Write-Host "Verificando $servidor..."
    $tipo = Get-ServidorTipo -Servidor $servidor

    switch ($tipo) {
        "linux" {
            $totalLinux++
            Write-Ok "Servidor Linux detectado"
            Write-Host "  Copiando com compressao..."
            if (Copiar-ComCompressao -Servidor $servidor -DestinoLocal $Destino -NomeServidor $nomeServidor) {
                Write-Ok "Sucesso na copia com compressao"
                $totalSucesso++
            } else {
                Write-Fail "Falha na copia"
                $totalFalha++
            }
        }
        "windows" {
            $totalWindows++
            Write-Fail "Servidor Windows detectado (pulado - nao suporta SSH)"
        }
        "indisponivel" {
            $totalIndisponivel++
            Write-Fail "Servidor indisponivel (sem resposta ao ping)"
        }
        "indeterminado" {
            $totalIndeterminado++
            Write-Warn "TTL fora das faixas esperadas - tentando conexao SSH..."
            if (Copiar-ComCompressao -Servidor $servidor -DestinoLocal $Destino -NomeServidor $nomeServidor) {
                Write-Ok "Sucesso na copia com compressao (possivelmente Linux)"
                $totalSucesso++
            } else {
                Write-Fail "Falha na copia"
                $totalFalha++
            }
        }
        default {
            Write-Fail "Nao foi possivel determinar o sistema operacional"
        }
    }

    Write-Dash
}

# ── Limpar askpass temporario ─────────────────────────────────────────────────
Remove-Item $askpassDir -Recurse -Force -ErrorAction SilentlyContinue
$SenhaPlana = $null
[System.GC]::Collect()

# ── Relatorio final ───────────────────────────────────────────────────────────
Write-Sep
Write-Host "STATUS FINAL - RELATORIO DE PROCESSAMENTO"
Write-Sep
Write-Host "Total de servidores processados: $totalServidores"
Write-Host ""
Write-Host "Classificacao por tipo:"
Write-Ok   "Linux identificados: $totalLinux"
Write-Fail "Windows identificados: $totalWindows"
Write-Fail "Indisponiveis (sem ping): $totalIndisponivel"
Write-Warn "Indeterminados (TTL fora da faixa): $totalIndeterminado"
Write-Host ""
Write-Host "Resultado das copias (Linux + Indeterminados):"
Write-Ok   "Copias com sucesso: $totalSucesso"
Write-Fail "Copias com falha: $totalFalha"
Write-Host ""

$tentativas = $totalLinux + $totalIndeterminado
if ($tentativas -gt 0) {
    $taxa = [math]::Round($totalSucesso * 100 / $tentativas, 1)
    Write-Host "Taxa de sucesso: $taxa%"
}

Write-Host ""
Write-Host "Economia de banda com compressao:"
if ($script:TotalOriginal -gt 0) {
    $pct = [math]::Round($script:TotalBaixado * 100 / $script:TotalOriginal, 1)
    Write-Ok ("Total original (remoto):  {0}" -f (Format-Bytes $script:TotalOriginal))
    Write-Ok ("Total transferido:        {0} ({1}% do original)" -f (Format-Bytes $script:TotalBaixado), $pct)
    Write-Ok ("Total economizado:        {0}" -f (Format-Bytes $script:TotalEconomiaBanda))
} else {
    Write-Ok ("Total transferido: {0}" -f (Format-Bytes $script:TotalBaixado))
}
Write-Sep
Write-Host "Processamento concluido!" -ForegroundColor Green
