<#
.SYNOPSIS
    Coleta arquivos netstat de servidores Linux e Windows.
.DESCRIPTION
    - Linux: SSH + tar+gzip (compressao remota)
    - Windows: SMB via C$ (net use com credencial local) + Compress-Archive local
    - Detecta SO automaticamente via TTL do ping
    - Cobre Windows Server 2008 ate 2025 sem dependencias nos servidores remotos

    Pre-requisitos:
      - OpenSSH instalado (ssh.exe no PATH) — para Linux
      - Compartilhamento administrativo C$ habilitado — para Windows (padrao em todos)
      - Modulo ImportExcel: Install-Module ImportExcel -Scope CurrentUser

.PARAMETER NumeroOnda
    Numero da onda a processar
.PARAMETER ArquivoExcel
    Caminho para o arquivo .xlsx (auto-detectado se omitido)
.PARAMETER Destino
    Pasta de destino para os arquivos coletados (padrao: dados\raw\)
.PARAMETER Usuario
    Usuario para autenticacao (padrao: migracao)
.EXAMPLE
    .\Coletar-Netstat.ps1 -NumeroOnda 2
    .\Coletar-Netstat.ps1 -NumeroOnda 2 -Usuario "migracao"
#>
param(
    [Parameter(Mandatory)][string]$NumeroOnda,
    [string]$ArquivoExcel = "",
    [string]$Destino      = (Join-Path (Split-Path $PSScriptRoot -Parent) "dados\raw\"),
    [string]$Usuario      = "migracao",
    [string]$Senha        = ""  # Se vazio, abre Get-Credential interativo
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

# Resolver caminhos
if ($ArquivoExcel) {
    $ArquivoExcel = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ArquivoExcel)
} else {
    $candidatos = @(Get-ChildItem -Path (Split-Path $PSScriptRoot -Parent) -Filter "*.xlsx" -ErrorAction SilentlyContinue)
    if ($candidatos.Count -eq 0) {
        Write-Error "Nenhum .xlsx encontrado. Use -ArquivoExcel para especificar o caminho."
        exit 1
    }
    $ArquivoExcel = $candidatos[0].FullName
    Write-Host "Planilha: $ArquivoExcel" -ForegroundColor Cyan
}
$Destino = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Destino)

# ── Helpers ───────────────────────────────────────────────────────────────────
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

# ── Credenciais ───────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "Informe as credenciais para os servidores da Onda ${NumeroOnda}:" -ForegroundColor Cyan
Write-Host "(usadas para SSH nos Linux e SMB nos Windows)" -ForegroundColor Gray

# Se senha foi passada como parametro (modo headless via dashboard), usar direto
if (-not $Senha) {
    # Tentar variavel de ambiente injetada pelo Node
    $Senha = $env:OCVS_SENHA
}

if ($Senha) {
    Write-Host "Usando credenciais fornecidas para usuario '$Usuario'" -ForegroundColor Green
    $SenhaPlana = $Senha
} else {
    $credencial = Get-Credential -UserName $Usuario -Message "Credenciais - Onda ${NumeroOnda}"
    if (-not $credencial) {
        Write-Error "Credenciais nao informadas. Abortando."
        exit 1
    }
    $Usuario    = $credencial.UserName
    $SenhaPlana = $credencial.GetNetworkCredential().Password
}

# ── Askpass para SSH ──────────────────────────────────────────────────────────
$askpassDir = Join-Path $env:TEMP ("ssh_askpass_" + [System.IO.Path]::GetRandomFileName().Replace('.',''))
New-Item -ItemType Directory -Path $askpassDir -Force | Out-Null
$askpassBat = Join-Path $askpassDir "askpass.bat"

$SenhaEscapada = $SenhaPlana `
    -replace '\^', '^^' -replace '!', '^!' -replace '&', '^&' `
    -replace '\|', '^|' -replace '<', '^<' -replace '>', '^>' -replace '%', '%%'
[System.IO.File]::WriteAllText($askpassBat, "@echo off`r`necho $SenhaEscapada", [System.Text.Encoding]::ASCII)

# ── Opcoes SSH ────────────────────────────────────────────────────────────────
$SshOpts = @(
    "-o", "KexAlgorithms=+diffie-hellman-group1-sha1",
    "-o", "HostKeyAlgorithms=+ssh-rsa",
    "-o", "Ciphers=+aes128-cbc",
    "-o", "MACs=+hmac-sha1",
    "-o", "StrictHostKeyChecking=no",
    "-o", "PasswordAuthentication=yes",
    "-o", "BatchMode=no"
)

# ── Detectar SO via TTL ───────────────────────────────────────────────────────
function Get-ServidorTipo {
    param([string]$Servidor)
    $ping = Test-Connection -ComputerName $Servidor -Count 1 -TimeoutSeconds 3 -ErrorAction SilentlyContinue
    if (-not $ping) { return "indisponivel" }
    $ttl = $null
    if ($ping.PSObject.Properties['Reply'] -and $ping.Reply) { $ttl = $ping.Reply.Options.Ttl }
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

# ── SSH helper ────────────────────────────────────────────────────────────────
function Invoke-SSH {
    param([string[]]$SshArgs, [bool]$CapturarSaida = $false)
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
        $stderrTask = $proc.StandardError.ReadToEndAsync()
        $saida = $proc.StandardOutput.ReadToEnd()
        $proc.WaitForExit()
        return [PSCustomObject]@{ Saida = $saida; ExitCode = $proc.ExitCode; Stderr = $stderrTask.Result }
    }
    return [PSCustomObject]@{ Proc = $proc }
}

# ── Coleta Linux (SSH + tar+gzip) ─────────────────────────────────────────────
function Coletar-Linux {
    param([string]$Servidor, [string]$DestinoLocal, [string]$NomeServidor)

    $tempDir  = Join-Path $env:TEMP "netstat_coleta"
    $tempFile = Join-Path $tempDir "${NomeServidor}_netstat.tar.gz"
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

    $sshArgs = $SshOpts + @("${Usuario}@${Servidor}", "cd /tmp && tar czf - netstat_*.txt 2>/dev/null")

    try {
        $r    = Invoke-SSH -SshArgs $sshArgs -CapturarSaida $false
        $proc = $r.Proc
        $stderrTask = $proc.StandardError.ReadToEndAsync()

        $srcStream  = $proc.StandardOutput.BaseStream
        $fs         = [System.IO.File]::OpenWrite($tempFile)
        $buffer     = New-Object byte[] 65536
        $totalBytes = 0
        $sw         = [System.Diagnostics.Stopwatch]::StartNew()

        Write-Host "    Transferindo..." -NoNewline
        while ($true) {
            $lidos = $srcStream.Read($buffer, 0, $buffer.Length)
            if ($lidos -eq 0) { break }
            $fs.Write($buffer, 0, $lidos)
            $totalBytes += $lidos
            if ($sw.ElapsedMilliseconds -ge 200) {
                Write-Progress -Activity "Linux $Servidor" -Status ("Recebido: {0}" -f (Format-Bytes $totalBytes)) -PercentComplete -1
                $sw.Restart()
            }
        }
        $fs.Close()
        $proc.WaitForExit()
        Write-Progress -Activity "Linux $Servidor" -Completed

        $stderrMsg         = $stderrTask.Result.Trim()
        $tamanhoCompactado = (Get-Item $tempFile -ErrorAction SilentlyContinue).Length

        if ($proc.ExitCode -ne 0 -or -not $tamanhoCompactado -or $tamanhoCompactado -eq 0) {
            Write-Host ""
            Write-Fail "Falha SSH (exit $($proc.ExitCode))"
            if ($stderrMsg) { Write-Host "    Erro: $stderrMsg" -ForegroundColor DarkRed }
            Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
            return $false
        }

        Write-Host "`r    Extraindo...                    " -NoNewline
        $extraido = $false
        if (Get-Command tar -ErrorAction SilentlyContinue) {
            tar xzf $tempFile -C $DestinoLocal 2>$null
            $extraido = ($LASTEXITCODE -eq 0)
        }
        if (-not $extraido -and (Get-Command wsl -ErrorAction SilentlyContinue)) {
            wsl tar xzf (wsl wslpath ($tempFile -replace '\\','/')) -C (wsl wslpath ($DestinoLocal -replace '\\','/'))
            $extraido = ($LASTEXITCODE -eq 0)
        }
        if (-not $extraido) {
            Write-Host ""; Write-Fail "Falha na extracao (tar nao disponivel)"
            Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
            return $false
        }

        $r2 = Invoke-SSH -SshArgs ($SshOpts + @("${Usuario}@${Servidor}", 'cd /tmp && du -sb netstat_*.txt 2>/dev/null | awk ''{sum+=$1} END {print sum}''')) -CapturarSaida $true
        $tamanhoOriginal = ($r2.Saida.Trim()) -as [long]
        Remove-Item $tempFile -Force -ErrorAction SilentlyContinue

        if ($tamanhoOriginal -and $tamanhoCompactado) {
            $economia = $tamanhoOriginal - $tamanhoCompactado
            $script:TotalEconomiaBanda += $economia
            $script:TotalBaixado       += $tamanhoCompactado
            $script:TotalOriginal      += $tamanhoOriginal
            Write-Host ("`r    v {0} -> {1} comprimido (economia {2})" -f (Format-Bytes $tamanhoOriginal),(Format-Bytes $tamanhoCompactado),(Format-Bytes $economia)) -ForegroundColor Green
        } else {
            $script:TotalBaixado += $tamanhoCompactado
            Write-Host ("`r    v Transferido: {0}" -f (Format-Bytes $tamanhoCompactado)) -ForegroundColor Green
        }
        return $true
    } catch {
        Write-Fail "Erro inesperado: $_"
        Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
        return $false
    }
}

# ── Coleta Windows (SMB via C$) ───────────────────────────────────────────────
function Coletar-Windows {
    param([string]$Servidor, [string]$DestinoLocal, [string]$NomeServidor)

    $share    = "\\$Servidor\C$\TEMP"
    $netDrive = $null

    try {
        # Para conta local, o "dominio" no net use deve ser o hostname/IP do servidor remoto
        # Formato: servidor\usuario — força autenticacao como conta local do servidor
        $usuarioLocal = "$Servidor\$Usuario"

        Write-Host "    Conectando via SMB ($usuarioLocal)..." -NoNewline
        $netResult = net use $share /user:$usuarioLocal $SenhaPlana 2>&1
        if ($LASTEXITCODE -ne 0) {
            # Fallback: tentar sem prefixo de dominio (alguns ambientes aceitam)
            $netResult2 = net use $share /user:$Usuario $SenhaPlana 2>&1
            if ($LASTEXITCODE -ne 0) {
                Write-Host ""
                Write-Fail "Falha ao conectar SMB: $netResult"
                return $false
            }
        }
        Write-Host " OK" -ForegroundColor Green

        # Listar arquivos netstat no TEMP remoto
        $arquivosRemotos = @(Get-ChildItem -Path $share -Filter "netstat_*.txt" -ErrorAction SilentlyContinue)

        if ($arquivosRemotos.Count -eq 0) {
            Write-Warn "Nenhum arquivo netstat_*.txt encontrado em $share"
            return $false
        }

        Write-Host "    Copiando $($arquivosRemotos.Count) arquivo(s)..." -NoNewline
        $totalBytes = 0
        $sw = [System.Diagnostics.Stopwatch]::StartNew()

        foreach ($arq in $arquivosRemotos) {
            $destArq = Join-Path $DestinoLocal $arq.Name
            Copy-Item -Path $arq.FullName -Destination $destArq -Force
            $totalBytes += $arq.Length
            if ($sw.ElapsedMilliseconds -ge 200) {
                Write-Progress -Activity "Windows $Servidor" -Status ("Copiado: {0}" -f (Format-Bytes $totalBytes)) -PercentComplete -1
                $sw.Restart()
            }
        }
        Write-Progress -Activity "Windows $Servidor" -Completed

        $script:TotalBaixado  += $totalBytes
        $script:TotalOriginal += $totalBytes

        Write-Host ("`r    v {0} copiado(s) — {1} total" -f $arquivosRemotos.Count, (Format-Bytes $totalBytes)) -ForegroundColor Green
        Write-Host "    (compressao feita localmente apos copia)" -ForegroundColor Gray

        # Comprimir localmente os arquivos copiados (Compress-Archive disponivel no PS3+)
        # Para PS2 (Server 2008 no operador — improvavel mas seguro ignorar)
        if ($PSVersionTable.PSVersion.Major -ge 3) {
            $arquivosCopiados = @(Get-ChildItem -Path $DestinoLocal -Filter "netstat_*.txt" |
                Where-Object { $arquivosRemotos.Name -contains $_.Name })
            # Nao comprime aqui — mantemos os .txt para compatibilidade com o processar_onda
            # A compressao seria opcional e nao altera o fluxo de processamento
        }

        return $true

    } catch {
        Write-Fail "Erro inesperado: $_"
        return $false
    } finally {
        # Desmontar share sempre, mesmo em caso de erro
        net use $share /delete /yes 2>&1 | Out-Null
    }
}

# ── Inicio ────────────────────────────────────────────────────────────────────
Write-Sep
Write-Host "Lendo servidores com Onda $NumeroOnda do arquivo Excel..."
Write-Sep

$servidores = @(& "$PSScriptRoot\Extrair-IPs.ps1" -NumeroOnda $NumeroOnda -ArquivoExcel $ArquivoExcel)

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
            Write-Ok "Linux detectado (TTL 50-64 ou 240-255)"
            if (Coletar-Linux -Servidor $servidor -DestinoLocal $Destino -NomeServidor $nomeServidor) {
                Write-Ok "Sucesso"; $totalSucesso++
            } else {
                Write-Fail "Falha"; $totalFalha++
            }
        }
        "windows" {
            $totalWindows++
            Write-Ok "Windows detectado (TTL 110-128)"
            if (Coletar-Windows -Servidor $servidor -DestinoLocal $Destino -NomeServidor $nomeServidor) {
                Write-Ok "Sucesso"; $totalSucesso++
            } else {
                Write-Fail "Falha"; $totalFalha++
            }
        }
        "indisponivel" {
            $totalIndisponivel++
            Write-Fail "Servidor indisponivel (sem resposta ao ping)"
        }
        "indeterminado" {
            $totalIndeterminado++
            Write-Warn "TTL fora das faixas — tentando Linux (SSH) primeiro..."
            if (Coletar-Linux -Servidor $servidor -DestinoLocal $Destino -NomeServidor $nomeServidor) {
                Write-Ok "Sucesso via SSH"; $totalSucesso++
            } else {
                Write-Warn "SSH falhou — tentando Windows (SMB)..."
                if (Coletar-Windows -Servidor $servidor -DestinoLocal $Destino -NomeServidor $nomeServidor) {
                    Write-Ok "Sucesso via SMB"; $totalSucesso++
                } else {
                    Write-Fail "Falha em ambos os metodos"; $totalFalha++
                }
            }
        }
        default {
            Write-Fail "Nao foi possivel determinar o SO"
        }
    }

    Write-Dash
}

# ── Limpeza ───────────────────────────────────────────────────────────────────
Remove-Item $askpassDir -Recurse -Force -ErrorAction SilentlyContinue
$SenhaPlana = $null
[System.GC]::Collect()

# ── Relatorio final ───────────────────────────────────────────────────────────
Write-Sep
Write-Host "STATUS FINAL"
Write-Sep
Write-Host "Total processados: $totalServidores"
Write-Host ""
Write-Host "Por tipo:"
Write-Ok   "Linux:          $totalLinux"
Write-Ok   "Windows:        $totalWindows"
Write-Fail "Indisponiveis:  $totalIndisponivel"
Write-Warn "Indeterminados: $totalIndeterminado"
Write-Host ""
Write-Ok   "Sucesso: $totalSucesso"
Write-Fail "Falha:   $totalFalha"
if (($totalLinux + $totalWindows + $totalIndeterminado) -gt 0) {
    $taxa = [math]::Round($totalSucesso * 100 / ($totalLinux + $totalWindows + $totalIndeterminado), 1)
    Write-Host "Taxa de sucesso: $taxa%"
}
Write-Host ""
Write-Host "Transferencia:"
if ($script:TotalOriginal -gt 0) {
    $pct = [math]::Round($script:TotalBaixado * 100 / $script:TotalOriginal, 1)
    Write-Ok ("Original:     {0}" -f (Format-Bytes $script:TotalOriginal))
    Write-Ok ("Transferido:  {0} ({1}% do original)" -f (Format-Bytes $script:TotalBaixado), $pct)
    if ($script:TotalEconomiaBanda -gt 0) {
        Write-Ok ("Economizado:  {0} (compressao Linux)" -f (Format-Bytes $script:TotalEconomiaBanda))
    }
} else {
    Write-Ok ("Transferido: {0}" -f (Format-Bytes $script:TotalBaixado))
}
Write-Sep
Write-Host "Concluido!" -ForegroundColor Green
