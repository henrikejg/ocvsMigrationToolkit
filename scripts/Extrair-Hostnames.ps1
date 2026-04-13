<#
.SYNOPSIS
    Extrai hostnames (coluna VM) de uma planilha Excel para uma onda especifica.
.DESCRIPTION
    Equivalente PowerShell de extrair_hostnames.py — aba vInfo
.PARAMETER NumeroOnda
    Numero da onda a filtrar (ex: 2 para "Onda 2")
.PARAMETER ArquivoExcel
    Caminho absoluto para o arquivo .xlsx
.EXAMPLE
    $hosts = .\Extrair-Hostnames.ps1 -NumeroOnda 2 -ArquivoExcel "C:\dados\planilha.xlsx"
#>
param(
    [Parameter(Mandatory)][string]$NumeroOnda,
    [Parameter(Mandatory)][string]$ArquivoExcel
)

if (-not (Test-Path $ArquivoExcel)) {
    Write-Error "Arquivo Excel nao encontrado: $ArquivoExcel"
    exit 1
}

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Error "Modulo 'ImportExcel' nao encontrado. Instale com: Install-Module ImportExcel -Scope CurrentUser"
    exit 1
}

Import-Module ImportExcel -ErrorAction Stop *>$null

# Copiar para subpasta temp unica — robocopy le arquivos com lock do Excel
$tempSubDir = Join-Path $env:TEMP ("excel_" + [System.IO.Path]::GetRandomFileName().Replace('.',''))
New-Item -ItemType Directory -Path $tempSubDir -Force | Out-Null
$srcDir  = Split-Path $ArquivoExcel -Parent
$srcFile = Split-Path $ArquivoExcel -Leaf

$null = robocopy $srcDir $tempSubDir $srcFile /NFL /NDL /NJH /NJS 2>&1
$tempExcel = Join-Path $tempSubDir $srcFile

if (-not (Test-Path $tempExcel)) {
    Remove-Item $tempSubDir -Recurse -Force -ErrorAction SilentlyContinue
    Write-Error "Erro ao copiar Excel (robocopy falhou)"
    exit 1
}

try {
    $dados = Import-Excel -Path $tempExcel -NoHeader -WorksheetName "vInfo" -ErrorAction Stop
} catch {
    Write-Error "Erro ao abrir Excel: $_"
    exit 1
} finally {
    Remove-Item $tempSubDir -Recurse -Force -ErrorAction SilentlyContinue
}

if (-not $dados -or $dados.Count -lt 2) {
    Write-Error "Planilha vazia ou sem dados."
    exit 1
}

# Identificar colunas pelo cabecalho (linha 0)
$cabecalho = $dados[0]
$colVM = $null; $colIP = $null; $colOnda = $null

foreach ($prop in $cabecalho.PSObject.Properties.Name) {
    $val = [string]$cabecalho.$prop
    if     ($val -match "^VM$")                        { $colVM   = $prop }
    elseif ($val -match "IP|Address")                  { $colIP   = $prop }
    elseif ($val -match "^ONDA$")                      { $colOnda = $prop }
}

if (-not $colVM -or -not $colOnda) {
    Write-Error "Nao foi possivel identificar colunas VM ou ONDA. ColVM=$colVM ColOnda=$colOnda"
    exit 1
}

# Filtrar linhas pela onda
$hostnames = for ($i = 1; $i -lt $dados.Count; $i++) {
    $hostname = [string]$dados[$i].$colVM
    $ondaVal  = [string]$dados[$i].$colOnda

    if ($ondaVal -match "Onda $NumeroOnda" -and
        $hostname -and $hostname -ne "A definir" -and $hostname -ne "None") {
        $hostname
    }
}

if (-not $hostnames -or @($hostnames).Count -eq 0) {
    Write-Warning "Nenhum servidor com Onda $NumeroOnda encontrado."
    exit 1
}

# Retornar apenas os hostnames — sem nenhum outro output
$hostnames
