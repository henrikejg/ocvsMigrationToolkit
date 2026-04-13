<#
.SYNOPSIS
    Analisa dependências entre servidores OCVS com threshold configurável.
.DESCRIPTION
    Equivalente PowerShell de analise.sh
    Lê um arquivo CSV de comunicações (separado por ';') com 17 colunas:
      Col  2: servidor_origem
      Col 12: ip_destino
      Col 15: grupo (OCVS ou FORA DO OCVS)
      Col 17: contador de conexões aglutinadas

    Modos de operação:
      - Com arquivo de servidores da onda: identifica dependências externas
        e aplica threshold mínimo de conexões
      - Sem arquivo de servidores: análise completa, exibe top 10 origem/destino

    Saídas geradas:
      - dependencias_ocvs_YYYYMMDD_HHmmss.txt  (relatório completo)
      - servidores_recomendados_onda.txt        (lista de servidores a incluir)

.PARAMETER ArquivoComunicacoes
    Arquivo CSV com as comunicações (obrigatório)
.PARAMETER OndaServidores
    Arquivo .txt com um servidor por linha (opcional)
.PARAMETER Threshold
    Número mínimo de conexões para considerar dependência (padrão: 10)
.EXAMPLE
    .\Analisar-Dependencias.ps1 -ArquivoComunicacoes comunicacoes.csv -OndaServidores onda2.txt -Threshold 10
    .\Analisar-Dependencias.ps1 -ArquivoComunicacoes comunicacoes.csv
#>
param(
    [Parameter(Mandatory)][string]$ArquivoComunicacoes,
    [string]$OndaServidores = "",
    [int]$Threshold = 10
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

# ── Helpers de cor ────────────────────────────────────────────────────────────
function Write-Red     { param($m) Write-Host $m -ForegroundColor Red     }
function Write-Green   { param($m) Write-Host $m -ForegroundColor Green   }
function Write-Yellow  { param($m) Write-Host $m -ForegroundColor Yellow  }
function Write-Blue    { param($m) Write-Host $m -ForegroundColor Cyan    }
function Write-Magenta { param($m) Write-Host $m -ForegroundColor Magenta }
function Write-Sep     { Write-Host ("=" * 41) -ForegroundColor Yellow    }

# ── Validações ────────────────────────────────────────────────────────────────
if (-not (Test-Path $ArquivoComunicacoes)) {
    Write-Red "Erro: Arquivo de comunicações '$ArquivoComunicacoes' não encontrado"
    exit 1
}

$modoOnda = $false
$servidoresOnda = @()

if ($OndaServidores -and (Test-Path $OndaServidores)) {
    $modoOnda = $true
    $servidoresOnda = Get-Content $OndaServidores |
        ForEach-Object { $_.Trim().ToUpper() -replace '\r','' } |
        Where-Object { $_ -ne '' }
}

# ── Cabeçalho ─────────────────────────────────────────────────────────────────
Write-Sep
Write-Host "Análise de Dependências OCVS" -ForegroundColor Yellow
Write-Sep
Write-Host "Arquivo de comunicações: $ArquivoComunicacoes"
if ($modoOnda) {
    Write-Host "Servidores na onda: $($servidoresOnda.Count)"
    Write-Host "Arquivo de onda: $OndaServidores"
    Write-Host "Threshold mínimo: $Threshold conexões (considerando contador da coluna 17)"
} else {
    Write-Host "Modo: Análise completa (todas comunicações OCVS)"
}
Write-Host "Data: $(Get-Date)"
Write-Sep
Write-Host ""

# ── Estruturas de dados ───────────────────────────────────────────────────────
$comunicacoesOrigem  = @{}
$comunicacoesDestino = @{}
$contagem            = @{}
$servidoresExternos  = @{}

$totalLinhas          = 0
$totalOcvs            = 0
$totalConexoesOcvs    = 0
$totalDepFora         = 0
$totalConexoesFora    = 0
$totalDepIgnoradas    = 0
$totalConexoesIgnoradas = 0

# ── Normalizar chave (remover [], (), espaços) ────────────────────────────────
function Normalize-Key { param($s) return ($s -replace '[\[\]\(\)\s]', '') }

# ── Processar arquivo linha por linha ─────────────────────────────────────────
Write-Blue "Analisando arquivo de comunicações..."
Write-Blue "Considerando coluna 15 (grupo) e coluna 17 (contador)"
Write-Host ""

$reader = [System.IO.StreamReader]::new($ArquivoComunicacoes, [System.Text.Encoding]::UTF8)

while ($null -ne ($linha = $reader.ReadLine())) {
    $totalLinhas++
    $cols = $linha -split ';'

    # Garantir que temos pelo menos 17 colunas
    if ($cols.Count -lt 17) { continue }

    # Índices base-0: col2=idx1, col12=idx11, col15=idx14, col17=idx16
    $servidorOrigem = $cols[1].Trim() -replace '\r',''
    $ipDestino      = $cols[11].Trim() -replace '\r',''
    $grupo          = $cols[14].Trim() -replace '\r',''
    $contadorRaw    = $cols[16].Trim() -replace '\r','' -replace '[^\d]',''

    $contador = if ($contadorRaw -match '^\d+$') { [int]$contadorRaw } else { 1 }

    if ($grupo -ne "OCVS") { continue }

    $totalOcvs++
    $totalConexoesOcvs += $contador

    $servOrigem  = Normalize-Key ($servidorOrigem.ToUpper())
    # ip_destino pode ter porta: "10.0.0.1:443" → pegar só o IP
    $ipLimpo     = ($ipDestino -split ':')[0] -replace '\r',''
    $servDestino = Normalize-Key $ipLimpo

    # Acumular comunicações
    if ($comunicacoesOrigem.ContainsKey($servOrigem))  { $comunicacoesOrigem[$servOrigem]  += $contador }
    else                                               { $comunicacoesOrigem[$servOrigem]   = $contador }

    if ($comunicacoesDestino.ContainsKey($servDestino)) { $comunicacoesDestino[$servDestino] += $contador }
    else                                                { $comunicacoesDestino[$servDestino]  = $contador }

    if (-not $modoOnda) { continue }

    # Verificar se origem está na onda
    $origemNaOnda  = $servidoresOnda -contains $servOrigem
    $destinoNaOnda = $servidoresOnda -contains $servDestino

    if ($origemNaOnda -and -not $destinoNaOnda) {
        $totalDepFora++
        $totalConexoesFora += $contador

        $chave = "$servOrigem -> $servDestino"
        if ($contagem.ContainsKey($chave))          { $contagem[$chave]           += $contador }
        else                                        { $contagem[$chave]            = $contador }

        if ($servidoresExternos.ContainsKey($servDestino)) { $servidoresExternos[$servDestino] += $contador }
        else                                               { $servidoresExternos[$servDestino]  = $contador }
    }
}

$reader.Close()

Write-Green "✓ Processamento concluído"
Write-Host ""

# ── Estatísticas gerais ───────────────────────────────────────────────────────
Write-Blue "ESTATÍSTICAS GERAIS:"
Write-Host "Total de linhas processadas: $totalLinhas"
Write-Host "Total de comunicações OCVS (linhas aglutinadas): $totalOcvs"
Write-Host "Total de conexões OCVS (soma dos contadores - coluna 17): $totalConexoesOcvs"
Write-Host "Total de servidores origem (OCVS): $($comunicacoesOrigem.Count)"
Write-Host "Total de servidores destino (OCVS): $($comunicacoesDestino.Count)"
Write-Host ""

# ── Modo onda: análise de dependências ───────────────────────────────────────
if ($modoOnda) {

    Write-Sep
    Write-Yellow "DEPENDÊNCIAS COM SERVIDORES FORA DA ONDA"
    Write-Yellow "(Threshold mínimo: $Threshold conexões - considerando contador da coluna 17)"
    Write-Sep

    # Filtrar pelo threshold
    $externosFiltrados = @{}
    $totalConexoesAcimaThreshold = 0

    foreach ($serv in $servidoresExternos.Keys) {
        $qtd = $servidoresExternos[$serv]
        if ($qtd -ge $Threshold) {
            $externosFiltrados[$serv] = $qtd
            $totalConexoesAcimaThreshold += $qtd
        } else {
            $totalDepIgnoradas++
            $totalConexoesIgnoradas += $qtd
        }
    }

    if ($externosFiltrados.Count -gt 0) {

        Write-Red "⚠️  ATENÇÃO: Foram encontradas dependências externas significativas!"
        Write-Host ""
        Write-Blue "Total de conexões externas analisadas: $totalConexoesFora"
        Write-Blue "Conexões acima do threshold (${Threshold}+): $totalConexoesAcimaThreshold"
        Write-Blue "Conexões ignoradas (abaixo de $Threshold): $totalConexoesIgnoradas"
        Write-Host ""

        Write-Magenta "SERVIDORES EXTERNOS ACIMA DO THRESHOLD (${Threshold}+ conexões):"
        Write-Host ("-" * 56)
        $externosFiltrados.GetEnumerator() |
            Sort-Object Value -Descending |
            ForEach-Object {
                Write-Host ("{0,6} conexões - {1}" -f $_.Value, $_.Key) -ForegroundColor Yellow
            }

        Write-Host ""
        Write-Magenta "PRINCIPAIS DEPENDÊNCIAS (acima do threshold):"
        Write-Host ("-" * 56)
        $contagem.GetEnumerator() |
            Where-Object {
                $dest = ($_.Key -split ' -> ')[1]
                $externosFiltrados.ContainsKey($dest)
            } |
            Sort-Object Value -Descending |
            ForEach-Object {
                Write-Host ("{0,6} conexões - {1}" -f $_.Value, $_.Key) -ForegroundColor Green
            }

        # ── Gerar relatório ───────────────────────────────────────────────────
        $timestamp    = Get-Date -Format "yyyyMMdd_HHmmss"
        $arquivoSaida = "dependencias_ocvs_${timestamp}.txt"
        $listaRecomendados = "servidores_recomendados_onda.txt"

        $sb = [System.Text.StringBuilder]::new()
        $sb.AppendLine("=========================================")          | Out-Null
        $sb.AppendLine("RELATÓRIO DE DEPENDÊNCIAS OCVS")                    | Out-Null
        $sb.AppendLine("Gerado em: $(Get-Date)")                            | Out-Null
        $sb.AppendLine("Threshold mínimo: $Threshold conexões (considerando contador da coluna 17)") | Out-Null
        $sb.AppendLine("=========================================")          | Out-Null
        $sb.AppendLine("")                                                   | Out-Null
        $sb.AppendLine("SERVIDORES NA ONDA:")                               | Out-Null
        $servidoresOnda | ForEach-Object { $sb.AppendLine($_) | Out-Null }
        $sb.AppendLine("")                                                   | Out-Null
        $sb.AppendLine("=========================================")          | Out-Null
        $sb.AppendLine("RESUMO ESTATÍSTICO")                                | Out-Null
        $sb.AppendLine("=========================================")          | Out-Null
        $sb.AppendLine("Total de comunicações OCVS (linhas aglutinadas): $totalOcvs")          | Out-Null
        $sb.AppendLine("Total de conexões OCVS (soma dos contadores - coluna 17): $totalConexoesOcvs") | Out-Null
        $sb.AppendLine("Total de dependências externas (linhas): $totalDepFora")               | Out-Null
        $sb.AppendLine("Total de conexões externas (soma dos contadores): $totalConexoesFora") | Out-Null
        $sb.AppendLine("Total de conexões ignoradas (abaixo de $Threshold): $totalConexoesIgnoradas") | Out-Null
        $sb.AppendLine("Total de servidores externos acima do threshold: $($externosFiltrados.Count)") | Out-Null
        $sb.AppendLine("")                                                   | Out-Null
        $sb.AppendLine("=========================================")          | Out-Null
        $sb.AppendLine("SERVIDORES EXTERNOS CRÍTICOS (acima do threshold)") | Out-Null
        $sb.AppendLine("=========================================")          | Out-Null
        $sb.AppendLine("")                                                   | Out-Null

        $externosFiltrados.GetEnumerator() |
            Sort-Object Value -Descending |
            ForEach-Object { $sb.AppendLine("$($_.Key) - $($_.Value) conexões") | Out-Null }

        $sb.AppendLine("")                                                   | Out-Null
        $sb.AppendLine("=========================================")          | Out-Null
        $sb.AppendLine("RECOMENDAÇÕES")                                     | Out-Null
        $sb.AppendLine("=========================================")          | Out-Null
        $sb.AppendLine("")                                                   | Out-Null
        $sb.AppendLine("Com base no threshold de $Threshold conexões (considerando o contador da coluna 17):") | Out-Null
        $sb.AppendLine("")                                                   | Out-Null

        $externosFiltrados.GetEnumerator() |
            Sort-Object Value -Descending |
            ForEach-Object { $sb.AppendLine("✓ INCLUIR NA ONDA: $($_.Key) ($($_.Value) conexões)") | Out-Null }

        $sb.AppendLine("")                                                   | Out-Null
        $sb.AppendLine("Servidores com menos de $Threshold conexões foram ignorados") | Out-Null
        $sb.AppendLine("para evitar crescimento excessivo da onda de migração.")      | Out-Null

        $sb.ToString() | Set-Content -Path $arquivoSaida -Encoding UTF8

        # Lista de recomendados
        $linhasRec = @("# Servidores recomendados para incluir na onda - $(Get-Date)")
        $linhasRec += "# Threshold: $Threshold conexões mínimas (considerando contador da coluna 17)"
        $linhasRec += "#"
        $externosFiltrados.GetEnumerator() |
            Sort-Object Value -Descending |
            ForEach-Object { $linhasRec += "$($_.Key) # $($_.Value) conexões" }
        $linhasRec | Set-Content -Path $listaRecomendados -Encoding UTF8

        Write-Host ""
        Write-Green "Relatório detalhado salvo em: $arquivoSaida"
        Write-Green "Lista de servidores recomendados: $listaRecomendados"

    } else {
        Write-Green "✓ Nenhuma dependência externa significativa encontrada!"
        Write-Host "   Todas as dependências externas têm menos de $Threshold conexões."
        Write-Host "   Pode prosseguir com a onda sem adicionar novos servidores."
    }

} else {
    # ── Modo análise completa ─────────────────────────────────────────────────
    Write-Blue "TOP 10 SERVIDORES ORIGEM (OCVS) por volume de conexões:"
    Write-Host ("-" * 56)
    $comunicacoesOrigem.GetEnumerator() |
        Sort-Object Value -Descending |
        Select-Object -First 10 |
        ForEach-Object { Write-Host ("{0,6} conexões - {1}" -f $_.Value, $_.Key) }

    Write-Host ""
    Write-Blue "TOP 10 SERVIDORES DESTINO (OCVS) por volume de conexões:"
    Write-Host ("-" * 56)
    $comunicacoesDestino.GetEnumerator() |
        Sort-Object Value -Descending |
        Select-Object -First 10 |
        ForEach-Object { Write-Host ("{0,6} conexões - {1}" -f $_.Value, $_.Key) }
}

Write-Host ""
Write-Sep
Write-Host "Análise concluída!" -ForegroundColor Green
Write-Sep
