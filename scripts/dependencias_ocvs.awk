#!/usr/bin/awk -f

BEGIN {
    FS = OFS = ";"

    # Lista de portas de servi├ºos conhecidos (servidores)
    # Para identificar conex├Áes de ENTRADA (quem est├í ouvindo na porta)
    split("22,23,25,53,80,88,110,111,135,139,143,389,443,445,993,995,8080,3128,3306,3389,5432,5671,1433,1521,27017,6379,11211,9200,9300,5044,5601,10050,8680,587",
          portas_servidor, ",")

    for(i in portas_servidor) {
        portas_servico[portas_servidor[i]] = 1
    }

    # Lista completa de todos os /24 permitidos (expandido dos /22)
    split("160,161,162,163,164,165,166,167,168,169,170,176,177,178,179,180,181,182,183,184",
          redes_temp, ",")

    for(i in redes_temp) {
        redes_permitidas[redes_temp[i]] = 1
    }
}

# Fun├º├úo para extrair IP e porta de uma string "IP:PORTA"
function extrair_ip_porta(campo,    ip_porta, ip, porta) {
    split(campo, ip_porta, ":")
    ip = ip_porta[1]
    porta = ip_porta[2]
    return ip "|" porta
}

# Fun├º├úo para determinar dire├º├úo da conex├úo (IN ou OUT)
function determinar_direcao(porta_local, porta_remota,    servico_local, servico_remoto) {
    # Verificar se a porta local ├® de servi├ºo conhecido
    if(porta_local in portas_servico) {
        return "IN"  # Conex├úo entrando no servidor local
    }

    # Verificar se a porta remota ├® de servi├ºo conhecido
    if(porta_remota in portas_servico) {
        return "OUT"  # Conex├úo saindo do servidor local
    }

    # Portas altas (>1024) em ambos os lados - n├úo d├í pra determinar
    return "ANALISAR"
}

{

    # Remover caracteres CR (\r) do final da linha (Windows line endings)
    gsub(/\r$/, "", $0)

    # Extrair informa├º├Áes da linha original
    data_hora = $1
    hostname = $2
    protocolo = $3
    local = $4
    remoto = $5
    estado = $6
    pid = $7
    processo = $8

    # Extrair IP e porta local
    split(local, local_ip_porta, ":")
    ip_local = local_ip_porta[1]
    porta_local = local_ip_porta[2]

    # Extrair IP e porta remota
    split(remoto, remoto_ip_porta, ":")
    ip_remoto = remoto_ip_porta[1]
    porta_remota = remoto_ip_porta[2]

    # Coluna 1 extra: em branco (para fun├º├úo Excel)
    col_extra1 = ""

    # Coluna 2: IP local sem porta
    col_ip_local = ip_local

    # Coluna 3: porta local
    col_porta_local = porta_local

    # Coluna 4: IP remoto
    col_ip_remoto = ip_remoto

    # Coluna 5: porta remota
    col_porta_remota = porta_remota

    # Coluna 6: dire├º├úo (IN/OUT/ANALISAR)
    direcao = determinar_direcao(porta_local, porta_remota)

    # Coluna 7: verificar se IP remoto est├í nos ranges permitidos
    split(ip_remoto, oct, ".")

    # Verificar se o terceiro octeto est├í nas redes permitidas
    # Tamb├®m validar se ├® IP 10.62.x.x
    if(oct[1] == 10 && oct[2] == 62 && (oct[3] in redes_permitidas)) {
        ocvs = "OCVS"
    } else {
        ocvs = "FORA DO OCVS"
    }

    # Coluna 8: Porta Local e Remota formatada (NOVA COLUNA)
    col_portas = "L " porta_local " | R " porta_remota

    # Imprimir linha original + todas as colunas extras (agora com 8 colunas extras)
    print $0, col_extra1, col_ip_local, col_porta_local, col_ip_remoto, col_porta_remota, direcao, ocvs, col_portas
}
