#!/usr/bin/awk -f

BEGIN {
    FS = OFS = ";"
}

{
    # VALIDA├ç├âO 1: Ignorar conex├Áes IPv6
    # Verificar se o IP local ou remoto cont├®m ":" (caracter├¡stico de IPv6)
    if($10 ~ /:/ || $12 ~ /:/) {
        next  # Pular esta linha, n├úo processar
    }
    
    # VALIDA├ç├âO 2: Ignorar conex├Áes onde IP Local ├® igual a IP Remoto
    if($10 == $12) {
        next  # Pular esta linha, n├úo processar
    }
    
    # A sa├¡da do v1 tem:
    # $10 = IP Local
    # $11 = Porta Local
    # $12 = IP Remoto
    # $13 = Porta Remota
    # $14 = Dire├º├úo (IN/OUT/ANALISAR)

    direcao = $14
    
    # Chave de agrupamento baseada na dire├º├úo
    if(direcao == "IN") {
        # IN: agrupar por IP Local + Porta Local (servi├ºo) + IP Remoto
        chave = $10 "|" $11 "|" $12 "|" direcao "|" $15
    } 
    else if(direcao == "OUT") {
        # OUT: agrupar por IP Local + IP Remoto + Porta Remota (servi├ºo)
        chave = $10 "|" $12 "|" $13 "|" direcao "|" $15
    }
    else {
        # ANALISAR: agrupar por tudo
        chave = $10 "|" $11 "|" $12 "|" $13 "|" direcao "|" $15
    }
    
    if(!(chave in dados)) {
        dados[chave] = $0
        contador[chave] = 1
    } else {
        contador[chave]++
    }
}

END {
    for(chave in dados) {
        split(dados[chave], campos, ";")
        
        data_hora = campos[1]
        hostname = campos[2]
        protocolo = campos[3]
        local = campos[4]
        remoto = campos[5]
        estado = campos[6]
        pid = campos[7]
        processo = campos[8]
        col_extra1 = campos[9]
        ip_local = campos[10]
        porta_local = campos[11]
        ip_remoto = campos[12]
        porta_remota = campos[13]
        direcao = campos[14]
        ocvs = campos[15]
        portas_formatadas = campos[16]
        
        qtd = contador[chave]
        
        # Criar nova coluna de portas formatadas com contagem
        if(direcao == "IN") {
            novas_portas = "L " porta_local " | R (" qtd " conexoes)"
        } else if(direcao == "OUT") {
            novas_portas = "L (" qtd " conexoes) | R " porta_remota
        } else {
            novas_portas = portas_formatadas " (+" qtd ")"
        }
        
        print data_hora, hostname, protocolo, local, remoto, estado, pid, processo,
              col_extra1, ip_local, porta_local, ip_remoto, porta_remota, direcao, ocvs, novas_portas, qtd
    }
}
