#!/bin/bash
# Script de coleta de netstat para servidores Linux
# Executar no servidor alvo — gera /tmp/netstat_$(hostname).txt
# Compatível com netstat em português (RHEL/CentOS pt_BR) e inglês

date="$(date '+%Y-%m-%d %H:%M:%S')"
host="$(hostname)"

netstat -nlpta 2>/dev/null | awk -v dt="$date" -v hn="$host" '
!/^Active/ && !/^Conex/ && !/^Proto/ && NF>0 {

    # Identificar colunas — netstat pode ter 6 ou 7 campos dependendo do estado
    proto = $1
    local_addr = $4
    remote_addr = $5
    state = $6
    pidprog = $7

    # Se estado está vazio (ex: UDP), ajustar
    if (pidprog == "" && state ~ /[0-9]/) {
        pidprog = state
        state = "-"
    }

    # Traduzir estados do português para inglês
    if (state == "ESTABELECIDA" || state == "ESTABELECIDO") state = "ESTABLISHED"
    else if (state == "OUÇA" || state == "OUÇA") state = "LISTEN"
    else if (state == "TEMPO_ESPERA" || state == "TIME_WAIT") state = "TIME_WAIT"
    else if (state == "ESPERA_FECHAR" || state == "CLOSE_WAIT") state = "CLOSE_WAIT"
    else if (state == "FECHADO" || state == "CLOSED") state = "CLOSED"
    else if (state == "SIN_ENVIADO" || state == "SYN_SENT") state = "SYN_SENT"
    else if (state == "SIN_RECEBIDO" || state == "SYN_RECV") state = "SYN_RECV"
    else if (state == "ÚLTIMO_ACK" || state == "LAST_ACK") state = "LAST_ACK"
    else if (state == "ESPERA_FIN1" || state == "FIN_WAIT1") state = "FIN_WAIT1"
    else if (state == "ESPERA_FIN2" || state == "FIN_WAIT2") state = "FIN_WAIT2"
    else if (state == "FECHANDO" || state == "CLOSING") state = "CLOSING"

    # Filtrar LISTEN, TIME_WAIT e loopback
    if (state == "LISTEN") next
    if (state == "TIME_WAIT") next
    if (local_addr ~ /^127\.0\.0\.1/) next

    # Extrair PID e processo
    split(pidprog, pp, "/")
    pid = pp[1]
    proc = pp[2]
    # Remover ":" e tudo depois no nome do processo
    gsub(/:.*$/, "", proc)

    print dt ";" hn ";TCP;" local_addr ";" remote_addr ";" state ";" pid ";" proc
}' >> /tmp/netstat_${host}.txt
