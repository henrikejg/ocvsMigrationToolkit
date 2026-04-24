#!/usr/bin/awk -f

# Pré-filtro: separa linhas com IP remoto privado vs público
# Uso: awk -f filtrar_ip.awk -v PRIVADO=saida_privado.txt -v PUBLICO=saida_publico.txt entrada.txt

BEGIN {
    FS = ";"
}

{
    # Remover CR (Windows line endings)
    gsub(/\r$/, "", $0)

    # Campo 5 = endereço remoto (IP:PORTA)
    split($5, remoto, ":")
    ip = remoto[1]

    split(ip, oct, ".")
    o1 = int(oct[1])
    o2 = int(oct[2])

    # RFC1918 + loopback + link-local
    privado = 0
    if (o1 == 10) privado = 1
    else if (o1 == 172 && o2 >= 16 && o2 <= 31) privado = 1
    else if (o1 == 192 && o2 == 168) privado = 1
    else if (o1 == 127) privado = 1
    else if (o1 == 169 && o2 == 254) privado = 1

    if (privado) {
        print $0 > PRIVADO
    } else {
        print $0 > PUBLICO
    }
}
