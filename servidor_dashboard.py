"""
Servidor HTTP simples que deixa o dashboard de divergências acessível por um
link fixo na rede local, enquanto este computador estiver ligado e este
script rodando.

Serve APENAS a pasta dashboard_site/ (nunca o projeto inteiro, pra não expor
o .env nem as planilhas de origem). O main.py atualiza o conteúdo dessa
pasta toda vez que a automação roda; este script só precisa ficar de pé.

Uso:
    python servidor_dashboard.py          (com console, pra ver os links na tela)
    pythonw servidor_dashboard.py         (sem console, roda "invisível" em segundo plano)

O link pra compartilhar com a equipe fica registrado em servidor_dashboard.log
toda vez que o servidor sobe (procure a linha "Na rede local:").
Rodando sem console (pythonw), pra parar use o Gerenciador de Tarefas e
finalize o processo "pythonw.exe". Rodando com console, Ctrl+C encerra.
"""

import http.server
import logging
import os
import socket
import socketserver
import sys

PORTA = int(os.getenv("SITE_PORTA", "8000"))
PASTA_SITE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dashboard_site")
LOG_ARQUIVO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "servidor_dashboard.log")

# Rodando com pythonw.exe (sem console) sys.stdout/stderr vêm como None; sem
# isso, o log de acesso do http.server tentaria escrever neles e quebraria.
TEM_CONSOLE = sys.stdout is not None
if not TEM_CONSOLE:
    sys.stdout = open(os.devnull, "w")
    sys.stderr = open(os.devnull, "w")

_handlers = [logging.FileHandler(LOG_ARQUIVO, encoding="utf-8")]
if TEM_CONSOLE:
    _handlers.append(logging.StreamHandler())

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(message)s", handlers=_handlers)
logger = logging.getLogger(__name__)


def descobrir_ip_local():
    """Obtém o IP local usado na rede (não o loopback), sem depender de internet."""
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("10.255.255.255", 1))
        ip = s.getsockname()[0]
    except Exception:
        ip = "127.0.0.1"
    finally:
        s.close()
    return ip


def main():
    os.makedirs(PASTA_SITE, exist_ok=True)
    if not os.path.exists(os.path.join(PASTA_SITE, "index.html")):
        with open(os.path.join(PASTA_SITE, "index.html"), "w", encoding="utf-8") as f:
            f.write("<h1>Dashboard ainda não gerado</h1><p>Rode main.py para gerar o conteúdo.</p>")

    handler = lambda *args, **kwargs: http.server.SimpleHTTPRequestHandler(
        *args, directory=PASTA_SITE, **kwargs
    )

    hostname = socket.gethostname()
    ip_local = descobrir_ip_local()

    logger.info("Servindo o dashboard de divergências.")
    logger.info(f"  Pasta servida: {PASTA_SITE}")
    logger.info(f"  Neste computador:  http://localhost:{PORTA}/")
    logger.info(f"  Na rede local:     http://{ip_local}:{PORTA}/")
    logger.info(f"  Pelo nome do PC:   http://{hostname}:{PORTA}/  (se a rede resolver o nome)")
    logger.info("  (o link só funciona pra outras máquinas depois de liberar a porta no Firewall do Windows)")
    logger.info("Pressione Ctrl+C para encerrar.")

    try:
        with socketserver.TCPServer(("0.0.0.0", PORTA), handler) as httpd:
            httpd.serve_forever()
    except KeyboardInterrupt:
        logger.info("Servidor encerrado.")
    except OSError as e:
        logger.error(f"Não consegui abrir a porta {PORTA}: {e}")
        raise


if __name__ == "__main__":
    main()
