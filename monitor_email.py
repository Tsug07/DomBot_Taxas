"""
DomBot - Monitor de E-mail
==========================
Monitora a caixa de entrada aguardando respostas do sindicato.
Notifica via Discord e salva anexos automaticamente.

Autor: Hugo L. Almeida
"""

import imaplib
import email
import email.header
import os
import time
import logging
import requests
import json
import threading
from datetime import datetime
from email import policy
from pathlib import Path

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

try:
    import pystray
    from PIL import Image, ImageDraw
    TRAY_DISPONIVEL = True
except ImportError:
    TRAY_DISPONIVEL = False

# ── Configurações ─────────────────────────────────────────────────────────────

IMAP_SERVER       = "imap.gmail.com"
IMAP_PORT         = 993
GMAIL_USER        = os.getenv("GMAIL_USER", "")
GMAIL_PASSWORD    = os.getenv("GMAIL_APP_PASSWORD", "")
REMETENTE_ALVO    = os.getenv("EMAIL_DESTINATARIO", "sindembar@uol.com.br")
DISCORD_WEBHOOK   = os.getenv("DISCORD_WEBHOOK_URL", "")
GMAIL_LABEL       = os.getenv("GMAIL_LABEL", "SINDEMBAR")
INTERVALO_SEG     = 15 * 60  # 15 minutos

# Raiz onde os anexos recebidos serão salvos (subpastas por competência criadas automaticamente)
ANEXOS_RAIZ       = Path(os.getenv("ANEXOS_RECEBIDOS_DIR", r"Z:\Pessoal\2026\TAXAS_BARES"))
BASE_DIR          = Path(__file__).parent
LOGS_DIR          = BASE_DIR / "logs"
ESTADO_FILE       = BASE_DIR / ".monitor_ids_vistos.json"

LOGS_DIR.mkdir(exist_ok=True)

# ── Logging ───────────────────────────────────────────────────────────────────

def _setup_logger():
    logger = logging.getLogger("MonitorEmail")
    logger.setLevel(logging.INFO)
    if logger.handlers:
        return logger

    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%Y-%m-%d %H:%M:%S")

    ch = logging.StreamHandler()
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    log_file = LOGS_DIR / f"monitor_{datetime.now().strftime('%Y-%m')}.log"
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    return logger

log = _setup_logger()

# ── Persistência de IDs já processados ────────────────────────────────────────

def _carregar_ids_vistos() -> set:
    try:
        if ESTADO_FILE.exists():
            with open(ESTADO_FILE, "r", encoding="utf-8") as f:
                return set(json.load(f))
    except Exception:
        pass
    return set()

def _salvar_ids_vistos(ids: set):
    try:
        with open(ESTADO_FILE, "w", encoding="utf-8") as f:
            json.dump(list(ids), f)
    except Exception as e:
        log.warning(f"Não foi possível salvar estado: {e}")

# ── Decodificação de cabeçalhos ───────────────────────────────────────────────

def _decodificar(valor: str) -> str:
    partes = email.header.decode_header(valor or "")
    resultado = []
    for parte, charset in partes:
        if isinstance(parte, bytes):
            resultado.append(parte.decode(charset or "utf-8", errors="replace"))
        else:
            resultado.append(parte)
    return " ".join(resultado)

# ── Notificação Discord ───────────────────────────────────────────────────────

def _notificar_discord(assunto: str, remetente: str, anexos_salvos: list[Path], corpo: str = ""):
    if not DISCORD_WEBHOOK:
        log.warning("DISCORD_WEBHOOK_URL não configurado — notificação ignorada")
        return
    try:
        if anexos_salvos:
            lista_anexos = "\n".join(f"  📎 `{p}`" for p in anexos_salvos)
        else:
            lista_anexos = "  _(sem anexos)_"
        corpo_trecho = ""
        if corpo:
            trecho = corpo.strip()[:800]
            if len(corpo.strip()) > 800:
                trecho += "…"
            corpo_trecho = f"\n💬 **Mensagem:**\n```\n{trecho}\n```"
        mensagem = (
            f"📬 **Resposta recebida do sindicato!**\n\n"
            f"👤 **De:** {remetente}\n"
            f"📋 **Assunto:** {assunto}\n"
            f"📂 **Anexos salvos:**\n{lista_anexos}"
            f"{corpo_trecho}\n\n"
            f"<@&1299044385899548752>"
        )
        requests.post(DISCORD_WEBHOOK, json={"content": mensagem}, timeout=10)
        log.info("Notificação enviada ao Discord")
    except Exception as e:
        log.warning(f"Erro ao notificar Discord: {e}")

# ── Extração de competência do assunto ───────────────────────────────────────

def _extrair_competencia(assunto: str) -> str:
    """Extrai MM-AAAA ou MM/AAAA do assunto. Ex: 'RE: Taxa Assistencial 04/2026' → '04-2026'."""
    import re
    m = re.search(r'(\d{2})[/-](\d{4})', assunto)
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    return datetime.now().strftime("%m-%Y")

# ── Download de anexos ────────────────────────────────────────────────────────

def _baixar_anexos(msg: email.message.Message, uid: str, assunto: str) -> list[Path]:
    competencia = _extrair_competencia(assunto)
    pasta_destino = ANEXOS_RAIZ / f"recebidos_{competencia}"
    try:
        pasta_destino.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        log.error(f"Não foi possível criar pasta '{pasta_destino}': {e}")
        return []

    data_recebimento = datetime.now().strftime("%Y%m%d")
    salvos: list[Path] = []

    for parte in msg.walk():
        if parte.get_content_maintype() == "multipart":
            continue
        if parte.get("Content-Disposition") is None:
            continue
        nome_arquivo = parte.get_filename()
        if not nome_arquivo:
            continue
        nome_arquivo = _decodificar(nome_arquivo)

        destino = pasta_destino / nome_arquivo
        if destino.exists():
            stem = Path(nome_arquivo).stem
            suffix = Path(nome_arquivo).suffix
            destino = pasta_destino / f"{stem}_{data_recebimento}_{uid}{suffix}"

        try:
            destino.write_bytes(parte.get_payload(decode=True))
            salvos.append(destino)
            log.info(f"Anexo salvo: {destino}")
        except Exception as e:
            log.error(f"Erro ao salvar anexo '{nome_arquivo}': {e}")

    return salvos

# ── Verificação da caixa de entrada ──────────────────────────────────────────

def _extrair_corpo(msg: email.message.Message) -> str:
    """Extrai o texto simples do corpo do email, ignorando histórico de respostas."""
    corpo = ""
    for parte in msg.walk():
        if parte.get_content_type() == "text/plain" and parte.get_content_disposition() != "attachment":
            try:
                texto = parte.get_payload(decode=True).decode(parte.get_content_charset() or "utf-8", errors="replace")
                for marcador in ("\nDe:", "\nFrom:", "\n-----", "\n>", "\nEm ", "\nOn "):
                    idx = texto.find(marcador)
                    if idx > 0:
                        texto = texto[:idx]
                corpo = texto.strip()
                break
            except Exception:
                pass
    return corpo


def verificar_emails(ids_vistos: set) -> set:
    novos_ids = set()
    try:
        with imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT) as imap:
            imap.login(GMAIL_USER, GMAIL_PASSWORD.replace(" ", ""))

            # Seleciona a label do Gmail em vez da INBOX genérica
            status_sel, _ = imap.select(f'"{GMAIL_LABEL}"')
            if status_sel != "OK":
                log.error(f"Não foi possível selecionar a label '{GMAIL_LABEL}' — verifique o nome exato no Gmail")
                return novos_ids

            # Busca apenas emails não lidos recebidos deles (exclui os que eu enviei)
            criterio = f'(FROM "{REMETENTE_ALVO}" UNSEEN)'
            status, dados = imap.search(None, criterio)
            if status != "OK":
                log.warning("Falha ao buscar e-mails")
                return novos_ids

            uids = dados[0].split()
            if not uids:
                log.info(f"Nenhum e-mail novo de {REMETENTE_ALVO} na label '{GMAIL_LABEL}'")
                return novos_ids

            log.info(f"{len(uids)} e-mail(s) novo(s) de {REMETENTE_ALVO} na label '{GMAIL_LABEL}'")

            for uid_bytes in uids:
                uid = uid_bytes.decode()
                if uid in ids_vistos:
                    continue

                status, msg_data = imap.fetch(uid_bytes, "(RFC822)")
                if status != "OK":
                    continue

                raw = msg_data[0][1]
                msg = email.message_from_bytes(raw, policy=policy.default)

                assunto   = _decodificar(msg.get("Subject", "(sem assunto)"))
                remetente = _decodificar(msg.get("From", ""))

                log.info(f"Processando: de='{remetente}' | assunto='{assunto}'")

                anexos = _baixar_anexos(msg, uid, assunto)
                corpo  = _extrair_corpo(msg)

                # Marcar como lido
                imap.store(uid_bytes, "+FLAGS", "\\Seen")
                log.info(f"E-mail {uid} marcado como lido")

                _notificar_discord(assunto, remetente, anexos, corpo)

                novos_ids.add(uid)

    except imaplib.IMAP4.error as e:
        log.error(f"Erro IMAP: {e}")
    except Exception as e:
        log.error(f"Erro inesperado: {e}")

    return novos_ids

# ── Bandeja do sistema (system tray) ─────────────────────────────────────────

_estado_tray = {
    "status": "iniciando",   # "monitorando" | "erro" | "iniciando"
    "ultima_verificacao": "—",
    "icone": None,
}

def _criar_imagem_icone(cor: tuple) -> "Image.Image":
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.ellipse((4, 4, 60, 60), fill=cor)
    return img

_COR_MONITORANDO = (128, 0, 128, 255)   # roxo
_COR_ERRO        = (120, 120, 120, 255) # cinza

def _tooltip_tray() -> str:
    s = _estado_tray["status"].capitalize()
    u = _estado_tray["ultima_verificacao"]
    return f"DomBot Monitor\nStatus: {s}\nÚltima verificação: {u}"

def _atualizar_icone_tray():
    icone = _estado_tray["icone"]
    if icone is None:
        return
    cor = _COR_MONITORANDO if _estado_tray["status"] == "monitorando" else _COR_ERRO
    icone.icon = _criar_imagem_icone(cor)
    icone.title = _tooltip_tray()

def _menu_tray():
    status = _estado_tray["status"].capitalize()
    ultima = _estado_tray["ultima_verificacao"]
    return (
        pystray.MenuItem(f"Status: {status}", None, enabled=False),
        pystray.MenuItem(f"Última verificação: {ultima}", None, enabled=False),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Sair", lambda icon, _: icon.stop()),
    )

def _iniciar_tray():
    img_inicial = _criar_imagem_icone(_COR_ERRO)
    icone = pystray.Icon(
        "DomBot",
        img_inicial,
        "DomBot Monitor\nStatus: Iniciando",
        menu=pystray.Menu(_menu_tray),
    )
    _estado_tray["icone"] = icone
    icone.run()  # bloqueia até icon.stop()

# ── Loop principal ────────────────────────────────────────────────────────────

def _loop_monitor():
    if not GMAIL_USER or not GMAIL_PASSWORD:
        log.error("GMAIL_USER e GMAIL_APP_PASSWORD devem estar configurados no .env")
        if _estado_tray["icone"]:
            _estado_tray["icone"].stop()
        return

    log.info("=" * 60)
    log.info("Monitor de E-mail iniciado")
    log.info(f"Monitorando respostas de: {REMETENTE_ALVO}")
    log.info(f"Intervalo de verificação: {INTERVALO_SEG // 60} minutos")
    log.info(f"Anexos salvos em: {ANEXOS_RAIZ}")
    log.info("=" * 60)

    ids_vistos = _carregar_ids_vistos()

    while True:
        log.info("Verificando caixa de entrada...")
        _estado_tray["status"] = "monitorando"
        _estado_tray["ultima_verificacao"] = datetime.now().strftime("%d/%m/%Y %H:%M")
        _atualizar_icone_tray()

        try:
            novos = verificar_emails(ids_vistos)
            if novos:
                ids_vistos.update(novos)
                _salvar_ids_vistos(ids_vistos)
                log.info(f"{len(novos)} resposta(s) processada(s)")
        except Exception as e:
            log.error(f"Erro no loop: {e}")
            _estado_tray["status"] = "erro"
            _atualizar_icone_tray()

        log.info(f"Próxima verificação em {INTERVALO_SEG // 60} minutos")
        time.sleep(INTERVALO_SEG)


def main():
    if TRAY_DISPONIVEL:
        t = threading.Thread(target=_loop_monitor, daemon=True)
        t.start()
        _iniciar_tray()  # bloqueia aqui; encerra quando usuário clicar em Sair
    else:
        log.warning("pystray/Pillow não instalados — rodando sem bandeja")
        _loop_monitor()


if __name__ == "__main__":
    main()
