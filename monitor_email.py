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
from datetime import datetime
from email import policy
from pathlib import Path

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ── Configurações ─────────────────────────────────────────────────────────────

IMAP_SERVER       = "imap.gmail.com"
IMAP_PORT         = 993
GMAIL_USER        = os.getenv("GMAIL_USER", "")
GMAIL_PASSWORD    = os.getenv("GMAIL_APP_PASSWORD", "")
REMETENTE_ALVO    = os.getenv("EMAIL_DESTINATARIO", "sindembar@uol.com.br")
DISCORD_WEBHOOK   = os.getenv("DISCORD_WEBHOOK_URL", "")
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

def _notificar_discord(assunto: str, remetente: str, anexos: list[str]):
    if not DISCORD_WEBHOOK:
        log.warning("DISCORD_WEBHOOK_URL não configurado — notificação ignorada")
        return
    try:
        lista_anexos = "\n".join(f"  📎 `{a}`" for a in anexos) if anexos else "  _(sem anexos)_"
        mensagem = (
            f"📬 **Resposta recebida do sindicato!**\n\n"
            f"👤 **De:** {remetente}\n"
            f"📋 **Assunto:** {assunto}\n"
            f"📂 **Anexos baixados:**\n{lista_anexos}\n\n"
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

def _baixar_anexos(msg: email.message.Message, uid: str, assunto: str) -> list[str]:
    competencia = _extrair_competencia(assunto)
    pasta_destino = ANEXOS_RAIZ / f"recebidos_{competencia}"
    try:
        pasta_destino.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        log.error(f"Não foi possível criar pasta '{pasta_destino}': {e}")
        return []

    data_recebimento = datetime.now().strftime("%Y%m%d")
    salvos = []

    for parte in msg.walk():
        if parte.get_content_maintype() == "multipart":
            continue
        if parte.get("Content-Disposition") is None:
            continue
        nome_arquivo = parte.get_filename()
        if not nome_arquivo:
            continue
        nome_arquivo = _decodificar(nome_arquivo)

        # Mantém o nome original do anexo; prefixo só se houver colisão
        destino = pasta_destino / nome_arquivo
        if destino.exists():
            stem = Path(nome_arquivo).stem
            suffix = Path(nome_arquivo).suffix
            destino = pasta_destino / f"{stem}_{data_recebimento}_{uid}{suffix}"

        try:
            destino.write_bytes(parte.get_payload(decode=True))
            salvos.append(destino.name)
            log.info(f"Anexo salvo: {destino}")
        except Exception as e:
            log.error(f"Erro ao salvar anexo '{nome_arquivo}': {e}")

    return salvos

# ── Verificação da caixa de entrada ──────────────────────────────────────────

def verificar_emails(ids_vistos: set) -> set:
    novos_ids = set()
    try:
        with imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT) as imap:
            imap.login(GMAIL_USER, GMAIL_PASSWORD.replace(" ", ""))
            imap.select("INBOX")

            # Busca e-mails do remetente alvo não lidos, enviados diretamente para esta caixa
            criterio = f'(FROM "{REMETENTE_ALVO}" TO "{GMAIL_USER}" UNSEEN)'
            status, dados = imap.search(None, criterio)
            if status != "OK":
                log.warning("Falha ao buscar e-mails")
                return novos_ids

            uids = dados[0].split()
            if not uids:
                log.info(f"Nenhum e-mail novo de {REMETENTE_ALVO}")
                return novos_ids

            log.info(f"{len(uids)} e-mail(s) novo(s) encontrado(s)")

            for uid_bytes in uids:
                uid = uid_bytes.decode()
                if uid in ids_vistos:
                    continue

                status, msg_data = imap.fetch(uid_bytes, "(RFC822)")
                if status != "OK":
                    continue

                raw = msg_data[0][1]
                msg = email.message_from_bytes(raw, policy=policy.default)

                assunto  = _decodificar(msg.get("Subject", "(sem assunto)"))
                remetente = _decodificar(msg.get("From", ""))

                log.info(f"Processando: de='{remetente}' | assunto='{assunto}'")

                anexos = _baixar_anexos(msg, uid, assunto)

                # Marcar como lido
                imap.store(uid_bytes, "+FLAGS", "\\Seen")
                log.info(f"E-mail {uid} marcado como lido")

                _notificar_discord(assunto, remetente, anexos)

                novos_ids.add(uid)

    except imaplib.IMAP4.error as e:
        log.error(f"Erro IMAP: {e}")
    except Exception as e:
        log.error(f"Erro inesperado: {e}")

    return novos_ids

# ── Loop principal ────────────────────────────────────────────────────────────

def main():
    if not GMAIL_USER or not GMAIL_PASSWORD:
        log.error("GMAIL_USER e GMAIL_APP_PASSWORD devem estar configurados no .env")
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
        novos = verificar_emails(ids_vistos)
        if novos:
            ids_vistos.update(novos)
            _salvar_ids_vistos(ids_vistos)
            log.info(f"{len(novos)} resposta(s) processada(s)")
        log.info(f"Próxima verificação em {INTERVALO_SEG // 60} minutos")
        time.sleep(INTERVALO_SEG)


if __name__ == "__main__":
    main()
