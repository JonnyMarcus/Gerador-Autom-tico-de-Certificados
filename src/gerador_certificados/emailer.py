"""Envio de certificados individuais por SMTP com validacao de destinatarios."""

from __future__ import annotations

from collections.abc import Sequence
from email.message import EmailMessage
from email.utils import parseaddr
import re
import smtplib
import ssl

_EMAIL_RE = re.compile(r"^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?)+$")


def validar_email(valor: str) -> bool:
    """Valida enderecos simples e rejeita display names ou quebras de cabecalho."""
    endereco = valor.strip()
    if not endereco or any(caractere in endereco for caractere in "\r\n"):
        return False
    nome, parseado = parseaddr(endereco)
    return not nome and parseado == endereco and bool(_EMAIL_RE.fullmatch(endereco))


def enviar_lote_smtp(config: dict[str, object], mensagens: Sequence[dict[str, object]]) -> list[dict[str, str | bool]]:
    """Envia uma mensagem individual por pessoa; erros de destinatario nao param o lote."""
    host = str(config["host"])
    porta = int(config.get("port", 587))
    remetente = str(config["from_email"])
    if not validar_email(remetente):
        raise ValueError("O remetente SMTP configurado nao e um endereco de e-mail valido.")

    usar_ssl = bool(config.get("use_ssl", False))
    usar_starttls = bool(config.get("starttls", True)) and not usar_ssl
    if usar_ssl:
        cliente = smtplib.SMTP_SSL(host, porta, timeout=30, context=ssl.create_default_context())
    else:
        cliente = smtplib.SMTP(host, porta, timeout=30)

    resultados = []
    with cliente as smtp:
        if usar_starttls:
            smtp.starttls(context=ssl.create_default_context())
        usuario = str(config.get("username", ""))
        senha = str(config.get("password", ""))
        if usuario:
            smtp.login(usuario, senha)

        for item in mensagens:
            destinatario = str(item["to"]).strip()
            if not validar_email(destinatario):
                resultados.append({"to": destinatario, "ok": False, "erro": "Endereco invalido"})
                continue
            mensagem = EmailMessage()
            mensagem["Subject"] = str(item.get("subject", "Seu certificado"))
            mensagem["From"] = remetente
            mensagem["To"] = destinatario
            mensagem.set_content(str(item["body"]))
            mensagem.add_attachment(
                bytes(item["attachment"]),
                maintype=str(item.get("maintype", "application")),
                subtype=str(item.get("subtype", "pdf")),
                filename=str(item["filename"]),
            )
            try:
                recusados = smtp.send_message(mensagem)
                if recusados:
                    resultados.append({"to": destinatario, "ok": False, "erro": "O servidor recusou o destinatario"})
                else:
                    resultados.append({"to": destinatario, "ok": True, "erro": ""})
            except (smtplib.SMTPException, OSError) as error:
                resultados.append({"to": destinatario, "ok": False, "erro": str(error)[:200]})
    return resultados
