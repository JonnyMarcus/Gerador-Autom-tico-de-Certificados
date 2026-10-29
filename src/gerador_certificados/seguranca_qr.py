"""Assinatura HMAC e validação dos QR Codes dos certificados."""

from __future__ import annotations

import base64
import hashlib
import hmac
import json
from urllib.parse import quote


def gerar_link_verificacao(dados: dict[str, object], segredo: str, url_publica: str) -> str:
    conteudo = json.dumps(dados, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")
    token = base64.urlsafe_b64encode(conteudo).decode("ascii").rstrip("=")
    assinatura = hmac.new(segredo.encode("utf-8"), token.encode("ascii"), hashlib.sha256).digest()
    assinatura_url = base64.urlsafe_b64encode(assinatura).decode("ascii").rstrip("=")
    return f"{url_publica.rstrip('/')}?verificar={quote(token + '.' + assinatura_url)}"


def verificar_link_verificacao(token_assinado: str, segredo: str) -> dict[str, object] | None:
    try:
        token, assinatura_recebida = token_assinado.split(".", 1)
        assinatura_esperada = base64.urlsafe_b64encode(
            hmac.new(segredo.encode("utf-8"), token.encode("ascii"), hashlib.sha256).digest()
        ).decode("ascii").rstrip("=")
        if not hmac.compare_digest(assinatura_recebida, assinatura_esperada):
            return None
        preenchimento = "=" * (-len(token) % 4)
        dados = json.loads(base64.urlsafe_b64decode(token + preenchimento).decode("utf-8"))
        return dados if isinstance(dados, dict) else None
    except (ValueError, UnicodeDecodeError, json.JSONDecodeError):
        return None
