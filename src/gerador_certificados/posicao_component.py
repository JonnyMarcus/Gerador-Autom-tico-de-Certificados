"""Componente Streamlit simples para mover o nome sobre a primeira página do PDF."""

from pathlib import Path

import streamlit.components.v1 as components

_COMPONENTE = components.declare_component(
    "editor_posicao_certificado",
    path=str(Path(__file__).parent / "editor_posicao"),
)


def editor_posicao(
    *, imagem: str, nome: str, x: int, y: int, tamanho: int, cor: str,
    fonte: str, chave: str, on_change=None,
):
    return _COMPONENTE(
        imagem=imagem,
        nome=nome,
        x=x,
        y=y,
        tamanho=tamanho,
        cor=cor,
        fonte=fonte,
        default=None,
        key=chave,
        on_change=on_change,
    )
