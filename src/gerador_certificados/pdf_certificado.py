from collections.abc import Callable
from io import BytesIO
from pathlib import Path

from pypdf import PdfReader, PdfWriter
from reportlab.lib.colors import HexColor
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas

from .exportador import nome_seguro

FONTES_PDF = (
    "Helvetica",
    "Helvetica-Bold",
    "Helvetica-Oblique",
    "Helvetica-BoldOblique",
    "Times-Roman",
    "Times-Bold",
    "Times-Italic",
    "Times-BoldItalic",
    "Courier",
    "Courier-Bold",
    "Courier-Oblique",
    "Courier-BoldOblique",
)


def _sobreposicao_nome(
    nome: str,
    largura: float,
    altura: float,
    posicao_x: float,
    posicao_y: float,
    fonte: str,
    tamanho_fonte: float,
    cor: str,
) -> PdfReader:
    buffer = BytesIO()
    desenho = canvas.Canvas(buffer, pagesize=(largura, altura), pageCompression=1)
    tamanho = tamanho_fonte
    while tamanho > 8 and stringWidth(nome, fonte, tamanho) > largura * 0.84:
        tamanho -= 1
    desenho.setFont(fonte, tamanho)
    desenho.setFillColor(HexColor(cor))
    x = largura * posicao_x / 100
    y = altura * (100 - posicao_y) / 100
    desenho.drawCentredString(x, y, nome)
    desenho.save()
    buffer.seek(0)
    return PdfReader(buffer)


def gerar_certificados_pdf(
    modelo: Path,
    saida: Path,
    participantes: list[dict[str, str]],
    coluna_nome: str = "nome",
    posicao_x: float = 50,
    posicao_y: float = 51,
    fonte: str = "Helvetica-Bold",
    tamanho_fonte: float = 32,
    cor: str = "#17233C",
    ao_gerar: Callable[[int, int], None] | None = None,
) -> list[Path]:
    """Aplica cada nome à primeira página do PDF, usando coordenadas percentuais."""
    if not modelo.is_file():
        raise FileNotFoundError(f"Modelo PDF não encontrado: {modelo}")
    if not 0 <= posicao_x <= 100 or not 0 <= posicao_y <= 100:
        raise ValueError("A posição do nome deve estar entre 0 e 100%.")
    if fonte not in FONTES_PDF:
        raise ValueError(f"Fonte não suportada. Opções: {', '.join(FONTES_PDF)}")
    if not 8 <= tamanho_fonte <= 72:
        raise ValueError("O tamanho da fonte deve estar entre 8 e 72 pontos.")
    try:
        HexColor(cor)
    except (ValueError, TypeError) as error:
        raise ValueError("A cor deve estar no formato hexadecimal, por exemplo #17233C.") from error

    saida.mkdir(parents=True, exist_ok=True)
    leitor = PdfReader(modelo)
    if not leitor.pages:
        raise ValueError("O PDF do certificado não contém páginas.")
    destinos = []

    for indice, participante in enumerate(participantes, start=1):
        nome = participante[coluna_nome.lower()].strip()
        escritor = PdfWriter()
        primeira_pagina = escritor.add_page(leitor.pages[0])
        largura = float(primeira_pagina.mediabox.width)
        altura = float(primeira_pagina.mediabox.height)
        camada_nome = _sobreposicao_nome(
            nome,
            largura,
            altura,
            posicao_x,
            posicao_y,
            fonte,
            tamanho_fonte,
            cor,
        )
        primeira_pagina.merge_page(camada_nome.pages[0])

        for pagina in leitor.pages[1:]:
            escritor.add_page(pagina)

        destino = saida / f"{indice:03d}_{nome_seguro(nome)}.pdf"
        with destino.open("wb") as arquivo:
            escritor.write(arquivo)
        destinos.append(destino)
        if ao_gerar:
            ao_gerar(indice, len(participantes))

    return destinos


def combinar_certificados_pdf(arquivos: list[Path]) -> bytes:
    """Combina certificados individuais em um único PDF para download."""
    if not arquivos:
        raise ValueError("Não há certificados para combinar.")
    escritor = PdfWriter()
    for arquivo in arquivos:
        escritor.append(arquivo)
    buffer = BytesIO()
    escritor.write(buffer)
    return buffer.getvalue()
