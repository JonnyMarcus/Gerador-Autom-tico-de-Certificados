from io import BytesIO
from pathlib import Path

from pypdf import PdfReader, PdfWriter
from reportlab.lib.colors import HexColor
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas

from .exportador import nome_seguro


def _sobreposicao_nome(nome: str, largura: float, altura: float) -> PdfReader:
    buffer = BytesIO()
    desenho = canvas.Canvas(buffer, pagesize=(largura, altura), pageCompression=1)
    fonte = "Helvetica-Bold"
    tamanho = 32.0
    while tamanho > 12 and stringWidth(nome, fonte, tamanho) > largura * 0.84:
        tamanho -= 1
    desenho.setFont(fonte, tamanho)
    desenho.setFillColor(HexColor("#17233C"))
    desenho.drawCentredString(largura / 2, altura * 0.49, nome)
    desenho.save()
    buffer.seek(0)
    return PdfReader(buffer)


def gerar_certificados_pdf(
    modelo: Path,
    saida: Path,
    participantes: list[dict[str, str]],
    coluna_nome: str = "nome",
) -> list[Path]:
    """Aplica cada nome ao centro da primeira página de um PDF de modelo."""
    if not modelo.is_file():
        raise FileNotFoundError(f"Modelo PDF não encontrado: {modelo}")
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
        camada_nome = _sobreposicao_nome(nome, largura, altura)
        primeira_pagina.merge_page(camada_nome.pages[0])

        for pagina in leitor.pages[1:]:
            escritor.add_page(pagina)

        destino = saida / f"{indice:03d}_{nome_seguro(nome)}.pdf"
        with destino.open("wb") as arquivo:
            escritor.write(arquivo)
        destinos.append(destino)

    return destinos
