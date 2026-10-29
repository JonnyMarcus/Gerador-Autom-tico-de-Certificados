import re
import unicodedata
from pathlib import Path
from .certificado import preencher_modelo

def nome_seguro(valor: str) -> str:
    normalized = unicodedata.normalize("NFKD", valor).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^A-Za-z0-9_-]+", "_", normalized).strip("_-") or "participante"

def gerar_certificados(
    modelo: Path,
    saida: Path,
    participantes: list[dict[str, str]],
    coluna_nome: str,
    imagens: list[dict[str, float | str | bytes]] | None = None,
) -> list[Path]:
    saida.mkdir(parents=True, exist_ok=True)
    destinos = []
    for index, pessoa in enumerate(participantes, start=1):
        destino = saida / f"{index:03d}_{nome_seguro(pessoa[coluna_nome.lower()])}.pptx"
        preencher_modelo(modelo, destino, pessoa, imagens=imagens)
        destinos.append(destino)
    return destinos
