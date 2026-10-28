from dataclasses import dataclass
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[2]

@dataclass(frozen=True)
class Settings:
    planilha: Path
    modelo: Path
    saida: Path
    coluna_nome: str = "nome"
    campos: tuple[str, ...] = ()
    aba: str | None = None

def default_settings() -> Settings:
    return Settings(planilha=ROOT_DIR / "data" / "participantes.xlsx", modelo=ROOT_DIR / "assets" / "certificado_base.pptx", saida=ROOT_DIR / "output")
