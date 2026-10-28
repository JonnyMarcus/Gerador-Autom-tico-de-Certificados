import re
from pathlib import Path
from pptx import Presentation
from pptx.shapes.group import GroupShape

TOKEN = re.compile(r"\{([A-Za-zÀ-ÿ0-9_ -]+)\}")

def _replace_in_shape(shape, replacements: dict[str, str]) -> None:
    if isinstance(shape, GroupShape):
        for child in shape.shapes:
            _replace_in_shape(child, replacements)
        return
    if not shape.has_text_frame:
        return
    for paragraph in shape.text_frame.paragraphs:
        original = "".join(run.text for run in paragraph.runs)
        if not original:
            continue
        updated = TOKEN.sub(lambda match: replacements.get(match.group(1).strip().lower(), match.group(0)), original)
        if updated != original and paragraph.runs:
            paragraph.runs[0].text = updated
            for run in paragraph.runs[1:]:
                run.text = ""

def preencher_modelo(modelo: Path, destino: Path, dados: dict[str, str]) -> None:
    """Preenche tokens {coluna} em todos os slides e salva uma cópia do modelo."""
    presentation = Presentation(modelo)
    replacements = {key.strip().lower(): value for key, value in dados.items()}
    for slide in presentation.slides:
        for shape in slide.shapes:
            _replace_in_shape(shape, replacements)
    presentation.save(destino)
