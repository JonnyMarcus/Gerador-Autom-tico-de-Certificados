from collections.abc import Sequence

def validar_participantes(participantes: Sequence[dict[str, str]], coluna_nome: str, campos: Sequence[str]) -> None:
    if not participantes:
        raise ValueError("Nenhum participante foi encontrado na planilha.")
    available = set(participantes[0])
    required = {coluna_nome.lower(), *(campo.lower() for campo in campos)}
    missing = sorted(required - available)
    if missing:
        raise ValueError(f"Colunas obrigatórias ausentes na planilha: {', '.join(missing)}")
    invalid = [str(index) for index, person in enumerate(participantes, start=2) if not person.get(coluna_nome.lower(), "").strip()]
    if invalid:
        raise ValueError(f"Nome vazio nas linhas: {', '.join(invalid)}")
