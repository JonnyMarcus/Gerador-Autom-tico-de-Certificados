import argparse
import logging
from pathlib import Path
from .config import default_settings
from .exportador import gerar_certificados
from .logging_config import configurar_logging
from .planilha import ler_participantes
from .validacao import validar_participantes

logger = logging.getLogger(__name__)

def criar_parser() -> argparse.ArgumentParser:
    defaults = default_settings()
    parser = argparse.ArgumentParser(description="Gera certificados PowerPoint a partir de uma planilha Excel.")
    parser.add_argument("--planilha", type=Path, default=defaults.planilha, help="Arquivo .xlsx dos participantes")
    parser.add_argument("--modelo", type=Path, default=defaults.modelo, help="Modelo .pptx com tokens entre chaves")
    parser.add_argument("--saida", type=Path, default=defaults.saida, help="Pasta de destino")
    parser.add_argument("--coluna-nome", default="nome", help="Cabeçalho usado para nomear os certificados")
    parser.add_argument("--campos", nargs="*", default=[], help="Outras colunas obrigatórias")
    parser.add_argument("--aba", help="Nome da aba do Excel (padrão: aba ativa)")
    parser.add_argument("--verbose", action="store_true", help="Exibe detalhes de execução")
    return parser

def main() -> int:
    args = criar_parser().parse_args()
    configurar_logging(args.verbose)
    try:
        if not args.modelo.is_file():
            raise FileNotFoundError(f"Modelo PowerPoint não encontrado: {args.modelo}")
        participantes = ler_participantes(args.planilha, args.aba)
        coluna_nome = args.coluna_nome.strip().lower()
        campos = [campo.strip().lower() for campo in args.campos]
        validar_participantes(participantes, coluna_nome, campos)
        arquivos = gerar_certificados(args.modelo, args.saida, participantes, coluna_nome)
    except (FileNotFoundError, ValueError, PermissionError) as error:
        logger.error("%s", error)
        return 1
    logger.info("%d certificado(s) criado(s) em %s", len(arquivos), args.saida.resolve())
    return 0
