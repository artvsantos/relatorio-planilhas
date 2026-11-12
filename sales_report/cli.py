"""Linha de comando: python -m sales_report PASTA [-o relatorio.xlsx]."""

import argparse
import sys
from pathlib import Path

from sales_report.generate import ReportError, generate_report


def parse_args(argv):
    parser = argparse.ArgumentParser(
        prog='python -m sales_report',
        description=(
            'Junta as planilhas de vendas de uma pasta (.xlsx e .csv) e gera '
            'um relatório em Excel, com as linhas problemáticas listadas à '
            'parte.'
        ),
    )
    parser.add_argument('pasta', help='pasta com as planilhas das lojas')
    parser.add_argument(
        '-o', '--saida', default='relatorio.xlsx',
        help='onde salvar o relatório (padrão: relatorio.xlsx)',
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        summary = generate_report(Path(args.pasta), Path(args.saida), print)
    except ReportError as error:
        print(error, file=sys.stderr)
        return 1
    print()
    for line in summary.lines():
        print(line)
    return 0
