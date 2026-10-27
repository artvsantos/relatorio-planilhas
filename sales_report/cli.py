"""Linha de comando: python -m sales_report PASTA [-o relatorio.xlsx]."""

import argparse
import sys
from pathlib import Path

from sales_report.cleaning import consolidate, load_file
from sales_report.parsing import format_brl
from sales_report.reader import list_spreadsheets
from sales_report.report import write_report


def plural(count: int, singular: str, plural_form: str) -> str:
    number = f'{count:,}'.replace(',', '.')
    return f'{number} {singular if count == 1 else plural_form}'


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
    folder = Path(args.pasta)
    if not folder.is_dir():
        return fail(f'Pasta não encontrada: {folder}')
    files = list_spreadsheets(folder)
    if not files:
        return fail(f'Nenhuma planilha (.xlsx ou .csv) em {folder}')

    print(f"Lendo {plural(len(files), 'planilha', 'planilhas')} em {folder}")
    results = [load_file(path) for path in files]
    for result in results:
        if result.error:
            print(f'  {result.source}: ignorada ({result.error})')
            continue
        count = plural(len(result.sales), 'venda', 'vendas')
        line = f'  {result.source}: {count}'
        if result.problems:
            line += ', ' + plural(len(result.problems),
                                  'linha com problema',
                                  'linhas com problema')
        print(line)

    sales, problems = consolidate(results)
    if not sales:
        return fail('Nenhuma venda válida encontrada. Veja os motivos acima.')

    output = Path(args.saida)
    try:
        write_report(output, sales, problems,
                     [result.source for result in results])
    except PermissionError:
        return fail(
            f'Não consegui salvar {output}. Ele está aberto no Excel? '
            'Feche o arquivo e rode de novo.',
        )

    revenue = sum(sale.total for sale in sales)
    print(f'\nRelatório salvo em {output}')
    print(f"{plural(len(sales), 'venda', 'vendas')}, "
          f'{format_brl(revenue)} de faturamento.')
    if problems:
        print(plural(len(problems), 'problema listado',
                     'problemas listados') + ' na aba "Problemas".')
    return 0


def fail(message: str) -> int:
    print(message, file=sys.stderr)
    return 1
