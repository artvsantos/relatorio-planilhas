"""O passo a passo do relatório, o mesmo na linha de comando e na janela.

Quem chama decide onde mostrar o andamento: a linha de comando imprime no
terminal; a janela mostra na tela.
"""

from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path

from sales_report.cleaning import FileResult, consolidate, load_file
from sales_report.parsing import format_brl
from sales_report.reader import list_spreadsheets
from sales_report.report import write_report


class ReportError(Exception):
    """O que impede o relatório de sair. A mensagem é para quem usa."""


@dataclass(frozen=True)
class Summary:
    output: Path
    sales: int
    revenue: float
    problems: int

    def lines(self) -> list[str]:
        lines = [
            f'Relatório salvo em {self.output}',
            f"{plural(self.sales, 'venda', 'vendas')}, "
            f'{format_brl(self.revenue)} de faturamento.',
        ]
        if self.problems:
            lines.append(plural(self.problems, 'problema listado',
                                'problemas listados')
                         + ' na aba "Problemas".')
        return lines


def plural(count: int, singular: str, plural_form: str) -> str:
    number = f'{count:,}'.replace(',', '.')
    return f'{number} {singular if count == 1 else plural_form}'


def generate_report(folder: Path, output: Path,
                    log: Callable[[str], None]) -> Summary:
    if not folder.is_dir():
        raise ReportError(f'Pasta não encontrada: {folder}')
    files = list_spreadsheets(folder)
    if output.exists():
        # O relatório anterior, salvo na mesma pasta, não é planilha de
        # vendas.
        files = [path for path in files if not path.samefile(output)]
    if not files:
        raise ReportError(f'Nenhuma planilha (.xlsx ou .csv) em {folder}')

    log(f"Lendo {plural(len(files), 'planilha', 'planilhas')} em {folder}")
    results = []
    for path in files:
        result = load_file(path)
        results.append(result)
        log(_file_line(result))

    sales, problems = consolidate(results)
    if not sales:
        raise ReportError(
            'Nenhuma venda válida encontrada. Veja os motivos acima.',
        )
    try:
        write_report(output, sales, problems,
                     [result.source for result in results])
    except PermissionError:
        raise ReportError(
            f'Não consegui salvar {output}. Ele está aberto no Excel? '
            'Feche o arquivo e tente de novo.',
        ) from None
    return Summary(output, sales=len(sales),
                   revenue=sum(sale.total for sale in sales),
                   problems=len(problems))


def _file_line(result: FileResult) -> str:
    if result.error:
        return f'  {result.source}: ignorada ({result.error})'
    line = f"  {result.source}: {plural(len(result.sales), 'venda', 'vendas')}"
    if result.problems:
        line += ', ' + plural(len(result.problems), 'linha com problema',
                              'linhas com problema')
    return line
