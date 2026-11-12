from pathlib import Path

import pytest

from sales_report import generate
from sales_report.generate import (
    ReportError,
    Summary,
    generate_report,
    plural,
)
from tests.helpers import write_csv, write_xlsx

HEADER = ['Data', 'Produto', 'Qtd', 'Preço']


@pytest.fixture
def folder(tmp_path):
    folder = tmp_path / 'vendas'
    folder.mkdir()
    write_xlsx(folder / 'loja_centro.xlsx', [
        HEADER,
        ['27/10/2026', 'Caneca', 2, 30],
        ['28/10/2026', 'Prato', 1, '1.000,00'],
    ])
    write_csv(folder / 'loja_norte.csv', [
        HEADER,
        ['27/10/2026', 'caneca', '1', '30,00'],
        ['ontem', 'Prato', '1', '10'],
    ])
    return folder


def ignore(line):
    pass


def test_tells_each_file_as_it_is_read(folder, tmp_path):
    lines = []

    generate_report(folder, tmp_path / 'r.xlsx', lines.append)

    assert lines == [
        f'Lendo 2 planilhas em {folder}',
        '  loja_centro.xlsx: 2 vendas',
        '  loja_norte.csv: 1 venda, 1 linha com problema',
    ]


def test_saves_the_report_and_returns_the_summary(folder, tmp_path):
    output = tmp_path / 'r.xlsx'

    summary = generate_report(folder, output, ignore)

    assert output.exists()
    assert summary == Summary(output, sales=3, revenue=1090.0, problems=1)


def test_ignored_file_is_explained(folder, tmp_path):
    write_xlsx(folder / 'estoque.xlsx', [['Produto', 'Saldo']])
    lines = []

    generate_report(folder, tmp_path / 'r.xlsx', lines.append)

    assert ('  estoque.xlsx: ignorada (faltam as colunas: Data, Quantidade, '
            'Preço unitário ou Total)') in lines


def test_report_saved_inside_the_folder_is_not_read_as_sales(folder):
    # Na segunda vez, o relatório da primeira já está na pasta; ele não é
    # planilha de vendas e não pode virar "arquivo ignorado".
    output = folder / 'relatorio.xlsx'
    generate_report(folder, output, ignore)
    lines = []

    summary = generate_report(folder, output, lines.append)

    assert lines[0] == f'Lendo 2 planilhas em {folder}'
    assert not any('relatorio.xlsx' in line for line in lines)
    assert summary.problems == 1


def test_missing_folder(tmp_path):
    with pytest.raises(ReportError, match='Pasta não encontrada'):
        generate_report(tmp_path / 'nao-existe', tmp_path / 'r.xlsx', ignore)


def test_folder_without_spreadsheets(tmp_path):
    (tmp_path / 'notas.txt').write_text('x')

    with pytest.raises(ReportError,
                       match=r'Nenhuma planilha \(.xlsx ou .csv\) em'):
        generate_report(tmp_path, tmp_path / 'r.xlsx', ignore)


def test_no_valid_sales_writes_nothing(tmp_path):
    write_xlsx(tmp_path / 'ruim.xlsx', [HEADER, ['ontem', '', 0, 1]])
    output = tmp_path / 'r.xlsx'
    lines = []

    with pytest.raises(ReportError, match='Nenhuma venda válida'):
        generate_report(tmp_path, output, lines.append)

    # Os motivos já foram contados, antes do erro.
    assert lines[-1].startswith('  ruim.xlsx: 0 vendas, 1 linha com')
    assert not output.exists()


def test_report_open_in_excel(folder, tmp_path, monkeypatch):
    # No Windows, o Excel trava o arquivo aberto e salvar dá
    # PermissionError.
    def locked(*args, **kwargs):
        raise PermissionError

    monkeypatch.setattr(generate, 'write_report', locked)
    output = tmp_path / 'r.xlsx'

    with pytest.raises(ReportError) as error:
        generate_report(folder, output, ignore)

    assert str(error.value) == (
        f'Não consegui salvar {output}. Ele está aberto no Excel? Feche o '
        'arquivo e tente de novo.'
    )


def test_summary_lines():
    summary = Summary(Path('r.xlsx'), sales=3, revenue=1090.0, problems=1)

    assert summary.lines() == [
        'Relatório salvo em r.xlsx',
        '3 vendas, R$ 1.090,00 de faturamento.',
        '1 problema listado na aba "Problemas".',
    ]


def test_summary_without_problems_has_no_problem_line():
    summary = Summary(Path('r.xlsx'), sales=1, revenue=10.0, problems=0)

    assert summary.lines() == [
        'Relatório salvo em r.xlsx',
        '1 venda, R$ 10,00 de faturamento.',
    ]


@pytest.mark.parametrize(('count', 'expected'), [
    (0, '0 vendas'),
    (1, '1 venda'),
    (2, '2 vendas'),
    (1234, '1.234 vendas'),
])
def test_plural(count, expected):
    assert plural(count, 'venda', 'vendas') == expected
