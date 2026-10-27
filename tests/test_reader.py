from datetime import datetime

import pytest

from sales_report.reader import ReadError, list_spreadsheets, read_rows
from tests.helpers import write_csv, write_xlsx


def test_reads_xlsx_keeping_numbers_and_dates(tmp_path):
    path = write_xlsx(tmp_path / 'loja.xlsx', [
        ['Data', 'Produto', 'Qtd', 'Preço'],
        [datetime(2026, 10, 27), 'Caneca', 2, 30.5],
    ])

    assert read_rows(path) == [
        ['Data', 'Produto', 'Qtd', 'Preço'],
        [datetime(2026, 10, 27), 'Caneca', 2, 30.5],
    ]


def test_reads_only_the_first_sheet(tmp_path):
    path = write_xlsx(tmp_path / 'loja.xlsx', [['primeira']],
                      extra_sheet=[['segunda']])

    assert read_rows(path) == [['primeira']]


def test_reads_csv_as_excel_in_portuguese_saves_it(tmp_path):
    path = write_csv(tmp_path / 'loja.csv', [
        ['Data', 'Produto', 'Qtd', 'Preço'],
        ['27/10/2026', 'Pão de açúcar', '2', '1.234,50'],
    ])

    assert read_rows(path) == [
        ['Data', 'Produto', 'Qtd', 'Preço'],
        ['27/10/2026', 'Pão de açúcar', '2', '1.234,50'],
    ]


@pytest.mark.parametrize('delimiter, encoding', [
    (',', 'utf-8-sig'),     # Google Planilhas e sistemas
    ('\t', 'utf-8'),
    (';', 'utf-8'),
])
def test_reads_other_csv_flavors(tmp_path, delimiter, encoding):
    path = write_csv(tmp_path / 'loja.csv', [
        ['Data', 'Produto', 'Preço'], ['27/10/2026', 'Açaí', '10'],
    ], delimiter=delimiter, encoding=encoding)

    assert read_rows(path)[1] == ['27/10/2026', 'Açaí', '10']


def test_empty_csv(tmp_path):
    path = tmp_path / 'vazio.csv'
    path.write_text('')

    assert read_rows(path) == []


def test_broken_xlsx_is_reported(tmp_path):
    path = tmp_path / 'quebrado.xlsx'
    path.write_text('não sou uma planilha')

    with pytest.raises(ReadError, match='não consegui abrir'):
        read_rows(path)


def test_lists_only_spreadsheets_in_alphabetical_order(tmp_path):
    for name in ['b.csv', 'a.xlsx', 'C.XLSX', 'notas.txt', 'foto.png',
                 '~$a.xlsx']:   # ~$: arquivo temporário do Excel aberto
        (tmp_path / name).write_text('x')
    (tmp_path / 'subpasta.csv').mkdir()

    names = [path.name for path in list_spreadsheets(tmp_path)]

    assert names == ['a.xlsx', 'b.csv', 'C.XLSX']


def test_csv_with_a_single_column(tmp_path):
    # Sem separador nenhum para detectar.
    path = tmp_path / 'uma.csv'
    path.write_text('Produto\nCaneca\n', encoding='utf-8')

    assert read_rows(path) == [['Produto'], ['Caneca']]


def test_csv_with_characters_that_only_cp1252_has(tmp_path):
    # O Excel troca "-" por "–" e aspas retas por curvas; em latin-1 esses
    # caracteres viram lixo.
    path = write_csv(tmp_path / 'loja.csv', [
        ['Produto'], ['Café – 1 kg “especial”'],
    ])

    assert read_rows(path)[1] == ['Café – 1 kg “especial”']
