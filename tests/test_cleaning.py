from datetime import date, datetime

import pytest

from sales_report.cleaning import (
    FileResult,
    Problem,
    Sale,
    clean_text,
    consolidate,
    load_file,
    store_from_filename,
)
from tests.helpers import write_csv, write_xlsx

HEADER = ['Data', 'Produto', 'Qtd', 'Preço']


def load(tmp_path, rows, name='loja_centro.xlsx'):
    return load_file(write_xlsx(tmp_path / name, rows))


def test_valid_rows_become_sales(tmp_path):
    result = load(tmp_path, [
        HEADER,
        [datetime(2026, 10, 27), 'Caneca', 2, 30.5],
        ['28/10/2026', '  Camiseta   preta ', '1', '59,90'],
    ])

    assert result.error is None
    assert result.problems == []
    assert result.sales == [
        Sale('loja_centro.xlsx', 2, 'Loja Centro', date(2026, 10, 27),
             'Caneca', 2.0, 30.5, 61.0),
        Sale('loja_centro.xlsx', 3, 'Loja Centro', date(2026, 10, 28),
             'Camiseta preta', 1.0, 59.9, 59.9),
    ]


def test_line_numbers_match_the_spreadsheet(tmp_path):
    # Quem abrir a planilha precisa achar a linha do problema.
    result = load(tmp_path, [
        ['Vendas de outubro'],
        [],
        HEADER,
        ['27/10/2026', 'Caneca', 1, 10],
    ])

    assert result.sales[0].line == 4


def test_store_column_wins_over_the_file_name(tmp_path):
    result = load(tmp_path, [
        ['Loja', *HEADER],
        ['Filial Norte', '27/10/2026', 'Caneca', 1, 10],
        ['', '27/10/2026', 'Caneca', 1, 10],
    ])

    assert [sale.store for sale in result.sales] == [
        'Filial Norte', 'Loja Centro',
    ]


@pytest.mark.parametrize('name, expected', [
    ('loja_centro.xlsx', 'Loja Centro'),
    ('vendas-shopping-norte.csv', 'Vendas Shopping Norte'),
    ('LOJA 3.xlsx', 'Loja 3'),
])
def test_store_from_filename(name, expected):
    assert store_from_filename(name) == expected


def test_total_alone_gives_the_unit_price(tmp_path):
    result = load(tmp_path, [
        ['Data', 'Produto', 'Quantidade', 'Valor total'],
        ['27/10/2026', 'Caneca', 4, '100,00'],
    ])

    sale = result.sales[0]
    assert (sale.unit_price, sale.total) == (25.0, 100.0)


def test_total_that_matches_is_accepted(tmp_path):
    result = load(tmp_path, [
        [*HEADER, 'Total'],
        ['27/10/2026', 'Caneca', 3, '33,33', '99,99'],
    ])

    assert result.problems == []
    assert result.sales[0].total == 99.99


def test_total_that_does_not_match_is_a_problem(tmp_path):
    result = load(tmp_path, [
        [*HEADER, 'Total'],
        ['27/10/2026', 'Caneca', 3, 10, 50],
    ])

    assert result.sales == []
    assert result.problems == [Problem(
        'loja_centro.xlsx', 2,
        'o total (R$ 50,00) não bate com quantidade × preço (R$ 30,00)',
    )]


@pytest.mark.parametrize('row, message', [
    (['31/02/2026', 'Caneca', 1, 10],
     'Data: "31/02/2026" não é uma data válida'),
    (['', 'Caneca', 1, 10], 'Data vazia'),
    (['27/10/2026', '', 1, 10], 'Produto vazio'),
    (['27/10/2026', 'Caneca', 'dois', 10],
     'Quantidade: "dois" não é um número'),
    (['27/10/2026', 'Caneca', 0, 10],
     'Quantidade precisa ser maior que zero'),
    (['27/10/2026', 'Caneca', -1, 10],
     'Quantidade precisa ser maior que zero'),
    (['27/10/2026', 'Caneca', 1, ''], 'Preço vazio'),
    (['27/10/2026', 'Caneca', 1, 'R$ -5,00'],
     'Preço não pode ser negativo'),
])
def test_invalid_rows_become_problems(tmp_path, row, message):
    result = load(tmp_path, [HEADER, row])

    assert result.sales == []
    assert result.problems == [Problem('loja_centro.xlsx', 2, message)]


def test_all_problems_of_a_row_are_listed_together(tmp_path):
    result = load(tmp_path, [HEADER, ['ontem', '', 0, 10]])

    assert result.problems[0].message == (
        'Data: "ontem" não é uma data válida; Produto vazio; '
        'Quantidade precisa ser maior que zero'
    )


def test_bad_rows_do_not_stop_the_good_ones(tmp_path):
    result = load(tmp_path, [
        HEADER,
        ['27/10/2026', 'Caneca', 1, 10],
        ['27/10/2026', 'Caneca', 'x', 10],
        ['28/10/2026', 'Prato', 2, 15],
    ])

    assert [sale.line for sale in result.sales] == [2, 4]
    assert [problem.line for problem in result.problems] == [3]


def test_blank_rows_are_skipped_silently(tmp_path):
    # O Excel costuma deixar linhas vazias (ou só com espaço) no fim.
    result = load(tmp_path, [
        HEADER,
        ['27/10/2026', 'Caneca', 1, 10],
        [None, None, None, None],
        ['', '  ', None, ''],
    ])

    assert len(result.sales) == 1
    assert result.problems == []


def test_missing_columns_reject_the_file(tmp_path):
    result = load(tmp_path, [['Data', 'Produto', 'Preço'],
                             ['27/10/2026', 'Caneca', 10]])

    assert result == FileResult(
        'loja_centro.xlsx', [], [], 'faltam as colunas: Quantidade',
    )


def test_header_after_many_blank_rows_is_found(tmp_path):
    result = load(tmp_path, [[None]] * 20 + [
        HEADER, ['27/10/2026', 'Caneca', 1, 10],
    ])

    assert result.error is None
    assert result.sales[0].line == 22


def test_empty_file_is_rejected(tmp_path):
    path = tmp_path / 'vazia.csv'
    path.write_text('')

    assert load_file(path).error == 'a planilha está vazia'


def test_unreadable_file_is_rejected(tmp_path):
    path = tmp_path / 'quebrada.xlsx'
    path.write_text('lixo')

    assert load_file(path).error.startswith('não consegui abrir')


def test_csv_with_brazilian_formats(tmp_path):
    path = write_csv(tmp_path / 'loja.csv', [
        ['Dt. venda', 'Descrição', 'Qtde', 'Valor unitário'],
        ['27/10/26', 'Pão de queijo', '12', 'R$ 1,50'],
    ])

    [sale] = load_file(path).sales

    assert (sale.date, sale.product, sale.quantity, sale.total) == (
        date(2026, 10, 27), 'Pão de queijo', 12.0, 18.0,
    )


def test_consolidate_unifies_names_written_differently():
    # "caneca azul" e "Caneca Azul" são o mesmo produto no relatório; fica
    # a grafia que apareceu primeiro.
    def sale(store, product):
        return Sale('a.xlsx', 2, store, date(2026, 10, 27), product,
                    1.0, 10.0, 10.0)

    results = [
        FileResult('a.xlsx', [sale('Loja Centro', 'Caneca Azul')],
                   [Problem('a.xlsx', 3, 'Produto vazio')], None),
        FileResult('b.csv', [sale('loja centro', 'caneca azul'),
                             sale('Loja Norte', 'Prato')], [], None),
        FileResult('c.csv', [], [], 'a planilha está vazia'),
    ]

    sales, problems = consolidate(results)

    assert [(s.store, s.product) for s in sales] == [
        ('Loja Centro', 'Caneca Azul'),
        ('Loja Centro', 'Caneca Azul'),
        ('Loja Norte', 'Prato'),
    ]
    # O arquivo recusado entra na lista de problemas, sem número de linha.
    assert problems == [
        Problem('a.xlsx', 3, 'Produto vazio'),
        Problem('c.csv', None, 'arquivo ignorado: a planilha está vazia'),
    ]


@pytest.mark.parametrize('value, expected', [
    (123.0, '123'),        # código de produto que o Excel leu como número
    (12.5, '12.5'),
    ('  Caneca   azul ', 'Caneca azul'),
])
def test_clean_text(value, expected):
    assert clean_text(value) == expected


@pytest.mark.parametrize('header, row, message', [
    (HEADER, ['27/10/2026', 'Caneca', None, 10], 'Quantidade vazia'),
    (['Data', 'Produto', 'Qtd', 'Total'],
     ['27/10/2026', 'Caneca', 1, '-10,00'], 'Total não pode ser negativo'),
    (['Data', 'Produto', 'Qtd', 'Total'],
     ['27/10/2026', 'Caneca', 1, ''], 'Total vazio'),
])
def test_more_invalid_rows(tmp_path, header, row, message):
    result = load(tmp_path, [header, row])

    assert result.problems[0].message == message
