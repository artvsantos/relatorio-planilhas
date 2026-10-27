import pytest

from sales_report.columns import (
    MissingColumnsError,
    find_header,
    map_columns,
    normalize_header,
)


@pytest.mark.parametrize('header, expected', [
    ('Preço Unitário', 'preco unitario'),
    ('  QTD. ', 'qtd'),
    ('Valor_Total', 'valor total'),
    ('DATA DA VENDA', 'data da venda'),
    (None, ''),
    (2026, '2026'),
])
def test_normalize_header(header, expected):
    assert normalize_header(header) == expected


def test_maps_the_usual_names():
    headers = ['Data', 'Produto', 'Qtd', 'Preço unitário']

    assert map_columns(headers) == {
        'date': 0, 'product': 1, 'quantity': 2, 'unit_price': 3,
    }


def test_maps_other_spellings_of_the_same_column():
    headers = ['Loja', 'Dt. venda', 'Descrição do produto', 'Quantidade',
               'Valor total', 'Observação']

    assert map_columns(headers) == {
        'store': 0, 'date': 1, 'product': 2, 'quantity': 3, 'total': 4,
    }


def test_first_column_wins_when_a_name_repeats():
    assert map_columns(['Data', 'Produto', 'Qtd', 'Total', 'Data'])[
        'date'] == 0


@pytest.mark.parametrize('headers, missing', [
    (['Data', 'Produto', 'Preço'], ['Quantidade']),
    (['Produto', 'Qtd', 'Total'], ['Data']),
    # Sem preço e sem total não dá para saber quanto foi vendido.
    (['Data', 'Produto', 'Qtd'], ['Preço unitário ou Total']),
])
def test_missing_required_columns(headers, missing):
    with pytest.raises(MissingColumnsError) as error:
        map_columns(headers)

    assert error.value.missing == missing
    assert str(error.value) == f"faltam as colunas: {', '.join(missing)}"


HEADER = ['Data', 'Produto', 'Qtd', 'Preço']
HEADER_COLUMNS = {'date': 0, 'product': 1, 'quantity': 2, 'unit_price': 3}


def test_header_can_be_below_a_title():
    # Muita planilha começa com título, data de emissão e uma linha vazia.
    rows = [
        ['Relatório de vendas - Loja Centro', None, None, None],
        ['Emitido em 01/11/2026', None, None, None],
        [None, None, None, None],
        HEADER,
        ['27/10/2026', 'Caneca', 2, 30],
    ]

    assert find_header(rows) == (3, HEADER_COLUMNS)


def test_blank_rows_do_not_count_in_the_search():
    rows = [[None, '']] * 30 + [HEADER]

    assert find_header(rows) == (30, HEADER_COLUMNS)


def test_search_stops_after_15_filled_rows():
    rows = [['texto']] * 15 + [HEADER]

    with pytest.raises(MissingColumnsError):
        find_header(rows)


def test_header_not_found_tells_what_the_first_row_lacks():
    rows = [['Data', 'Produto', 'Preço'], ['nada', 'aqui', 'x']]

    with pytest.raises(MissingColumnsError) as error:
        find_header(rows)

    assert error.value.missing == ['Quantidade']


def test_no_filled_rows_lacks_everything():
    with pytest.raises(MissingColumnsError) as error:
        find_header([[None], ['  ']])

    assert error.value.missing == [
        'Data', 'Produto', 'Quantidade', 'Preço unitário ou Total',
    ]
