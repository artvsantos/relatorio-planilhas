"""Descobre qual coluna é qual, seja qual for o nome que a loja usou."""

import re
import unicodedata

from sales_report.parsing import is_empty

# Nomes já normalizados (minúsculas, sem acento e sem pontuação).
ALIASES = {
    'date': {'data', 'data da venda', 'data venda', 'dt venda', 'dt', 'dia'},
    'product': {'produto', 'item', 'descricao', 'descricao do produto',
                'nome do produto', 'mercadoria'},
    'quantity': {'quantidade', 'qtd', 'qtde', 'quant', 'qt'},
    'unit_price': {'preco', 'preco unitario', 'preco unit', 'valor unitario',
                   'valor unit', 'vl unit', 'vl unitario', 'unitario'},
    'total': {'total', 'valor total', 'vl total', 'total da venda',
              'valor da venda', 'valor'},
    'store': {'loja', 'filial', 'unidade', 'ponto de venda'},
}
LABELS = {'date': 'Data', 'product': 'Produto', 'quantity': 'Quantidade'}
# Quantas linhas preenchidas do topo procurar pelo cabeçalho.
HEADER_SEARCH_ROWS = 15


class MissingColumnsError(ValueError):
    def __init__(self, missing: list[str]):
        self.missing = missing
        super().__init__(f"faltam as colunas: {', '.join(missing)}")


def normalize_header(header) -> str:
    """"Preço Unitário" → "preco unitario"."""
    if header is None:
        return ''
    text = unicodedata.normalize('NFKD', str(header))
    text = ''.join(char for char in text if not unicodedata.combining(char))
    return ' '.join(re.sub(r'[^a-z0-9]+', ' ', text.lower()).split())


def map_columns(headers: list) -> dict[str, int]:
    """Posição de cada coluna conhecida. Levanta MissingColumnsError se
    faltar alguma obrigatória."""
    found = {}
    for position, header in enumerate(headers):
        name = normalize_header(header)
        for column, aliases in ALIASES.items():
            if name in aliases and column not in found:
                found[column] = position

    missing = [label for column, label in LABELS.items()
               if column not in found]
    if 'unit_price' not in found and 'total' not in found:
        missing.append('Preço unitário ou Total')
    if missing:
        raise MissingColumnsError(missing)
    return found


def find_header(rows: list[list]) -> tuple[int, dict[str, int]]:
    """Acha o cabeçalho: a primeira linha com todas as colunas obrigatórias
    (o título que costuma vir acima não tem). Devolve o índice da linha e
    as colunas.

    Se não achar, levanta o erro da primeira linha preenchida, que é a
    mais provável de ser o cabeçalho, para dizer o que falta nela.
    """
    first_error = None
    filled = 0
    for index, row in enumerate(rows):
        if all(is_empty(value) for value in row):
            continue
        try:
            return index, map_columns(row)
        except MissingColumnsError as error:
            first_error = first_error or error
        filled += 1
        if filled == HEADER_SEARCH_ROWS:
            break
    raise first_error or MissingColumnsError(
        [*LABELS.values(), 'Preço unitário ou Total'],
    )
