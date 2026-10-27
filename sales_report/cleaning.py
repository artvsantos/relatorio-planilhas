"""Transforma as linhas das planilhas em vendas, ou em problemas.

Uma linha com erro não derruba o relatório: ela vai para a lista de
problemas, com o arquivo, a linha e o motivo, para alguém corrigir na
origem. Melhor do que sumir em silêncio ou travar tudo.
"""

import re
import unicodedata
from dataclasses import dataclass, replace
from datetime import date
from pathlib import Path

from sales_report.columns import MissingColumnsError, find_header
from sales_report.parsing import format_brl, is_empty, parse_date, parse_number
from sales_report.reader import ReadError, read_rows

# Diferença aceita entre o total da planilha e quantidade × preço
# (arredondamento de centavos).
TOTAL_TOLERANCE = 0.01


@dataclass(frozen=True)
class Sale:
    source: str
    line: int
    store: str
    date: date
    product: str
    quantity: float
    unit_price: float
    total: float


@dataclass(frozen=True)
class Problem:
    source: str
    line: int | None  # None: o problema é o arquivo inteiro
    message: str


@dataclass(frozen=True)
class FileResult:
    source: str
    sales: list[Sale]
    problems: list[Problem]
    error: str | None  # arquivo recusado, e por quê


def store_from_filename(name: str) -> str:
    """"loja_centro.xlsx" → "Loja Centro"."""
    words = re.split(r'[\s_-]+', Path(name).stem)
    return ' '.join(word.capitalize() for word in words if word)


def clean_text(value) -> str:
    if isinstance(value, float) and value.is_integer():
        value = int(value)  # código de produto 123 lido como 123.0
    return ' '.join(str(value).split())


def load_file(path: Path) -> FileResult:
    source = path.name
    try:
        rows = read_rows(path)
    except ReadError as error:
        return FileResult(source, [], [], str(error))

    if all(_is_blank(row) for row in rows):
        return FileResult(source, [], [], 'a planilha está vazia')
    try:
        header_index, columns = find_header(rows)
    except MissingColumnsError as error:
        return FileResult(source, [], [], str(error))

    sales, problems = [], []
    default_store = store_from_filename(source)
    for offset, row in enumerate(rows[header_index + 1:]):
        if _is_blank(row):
            continue
        # +1 porque a planilha começa da linha 1, +1 por causa do cabeçalho.
        line = header_index + 2 + offset
        result = _parse_row(row, columns, source, line, default_store)
        (sales if isinstance(result, Sale) else problems).append(result)
    return FileResult(source, sales, problems, None)


def _is_blank(row: list) -> bool:
    return all(is_empty(value) for value in row)


def _parse_row(row, columns, source, line, default_store) -> Sale | Problem:
    def cell(column):
        position = columns.get(column)
        if position is None or position >= len(row):
            return None
        value = row[position]
        return None if is_empty(value) else value

    reasons = []

    def number(column, label):
        value = cell(column)
        if value is None:
            return None
        try:
            return parse_number(value)
        except ValueError as error:
            reasons.append(f'{label}: {error}')
            return None

    sold_on = None
    if (raw_date := cell('date')) is None:
        reasons.append('Data vazia')
    else:
        try:
            sold_on = parse_date(raw_date)
        except ValueError as error:
            reasons.append(f'Data: {error}')

    product = cell('product')
    if product is None:
        reasons.append('Produto vazio')

    has_quantity = cell('quantity') is not None
    quantity = number('quantity', 'Quantidade')
    if not has_quantity:
        reasons.append('Quantidade vazia')
    elif quantity is not None and quantity <= 0:
        reasons.append('Quantidade precisa ser maior que zero')

    had_price = cell('unit_price') is not None or cell('total') is not None
    price = number('unit_price', 'Preço')
    total = number('total', 'Total')
    if not had_price:
        reasons.append(
            'Preço vazio' if 'unit_price' in columns else 'Total vazio',
        )
    if price is not None and price < 0:
        reasons.append('Preço não pode ser negativo')
    if total is not None and total < 0:
        reasons.append('Total não pode ser negativo')

    if reasons:
        return Problem(source, line, '; '.join(reasons))

    if price is not None:
        expected = round(quantity * price, 2)
        if total is not None and abs(total - expected) > TOTAL_TOLERANCE:
            return Problem(source, line, (
                f'o total ({format_brl(total)}) não bate com quantidade × '
                f'preço ({format_brl(expected)})'
            ))
        total = expected if total is None else total
    else:
        price = round(total / quantity, 2)

    store = cell('store')
    return Sale(
        source=source,
        line=line,
        store=clean_text(store) if store is not None else default_store,
        date=sold_on,
        product=clean_text(product),
        quantity=quantity,
        unit_price=price,
        total=round(total, 2),
    )


def _name_key(name: str) -> str:
    # "Pão de Queijo", "pao de queijo" e "PÃO DE QUEIJO" são o mesmo.
    text = unicodedata.normalize('NFKD', name.casefold())
    return ''.join(char for char in text if not unicodedata.combining(char))


def consolidate(results: list[FileResult]) -> tuple[list[Sale],
                                                    list[Problem]]:
    """Junta as vendas de todos os arquivos, com nomes de produto e de loja
    unificados (fica a grafia que apareceu primeiro)."""
    products: dict[str, str] = {}
    stores: dict[str, str] = {}
    sales, problems = [], []
    for result in results:
        for sale in result.sales:
            sales.append(replace(
                sale,
                store=stores.setdefault(_name_key(sale.store), sale.store),
                product=products.setdefault(_name_key(sale.product),
                                            sale.product),
            ))
        problems.extend(result.problems)
        if result.error:
            problems.append(Problem(
                result.source, None, f'arquivo ignorado: {result.error}',
            ))
    return sales, problems
