"""Leitura de números e datas como aparecem nas planilhas brasileiras.

Cada loja exporta de um jeito: "1.234,56" digitado à mão, "1234.56" vindo
de um sistema, "R$ 10,00" com o símbolo, datas em dd/mm/aaaa ou como o
número de série do Excel. Aqui tudo vira float e date.
"""

import math
import re
from datetime import date, datetime, timedelta

# O Excel conta os dias a partir de 30/12/1899 (por causa de um erro antigo
# do Lotus 1-2-3 com o ano bissexto de 1900, que o Excel copiou).
EXCEL_EPOCH = date(1899, 12, 30)
# Números nessa faixa são datas entre 1954 e 2119: fora dela, é mais
# provável que a célula tenha um número qualquer, e não uma data.
EXCEL_SERIAL_RANGE = (20_000, 80_000)

DAY_FIRST = re.compile(r'(\d{1,2})[/.-](\d{1,2})[/.-](\d{4}|\d{2})(?:\s.*)?')
ISO = re.compile(r'(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T].*)?')


def is_empty(value) -> bool:
    if value is None:
        return True
    if isinstance(value, float):
        return math.isnan(value)
    return isinstance(value, str) and not value.strip()


def parse_number(value) -> float:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = str(value).replace('R$', '').replace('\xa0', '').replace(' ', '')
    error = ValueError(f'"{value}" não é um número')
    if isinstance(value, bool) or not re.fullmatch(r'-?[\d.,]*\d[\d.,]*',
                                                   text):
        raise error

    if ',' in text and '.' in text:
        # O separador que vem por último é o decimal: 1.234,56 ou 1,234.56.
        decimal = ',' if text.rfind(',') > text.rfind('.') else '.'
    elif ',' in text:
        decimal = ','
    elif text.count('.') == 1 and len(text.rsplit('.', 1)[1]) != 3:
        decimal = '.'
    else:
        # "1.234" ou "1.234.567": no Brasil, ponto sozinho é milhar.
        decimal = None

    if decimal is None:
        integer, fraction = text, ''
    else:
        integer, _, fraction = text.rpartition(decimal)
    thousands = ',' if decimal == '.' else '.'
    integer = integer.replace(thousands, '')
    if not re.fullmatch(r'-?\d*', integer) or not re.fullmatch(r'\d*',
                                                              fraction):
        raise error
    return float(f'{integer or 0}.{fraction or 0}')


def parse_date(value) -> date:
    error = ValueError(f'"{value}" não é uma data válida')
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        low, high = EXCEL_SERIAL_RANGE
        if not low <= value <= high:
            raise error
        return EXCEL_EPOCH + timedelta(days=int(value))

    text = str(value).strip()
    if match := DAY_FIRST.fullmatch(text):
        day, month, year = (int(part) for part in match.groups())
        if year < 100:
            year += 2000
    elif match := ISO.fullmatch(text):
        year, month, day = (int(part) for part in match.groups())
    else:
        raise error
    try:
        return date(year, month, day)
    except ValueError:
        raise error from None


def format_brl(value: float) -> str:
    """1234.5 → "R$ 1.234,50"."""
    text = f'{abs(value):,.2f}'.replace(',', '_').replace('.', ',')
    sign = '-' if value < 0 else ''
    return f"{sign}R$ {text.replace('_', '.')}"
