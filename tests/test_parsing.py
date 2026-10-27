from datetime import date, datetime

import pytest

from sales_report.parsing import (
    format_brl,
    is_empty,
    parse_date,
    parse_number,
)


@pytest.mark.parametrize('value, expected', [
    # Como o brasileiro digita
    ('1.234,56', 1234.56),
    ('1234,56', 1234.56),
    ('R$ 1.234,56', 1234.56),
    ('R$1.234,56', 1234.56),
    ('  12,5 ', 12.5),
    ('-10,00', -10.0),
    ('1.234.567', 1234567.0),
    # Ponto sozinho com 3 dígitos depois: milhar, como no Brasil
    ('1.234', 1234.0),
    # Como o sistema exporta (padrão americano)
    ('10.5', 10.5),
    ('1,234.56', 1234.56),
    ('3', 3.0),
    # O Excel já entrega número
    (12, 12.0),
    (12.75, 12.75),
])
def test_parse_number(value, expected):
    assert parse_number(value) == pytest.approx(expected)


@pytest.mark.parametrize('value', ['abc', '12,3,4', 'R$', '1.2.3,4,5', True])
def test_invalid_number(value):
    with pytest.raises(ValueError, match='não é um número'):
        parse_number(value)


@pytest.mark.parametrize('value, expected', [
    ('27/10/2026', date(2026, 10, 27)),
    ('7/3/2026', date(2026, 3, 7)),
    ('27/10/26', date(2026, 10, 27)),
    ('27-10-2026', date(2026, 10, 27)),
    ('27.10.2026', date(2026, 10, 27)),
    ('2026-10-27', date(2026, 10, 27)),
    ('27/10/2026 14:35', date(2026, 10, 27)),
    ('2026-10-27 14:35:00', date(2026, 10, 27)),
    # O Excel já entrega data
    (datetime(2026, 10, 27, 9, 30), date(2026, 10, 27)),
    (date(2026, 10, 27), date(2026, 10, 27)),
    # Célula de data sem formatação: o Excel mostra o número de série
    (46322, date(2026, 10, 27)),
    (46322.5, date(2026, 10, 27)),
])
def test_parse_date(value, expected):
    assert parse_date(value) == expected


@pytest.mark.parametrize('value', [
    '31/02/2026',     # não existe
    '10/27/2026',     # mês 27: formato americano não é aceito
    'ontem',
    '2026',
    12,               # número pequeno demais para ser data do Excel
])
def test_invalid_date(value):
    with pytest.raises(ValueError, match='não é uma data válida'):
        parse_date(value)


@pytest.mark.parametrize('value, expected', [
    (None, True),
    ('', True),
    ('   ', True),
    (float('nan'), True),
    ('x', False),
    (0, False),
])
def test_is_empty(value, expected):
    assert is_empty(value) is expected


@pytest.mark.parametrize('value, expected', [
    (1234.5, 'R$ 1.234,50'),
    (0, 'R$ 0,00'),
    (1234567.891, 'R$ 1.234.567,89'),
    (-5.5, '-R$ 5,50'),
])
def test_format_brl(value, expected):
    assert format_brl(value) == expected
