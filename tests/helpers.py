"""Cria planilhas de teste no disco."""

import csv

from openpyxl import Workbook


def write_xlsx(path, rows, *, sheet_title='Vendas', extra_sheet=None):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = sheet_title
    for row in rows:
        sheet.append(list(row))
    if extra_sheet is not None:
        other = workbook.create_sheet('Outra')
        for row in extra_sheet:
            other.append(list(row))
    workbook.save(path)
    return path


def write_csv(path, rows, *, delimiter=';', encoding='cp1252'):
    # cp1252 e ponto e vírgula: é como o Excel em português salva CSV.
    with open(path, 'w', newline='', encoding=encoding) as file:
        csv.writer(file, delimiter=delimiter).writerows(rows)
    return path
