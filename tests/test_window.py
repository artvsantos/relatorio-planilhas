import gc
import sys
import threading
import time
import tkinter as tk
from pathlib import Path
from types import SimpleNamespace

import pytest

from sales_report import window as window_module
from sales_report.generate import ReportError
from sales_report.window import (
    TITLE,
    ReportWindow,
    default_report_name,
    enable_sharp_text,
    main,
    open_path,
    show_in_folder,
)
from tests.helpers import write_xlsx

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
    write_xlsx(folder / 'loja_norte.xlsx', [
        HEADER,
        ['27/10/2026', 'caneca', 1, 30],
        ['ontem', 'Prato', 1, 10],
    ])
    return folder


class Screen:
    """O que a janela faz fora dela: diálogos, avisos e arquivos abertos.
    Nos testes, nada disso aparece de verdade."""

    def __init__(self):
        self.opened = []
        self.shown_in_folder = []
        self.errors = []
        self.warnings = []
        self.folder_answer = ''
        self.save_answer = ''
        self.save_options = None

    def askdirectory(self, **options):
        return self.folder_answer

    def asksaveasfilename(self, **options):
        self.save_options = options
        return self.save_answer

    def showerror(self, title, message, **options):
        self.errors.append(message)

    def showwarning(self, title, message, **options):
        self.warnings.append(message)


@pytest.fixture
def screen(monkeypatch):
    screen = Screen()
    monkeypatch.setattr(window_module, 'open_path', screen.opened.append)
    monkeypatch.setattr(window_module, 'show_in_folder',
                        screen.shown_in_folder.append)
    for name in ('askdirectory', 'asksaveasfilename'):
        monkeypatch.setattr(window_module.filedialog, name,
                            getattr(screen, name))
    for name in ('showerror', 'showwarning'):
        monkeypatch.setattr(window_module.messagebox, name,
                            getattr(screen, name))
    return screen


@pytest.fixture(autouse=True)
def collect_old_windows():
    # Janelas de testes anteriores são apagadas aqui, na thread principal.
    # Se o coletor de lixo rodasse na thread do relatório, o Tcl seria
    # mexido fora da thread dele e a próxima janela poderia nem abrir.
    gc.collect()


@pytest.fixture
def app(screen):
    root = tk.Tk()
    root.withdraw()  # a janela funciona sem aparecer na tela
    yield ReportWindow(root)
    root.destroy()


def text(app):
    return app.log_box.get('1.0', 'end-1c')


def status(app):
    return app.status.cget('text')


def status_color(app):
    return str(app.status.cget('foreground'))


def enabled(button):
    return not button.instate(['disabled'])


def wait_until(condition, app, timeout=30):
    deadline = time.monotonic() + timeout
    while not condition():
        assert time.monotonic() < deadline, 'a janela não respondeu a tempo'
        app.root.update()
        time.sleep(0.01)


def generate(app, screen, output):
    screen.save_answer = str(output)
    app.generate_button.invoke()
    wait_until(lambda: not app.busy, app)


def test_starts_asking_for_the_folder(app):
    assert app.root.title() == TITLE
    assert status(app) == 'Escolha a pasta onde estão as planilhas das lojas.'
    assert not enabled(app.generate_button)
    assert not enabled(app.open_report_button)
    assert not enabled(app.show_button)


def test_choosing_the_folder_counts_the_spreadsheets(app, screen, folder):
    screen.folder_answer = str(folder)

    app.choose_button.invoke()

    assert app.folder == folder
    assert app.folder_var.get() == str(folder)
    assert status(app) == '2 planilhas encontradas.'
    assert enabled(app.generate_button)


def test_cancelling_the_folder_dialog_keeps_the_folder(app, screen, folder):
    app.set_folder(folder)
    screen.folder_answer = ''

    app.choose_button.invoke()

    assert app.folder == folder
    assert enabled(app.generate_button)


def test_folder_without_spreadsheets(app, tmp_path):
    (tmp_path / 'notas.txt').write_text('x')

    app.set_folder(tmp_path)

    assert status(app) == 'Nenhuma planilha (.xlsx ou .csv) nesta pasta.'
    assert not enabled(app.generate_button)


def test_missing_folder(app, tmp_path):
    missing = tmp_path / 'nao-existe'

    app.set_folder(missing)

    assert status(app) == f'Pasta não encontrada: {missing}'
    assert status_color(app) == window_module.RED
    assert not enabled(app.generate_button)


def test_dropped_spreadsheet_means_its_folder(app, folder):
    # Quem arrasta uma planilha para cima do programa quer a pasta dela.
    app.set_folder(folder / 'loja_centro.xlsx')

    assert app.folder == folder
    assert status(app) == '2 planilhas encontradas.'


def test_generates_the_report_and_opens_it(app, screen, folder, tmp_path):
    output = tmp_path / 'saida.xlsx'
    app.set_folder(folder)

    generate(app, screen, output)

    assert output.exists()
    assert screen.opened == [output]
    assert text(app) == '\n'.join([
        f'Lendo 2 planilhas em {folder}',
        '  loja_centro.xlsx: 2 vendas',
        '  loja_norte.xlsx: 1 venda, 1 linha com problema',
        '',
        f'Relatório salvo em {output}',
        '3 vendas, R$ 1.090,00 de faturamento.',
        '1 problema listado na aba "Problemas".',
    ])
    assert status(app) == 'Relatório pronto.'
    assert status_color(app) == window_module.GREEN
    assert enabled(app.open_report_button)
    assert enabled(app.show_button)
    assert enabled(app.generate_button)


def test_suggests_saving_next_to_the_folder(app, screen, folder):
    app.set_folder(folder)

    app.generate_button.invoke()

    assert screen.save_options['initialdir'] == folder.parent
    assert screen.save_options['initialfile'] == 'Relatório - vendas.xlsx'
    assert screen.save_options['defaultextension'] == '.xlsx'


def test_cancelling_the_save_dialog_generates_nothing(app, screen, folder):
    app.set_folder(folder)
    screen.save_answer = ''

    app.generate_button.invoke()

    assert not app.busy
    assert text(app) == ''
    assert screen.opened == []


def test_new_report_clears_the_previous_messages(app, screen, folder,
                                                 tmp_path):
    app.set_folder(folder)

    generate(app, screen, tmp_path / 'outubro.xlsx')
    generate(app, screen, tmp_path / 'novembro.xlsx')

    assert text(app).count('Lendo 2 planilhas') == 1
    assert screen.opened == [tmp_path / 'outubro.xlsx',
                             tmp_path / 'novembro.xlsx']


def test_window_keeps_working_while_the_report_is_made(app, screen, folder,
                                                       monkeypatch):
    # O relatório é feito em segundo plano: a janela não congela, mostra o
    # andamento e não deixa começar outro ao mesmo tempo.
    release = threading.Event()

    def slow_report(folder, output, log):
        log('Lendo 2 planilhas')
        release.wait(30)
        raise ReportError('parou')

    monkeypatch.setattr(window_module, 'generate_report', slow_report)
    app.set_folder(folder)
    screen.save_answer = str(folder.parent / 'r.xlsx')

    app.generate_button.invoke()
    wait_until(lambda: 'Lendo 2 planilhas' in text(app), app)

    assert app.busy
    assert status(app) == 'Gerando o relatório...'
    assert not enabled(app.choose_button)
    assert not enabled(app.generate_button)
    release.set()
    wait_until(lambda: not app.busy, app)
    assert enabled(app.choose_button)


def test_report_error_is_shown(app, screen, tmp_path):
    write_xlsx(tmp_path / 'ruim.xlsx', [HEADER, ['ontem', '', 0, 1]])
    app.set_folder(tmp_path)

    generate(app, screen, tmp_path / 'r.xlsx')

    message = 'Nenhuma venda válida encontrada. Veja os motivos acima.'
    assert screen.errors == [message]
    assert text(app).endswith('\n' + message)
    assert status(app) == 'O relatório não foi gerado.'
    assert status_color(app) == window_module.RED
    assert screen.opened == []
    assert not enabled(app.open_report_button)
    assert not enabled(app.show_button)
    assert enabled(app.generate_button)


def test_unexpected_error_is_shown_instead_of_hanging(app, screen, folder,
                                                      monkeypatch):
    def broken(folder, output, log):
        raise ValueError('deu ruim')

    monkeypatch.setattr(window_module, 'generate_report', broken)
    app.set_folder(folder)

    generate(app, screen, folder.parent / 'r.xlsx')

    assert screen.errors == ['Erro inesperado (ValueError): deu ruim']
    assert enabled(app.generate_button)


def test_open_report_and_show_in_folder(app, screen, folder, tmp_path):
    output = tmp_path / 'r.xlsx'
    app.set_folder(folder)
    generate(app, screen, output)
    screen.opened.clear()

    app.open_report_button.invoke()
    app.show_button.invoke()

    assert screen.opened == [output]
    assert screen.shown_in_folder == [output]


def test_warns_when_the_report_cannot_be_opened(app, screen, folder,
                                                tmp_path, monkeypatch):
    def no_excel(path):
        raise OSError('Nenhum aplicativo associado')

    monkeypatch.setattr(window_module, 'open_path', no_excel)
    output = tmp_path / 'r.xlsx'
    app.set_folder(folder)

    generate(app, screen, output)

    assert screen.warnings == [
        f'Não consegui abrir {output}.\n\nNenhum aplicativo associado',
    ]
    assert status(app) == 'Relatório pronto.'


def test_unexpected_error_in_a_button_is_shown(app, screen, folder,
                                               tmp_path, monkeypatch):
    # No .exe não há terminal: sem isso, o erro sumiria e o botão pareceria
    # não fazer nada.
    def broken(path):
        raise RuntimeError('quebrou')

    app.set_folder(folder)
    generate(app, screen, tmp_path / 'r.xlsx')
    monkeypatch.setattr(window_module, 'show_in_folder', broken)

    app.show_button.invoke()

    assert screen.errors == ['Erro inesperado (RuntimeError): quebrou']


def test_default_report_name():
    assert (default_report_name(Path('vendas', 'Outubro'))
            == 'Relatório - Outubro.xlsx')
    # A raiz de um pen drive (E:\) não tem nome.
    assert default_report_name(Path(Path.cwd().anchor)) == 'Relatório.xlsx'


def test_open_path_on_windows(monkeypatch):
    calls = []
    monkeypatch.setattr(sys, 'platform', 'win32')
    monkeypatch.setattr(window_module.os, 'startfile', calls.append,
                        raising=False)

    open_path(Path('r.xlsx'))

    assert calls == [Path('r.xlsx')]


@pytest.mark.parametrize(('platform', 'opener'), [
    ('darwin', 'open'),
    ('linux', 'xdg-open'),
])
def test_open_path_elsewhere(monkeypatch, platform, opener):
    calls = []
    monkeypatch.setattr(sys, 'platform', platform)
    monkeypatch.setattr(window_module.subprocess, 'Popen', calls.append)

    open_path(Path('r.xlsx'))

    assert calls == [[opener, 'r.xlsx']]


def test_show_in_folder_selects_the_file_on_windows(monkeypatch):
    calls = []
    monkeypatch.setattr(sys, 'platform', 'win32')
    monkeypatch.setattr(window_module.subprocess, 'Popen', calls.append)
    report = Path('relatorios', 'r.xlsx')

    show_in_folder(report)

    assert calls == [['explorer', '/select,', str(report)]]


def test_show_in_folder_elsewhere_opens_the_folder(monkeypatch):
    calls = []
    monkeypatch.setattr(sys, 'platform', 'linux')
    monkeypatch.setattr(window_module, 'open_path', calls.append)

    show_in_folder(Path('relatorios', 'r.xlsx'))

    assert calls == [Path('relatorios')]


def test_sharp_text_on_windows(monkeypatch):
    calls = []
    shcore = SimpleNamespace(SetProcessDpiAwareness=calls.append)
    monkeypatch.setattr(sys, 'platform', 'win32')
    monkeypatch.setattr(window_module.ctypes, 'windll',
                        SimpleNamespace(shcore=shcore), raising=False)

    enable_sharp_text()

    assert calls == [1]


def test_sharp_text_is_only_for_windows(monkeypatch):
    monkeypatch.setattr(sys, 'platform', 'linux')
    monkeypatch.delattr(window_module.ctypes, 'windll', raising=False)

    enable_sharp_text()  # sem o windll, que só existe no Windows


@pytest.fixture
def started(monkeypatch, screen):
    """Roda o main() sem prender o teste na janela aberta."""
    windows = []

    class Started(ReportWindow):
        def __init__(self, root):
            super().__init__(root)
            windows.append(self)

    monkeypatch.setattr(window_module, 'ReportWindow', Started)
    monkeypatch.setattr(window_module, 'enable_sharp_text', lambda: None)
    monkeypatch.setattr(tk.Tk, 'mainloop', lambda root, n=0: root.destroy())
    yield windows
    windows.clear()


def test_main_opens_the_window(started, monkeypatch):
    monkeypatch.setattr(sys, 'argv', ['RelatorioDeVendas.exe'])

    main()

    assert started[0].folder is None


def test_main_uses_the_folder_dropped_on_the_program(started, folder):
    main([str(folder)])

    assert started[0].folder == folder
