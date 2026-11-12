"""Janela para quem não usa a linha de comando: escolhe a pasta das
planilhas, clica em "Gerar relatório" e o relatório abre no Excel.

É esta janela que vira o programa para Windows (RelatorioDeVendas.exe).
"""

import ctypes
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from collections.abc import Callable
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from sales_report.generate import ReportError, generate_report, plural
from sales_report.reader import list_spreadsheets

TITLE = 'Relatório de vendas'
ICON = Path(__file__).with_name('icon.png')
# De quanto em quanto tempo (ms) a janela olha o andamento do relatório,
# que é feito em segundo plano.
POLL_MS = 50
MUTED = '#4B5563'
RED = '#B91C1C'
GREEN = '#15803D'


def open_path(path: Path) -> None:
    """Abre o arquivo no programa padrão: o relatório, no Excel."""
    if sys.platform == 'win32':
        os.startfile(path)
    else:
        opener = 'open' if sys.platform == 'darwin' else 'xdg-open'
        subprocess.Popen([opener, str(path)])


def show_in_folder(path: Path) -> None:
    """Abre a pasta do arquivo; no Windows, com ele já selecionado."""
    if sys.platform == 'win32':
        subprocess.Popen(['explorer', '/select,', str(path)])
    else:
        open_path(path.parent)


def default_report_name(folder: Path) -> str:
    # "Vendas\Outubro" vira "Relatório - Outubro.xlsx". A raiz de um pen
    # drive (E:\) não tem nome.
    if not folder.name:
        return 'Relatório.xlsx'
    return f'Relatório - {folder.name}.xlsx'


def enable_sharp_text() -> None:
    # Sem isso, com a escala de tela do Windows em 125% ou 150%, a janela é
    # esticada como uma imagem e o texto fica borrado.
    if sys.platform == 'win32':
        ctypes.windll.shcore.SetProcessDpiAwareness(1)


def _unexpected(error: BaseException) -> str:
    return f'Erro inesperado ({type(error).__name__}): {error}'


def _show_unexpected_error(kind, error, traceback) -> None:
    # No .exe não há terminal: um erro inesperado num botão vira um aviso
    # na tela, em vez de sumir sem explicação.
    messagebox.showerror(TITLE, _unexpected(error))


def _enable(button: ttk.Button, enabled: bool) -> None:
    button.state(['!disabled'] if enabled else ['disabled'])


class ReportWindow:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.folder: Path | None = None
        self.output: Path | None = None
        self.busy = False
        # A thread que gera o relatório não pode mexer na janela (o tkinter
        # só aceita a thread principal), então manda o andamento por aqui.
        self._progress: queue.Queue = queue.Queue()

        root.title(TITLE)
        self._icon = tk.PhotoImage(file=ICON)
        root.iconphoto(True, self._icon)
        # Uma função solta, e não um método: assim a janela não fica presa
        # num ciclo de referências com o root.
        root.report_callback_exception = _show_unexpected_error
        self._build()
        self._set_status('Escolha a pasta onde estão as planilhas das lojas.')

    def _build(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        frame = ttk.Frame(self.root, padding=(20, 16))
        frame.grid(sticky='nsew')
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(6, weight=1)
        ttk.Style(self.root).configure('Main.TButton', padding=(18, 6))

        ttk.Label(frame, text=TITLE, font=('Segoe UI', 15, 'bold')).grid(
            row=0, column=0, columnspan=2, sticky='w')
        ttk.Label(frame, foreground=MUTED, wraplength=540, text=(
            'Junta as planilhas de vendas das lojas (.xlsx e .csv) num '
            'relatório em Excel. As linhas com erro vão para uma aba à '
            'parte, com o motivo, para corrigir na origem.'
        )).grid(row=1, column=0, columnspan=2, sticky='w', pady=(2, 16))

        ttk.Label(frame, text='Pasta com as planilhas').grid(
            row=2, column=0, columnspan=2, sticky='w', pady=(0, 4))
        self.folder_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.folder_var, state='readonly',
                  width=60).grid(row=3, column=0, sticky='ew', padx=(0, 8))
        self.choose_button = ttk.Button(frame, text='Escolher pasta...',
                                        command=self.choose_folder)
        self.choose_button.grid(row=3, column=1)
        self.status = ttk.Label(frame)
        self.status.grid(row=4, column=0, columnspan=2, sticky='w',
                         pady=(4, 12))

        self.generate_button = ttk.Button(
            frame, text='Gerar relatório', style='Main.TButton',
            command=self.generate, state='disabled',
        )
        self.generate_button.grid(row=5, column=0, columnspan=2, sticky='w')

        box = ttk.Frame(frame)
        box.grid(row=6, column=0, columnspan=2, sticky='nsew', pady=(14, 10))
        box.columnconfigure(0, weight=1)
        box.rowconfigure(0, weight=1)
        self.log_box = tk.Text(box, height=11, width=70, wrap='word',
                               relief='solid', borderwidth=1, padx=10,
                               pady=8, font='TkDefaultFont',
                               state='disabled')
        self.log_box.grid(row=0, column=0, sticky='nsew')
        self.log_box.tag_configure('error', foreground=RED)
        scrollbar = ttk.Scrollbar(box, command=self.log_box.yview)
        scrollbar.grid(row=0, column=1, sticky='ns')
        self.log_box.configure(yscrollcommand=scrollbar.set)

        actions = ttk.Frame(frame)
        actions.grid(row=7, column=0, columnspan=2, sticky='e')
        self.open_report_button = ttk.Button(
            actions, text='Abrir relatório', command=self.open_report,
            state='disabled',
        )
        self.open_report_button.grid(row=0, column=0, padx=(0, 8))
        self.show_button = ttk.Button(
            actions, text='Mostrar na pasta', command=self.show_report,
            state='disabled',
        )
        self.show_button.grid(row=0, column=1)

    def choose_folder(self) -> None:
        chosen = filedialog.askdirectory(
            parent=self.root, mustexist=True, initialdir=self.folder,
            title='Escolha a pasta com as planilhas de vendas',
        )
        if chosen:
            self.set_folder(Path(chosen))

    def set_folder(self, folder: Path) -> None:
        if folder.is_file():
            # Arrastou uma planilha para cima do programa, em vez da pasta.
            folder = folder.parent
        self.folder = folder
        self.folder_var.set(str(folder))
        found = list_spreadsheets(folder) if folder.is_dir() else None
        if found is None:
            self._set_status(f'Pasta não encontrada: {folder}', RED)
        elif not found:
            self._set_status('Nenhuma planilha (.xlsx ou .csv) nesta pasta.',
                             RED)
        else:
            self._set_status(plural(len(found), 'planilha encontrada.',
                                    'planilhas encontradas.'))
        _enable(self.generate_button, bool(found))

    def generate(self) -> None:
        chosen = filedialog.asksaveasfilename(
            parent=self.root, title='Salvar o relatório como',
            initialdir=self.folder.parent,
            initialfile=default_report_name(self.folder),
            defaultextension='.xlsx',
            filetypes=[('Planilha do Excel', '*.xlsx')],
        )
        if not chosen:
            return
        self.output = Path(chosen)
        self.busy = True
        self.root.configure(cursor='watch')
        for button in (self.choose_button, self.generate_button,
                       self.open_report_button, self.show_button):
            _enable(button, False)
        self._set_status('Gerando o relatório...')
        self.log_box.configure(state='normal')
        self.log_box.delete('1.0', 'end')
        self.log_box.configure(state='disabled')
        threading.Thread(target=self._make_report, daemon=True,
                         args=(self.folder, self.output)).start()
        self.root.after(POLL_MS, self._check_progress)

    def _make_report(self, folder: Path, output: Path) -> None:
        # Roda em segundo plano, para a janela não congelar.
        def log(line: str) -> None:
            self._progress.put(('line', line))

        try:
            summary = generate_report(folder, output, log)
        except ReportError as error:
            self._progress.put(('error', str(error)))
        except Exception as error:
            # Melhor avisar do que deixar a janela esperando para sempre.
            self._progress.put(('error', _unexpected(error)))
        else:
            self._progress.put(('done', summary))

    def _check_progress(self) -> None:
        while not self._progress.empty():
            kind, value = self._progress.get()
            if kind == 'line':
                self._write(value)
            elif kind == 'error':
                self._finish(ok=False)
                self._write(value, 'error')
                messagebox.showerror(TITLE, value, parent=self.root)
            else:
                self._finish(ok=True)
                for line in ['', *value.lines()]:
                    self._write(line)
                self.open_report()
        if self.busy:
            self.root.after(POLL_MS, self._check_progress)

    def _finish(self, ok: bool) -> None:
        self.busy = False
        self.root.configure(cursor='')
        if ok:
            self._set_status('Relatório pronto.', GREEN)
        else:
            self._set_status('O relatório não foi gerado.', RED)
        _enable(self.choose_button, True)
        _enable(self.generate_button, True)
        _enable(self.open_report_button, ok)
        _enable(self.show_button, ok)

    def open_report(self) -> None:
        self._open_with(open_path)

    def show_report(self) -> None:
        self._open_with(show_in_folder)

    def _open_with(self, opener: Callable[[Path], None]) -> None:
        try:
            opener(self.output)
        except OSError as error:
            messagebox.showwarning(
                TITLE, f'Não consegui abrir {self.output}.\n\n{error}',
                parent=self.root,
            )

    def _write(self, line: str, tag: str = 'info') -> None:
        self.log_box.configure(state='normal')
        if self.log_box.get('1.0', 'end-1c'):
            self.log_box.insert('end', '\n')
        self.log_box.insert('end', line, tag)
        self.log_box.configure(state='disabled')
        self.log_box.see('end')

    def _set_status(self, text: str, color: str = MUTED) -> None:
        self.status.configure(text=text, foreground=color)


def main(argv: list[str] | None = None) -> None:
    args = sys.argv[1:] if argv is None else argv
    enable_sharp_text()
    root = tk.Tk()
    app = ReportWindow(root)
    if args:
        # A pasta (ou planilha) arrastada para cima do ícone do programa.
        app.set_folder(Path(args[0]))
    root.mainloop()
