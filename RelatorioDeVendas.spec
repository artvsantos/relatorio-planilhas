# Receita do PyInstaller para o programa de Windows: um .exe só, sem
# terminal, que quem recebe abre com dois cliques.
#
#     pip install -r requirements-build.txt
#     pyinstaller RelatorioDeVendas.spec
#
# O programa sai em dist\RelatorioDeVendas.exe.

analysis = Analysis(
    ['app.py'],
    # O ícone da janela vai junto, dentro do .exe.
    datas=[('sales_report/icon.png', 'sales_report')],
)

exe = EXE(
    PYZ(analysis.pure),
    analysis.scripts,
    analysis.binaries,
    analysis.datas,
    name='RelatorioDeVendas',
    icon='sales_report/icon.ico',
    console=False,
    # Executável comprimido com UPX é mais barrado por antivírus.
    upx=False,
)
