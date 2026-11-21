"""Estilos e funções auxiliares compartilhados por todas as planilhas do FinanTe."""
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

FONTE = "Arial"
AZUL_ESCURO = "14213D"
FORMATO_REAL = 'R$ #,##0.00;[Red]-R$ #,##0.00;"-"'
FORMATO_DATA = "dd/mm/yyyy"

FONTE_BASE = Font(name=FONTE, size=11)
FONTE_NEGRITO = Font(name=FONTE, size=11, bold=True)
FONTE_ENTRADA = Font(name=FONTE, size=11, color="0000FF")  # azul = célula para preencher
FONTE_TITULO = Font(name=FONTE, size=16, bold=True, color=AZUL_ESCURO)
FONTE_NOTA = Font(name=FONTE, size=10, italic=True, color="4A5873")
FONTE_EXEMPLO = Font(name=FONTE, size=10, italic=True, color="7A869E")
FONTE_CABECALHO = Font(name=FONTE, size=11, bold=True, color="FFFFFF")

PREENCHIMENTO_ENTRADA = PatternFill("solid", fgColor="FFF6CC")
PREENCHIMENTO_CABECALHO = PatternFill("solid", fgColor=AZUL_ESCURO)
PREENCHIMENTO_SUBTOTAL = PatternFill("solid", fgColor="E3EAF7")
PREENCHIMENTO_OK = PatternFill("solid", bgColor="C9F0DD")
PREENCHIMENTO_ALERTA = PatternFill("solid", bgColor="FFE0B8")

_LADO = Side(style="thin", color="B7C3DB")
BORDA = Border(left=_LADO, right=_LADO, top=_LADO, bottom=_LADO)


def titulo(ws, texto, legenda=None):
    """Escreve o título em A1 e, opcionalmente, a legenda de uso em A2."""
    ws["A1"] = texto
    ws["A1"].font = FONTE_TITULO
    if legenda:
        ws["A2"] = legenda
        ws["A2"].font = FONTE_NOTA


def larguras(ws, colunas):
    """Define a largura das colunas, ex.: {"A": 30, "B": 18}."""
    for letra, largura in colunas.items():
        ws.column_dimensions[letra].width = largura


def cabecalho(ws, linha, titulos, direita=()):
    """Linha de cabeçalho azul-escura; `direita` lista os índices (0, 1...) alinhados à direita."""
    for i, texto in enumerate(titulos):
        c = ws.cell(row=linha, column=i + 1, value=texto)
        c.font = FONTE_CABECALHO
        c.fill = PREENCHIMENTO_CABECALHO
        c.alignment = Alignment(horizontal="right" if i in direita else "left")


def marcar_entrada(celula, formato=None):
    """Aplica o visual de 'preencha aqui': texto azul e fundo amarelo claro."""
    celula.font = FONTE_ENTRADA
    celula.fill = PREENCHIMENTO_ENTRADA
    celula.border = BORDA
    if formato:
        celula.number_format = formato


def linha_total(ws, linha, rotulo, colunas_soma, primeira, ultima, formato):
    """Linha de total com SUM nas colunas indicadas (ex.: 'BCD')."""
    ws.cell(row=linha, column=1, value=rotulo)
    for letra in colunas_soma:
        c = ws[f"{letra}{linha}"]
        c.value = f"=SUM({letra}{primeira}:{letra}{ultima})"
        c.number_format = formato
    for col in range(1, ws.max_column + 1):
        c = ws.cell(row=linha, column=col)
        c.font = FONTE_NEGRITO
        c.fill = PREENCHIMENTO_SUBTOTAL
        c.border = BORDA


def configurar_pagina(ws, paisagem=False):
    """Esconde as linhas de grade e ajusta a impressão para caber na largura da página."""
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape" if paisagem else "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
