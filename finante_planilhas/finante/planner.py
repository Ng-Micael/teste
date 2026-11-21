"""Planner de estudos: matérias com meta de horas, grade semanal e cronograma de provas."""
import datetime

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.datavalidation import DataValidation

from . import estilos as e

DIAS = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
HORA_INICIAL, HORA_FINAL = 6, 22   # grade das 06:00 às 22:00, uma célula = 1 hora
N_MATERIAS = 10
N_PROVAS = 20
SITUACOES = ["A estudar", "Estudando", "Revisado"]

M_PRIMEIRA = 7                          # primeira linha de matérias (aba Matérias)
M_ULTIMA = M_PRIMEIRA + N_MATERIAS - 1
S_PRIMEIRA = 7                          # primeira linha da grade (aba Semana)
S_ULTIMA = S_PRIMEIRA + (HORA_FINAL - HORA_INICIAL)
GRADE = f"Semana!$B${S_PRIMEIRA}:$H${S_ULTIMA}"
LISTA_MATERIAS = f"=Matérias!$A${M_PRIMEIRA}:$A${M_ULTIMA}"


def _lista(formula, titulo, erro, aviso=True):
    """Validação em lista. Com aviso=True a pessoa pode digitar algo fora da lista, se insistir."""
    v = DataValidation(type="list", formula1=formula, allow_blank=True,
                       showErrorMessage=True, errorTitle=titulo, error=erro)
    if aviso:
        v.errorStyle = "warning"
    return v


def _aba_materias(ws):
    e.larguras(ws, {"A": 30, "B": 20, "C": 18, "D": 14, "E": 22})
    e.titulo(ws, "FinanTe · Planner de estudos: matérias",
             "Passo 1: liste suas matérias e a meta de horas por semana (células de texto azul e fundo "
             "amarelo claro). Passo 2: distribua as matérias na aba Semana. Passo 3: anote as provas.")
    ws["A4"] = "Exemplo (não entra nas somas):"
    ws["A4"].font = e.FONTE_NOTA
    ws["A5"], ws["B5"], ws["C5"], ws["D5"] = "Matemática", 4, 3, "=MAX(0,B5-C5)"
    for letra in "ABCD":
        ws[f"{letra}5"].font = e.FONTE_EXEMPLO
    e.cabecalho(ws, 6, ["Matéria", "Meta (horas/semana)", "Planejado (horas)", "Falta (horas)", "Situação"],
                direita=(1, 2, 3))
    for linha in range(M_PRIMEIRA, M_ULTIMA + 1):
        e.marcar_entrada(ws.cell(row=linha, column=1))
        e.marcar_entrada(ws.cell(row=linha, column=2), "0")
        ws.cell(row=linha, column=3,
                value=f'=IF(A{linha}="","",COUNTIF({GRADE},A{linha}))')
        ws.cell(row=linha, column=4, value=f'=IF(OR(A{linha}="",B{linha}=""),"",MAX(0,B{linha}-C{linha}))')
        ws.cell(row=linha, column=5, value=(
            f'=IF(A{linha}="","",IF(B{linha}="","Defina a meta",'
            f'IF(C{linha}>=B{linha},"Meta cumprida","Faltam "&(B{linha}-C{linha})&"h")))'))
        for col in (3, 4, 5):
            c = ws.cell(row=linha, column=col)
            c.font, c.border = e.FONTE_BASE, e.BORDA
            c.number_format = "0"
    total = M_ULTIMA + 1
    ws.cell(row=total, column=2, value=f"=SUM(B{M_PRIMEIRA}:B{M_ULTIMA})")
    ws.cell(row=total, column=3, value=f"=SUM(C{M_PRIMEIRA}:C{M_ULTIMA})")
    ws.cell(row=total, column=4, value=f"=SUM(D{M_PRIMEIRA}:D{M_ULTIMA})")
    e.linha_total(ws, total, "Total da semana", "", 0, 0, "0")
    for letra in "BCD":
        ws[f"{letra}{total}"].number_format = "0"
    faixa = f"E{M_PRIMEIRA}:E{M_ULTIMA}"
    ws.conditional_formatting.add(faixa, FormulaRule(
        formula=[f'$E{M_PRIMEIRA}="Meta cumprida"'], fill=e.PREENCHIMENTO_OK,
        font=Font(name=e.FONTE, bold=True, color="0B5D43")))
    ws.conditional_formatting.add(faixa, FormulaRule(
        formula=[f'LEFT($E{M_PRIMEIRA},6)="Faltam"'], fill=e.PREENCHIMENTO_ALERTA,
        font=Font(name=e.FONTE, color="8A4B00")))
    ws.cell(row=total + 2, column=1,
            value="Dica: sessões de 25 a 50 minutos com pausas curtas funcionam bem. "
                  "As metas e o exemplo são apenas sugestões.").font = e.FONTE_NOTA


def _aba_semana(ws):
    e.larguras(ws, {"A": 12, **{letra: 15 for letra in "BCDEFGH"}})
    e.titulo(ws, "FinanTe · Planner de estudos: semana",
             "Cada célula é 1 hora. Escolha a matéria na lista (ou escreva outra atividade, como "
             "Descanso). Só as matérias da aba Matérias entram na contagem de metas.")
    e.cabecalho(ws, 6, ["Horário"] + DIAS)
    for k, hora in enumerate(range(HORA_INICIAL, HORA_FINAL + 1)):
        linha = S_PRIMEIRA + k
        c = ws.cell(row=linha, column=1, value=f"{hora:02d}:00")
        c.font, c.border = e.FONTE_NEGRITO, e.BORDA
        for col in range(2, 9):
            e.marcar_entrada(ws.cell(row=linha, column=col))
            ws.cell(row=linha, column=col).alignment = Alignment(horizontal="center")
    total = S_ULTIMA + 1
    ws.cell(row=total, column=1, value="Horas no dia")
    for col in range(2, 9):
        letra = chr(64 + col)
        ws.cell(row=total, column=col, value=f"=COUNTA({letra}{S_PRIMEIRA}:{letra}{S_ULTIMA})")
    e.linha_total(ws, total, "Horas no dia", "", 0, 0, "0")
    for col in range(2, 9):
        ws.cell(row=total, column=col).alignment = Alignment(horizontal="center")
    v = _lista(LISTA_MATERIAS, "Matéria", "Essa matéria não está na lista. Continuar mesmo assim?")
    ws.add_data_validation(v)
    v.add(f"B{S_PRIMEIRA}:H{S_ULTIMA}")
    ws.freeze_panes = "B7"


def _aba_provas(ws):
    e.larguras(ws, {"A": 16, "B": 24, "C": 40, "D": 16, "E": 18})
    e.titulo(ws, "FinanTe · Planner de estudos: provas",
             "Anote cada prova ou entrega. Os dias restantes são calculados a partir de hoje; "
             "provas nos próximos 7 dias ficam destacadas.")
    ws["A4"] = "Exemplo (não entra nas somas):"
    ws["A4"].font = e.FONTE_NOTA
    ws["A5"], ws["B5"], ws["C5"], ws["D5"], ws["E5"] = (
        datetime.date(2026, 11, 20), "Matemática", "Funções e gráficos", '=A5-TODAY()', "Estudando")
    ws["A5"].number_format = e.FORMATO_DATA
    ws["D5"].number_format = "0"
    for letra in "ABCDE":
        ws[f"{letra}5"].font = e.FONTE_EXEMPLO
    e.cabecalho(ws, 6, ["Data da prova", "Matéria", "Conteúdo", "Dias restantes", "Situação"], direita=(3,))
    primeira, ultima = 7, 6 + N_PROVAS
    for linha in range(primeira, ultima + 1):
        for col in (1, 2, 3, 5):
            e.marcar_entrada(ws.cell(row=linha, column=col))
        ws.cell(row=linha, column=1).number_format = e.FORMATO_DATA
        d = ws.cell(row=linha, column=4, value=f'=IF(A{linha}="","",A{linha}-TODAY())')
        d.font, d.border, d.number_format = e.FONTE_BASE, e.BORDA, "0"
    v1 = _lista(LISTA_MATERIAS, "Matéria", "Essa matéria não está na lista. Continuar mesmo assim?")
    v2 = _lista('"' + ",".join(SITUACOES) + '"', "Situação", "Escolha uma situação da lista.", aviso=False)
    for v, col in ((v1, "B"), (v2, "E")):
        ws.add_data_validation(v)
        v.add(f"{col}{primeira}:{col}{ultima}")
    faixa = f"D{primeira}:D{ultima}"
    ws.conditional_formatting.add(faixa, FormulaRule(
        formula=[f'AND(ISNUMBER($D{primeira}),$D{primeira}>=0,$D{primeira}<=7,$E{primeira}<>"Revisado")'],
        fill=e.PREENCHIMENTO_ALERTA, font=Font(name=e.FONTE, bold=True, color="8A4B00")))
    ws.conditional_formatting.add(faixa, FormulaRule(
        formula=[f'AND(ISNUMBER($D{primeira}),$D{primeira}<0)'],
        font=Font(name=e.FONTE, color="7A869E")))
    ws.freeze_panes = "A7"


def criar(caminho):
    wb = Workbook()
    materias = wb.active
    materias.title = "Matérias"
    semana = wb.create_sheet("Semana")
    provas = wb.create_sheet("Provas")
    _aba_materias(materias)
    _aba_semana(semana)
    _aba_provas(provas)
    e.configurar_pagina(materias)
    e.configurar_pagina(semana, paisagem=True)
    e.configurar_pagina(provas)
    wb.save(caminho)
