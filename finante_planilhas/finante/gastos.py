"""Planilha de controle de gastos: lançamentos do dia a dia + resumo por categoria."""
import datetime

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.datavalidation import DataValidation

from . import estilos as e

CATEGORIAS = ["Moradia e contas", "Alimentação", "Transporte", "Estudos", "Lazer e vida social",
              "Saúde", "Dívidas e parcelas", "Outras despesas"]
PAGAMENTOS = ["Dinheiro", "Pix", "Débito", "Crédito", "Outro"]
PRIMEIRA, ULTIMA = 7, 206  # linhas da aba Gastos disponíveis para lançamentos (200 compras)
PRIMEIRA_CAT = 6           # primeira linha de categorias na aba Resumo


def _faixa(letra):
    """Intervalo absoluto de uma coluna da aba Gastos, ex.: Gastos!$E$7:$E$206."""
    return f"Gastos!${letra}${PRIMEIRA}:${letra}${ULTIMA}"


def _aba_gastos(ws):
    e.larguras(ws, {"A": 14, "B": 38, "C": 24, "D": 20, "E": 18})
    e.titulo(ws, "FinanTe · Controle de gastos",
             "Como usar: anote cada compra nas linhas abaixo (células de texto azul e fundo amarelo "
             "claro). Escolha a categoria e a forma de pagamento na lista. O resumo fica na aba Resumo.")
    ws["A4"] = "Exemplo (não entra nas somas):"
    ws["A4"].font = e.FONTE_NOTA
    exemplo = [datetime.date(2026, 11, 4), "Lanche na faculdade", "Alimentação", "Pix", 18.5]
    for i, valor in enumerate(exemplo):
        ws.cell(row=5, column=i + 1, value=valor).font = e.FONTE_EXEMPLO
    ws["A5"].number_format, ws["E5"].number_format = e.FORMATO_DATA, e.FORMATO_REAL
    e.cabecalho(ws, 6, ["Data", "Descrição", "Categoria", "Forma de pagamento", "Valor (R$)"], direita=(4,))

    for linha in range(PRIMEIRA, ULTIMA + 1):
        for col in range(1, 6):
            e.marcar_entrada(ws.cell(row=linha, column=col))
        ws.cell(row=linha, column=1).number_format = e.FORMATO_DATA
        ws.cell(row=linha, column=5).number_format = e.FORMATO_REAL

    validacoes = [
        (f"C{PRIMEIRA}:C{ULTIMA}", DataValidation(
            type="list", formula1=f"=Resumo!$A${PRIMEIRA_CAT}:$A${PRIMEIRA_CAT + len(CATEGORIAS) - 1}",
            allow_blank=True, showErrorMessage=True, errorTitle="Categoria",
            error="Escolha uma categoria da lista.")),
        (f"D{PRIMEIRA}:D{ULTIMA}", DataValidation(
            type="list", formula1='"' + ",".join(PAGAMENTOS) + '"', allow_blank=True,
            showErrorMessage=True, errorTitle="Pagamento", error="Escolha uma forma da lista.")),
        (f"E{PRIMEIRA}:E{ULTIMA}", DataValidation(
            type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
            showErrorMessage=True, errorTitle="Valor", error="Digite um valor numérico, sem R$.")),
    ]
    for faixa, validacao in validacoes:
        ws.add_data_validation(validacao)
        validacao.add(faixa)
    ws.freeze_panes = "A7"


def _aba_resumo(ws):
    e.larguras(ws, {"A": 34, "B": 20, "C": 18, "D": 14, "E": 18, "F": 22})
    e.titulo(ws, "FinanTe · Resumo dos gastos",
             "Preencha só o limite de cada categoria (opcional). O gasto vem da aba Gastos.")
    ws["A3"] = "Uma planilha por mês: no mês seguinte, salve uma cópia e apague as linhas da aba Gastos."
    ws["A3"].font = e.FONTE_NOTA
    e.cabecalho(ws, 5, ["Categoria", "Limite do mês (R$)", "Gasto (R$)", "% do total",
                        "Restante (R$)", "Situação"], direita=(1, 2, 3, 4))
    ultima_cat = PRIMEIRA_CAT + len(CATEGORIAS) - 1
    total = ultima_cat + 1
    for k, categoria in enumerate(CATEGORIAS):
        linha = PRIMEIRA_CAT + k
        ws.cell(row=linha, column=1, value=categoria)
        ws.cell(row=linha, column=3, value=f"=SUMIFS({_faixa('E')},{_faixa('C')},A{linha})")
        ws.cell(row=linha, column=4, value=f"=IF($C${total}=0,0,C{linha}/$C${total})")
        ws.cell(row=linha, column=5, value=f'=IF(B{linha}="","",B{linha}-C{linha})')
        ws.cell(row=linha, column=6,
                value=f'=IF(B{linha}="","",IF(C{linha}>B{linha},"Passou do limite","Dentro do limite"))')
        for col in range(1, 7):
            c = ws.cell(row=linha, column=col)
            c.border = e.BORDA
            if col == 2:
                e.marcar_entrada(c, e.FORMATO_REAL)
            else:
                c.font = e.FONTE_BASE
                c.number_format = "0%" if col == 4 else e.FORMATO_REAL
        ws.cell(row=linha, column=6).alignment = Alignment(horizontal="left")

    ws.cell(row=total, column=2, value=f"=SUM(B{PRIMEIRA_CAT}:B{ultima_cat})")
    ws.cell(row=total, column=3, value=f"=SUM({_faixa('E')})")
    ws.cell(row=total, column=4, value=f"=IF(C{total}=0,0,SUM(D{PRIMEIRA_CAT}:D{ultima_cat}))")
    e.linha_total(ws, total, "Total gasto", "", 0, 0, e.FORMATO_REAL)
    for letra in "BC":
        ws[f"{letra}{total}"].number_format = e.FORMATO_REAL
    ws[f"D{total}"].number_format = "0%"
    ws.cell(row=total + 1, column=1, value="Sem categoria (confira a aba Gastos)").font = e.FONTE_NOTA
    c = ws.cell(row=total + 1, column=3, value=f"=C{total}-SUM(C{PRIMEIRA_CAT}:C{ultima_cat})")
    c.number_format, c.font = e.FORMATO_REAL, e.FONTE_NOTA

    faixa_status = f"F{PRIMEIRA_CAT}:F{ultima_cat}"
    ws.conditional_formatting.add(faixa_status, FormulaRule(
        formula=[f'$F{PRIMEIRA_CAT}="Passou do limite"'], fill=e.PREENCHIMENTO_ALERTA,
        font=Font(name=e.FONTE, bold=True, color="8A4B00")))
    ws.conditional_formatting.add(faixa_status, FormulaRule(
        formula=[f'$F{PRIMEIRA_CAT}="Dentro do limite"'], fill=e.PREENCHIMENTO_OK,
        font=Font(name=e.FONTE, color="0B5D43")))

    cab = total + 4
    e.cabecalho(ws, cab, ["Forma de pagamento", "", "Gasto (R$)", "% do total"], direita=(2, 3))
    for k, forma in enumerate(PAGAMENTOS):
        linha = cab + 1 + k
        ws.cell(row=linha, column=1, value=forma)
        ws.cell(row=linha, column=3, value=f"=SUMIFS({_faixa('E')},{_faixa('D')},A{linha})")
        ws.cell(row=linha, column=4, value=f"=IF($C${total}=0,0,C{linha}/$C${total})")
        for col in range(1, 5):
            c = ws.cell(row=linha, column=col)
            c.font, c.border = e.FONTE_BASE, e.BORDA
            if col == 3:
                c.number_format = e.FORMATO_REAL
            elif col == 4:
                c.number_format = "0%"
    fim = cab + len(PAGAMENTOS) + 2
    ws.cell(row=fim, column=1,
            value="Atenção ao crédito: parcelar vira dívida. Veja quanto do total foi no cartão.").font = e.FONTE_NOTA
    ws.cell(row=fim + 1, column=1,
            value="Os limites e exemplos são apenas sugestões, não recomendação financeira.").font = e.FONTE_NOTA


def criar(caminho):
    wb = Workbook()
    gastos = wb.active
    gastos.title = "Gastos"
    resumo = wb.create_sheet("Resumo")
    _aba_gastos(gastos)
    _aba_resumo(resumo)
    e.configurar_pagina(gastos, paisagem=True)
    e.configurar_pagina(resumo)
    wb.save(caminho)
