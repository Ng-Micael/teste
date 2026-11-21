"""Planilha de orçamento mensal: receitas, despesas e meta de economia."""
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font

from . import estilos as e

RECEITAS = ["Mesada / ajuda da família", "Estágio ou trabalho", "Bolsa ou auxílio", "Outras receitas"]
DESPESAS = [
    "Moradia e contas (aluguel, luz, internet, celular)", "Alimentação", "Transporte",
    "Estudos (curso, material, mensalidade)", "Lazer e vida social", "Saúde",
    "Dívidas e parcelas", "Outras despesas",
]
META_PADRAO = 0.10  # sugestão: guardar 10% da receita


def _itens(ws, inicio, rotulos, receita):
    """Escreve as linhas de itens e devolve o número da linha seguinte ao último item."""
    for k, rotulo in enumerate(rotulos):
        linha = inicio + k
        ws.cell(row=linha, column=1, value=rotulo).font = e.FONTE_BASE
        ws.cell(row=linha, column=1).border = e.BORDA
        for col in (2, 3):
            e.marcar_entrada(ws.cell(row=linha, column=col), e.FORMATO_REAL)
        formula = f"=C{linha}-B{linha}" if receita else f"=B{linha}-C{linha}"
        dif = ws.cell(row=linha, column=4, value=formula)
        dif.font, dif.border, dif.number_format = e.FONTE_BASE, e.BORDA, e.FORMATO_REAL
    return inicio + len(rotulos)


def criar(caminho):
    wb = Workbook()
    ws = wb.active
    ws.title = "Orçamento mensal"
    e.larguras(ws, {"A": 44, "B": 18, "C": 18, "D": 18})
    e.titulo(ws, "FinanTe · Orçamento mensal",
             "Como usar: preencha só as células com texto azul e fundo amarelo claro. "
             "O restante calcula sozinho.")

    ws["A4"], ws["A5"] = "Mês de referência", "Meta de economia (% da receita)"
    ws["A4"].font = ws["A5"].font = e.FONTE_NEGRITO
    ws["B4"], ws["B5"] = "Novembro/2026", META_PADRAO
    e.marcar_entrada(ws["B4"])
    e.marcar_entrada(ws["B5"], "0%")
    ws["B4"].alignment = ws["B5"].alignment = Alignment(horizontal="right")
    ws["C5"] = "Sugestão: 10%. Mude para o que fizer sentido para você."
    ws["C5"].font = Font(name=e.FONTE, size=9, italic=True, color="4A5873")

    # Linha de exemplo (não entra em nenhuma soma)
    ws["A7"], ws["B7"], ws["C7"], ws["D7"] = (
        "Exemplo de preenchimento (não entra nas somas): Estágio", 900, 850, "=C7-B7")
    for letra in "ABCD":
        ws[f"{letra}7"].font = e.FONTE_EXEMPLO
    for letra in "BCD":
        ws[f"{letra}7"].number_format = e.FORMATO_REAL

    e.cabecalho(ws, 9, ["RECEITAS (o que entra)", "Previsto (R$)", "Real (R$)",
                        "Diferença (Real − Previsto)"], direita=(1, 2, 3))
    fim = _itens(ws, 10, RECEITAS, receita=True)
    e.linha_total(ws, fim, "Total de receitas", "BCD", 10, fim - 1, e.FORMATO_REAL)
    total_rec = fim

    inicio_desp = total_rec + 3
    e.cabecalho(ws, inicio_desp - 1, ["DESPESAS (o que sai)", "Previsto (R$)", "Real (R$)",
                                      "Diferença (Previsto − Real)"], direita=(1, 2, 3))
    fim = _itens(ws, inicio_desp, DESPESAS, receita=False)
    e.linha_total(ws, fim, "Total de despesas", "BCD", inicio_desp, fim - 1, e.FORMATO_REAL)
    total_desp = fim

    r0 = total_desp + 2
    e.cabecalho(ws, r0, ["RESULTADO DO MÊS", "Previsto", "Real", ""], direita=(1, 2, 3))
    sobra, meta, situacao = r0 + 1, r0 + 2, r0 + 3
    ws[f"A{sobra}"] = "Sobra do mês (receitas − despesas)"
    ws[f"A{meta}"] = "Meta de economia (receitas × % da meta)"
    ws[f"A{situacao}"] = "Situação"
    for letra in "BC":
        ws[f"{letra}{sobra}"] = f"={letra}{total_rec}-{letra}{total_desp}"
        ws[f"{letra}{meta}"] = f"={letra}{total_rec}*$B$5"
    ws[f"C{situacao}"] = (f'=IF(C{total_rec}=0,"Preencha as receitas",'
                          f'IF(C{sobra}>=C{meta},"Meta alcançada","Abaixo da meta"))')
    for linha in (sobra, meta, situacao):
        for col in range(1, 5):
            c = ws.cell(row=linha, column=col)
            c.font = e.FONTE_BASE if linha == meta else e.FONTE_NEGRITO
            c.border = e.BORDA
        for letra in "BC":
            ws[f"{letra}{linha}"].number_format = e.FORMATO_REAL
    ws[f"C{situacao}"].alignment = Alignment(horizontal="right")
    ws.conditional_formatting.add(f"C{situacao}", FormulaRule(
        formula=[f'$C${situacao}="Meta alcançada"'], fill=e.PREENCHIMENTO_OK,
        font=Font(name=e.FONTE, bold=True, color="0B5D43")))
    ws.conditional_formatting.add(f"C{situacao}", FormulaRule(
        formula=[f'$C${situacao}="Abaixo da meta"'], fill=e.PREENCHIMENTO_ALERTA,
        font=Font(name=e.FONTE, bold=True, color="8A4B00")))

    notas = [
        "Como ler a coluna Diferença:",
        "• Receitas: positivo = você recebeu mais do que previu. Negativo = recebeu menos.",
        "• Despesas: positivo = você gastou menos do que previu (economizou). Negativo = gastou mais.",
        "Dica: preencha o Previsto no começo do mês e o Real no final. Compare e ajuste o mês seguinte.",
        "Os valores de exemplo e a meta de 10% são apenas sugestões, não recomendação financeira.",
    ]
    for i, texto in enumerate(notas):
        c = ws.cell(row=situacao + 2 + i, column=1, value=texto)
        c.font = Font(name=e.FONTE, size=10, bold=(i == 0), color="4A5873")

    e.configurar_pagina(ws)
    wb.save(caminho)
