# FinanTe · Gerador de planilhas

Código Python que gera as planilhas do FinanTe (.xlsx):

- `FinanTe-orcamento-mensal.xlsx`: receitas, despesas e meta de economia
- `FinanTe-controle-de-gastos.xlsx`: lançamentos do dia a dia e resumo por categoria
- `FinanTe-planner-de-estudos.xlsx`: matérias, grade semanal e cronograma de provas

## Como usar

```bash
pip install -r requirements.txt
python gerar.py            # cria os arquivos em ./saida
python gerar.py --pasta x  # cria em ./x
```

## Estrutura

```
gerar.py             ponto de entrada: gera todas as planilhas
finante/estilos.py   cores, fontes e funções compartilhadas
finante/orcamento.py orçamento mensal
finante/gastos.py    controle de gastos
finante/planner.py   planner de estudos
saida/               planilhas já geradas
```

Para mudar categorias, matérias ou horários, edite as listas no topo de cada módulo
(ex.: `CATEGORIAS` em `gastos.py`, `HORA_INICIAL` em `planner.py`) e rode `gerar.py` de novo.
