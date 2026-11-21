"""Gera todas as planilhas do FinanTe.

Uso:
    python gerar.py                 # cria os arquivos na pasta ./saida
    python gerar.py --pasta meus    # cria na pasta ./meus
"""
import argparse
from pathlib import Path

from finante import gastos, orcamento, planner

PLANILHAS = {
    "FinanTe-orcamento-mensal.xlsx": orcamento.criar,
    "FinanTe-controle-de-gastos.xlsx": gastos.criar,
    "FinanTe-planner-de-estudos.xlsx": planner.criar,
}


def main():
    parser = argparse.ArgumentParser(description="Gera as planilhas do FinanTe.")
    parser.add_argument("--pasta", default="saida", help="pasta de destino (padrão: saida)")
    pasta = Path(parser.parse_args().pasta)
    pasta.mkdir(parents=True, exist_ok=True)
    for nome, criar in PLANILHAS.items():
        caminho = pasta / nome
        criar(caminho)
        print(f"Criada: {caminho}")


if __name__ == "__main__":
    main()
