"""
PROJETO GENESIS - Repasse de Parceiros Digitais e Polos
=======================================================
Calcula, a partir do relatório "Receita Polo" do Pincel, o repasse de cada Parceiro Digital (Parceiro
Comercial) e de cada Polo, seguindo o Manual de Auditoria - Repasse e Premiação, e gera:

  * PLANILHA GERAL - <MÊS>-<ANO>.xlsx       a planilha de controle de sempre, com o Valor de cada parceiro;
  * RELATÓRIOS PARCEIROS DIGITAIS/*.xlsx    um relatório por parceiro, igual ao feito à mão;
  * RELATÓRIO DE PAGAMENTOS EM POLOS .../    um relatório por polo (layout de Açailândia);
  * TOTAL REPASSE - <MÊS>-<ANO>.xlsx        resumo, RPA (RH), conferências, pendências e memória de cálculo;
  * F_pagamentos.csv                        base linha a linha para o Power BI.

Todo mês:
  1. Pincel: Gestão > Receita Polo > Geral > Mês/Ano > Gerar Relatório.
  2. Salve o arquivo na pasta "entrada" como "relatorioReceitaPolo_convertido.xlsx".
  3. Rode:  python main.py      (primeira vez: pip install -r requirements.txt)
"""
import sys
import time

import config
from genesis.processo import gerar_saidas, imprimir_resumo, processar


def main():
    inicio = time.time()
    print("=" * 74 + "\n PROJETO GENESIS - Repasse de Parceiros Digitais e Polos\n" + "=" * 74)
    print(f"[Pasta] {config.PASTA_PROJETO}")
    r = processar()
    pasta, consolidado, planilha = gerar_saidas(r)
    imprimir_resumo(r)
    print(f"\n[Saída] {planilha}\n[Saída] {consolidado}\n[Saída] {pasta}")
    print(f"\nConcluído em {time.time() - inicio:.0f}s.")


if __name__ == "__main__":
    try:
        main()
    except (FileNotFoundError, ValueError, PermissionError) as erro:
        print(f"\nERRO: {erro}")
        if isinstance(erro, PermissionError):
            print("Feche no Excel os arquivos da pasta de saída e rode de novo.")
        sys.exit(1)
