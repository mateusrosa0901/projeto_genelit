"""
CONFIGURAÇÕES DO PROJETO GENESIS
================================
Tudo o que costuma mudar fica aqui ou nos arquivos da pasta "dados":

  dados/D_Comissao.csv            percentuais por vigência (data de matrícula), modalidade e tipo de parcela
  dados/D_Excecoes.csv            exceções do manual (Wilhan, Ulisses, Altivo, Danubia, Marcieni, Integral)
  dados/D_Polos.csv               cadastro de polos e responsáveis
  dados/PLANILHA GERAL.xlsx       cadastro/controle dos Parceiros Digitais (a mesma planilha usada à mão)
  dados/D_Parceiros_Ajustes.csv   apelidos no Pincel, data de início da parceria e parceiros fora da planilha

Os textos comparados (situações, palavras etc.) são conferidos sem acento e sem diferença
de maiúsculas/minúsculas.
"""
from pathlib import Path

# ===================================== PASTAS ======================================
PASTA_PROJETO = Path(__file__).resolve().parent
PASTA_DADOS = PASTA_PROJETO / "dados"
PASTA_ENTRADA = PASTA_PROJETO / "entrada"            # relatório do Pincel do mês
PASTA_REFERENCIAS = PASTA_PROJETO / "referencias"    # opcional: planilhas calculadas à mão (só conferência)
PASTA_SAIDA = PASTA_PROJETO / "saida"                # resultado em saida/<ANO>/<MM>. REF. <MÊS>/

# ===================================== ARQUIVOS =====================================
# Pincel: Gestão > Receita Polo > Geral > Mês/Ano > Gerar Relatório. Se este nome não existir na
# pasta "entrada", o programa usa o arquivo "relatorioReceitaPolo*" mais recente dessa pasta.
ARQ_RECEITA = "relatorioReceitaPolo_convertido.xlsx"
ARQ_POLOS = "D_Polos.csv"
ARQ_COMISSAO = "D_Comissao.csv"
ARQ_EXCECOES = "D_Excecoes.csv"
ARQ_PLANILHA_GERAL = "PLANILHA GERAL.xlsx"
ARQ_AJUSTES_PARCEIROS = "D_Parceiros_Ajustes.csv"

# ================================ PARCEIROS DIGITAIS ================================
# Quem é Parceiro Digital: quem está na PLANILHA GERAL (ou em D_Parceiros_Ajustes.csv).
# Coluna "PF / PJ" e coluna de observação (sem título, logo após a tabela de códigos) decidem
# como cada parceiro recebe:
TIPOS_RPA = ["RPA"]                                   # PF / PJ = RPA -> pago pelo RH (aba "RPA (RH)")
SITUACOES_RPA = ["RETORNOU COMO PARCEIRO"]
TIPOS_NAO_PAGAR = ["POLO"]                            # PF / PJ = POLO -> recebe como polo, não como parceiro
SITUACOES_NAO_PAGAR = ["POLO", "DISTRATO", "NAO ENVIAR REPASSE", "INATIVO", "NAO PAGAR"]   # "começa com"
COLUNA_OBSERVACAO_PLANILHA = "Q"                      # coluna da PLANILHA GERAL com DISTRATO, POLO, RPA...
RPA_NA_PLANILHA_GERAL = False      # False = como no manual: RPA fica zerado na PLANILHA GERAL (vai para o RH)
INSS_PATRONAL_RPA = 0.20           # manual, pág. 10: desconto de 20% de INSS patronal na planilha do RH

# Regras confirmadas nos relatórios manuais de SET/2026 (2.623 linhas conferidas):
PARCELAS_SEM_REPASSE_PARCEIRO = [1]   # parcela 1 não gera repasse ao parceiro (MAT-R já é 0% no D_Comissao)
PARCELA_PRIMEIRA_MENSALIDADE = 2      # "1ª mensalidade" das exceções = PARCELA 2 (a parcela 1 é a matrícula)
EXCLUIR_VENDA_NO_PROPRIO_POLO = True  # venda do parceiro no polo ATIVO do qual é responsável = repasse do polo

# Cruzamento do "NOME PARCEIRO" do Pincel com a PLANILHA GERAL (nome igual, CPF/CNPJ, nome contido e,
# se ligado, nome muito parecido). Todo cruzamento que não for "nome igual" aparece na aba
# "Cruzamento de Nomes" para conferência.
CRUZAR_NOME_SEMELHANTE = True
SEMELHANCA_MINIMA = 0.90

GERAR_RELATORIO_PARCEIRO_SEM_VALOR = False   # True = gera relatório também para quem ficou com R$ 0,00

# Responsável de polo (D_Polos) que vende como parceiro, cadastrando o aluno em OUTRO polo, é Parceiro Digital
# mesmo sem estar na PLANILHA GERAL: entra no fim dela e recebe pelas vendas em outros polos (a venda no próprio
# polo continua no repasse do polo). Ligação pelo CPF/CNPJ do responsável ou pelo nome (só com EQUIPE de
# parceiro no Pincel). Para não pagar alguém, use situacao = NÃO PAGAR em D_Parceiros_Ajustes.csv.
RESPONSAVEL_DE_POLO_E_PARCEIRO = True
STATUS_RESPONSAVEL_PARCEIRO = ["ATIVO", "PARCEIRO_SEM_POLO"]   # status do polo em D_Polos

# Dados fixos da PLANILHA GERAL (usados nas linhas novas; as existentes mantêm o que já está lá)
PLANILHA_GERAL_FIXOS = {"Instituição": "EDITORA PROMINAS LTDA", "CPF/CNPJ": "07.435.771/0001-50",
                        "Descrição": "Repasse referente ao Pincel e Prominas Acadêmico",
                        "Cidade": "Coronel Fabriciano"}

# ====================================== POLOS ======================================
GERAR_POLOS = True                          # False = calcula/gera só Parceiros Digitais
# Manual, pág. 7: quem não é parceiro (equipes internas e vendedores internos)
EQUIPES_INTERNAS = ["COPESE", "GRADUAÇÃO", "PÓS-GRADUAÇÃO", "SÃO JOSÉ DOS CAMPOS", "SEGUNDA GRADUAÇÃO"]
PALAVRAS_VENDEDOR_INTERNO = ["00", "EQUIPE", "COMERCIAL", "COPESE", "POLO", "SEDE"]
EXCLUIR_NOME_COM_PARENTESES = True
VENDEDORES_INTERNOS = ["JEFFERSON QUIRINO", "MATEUS GONZAGA OLIVEIRA"]   # sem EQUIPE no Pincel

# Unidades próprias (regex no nome sem acento, em maiúsculas) e nomes do Pincel que não batem com D_Polos
PADROES_UNIDADE_PROPRIA = [r"^UNIDADE\b", r"TRANSITORIO", r"\(PROPRIO\)", r"^IPATINGA\b",
                           r"^SAO JOSE DOS CAMPOS \(CENTRO\)"]
POLO_APELIDOS = {"NOVO ARIPUANA - AM": "NOVO ARIPUANA (MARIA LUIZA) (COM TRACO) - AM",
                 "NOVO ARIPUANA AM": "NOVO ARIPUANA (SEM TRACO) - AM",
                 "VICOSA (MARIA LUCIA) - MG": "VICOSA - MG"}
GERAR_RELATORIO_POLOS_DESATIVADOS = True    # manual, pág. 14: ficam "aguardando autorização"
GERAR_RELATORIO_POLO_SEM_REPASSE = False
LIMITAR_REPASSE_A_100_POR_CENTO = True      # polo + parceiro nunca passam de 100% da parcela

# ===================================== SAÍDAS ======================================
# Relatórios de parceiro e de polo divididos pela coluna PARCELA: "<nome>.xlsx" com as demais parcelas e
# "<nome> - MAT-R E PARCELA 1.xlsx" só com as parcelas abaixo (valores pela regra de D_Comissao, como os demais).
SEPARAR_PARCELAS = True
PARCELAS_RELATORIO_SEPARADO = ["MAT-R", 1]
SUFIXO_RELATORIO_SEPARADO = " - MAT-R E PARCELA 1"
GERAR_MEMORIA_DE_CALCULO = True             # aba linha a linha no TOTAL REPASSE (deixa o arquivo maior)
GERAR_CSV_POWER_BI = True                   # F_pagamentos.csv
