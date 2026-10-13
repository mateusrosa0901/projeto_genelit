"""Sequência do cálculo e geração dos arquivos do mês."""
import pandas as pd
from openpyxl import Workbook

import config
from .conferencia import conferir_parceiros, conferir_polos
from .consolidado import COLUNA_ARQUIVO_SEPARADO, gerar_consolidado
from .entradas import (COLUNAS_PINCEL, ler_ajustes_parceiros, ler_comissao, ler_excecoes, ler_planilha_geral,
                       ler_polos, ler_receita)
from .excel import relatorio_parceiro, relatorio_polo, tabela_relatorio_polo, planilha_geral
from .parceiros import (C_NAO_CAD, ORIGEM_RESPONSAVEL, PAGAR, RPA, calcular_parceiros, cruzamentos, montar_cadastro,
                        nao_cadastrados, nao_recebem, nome_relatorio, responsaveis_de_polo, resumo_parceiros, tabela_rpa)
from .polos import (classificar_vendedores, identificar_polos, limitar_100, percentuais_polo, pendencias_polos,
                    resumo_polos, valores_polo)
from .util import MESES, moeda, nome_arquivo_seguro, norm, txt_int, vazio

SUBPASTA_PARCEIROS = "RELATÓRIOS PARCEIROS DIGITAIS"
SUBPASTA_POLOS = {"ATIVO - A PAGAR": "RELATÓRIO DE PAGAMENTOS EM POLOS ATIVOS",
                  "DESATIVADO - AGUARDANDO AUTORIZAÇÃO":
                      "RELATÓRIO DE PAGAMENTOS EM POLOS DESATIVADOS (AGUARDANDO AUTORIZAÇÃO)"}


def processar():
    arq, df = ler_receita()
    print(f"[Entrada] {arq.name}")
    polos, principal, por_tokens = ler_polos()
    regras, excecoes = ler_comissao(), ler_excecoes()
    planilha, colunas = ler_planilha_geral()
    ajustes = ler_ajustes_parceiros()
    cad = montar_cadastro(planilha, ajustes, principal)
    novos = int(cad["NOVO"].astype(bool).sum())
    texto = (f"[Entrada] {len(df):,} pagamentos | {len(principal)} polos | {len(regras)} regras de comissão | "
             f"{len(excecoes)} exceções | {len(cad) - novos} parceiros na PLANILHA GERAL"
             + (f" (+{novos} de D_Parceiros_Ajustes)" if novos else ""))
    print(texto.replace(",", "."))

    identificar_polos(df, principal, por_tokens)   # polo de cada pagamento (D_Polos)
    classificar_vendedores(df)                     # venda do responsável do polo x de outra pessoa
    if config.RESPONSAVEL_DE_POLO_E_PARCEIRO:      # responsáveis de polo fora da PLANILHA GERAL entram no fim dela
        extra = responsaveis_de_polo(df, cad, polos)
        if len(extra):
            cad = montar_cadastro(planilha, pd.concat([a for a in (ajustes, extra) if len(a)], ignore_index=True),
                                  principal)
            print(f"[Entrada] {int(cad['ORIGEM'].eq(ORIGEM_RESPONSAVEL).sum())} responsáveis de polo fora da "
                  "PLANILHA GERAL incluídos como Parceiro Digital (vendas em outros polos)")
    percentuais_polo(df, regras, excecoes)         # % do polo (D_Comissao + exceções de polo)
    cruz = calcular_parceiros(df, cad, regras, excecoes, polos)   # Parceiros Digitais
    limitar_100(df)                                # polo + parceiro <= 100% da parcela
    valores_polo(df)

    periodo = df["DATA PAGTO"].dropna().dt.to_period("M")  # mês de referência = mês com mais pagamentos
    referencia = periodo.mode().iloc[0] if len(periodo) else pd.Timestamp.today().to_period("M")
    resumo = resumo_parceiros(df, cad)
    total = dict(zip(resumo["_id"], resumo["TOTAL REPASSE"])) if len(resumo) else {}
    valores = {}
    for i, sit in zip(cad.index, cad["SITUAÇÃO"]):
        recebe = sit == PAGAR or (sit == RPA and config.RPA_NA_PLANILHA_GERAL)
        valores[i] = float(total.get(i, 0.0)) if recebe else 0.0
    r = {"arquivo_receita": arq, "dados": df, "regras": regras, "excecoes": excecoes, "cadastro": cad,
         "colunas_planilha": colunas, "cruzador": cruz, "referencia": referencia, "resumo_parceiros": resumo,
         "valores_planilha": valores, "rpa": tabela_rpa(resumo), "nao_recebem": nao_recebem(df, cad),
         "nao_cadastrados": nao_cadastrados(df, cruz), "cruzamentos": cruzamentos(df, cad)}
    if config.GERAR_POLOS:
        r["resumo_polos"], r["pendencias"] = resumo_polos(df, principal), pendencias_polos(df, principal)
    else:
        r["resumo_polos"], r["pendencias"] = pd.DataFrame(), pd.DataFrame()
    return r


def ordenar_relatorio(linhas):
    """Como nos relatórios feitos à mão: pela DATA DE MATRÍCULA, mantendo a ordem do Pincel nos empates."""
    return linhas.sort_values(["DATA DE MATRÍCULA", "LINHA ORIGEM"], kind="mergesort")


def partes_relatorio(linhas):
    """[(sufixo do arquivo, linhas)]: as demais parcelas e, num arquivo à parte, MAT-R e parcela 1
    (config.PARCELAS_RELATORIO_SEPARADO). Parte sem linhas não gera arquivo."""
    if not config.SEPARAR_PARCELAS:
        return [("", linhas)]
    alvo = {norm(txt_int(p)) for p in config.PARCELAS_RELATORIO_SEPARADO}
    sep = linhas["PARCELA"].map(lambda v: not vazio(v) and norm(txt_int(v)) in alvo).astype(bool)
    return [(s, l) for s, l in (("", linhas[~sep]), (config.SUFIXO_RELATORIO_SEPARADO, linhas[sep])) if len(l)]


def _nome_livre(destino, base, usados):
    nome, n = f"{base}.xlsx", 2
    while (destino / nome).as_posix().upper() in usados:
        nome, n = f"{base} ({n}).xlsx", n + 1
    usados.add((destino / nome).as_posix().upper())
    return nome


def gerar_relatorios_parceiros(r, pasta):
    """Devolve ({id: relatório das demais parcelas}, {id: relatório de MAT-R e parcela 1})."""
    df, resumo = r["dados"], r["resumo_parceiros"]
    arquivos, separados, usados = {}, {}, set()
    if resumo.empty:
        return arquivos, separados
    pagos = df[df["PARCEIRO PAGO"]]
    for _, p in resumo.iterrows():
        if p["TOTAL REPASSE"] <= 0 and not config.GERAR_RELATORIO_PARCEIRO_SEM_VALOR:
            continue
        linhas = ordenar_relatorio(pagos[pagos["_parceiro_id"] == p["_id"]])
        destino = pasta / SUBPASTA_PARCEIROS / ("RPA (RH)" if p["SITUAÇÃO"] == RPA else "")
        destino.mkdir(parents=True, exist_ok=True)
        base = nome_arquivo_seguro(nome_relatorio(linhas) or p["NOME PARCEIRO (PLANILHA GERAL)"])
        for sufixo, parte in partes_relatorio(linhas):
            nome = _nome_livre(destino, base + sufixo, usados)
            relatorio_parceiro(destino / nome, parte, parte["% PARCEIRO"].tolist())
            (separados if sufixo else arquivos)[p["_id"]] = str((destino / nome).relative_to(pasta))
    return arquivos, separados


def gerar_relatorios_polos(r, pasta):
    df, polos = r["dados"], r["resumo_polos"].copy()
    usados, caminhos, caminhos_sep = set(), [], []
    for _, p in polos.iterrows():
        caminho = {"": "", config.SUFIXO_RELATORIO_SEPARADO: ""}
        if ((p["TOTAL REPASSE"] > 0 or config.GERAR_RELATORIO_POLO_SEM_REPASSE)
                and (config.GERAR_RELATORIO_POLOS_DESATIVADOS or not p["STATUS"].startswith("DESATIVADO"))):
            destino = pasta / SUBPASTA_POLOS[p["STATUS"]]
            destino.mkdir(parents=True, exist_ok=True)
            for sufixo, parte in partes_relatorio(df[df["_polo_key"] == p["_polo_key"]]):
                nome = _nome_livre(destino, p["ARQUIVO"][:-len(".xlsx")] + sufixo, usados)
                wb = Workbook()
                wb.active.title = "Relatório"
                relatorio_polo(wb.active, tabela_relatorio_polo(parte))
                wb.save(destino / nome)
                caminho[sufixo] = str((destino / nome).relative_to(pasta))
        caminhos.append(caminho[""])
        caminhos_sep.append(caminho[config.SUFIXO_RELATORIO_SEPARADO])
    if len(polos):
        polos["ARQUIVO"] = caminhos
        if config.SEPARAR_PARCELAS:
            polos[COLUNA_ARQUIVO_SEPARADO] = caminhos_sep
    r["resumo_polos"] = polos


def memoria(df):
    cols = ["LINHA ORIGEM", "MATRÍCULA", "NOME (ALUNO)", "DATA DE MATRÍCULA", "NÍVEL", "CURSO", "POLO",
            "POLO (CADASTRO)", "STATUS POLO", "RESPONSÁVEL DO POLO", "PARCELA", "TIPO PARCELA", "NOME PARCEIRO",
            "CPF PARCEIRO", "EQUIPE", "PARCEIRO DIGITAL", "CRUZAMENTO", "CLASSIFICAÇÃO PARCEIRO", "MOTIVO PARCEIRO",
            "VIGÊNCIA PARCEIRO", "EXCEÇÃO PARCEIRO", "DATA PAGTO", "VALOR PAGTO", "% PARCEIRO",
            "VALOR DO PARCEIRO CALCULADO", "VALOR DO PARCEIRO", "DIF. PARCEIRO (CALC - PINCEL)",
            "CLASSIFICAÇÃO VENDEDOR", "PAPEL NA VENDA", "VIGÊNCIA DA REGRA", "EXCEÇÃO POLO", "% POLO",
            "VALOR DO POLO CALCULADO", "VALOR DO POLO", "DIF. POLO (CALC - PINCEL)", "SITUAÇÃO", "OBS PARCEIRO"]
    m = df[cols].rename(columns={"VALOR DO POLO": "VALOR DO POLO (PINCEL)",
                                 "VALOR DO PARCEIRO": "VALOR DO PARCEIRO (PINCEL)",
                                 "SITUAÇÃO": "SITUAÇÃO ALUNO"}).copy()
    m["CPF PARCEIRO"] = m["CPF PARCEIRO"].map(lambda v: None if vazio(v) else txt_int(v))
    if not config.GERAR_POLOS:
        m = m.drop(columns=["CLASSIFICAÇÃO VENDEDOR", "PAPEL NA VENDA", "VIGÊNCIA DA REGRA", "EXCEÇÃO POLO",
                            "% POLO", "VALOR DO POLO CALCULADO", "VALOR DO POLO (PINCEL)",
                            "DIF. POLO (CALC - PINCEL)"])
    return m


def gerar_saidas(r, caminho_completo=True):
    """caminho_completo=False: a aba Resumo mostra só o nome da pasta do mês (versão web, que usa pasta temporária)."""
    ref = r["referencia"]
    mes = f"{MESES[ref.month - 1]}-{ref.year}"
    pasta = config.PASTA_SAIDA / str(ref.year) / f"{ref.month:02d}. REF. {MESES[ref.month - 1]}"
    pasta.mkdir(parents=True, exist_ok=True)

    print("[Saída] Relatórios dos Parceiros Digitais...")
    r["arquivos_parceiros"], r["arquivos_parceiros_separados"] = gerar_relatorios_parceiros(r, pasta)
    destino_pg = pasta / f"PLANILHA GERAL - {mes}.xlsx"
    planilha_geral(config.PASTA_DADOS / config.ARQ_PLANILHA_GERAL, destino_pg, r["cadastro"], r["valores_planilha"],
                   ref, r["colunas_planilha"])
    if config.GERAR_POLOS:
        print("[Saída] Relatórios dos polos...")
        gerar_relatorios_polos(r, pasta)

    print("[Conferência] Comparando com as planilhas feitas à mão (pasta referencias)...")
    r["conferencia_parceiros"], r["conferencia_parceiros_linhas"] = conferir_parceiros(r, config.PASTA_REFERENCIAS)
    if config.GERAR_POLOS:
        r["conferencia_polos_linhas"], r["conferencia_polos_resumo"] = conferir_polos(r, config.PASTA_REFERENCIAS)

    r["memoria"] = memoria(r["dados"])
    info = [("Pasta", str(pasta) if caminho_completo else pasta.name), ("PLANILHA GERAL do mês", destino_pg.name),
            ("Relatórios de parceiros", f"{SUBPASTA_PARCEIROS} ({len(r['arquivos_parceiros'])} arquivos + "
                                        f"{len(r['arquivos_parceiros_separados'])} de MAT-R e parcela 1)")]
    if config.GERAR_CSV_POWER_BI:
        fato = r["dados"][["LINHA ORIGEM"] + COLUNAS_PINCEL].copy()
        fato["CPF PARCEIRO"] = fato["CPF PARCEIRO"].map(lambda v: None if vazio(v) else txt_int(v))
        extra = r["memoria"].drop(columns=[c for c in r["memoria"].columns if c in fato.columns or
                                           c in ("VALOR DO POLO (PINCEL)", "VALOR DO PARCEIRO (PINCEL)",
                                                 "SITUAÇÃO ALUNO")])
        fato = pd.concat([fato, extra], axis=1)
        fato.to_csv(pasta / "F_pagamentos.csv", index=False, encoding="utf-8-sig", date_format="%d/%m/%Y")
        info.append(("Base para Power BI", "F_pagamentos.csv"))
    r["info_arquivos"] = info
    consolidado = pasta / f"TOTAL REPASSE - {mes}.xlsx"
    print("[Saída] Consolidado...")
    gerar_consolidado(r, consolidado)
    return pasta, consolidado, destino_pg


def imprimir_resumo(r):
    """Resumo do mês (terminal e registro da versão web)."""
    ref, val = r["referencia"], r["valores_planilha"]
    print(f"\nREFERÊNCIA: {MESES[ref.month - 1]}/{ref.year}  ({len(r['dados']):,} pagamentos)".replace(",", "."))
    pagos = {i: v for i, v in val.items() if v > 0}
    print(f"  Parceiros Digitais - PLANILHA GERAL ..... {moeda(sum(pagos.values())):>18}  ({len(pagos)} parceiros)")
    rpa = r["rpa"]
    if len(rpa):
        print(f"  Parceiros Digitais - RPA (RH) .......... {moeda(rpa['REPASSE PINCEL'].sum()):>18}  "
              f"({len(rpa)} parceiros)")
    print(f"  Relatórios de parceiros gerados ........ {len(r['arquivos_parceiros']):>18}")
    if config.SEPARAR_PARCELAS:
        print(f"  Relatórios de MAT-R e parcela 1 ........ {len(r['arquivos_parceiros_separados']):>18}  (parceiros)")
    nc = r["nao_cadastrados"]
    if len(nc):
        verificar = nc[nc["CLASSIFICAÇÃO"] == C_NAO_CAD]
        print(f"  Vendedores fora da PLANILHA GERAL ...... {moeda(verificar['VALOR PELA REGRA'].sum()):>18}  "
              f"({len(verificar)} nomes para verificar - aba Vendedores Não Cadastrados)")
    polos = r["resumo_polos"]
    if len(polos):
        for rotulo, prefixo in [("Polos ATIVOS (a pagar) .................", "ATIVO"),
                                ("Polos DESATIVADOS (aguardando) .........", "DESATIVADO")]:
            p = polos[polos["STATUS"].str.startswith(prefixo)]
            print(f"  {rotulo} {moeda(p['TOTAL REPASSE'].sum()):>18}  ({int((p['TOTAL REPASSE'] > 0).sum())} polos)")
    conf = r.get("conferencia_parceiros")
    if conf is not None and len(conf):
        print("\nCONFERÊNCIA COM AS PLANILHAS FEITAS À MÃO:")
        for _, c in conf[conf["TIPO"].eq("RELATÓRIO DO PARCEIRO")].iterrows():
            print(f"  {c['ARQUIVO']}: {c['RESULTADO']} - {int(c['LINHAS IGUAIS'])}/{int(c['LINHAS NO MANUAL'])} linhas "
                  f"| manual {moeda(c['TOTAL MANUAL'])} x calculado {moeda(c['TOTAL CALCULADO'])}")
        pg = conf[conf["TIPO"].eq("PLANILHA GERAL")]
        if len(pg):
            ok = int(pg["RESULTADO"].str.startswith("OK").sum())
            print(f"  PLANILHA GERAL: {ok} de {len(pg)} parceiros iguais (ou só centavos) | "
                  f"manual {moeda(pg['TOTAL MANUAL'].sum())} x calculado {moeda(pg['TOTAL CALCULADO'].sum())} "
                  "(detalhes na aba Conferência Parceiros)")
