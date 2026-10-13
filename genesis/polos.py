"""Repasse dos POLOS (manual, págs. 4-6 e 14): polo de cada pagamento, papel da venda, percentuais e totais."""
import re

import numpy as np
import pandas as pd

import config
from .util import (ap, arred, docs_compativeis, documento, mesmo_nome, nome_base, norm, palavras, sem_acento,
                   tokens)

ATIVO, DESATIVADO, PARCEIRO_SEM_POLO = "ATIVO", "DESATIVADO", "PARCEIRO SEM POLO"
UNIDADE_PROPRIA, NAO_CADASTRADO, SEM_POLO = "UNIDADE PRÓPRIA", "NÃO CADASTRADO", "SEM POLO"
COM_REPASSE = (ATIVO, DESATIVADO)
SEM_PARCEIRO, RESPONSAVEL, OUTRO_VENDEDOR = "SEM PARCEIRO", "RESPONSÁVEL DO POLO", "OUTRO VENDEDOR"
MODALIDADES = ["Graduação", "Segunda Graduação", "Tecnólogo", "Curso Técnico", "Disciplina Isolada"]
SIGLAS = {"MG", "SP", "RJ", "BA", "PR", "SC", "RS", "GO", "MT", "MS", "PA", "AM", "CE", "PE", "PB", "RN", "AL", "SE",
          "PI", "MA", "TO", "RO", "AC", "AP", "RR", "ES", "DF", "PROPRIO", "CENTRO"}


# ============================ 1. POLO E VENDEDOR DE CADA PAGAMENTO ============================
def identificar_polos(df, principal, por_tokens):
    """Liga a coluna POLO do Pincel ao cadastro D_Polos (tolerando grafia diferente)."""
    apelidos = {norm(k): norm(v) for k, v in config.POLO_APELIDOS.items()}
    padroes = [re.compile(p) for p in config.PADROES_UNIDADE_PROPRIA]

    def identificar(nome):
        if not nome:
            return None, SEM_POLO
        chave = apelidos.get(nome, nome)
        if chave not in principal.index:
            chave = por_tokens.get(tokens(chave), chave)
        if chave in principal.index:
            st = principal.at[chave, "_status"]
            return chave, {"ATIVO": ATIVO, "PARCEIRO_SEM_POLO": PARCEIRO_SEM_POLO}.get(st, DESATIVADO)
        return None, UNIDADE_PROPRIA if any(p.search(chave) for p in padroes) else NAO_CADASTRADO
    nomes = ap(df["POLO"], norm)
    mapa = {n: identificar(n) for n in set(nomes)}
    df["_polo_key"] = ap(nomes, lambda n: mapa[n][0])
    df["STATUS POLO"] = ap(nomes, lambda n: mapa[n][1])
    for coluna, origem in [("POLO (CADASTRO)", "polo_pagamentos"), ("RESPONSÁVEL DO POLO", "nome_responsavel"),
                           ("_resp_doc", "_doc"), ("_resp_nomes", "_nomes")]:
        dic = principal[origem].to_dict()
        df[coluna] = pd.Series([dic.get(k) if isinstance(k, str) else None for k in df["_polo_key"]],
                               index=df.index, dtype=object)


def motivo_vendedor_interno(df):
    """Manual, pág. 7: equipes internas, nomes com parênteses ou com '00', EQUIPE, COMERCIAL, COPESE, POLO, SEDE."""
    equipes = {norm(e) for e in config.EQUIPES_INTERNAS}
    internos = {nome_base(n) for n in config.VENDEDORES_INTERNOS}
    palavras_proibidas = [norm(w) for w in config.PALAVRAS_VENDEDOR_INTERNO if norm(w)]
    proibidas = re.compile("|".join(re.escape(p) if p.isdigit() else rf"\b{re.escape(p)}\b"
                                    for p in palavras_proibidas)) if palavras_proibidas else None

    def motivo(nome, equipe):
        if equipe and equipe in equipes:
            return f"EQUIPE '{equipe}' não é de parceiros (manual, pág. 7)"
        if nome:
            if config.EXCLUIR_NOME_COM_PARENTESES and "(" in nome:
                return "nome com parênteses (vendedor/equipe interna - manual, pág. 7)"
            achou = proibidas.search(nome) if proibidas is not None else None
            if achou:
                return f"nome contém '{achou.group(0)}' (manual, pág. 7)"
            if nome_base(nome) in internos:
                return "listado em VENDEDORES_INTERNOS (config.py)"
        return None
    return pd.Series([motivo(n, e) for n, e in zip(df["_vend_norm"], df["_equipe_norm"])], index=df.index,
                     dtype=object)


def classificar_vendedores(df):
    """Papel da venda para o POLO: sem parceiro ou venda do responsável = SELF_SALE (o polo recebe o percentual
    total - manual, Obs 1); venda de qualquer outra pessoa = POLO_OUTRO_VENDEDOR."""
    df["_vend_norm"], df["_vend_base"] = ap(df["NOME PARCEIRO"], norm), ap(df["NOME PARCEIRO"], nome_base)
    df["_vend_doc"], df["_equipe_norm"] = ap(df["CPF PARCEIRO"], documento), ap(df["EQUIPE"], norm)
    df["_motivo_interno"] = motivo_vendedor_interno(df)

    def quem(nome, base, doc, status, resp_doc, resp_nomes):
        if not isinstance(nome, str):
            return SEM_PARCEIRO
        if status in COM_REPASSE and ((isinstance(doc, str) and docs_compativeis(doc, resp_doc)) or
                                      (isinstance(resp_nomes, list) and any(mesmo_nome(base, r) for r in resp_nomes))):
            return RESPONSAVEL
        return OUTRO_VENDEDOR
    df["CLASSIFICAÇÃO VENDEDOR"] = [quem(*t) for t in zip(df["_vend_norm"], df["_vend_base"], df["_vend_doc"],
                                                          df["STATUS POLO"], df["_resp_doc"], df["_resp_nomes"])]
    df["PAPEL NA VENDA"] = np.where(df["CLASSIFICAÇÃO VENDEDOR"].isin([SEM_PARCEIRO, RESPONSAVEL]),
                                    "SELF_SALE", "POLO_OUTRO_VENDEDOR")


# ============================== 2. PERCENTUAIS (D_Comissao) ==============================
def tipo_parcela(p):
    """Parcelas numeradas (1, 2, ... 60) são mensalidades; MAT, MAT-R e NEG ficam como estão."""
    if isinstance(p, (int, np.integer)):
        return "MENSALIDADE"
    t = norm(p)
    return "MENSALIDADE" if t and t.isdigit() else (t or "")


def percentual(df, regras, papel):
    """(percentual, vigência) da regra vigente na DATA DE MATRÍCULA para o papel; NaN = sem regra."""
    base = pd.DataFrame({"_linha": df.index, "_mod": df["_mod"].values, "_tipo": df["TIPO PARCELA"].values,
                         "_data": pd.to_datetime(df["DATA DE MATRÍCULA"]).values})
    m = base.merge(regras.loc[regras["_papel"] == papel, ["_mod", "_tipo", "_ini", "_fim", "_pct"]],
                   on=["_mod", "_tipo"])
    m = m[(m["_data"] >= m["_ini"]) & (m["_data"] <= m["_fim"])]
    m = m.drop_duplicates("_linha", keep="last").set_index("_linha")
    vig = m["_ini"].dt.strftime("%d/%m/%Y") + " a " + m["_fim"].dt.strftime("%d/%m/%Y")
    return m["_pct"].reindex(df.index), vig.reindex(df.index)


def eh_primeira_mensalidade(df):
    return df["PARCELA"].map(lambda p: isinstance(p, (int, np.integer)) and p == config.PARCELA_PRIMEIRA_MENSALIDADE)


def vendedor_da_excecao(df, e):
    """Linhas vendidas pela pessoa da exceção (nome igual ou mesmo CPF/CNPJ)."""
    m = pd.Series([b == e["_nome"] for b in df["_vend_base"]], index=df.index, dtype=bool)
    if e["_doc"]:
        m |= pd.Series([isinstance(d, str) and docs_compativeis(d, e["_doc"]) for d in df["_vend_doc"]],
                       index=df.index, dtype=bool)
    return m


def filtro_excecao(df, e, quem):
    """Linhas em que a exceção 'e' (linha de D_Excecoes) se aplica; 'quem' = linhas da pessoa da exceção."""
    m = pd.Series(quem, index=df.index).fillna(False).astype(bool)
    if e["_aplica"] == "PRIMEIRA_MENSALIDADE":
        m &= eh_primeira_mensalidade(df).astype(bool)
    else:
        m &= df["TIPO PARCELA"].isin(["MENSALIDADE", "NEG"])
    if e["_mods"]:
        m &= df["_mod"].isin(e["_mods"])
    if e["_mods_excl"]:
        m &= ~df["_mod"].isin(e["_mods_excl"])
    if pd.notna(e["_desde"]):
        m &= pd.to_datetime(df["DATA DE MATRÍCULA"]) >= e["_desde"]
    if e["_polos"]:
        m &= df["_polo_key"].isin(e["_polos"])
    return m


def percentuais_polo(df, regras, excecoes):
    """% POLO de cada pagamento (D_Comissao) + exceções do tipo POLO (Wilhan, Ulisses, Altivo)."""
    df["TIPO PARCELA"], df["_mod"] = ap(df["PARCELA"], tipo_parcela), ap(df["NÍVEL"], norm)
    (p_self, v_self), (p_outro, v_outro) = (percentual(df, regras, p) for p in ("SELF_SALE", "POLO_OUTRO_VENDEDOR"))
    self_sale = df["PAPEL NA VENDA"].eq("SELF_SALE")
    df["% POLO"] = np.where(self_sale, p_self, p_outro).astype(float)
    df["VIGÊNCIA DA REGRA"] = np.where(self_sale, v_self, v_outro)
    df["_sem_regra_polo"] = pd.isna(df["% POLO"])
    df["% POLO"] = df["% POLO"].fillna(0.0)
    df["EXCEÇÃO POLO"] = None
    if excecoes is None or excecoes.empty:
        return
    for _, e in excecoes[excecoes["_papel"] == "POLO"].iterrows():
        m = filtro_excecao(df, e, vendedor_da_excecao(df, e)) & df["STATUS POLO"].isin(COM_REPASSE)
        df.loc[m, "% POLO"] = e["_pct"]
        df.loc[m, "CLASSIFICAÇÃO VENDEDOR"] = RESPONSAVEL
        df.loc[m, "PAPEL NA VENDA"] = "SELF_SALE"
        df.loc[m, "EXCEÇÃO POLO"] = f"{e['parceiro']}: {e['_pct']:.0%} ({e['aplica_em']})"


def limitar_100(df):
    """Polo + Parceiro Digital nunca passam de 100% da parcela (ex.: 1ª mensalidade 100% para Danubia/Integral)."""
    df["_limitado"] = False
    if not config.LIMITAR_REPASSE_A_100_POR_CENTO:
        return
    soma = df["% POLO"] + df["% PARCEIRO"]
    m = (soma > 1.0 + 1e-9) & (df["% PARCEIRO"] > 0)
    df.loc[m, "% POLO"] = (1.0 - df.loc[m, "% PARCEIRO"]).clip(lower=0.0)
    df.loc[m, "_limitado"] = True


def valores_polo(df):
    """VALOR DO POLO CALCULADO: só polos cadastrados como ATIVO ou DESATIVADO recebem."""
    recebe = df["STATUS POLO"].isin(COM_REPASSE)
    polo = arred(df["VALOR PAGTO"] * df["% POLO"])
    df["VALOR POLO POTENCIAL"] = polo.where(df["STATUS POLO"].isin([NAO_CADASTRADO, PARCEIRO_SEM_POLO]), 0.0).values
    df["% POLO"] = df["% POLO"].where(recebe, 0.0)
    df["VALOR DO POLO CALCULADO"] = polo.where(recebe, 0.0).values
    df["DIF. POLO (CALC - PINCEL)"] = arred(df["VALOR DO POLO CALCULADO"] - df["VALOR DO POLO"]).values


# ======================================== 3. RESUMOS ========================================
def nome_arquivo_polo(polo):
    return "POLO " + re.sub(r"\s+", " ", re.sub(r'[\\/:*?"<>|]', "-", str(polo).strip().upper())) + ".xlsx"


def resumo_polos(df, principal):
    base = df[df["STATUS POLO"].isin(COM_REPASSE)].copy()
    if base.empty:
        return pd.DataFrame()
    base["_mod_resumo"] = base["NÍVEL"].where(base["NÍVEL"].isin(MODALIDADES), "Outras")
    por_mod = base.pivot_table(index="_polo_key", columns="_mod_resumo", values="VALOR DO POLO CALCULADO",
                               aggfunc="sum", fill_value=0.0).reindex(columns=MODALIDADES + ["Outras"], fill_value=0.0)
    t = base.groupby("_polo_key").agg(**{
        "QTD PAGAMENTOS": ("VALOR PAGTO", "size"), "VALOR PAGO PELOS ALUNOS": ("VALOR PAGTO", "sum"),
        "TOTAL REPASSE": ("VALOR DO POLO CALCULADO", "sum"), "VALOR DO POLO (PINCEL)": ("VALOR DO POLO", "sum"),
        "STATUS": ("STATUS POLO", "first")})
    info = principal.loc[t.index]
    r = pd.DataFrame({
        "POLO": info["polo_pagamentos"].astype(str).str.upper(), "UF": info["estado"],
        "RESPONSÁVEL": info["nome_responsavel"], "CPF/CNPJ RESPONSÁVEL": info["cpf_responsavel"],
        "E-MAIL": info["email"],
        "STATUS": t["STATUS"].map({ATIVO: "ATIVO - A PAGAR", DESATIVADO: "DESATIVADO - AGUARDANDO AUTORIZAÇÃO"}),
        "QTD PAGAMENTOS": t["QTD PAGAMENTOS"], "VALOR PAGO PELOS ALUNOS": t["VALOR PAGO PELOS ALUNOS"]},
        index=t.index).join(por_mod)
    r["TOTAL REPASSE"] = arred(t["TOTAL REPASSE"]).values
    r["VALOR DO POLO (PINCEL)"] = arred(t["VALOR DO POLO (PINCEL)"]).values
    r["DIFERENÇA (CALC - PINCEL)"] = arred(r["TOTAL REPASSE"] - r["VALOR DO POLO (PINCEL)"]).values
    r["ARQUIVO"], r["_polo_key"] = r["POLO"].map(nome_arquivo_polo), r.index
    return r.sort_values(["STATUS", "POLO"]).reset_index(drop=True)


def sugerir_polos(nome, principal, limite=3):
    alvo = tokens(nome) - SIGLAS
    if not alvo:
        return ""
    cands = sorted(((len(alvo & tk) / len(alvo), rank, chave, st) for chave, tk, rank, st in
                    zip(principal.index, principal["_tokens"], principal["_rank"], principal["status_polo"])
                    if alvo & tk), key=lambda c: (-c[0], c[1]))
    return " | ".join(f"{c[2]} ({c[3]})" for c in cands[:limite] if c[0] >= 0.5)


def pendencias_polos(df, principal):
    itens = []

    def add(tipo, descricao, linhas, valor, acao, sugestao=""):
        itens.append({"TIPO": tipo, "DESCRIÇÃO": descricao, "QTD PAGAMENTOS": len(linhas),
                      "VALOR PAGO PELOS ALUNOS": round(float(linhas["VALOR PAGTO"].sum()), 2),
                      "VALOR ENVOLVIDO": round(float(valor), 2), "O QUE FAZER": acao, "SUGESTÃO": sugestao})
    status = df["STATUS POLO"]
    x = df[status.eq(NAO_CADASTRADO)]
    for nome, linhas in x.groupby(ap(x["POLO"], norm), sort=True):
        add("POLO NÃO CADASTRADO EM D_POLOS", nome, linhas, linhas["VALOR POLO POTENCIAL"].sum(),
            "Cadastrar o polo em D_Polos.csv (ou mapear em POLO_APELIDOS / PADROES_UNIDADE_PROPRIA no config.py). "
            "Valor envolvido = repasse estimado se for um polo parceiro.", sugerir_polos(nome, principal))
    x = df[status.eq(SEM_POLO)]
    if len(x):
        add("PAGAMENTO SEM POLO", "Coluna POLO em branco no relatório do Pincel", x, x["VALOR DO POLO"].sum(),
            "Verificar no Pincel a qual polo o aluno pertence. Valor envolvido = VALOR DO POLO do Pincel. "
            "O repasse de Parceiro Digital dessas linhas já foi calculado normalmente.")
    x = df[df["_sem_regra_polo"] & status.isin(COM_REPASSE + (NAO_CADASTRADO,))]
    for (nivel, tipo), linhas in x.groupby([x["NÍVEL"].fillna("(vazio)"), "TIPO PARCELA"]):
        add("SEM REGRA EM D_COMISSAO (POLO)", f"{nivel} / {tipo}", linhas, linhas["VALOR DO POLO"].sum(),
            "Incluir a modalidade/tipo de parcela em D_Comissao.csv (papéis SELF_SALE e POLO_OUTRO_VENDEDOR). "
            "Enquanto isso o repasse do polo fica zerado. Valor envolvido = VALOR DO POLO do Pincel.")
    x = df[status.eq(DESATIVADO) & (df["VALOR DO POLO CALCULADO"] != 0)]
    for chave, linhas in x.groupby("_polo_key"):
        add("POLO DESATIVADO - AGUARDANDO AUTORIZAÇÃO", str(principal.at[chave, "polo_pagamentos"]).upper(),
            linhas, linhas["VALOR DO POLO CALCULADO"].sum(),
            "Pagar somente com autorização do setor de Expansão e Negócios (manual, pág. 14).",
            f"Responsável: {principal.at[chave, 'nome_responsavel']}")
    x = df[status.eq(PARCEIRO_SEM_POLO) & ~df["PARCEIRO PAGO"]]
    for chave, linhas in x.groupby("_polo_key"):
        add("POLO CONVERTIDO EM PARCEIRO SEM POLO", str(principal.at[chave, "polo_pagamentos"]).upper(),
            linhas, linhas["VALOR POLO POTENCIAL"].sum(),
            "Polo com distrato: o responsável recebe só como Parceiro Digital pelas vendas dele (já calculado, se "
            "estiver na PLANILHA GERAL). Estas linhas (vendas de terceiros ou sem parceiro) não geram repasse. "
            "Valor envolvido = quanto seria se ainda fosse polo.",
            f"Responsável: {principal.at[chave, 'nome_responsavel']}")
    # venda do responsável no próprio polo DESATIVADO, paga como Parceiro Digital: não pagar em dobro como polo
    x = df[status.eq(DESATIVADO) & df["PARCEIRO PAGO"] & df["CLASSIFICAÇÃO VENDEDOR"].eq(RESPONSAVEL)
           & (df["VALOR DO POLO CALCULADO"] != 0)]
    for chave, linhas in x.groupby("_polo_key"):
        add("RESPONSÁVEL DE POLO DESATIVADO PAGO COMO PARCEIRO", str(principal.at[chave, "polo_pagamentos"]).upper(),
            linhas, linhas["VALOR DO POLO CALCULADO"].sum(),
            "Estas vendas já entram no repasse de Parceiro Digital do responsável. Se Expansão e Negócios autorizar "
            "o pagamento deste polo como polo, retirar estas linhas para não pagar duas vezes (manual, pág. 14).",
            f"Responsável: {principal.at[chave, 'nome_responsavel']}")
    x = df[df["_limitado"] & status.isin(COM_REPASSE)]
    if len(x):
        add("PARCEIRO COM 100% - POLO SEM REPASSE NA PARCELA",
            "1ª mensalidade paga 100% ao Parceiro Digital (exceções Danubia / Integral)", x,
            x["VALOR DO PARCEIRO CALCULADO"].sum(),
            "Informativo: como o parceiro recebe 100% da parcela, o polo não recebe sobre ela "
            "(LIMITAR_REPASSE_A_100_POR_CENTO no config.py).")
    return pd.DataFrame(itens)


def ordenar_texto(serie):
    return serie.map(lambda n: sem_acento(str(n)).upper())


__all__ = ["identificar_polos", "classificar_vendedores", "percentuais_polo", "limitar_100", "valores_polo",
           "resumo_polos", "pendencias_polos", "tipo_parcela", "percentual", "filtro_excecao", "palavras",
           "ordenar_texto"]
