"""Repasse dos PARCEIROS DIGITAIS (Parceiro Comercial) - manual, págs. 6-7 e 10.

Regras (conferidas linha a linha com os relatórios calculados à mão em SET/2026):
  * Parceiro Digital = quem está na PLANILHA GERAL (ou em D_Parceiros_Ajustes.csv). O "NOME PARCEIRO" do Pincel
    é ligado ao cadastro pelo nome, pelo CPF/CNPJ ou por nome contido/parecido.
  * Responsável de polo (D_Polos) que vendeu como parceiro e não está no cadastro também é Parceiro Digital
    (config.RESPONSAVEL_DE_POLO_E_PARCEIRO): entra no fim da PLANILHA GERAL e recebe pelas vendas em outros polos.
  * Percentual do parceiro (D_Comissao, papel PARCEIRO_COMERCIAL, pela DATA DE MATRÍCULA):
      MAT = 100% | MAT-R = 0% | parcela 1 = 0% | mensalidades e NEG = 10% (matrícula até 31/01/2023) ou 20%.
  * Exceções (D_Excecoes, papel PARCEIRO): Danubia e Integral = 100% na 1ª mensalidade (PARCELA 2) de
    Graduação/Tecnólogo; Marcieni = 40% fora da Graduação para matrículas desde 01/2024.
  * Não entram no relatório do parceiro: vendas no polo ATIVO do qual ele é o responsável (já estão no repasse
    do polo) e matrículas anteriores à data de início da parceria (D_Parceiros_Ajustes.csv).
  * Situação na PLANILHA GERAL: POLO, DISTRATO e NÃO ENVIAR REPASSE não recebem; RPA ("Retornou como parceiro")
    recebe pelo RH (aba RPA).
"""
import numpy as np
import pandas as pd

import config
from .polos import filtro_excecao, percentual
from .util import (ap, arred, arred1, docs_compativeis, docs_conflitam, documento, nome_base, norm, palavras,
                   semelhanca, sem_acento, tipo_pessoa, txt_int, vazio)

PAGAR, RPA, NAO_RECEBE = "PAGAR", "RPA (RH)", "NÃO RECEBE"
C_SEM, C_INTERNO, C_NAO_CAD = "SEM PARCEIRO", "VENDEDOR INTERNO (regra do manual)", "NÃO CADASTRADO NA PLANILHA GERAL"
C_RESPONSAVEL, C_PARCEIRO_POLO = "RESPONSÁVEL DO POLO (venda no próprio polo)", "PARCEIRO-POLO (responsável de polo)"
C_SEM_EQUIPE = "SEM EQUIPE NO PINCEL (vendedor interno)"
FORA_DO_CADASTRO = [C_NAO_CAD, C_SEM_EQUIPE, C_INTERNO, C_PARCEIRO_POLO, C_RESPONSAVEL]   # ordem da aba
C_NAO_RECEBE, C_PROPRIO_POLO = "CADASTRADO - NÃO RECEBE", "VENDA NO PRÓPRIO POLO (repasse do polo)"
C_ANTES, C_DIGITAL, C_RPA = "MATRÍCULA ANTES DO INÍCIO DA PARCERIA", "PARCEIRO DIGITAL", "PARCEIRO DIGITAL - RPA (RH)"
PAGOS = (C_DIGITAL, C_RPA)
ORIGEM_AJUSTES, ORIGEM_RESPONSAVEL = "D_Parceiros_Ajustes", "RESPONSÁVEL DE POLO"   # parceiros fora da planilha


# ================================== 1. CADASTRO ==================================
def _lista(v, sep="|"):
    return [] if vazio(v) else [x.strip() for x in str(v).split(sep) if x.strip()]


def _nomes_do_cadastro(nome):
    """Nome da planilha + partes separadas por '/' ('HELENA MARA MACHADO/ MARCIEL PEREIRA FELIX')."""
    nomes = [nome_base(nome)]
    if "/" in str(nome):
        nomes += [nome_base(p) for p in str(nome).split("/") if len(palavras(nome_base(p))) >= 2]
    return [n for n in dict.fromkeys(nomes) if n]


def _situacao(tipo, obs, forcada):
    f = norm(forcada)
    if f:
        if f in ("PAGAR", "ATIVO", "PF", "PJ"):
            return PAGAR, "situação forçada em D_Parceiros_Ajustes.csv"
        if f.startswith("RPA"):
            return RPA, "situação forçada em D_Parceiros_Ajustes.csv"
        return NAO_RECEBE, f"{forcada} (D_Parceiros_Ajustes.csv)"
    t, o = norm(tipo), norm(obs)
    if t and t in {norm(x) for x in config.TIPOS_NAO_PAGAR}:
        return NAO_RECEBE, f"PF / PJ = {tipo}"
    if o and any(o.startswith(norm(s)) for s in config.SITUACOES_NAO_PAGAR):
        return NAO_RECEBE, str(obs)
    if (t and t in {norm(x) for x in config.TIPOS_RPA}) or (o and any(o.startswith(norm(s))
                                                                     for s in config.SITUACOES_RPA)):
        return RPA, "pago via RPA pelo RH (manual, pág. 10)"
    return PAGAR, ""


def montar_cadastro(planilha, ajustes, principal):
    """Cadastro de Parceiros Digitais = PLANILHA GERAL + D_Parceiros_Ajustes.csv."""
    regs = []
    for p in planilha.to_dict("records"):
        p.update({"NOVO": False, "_nomes": _nomes_do_cadastro(p["NOME"]), "_apelidos": [],
                  "_docs": [d for d in [documento(p["CPF"])] if d], "INÍCIO DA PARCERIA": pd.NaT,
                  "_forcada": None, "AJUSTE": None, "ORIGEM": None})
        regs.append(p)
    novos = 0
    for a in ajustes.to_dict("records"):
        base = nome_base(a["nome_planilha"])
        alvo = [p for p in regs if base in p["_nomes"]]
        if not alvo:   # parceiro que ainda não está na PLANILHA GERAL: entra no fim dela
            novos += 1
            p = {c: None for c in planilha.columns}
            p.update({"ID": f"N{novos}", "LINHA": None, "NOME": str(a["nome_planilha"]).strip(),
                      "CPF": a["cpf_cnpj"], "PF / PJ": a["pf_pj"], "EMAIL": a["email"],
                      "OBSERVAÇÃO": a["observacao"], "NOVO": True, "_nomes": _nomes_do_cadastro(a["nome_planilha"]),
                      "_apelidos": [], "_docs": [d for d in [documento(a["cpf_cnpj"])] if d],
                      "INÍCIO DA PARCERIA": pd.NaT, "_forcada": None, "AJUSTE": None,
                      "ORIGEM": a.get("_origem") if isinstance(a.get("_origem"), str) else ORIGEM_AJUSTES})
            regs.append(p)
            alvo = [p]
        for p in alvo:
            apel = [nome_base(x) for x in _lista(a["apelidos_pincel"]) if nome_base(x)]
            p["_apelidos"] = list(dict.fromkeys(p["_apelidos"] + apel))
            extra = [documento(x) for x in _lista(a["cpf_cnpj_extra"]) + _lista(a["cpf_cnpj"]) if documento(x)]
            p["_docs"] = list(dict.fromkeys(p["_docs"] + extra))
            if pd.notna(a["_inicio"]):
                p["INÍCIO DA PARCERIA"] = a["_inicio"]
            if not vazio(a["situacao"]):
                p["_forcada"] = a["situacao"]
            p["AJUSTE"] = a["observacao"] if not vazio(a["observacao"]) else "D_Parceiros_Ajustes.csv"
    cad = pd.DataFrame(regs).astype(object)
    cad["INÍCIO DA PARCERIA"] = pd.to_datetime(cad["INÍCIO DA PARCERIA"])
    cad = cad.set_index("ID", drop=False)
    sit = [_situacao(t, o, f) for t, o, f in zip(cad["PF / PJ"], cad["OBSERVAÇÃO"], cad["_forcada"])]
    cad["SITUAÇÃO"], cad["MOTIVO SITUAÇÃO"] = [s[0] for s in sit], [s[1] for s in sit]
    cad["_todos_nomes"] = [list(dict.fromkeys(n + a)) for n, a in zip(cad["_nomes"], cad["_apelidos"])]
    # polos ATIVOS dos quais o parceiro é o responsável (venda nesses polos = repasse do polo)
    ativos = principal[principal["_status"] == "ATIVO"]

    def proprios(docs, nomes):
        achou = []
        for chave, pdoc, pnomes in zip(ativos.index, ativos["_doc"], ativos["_nomes"]):
            if (pdoc and any(docs_compativeis(d, pdoc) for d in docs)) or \
                    any(_nomes_compativeis(n, pnomes) for n in nomes):
                achou.append(chave)
        return achou
    cad["POLOS PRÓPRIOS"] = [proprios(d, n) for d, n in zip(cad["_docs"], cad["_todos_nomes"])]
    return cad


def _nomes_compativeis(base, lista):
    """Mesmo 1º nome e um contido no outro (ou 2+ palavras em comum)."""
    pb = palavras(base)
    if len(pb) < 2:
        return False
    for n in lista:
        pn = palavras(n)
        if len(pn) >= 2 and pn[0] == pb[0] and (set(pn) <= set(pb) or set(pb) <= set(pn)):
            return True
    return False


# ====================== 2. CRUZAMENTO "NOME PARCEIRO" (Pincel) x CADASTRO ======================
class Cruzador:
    """Liga cada (NOME PARCEIRO, CPF PARCEIRO) do Pincel a um parceiro do cadastro."""

    def __init__(self, cad):
        self.cad = cad
        self.por_nome, self.por_apelido, self.por_doc, self.por_primeiro = {}, {}, {}, {}
        for i, nomes, apel, docs in zip(cad.index, cad["_nomes"], cad["_apelidos"], cad["_docs"]):
            for n in nomes:
                self.por_nome.setdefault(n, []).append(i)
            for n in apel:
                self.por_apelido.setdefault(n, []).append(i)
            for d in docs:
                self.por_doc.setdefault(d[:11], []).append(i)
            for n in nomes + apel:
                p = palavras(n)
                if p:
                    self.por_primeiro.setdefault(p[0], set()).add(i)

    def _docs(self, i):
        return self.cad.at[i, "_docs"]

    def _conflito(self, d, i):
        docs = self._docs(i)
        return bool(d) and bool(docs) and not any(docs_compativeis(d, x) for x in docs) and \
            any(docs_conflitam(d, x) for x in docs)

    def cruzar(self, nome, cpf):
        b, d = nome_base(nome), documento(cpf)
        if not b:
            return None, None
        for indice, rotulo in ((self.por_apelido, "APELIDO (D_Parceiros_Ajustes)"), (self.por_nome, "NOME IGUAL")):
            ids = indice.get(b, [])
            if ids:
                if len(ids) > 1 and d:
                    pref = [i for i in ids if any(docs_compativeis(d, x) for x in self._docs(i))]
                    ids = pref or ids
                return ids[0], rotulo
        pb = palavras(b)
        if d:
            cands = [i for i in dict.fromkeys(self.por_doc.get(d[:11], []))
                     if any(docs_compativeis(d, x) for x in self._docs(i))]

            def parecido(i):   # mesmo documento: basta o nome ser parecido (erro de digitação, sobrenome a mais)
                return _nomes_compativeis(b, self.cad.at[i, "_todos_nomes"]) or any(
                    (palavras(n)[:1] == pb[:1] and semelhanca(" ".join(pb), " ".join(palavras(n))) >= 0.75)
                    for n in self.cad.at[i, "_todos_nomes"])
            comp = [i for i in cands if parecido(i)]
            if len(comp) == 1:
                return comp[0], "CPF/CNPJ + NOME PARECIDO"
        if len(pb) < 2:
            return None, None
        grupo = sorted(self.por_primeiro.get(pb[0], set()), key=str)
        cont = [i for i in grupo if _nomes_compativeis(b, self.cad.at[i, "_todos_nomes"]) and not self._conflito(d, i)]
        if len(cont) == 1:
            return cont[0], "NOME CONTIDO"
        if config.CRUZAR_NOME_SEMELHANTE:
            melhores = {}
            for i in grupo:
                if self._conflito(d, i):
                    continue
                for n in self.cad.at[i, "_todos_nomes"]:
                    pn = palavras(n)
                    if len(pn) >= 2 and pn[0] == pb[0] and pn[-1] == pb[-1]:
                        s = semelhanca(" ".join(pb), " ".join(pn))
                        if s >= config.SEMELHANCA_MINIMA:
                            melhores[i] = max(s, melhores.get(i, 0))
            if melhores:
                ordem = sorted(melhores.items(), key=lambda x: -x[1])
                if len(ordem) == 1 or ordem[0][1] > ordem[1][1]:
                    return ordem[0][0], f"NOME SEMELHANTE ({ordem[0][1]:.0%})"
        return None, None

    def sugerir(self, nome, minimo=0.75):
        """Nome mais parecido do cadastro (para a aba de vendedores não cadastrados)."""
        b = nome_base(nome)
        if not b:
            return ""
        melhor = max(((semelhanca(b, n), i) for i, ns in zip(self.cad.index, self.cad["_todos_nomes"]) for n in ns),
                     default=(0, None))
        return f"{self.cad.at[melhor[1], 'NOME']} ({melhor[0]:.0%})" if melhor[0] >= minimo else ""


# =================== 2b. RESPONSÁVEIS DE POLO QUE VENDEM COMO PARCEIRO (fora do cadastro) ===================
def _polos_do_responsavel(resp, base, doc, interno, equipe):
    """Linhas de D_Polos das quais o vendedor é o responsável: pelo CPF/CNPJ ou, sem documento em comum, pelo
    nome - nesse caso só com EQUIPE de parceiro no Pincel e sem documento conflitante (evita homônimos)."""
    if isinstance(doc, str):
        achou = [k for k, d in zip(resp.index, resp["_doc"]) if isinstance(d, str) and docs_compativeis(doc, d)]
        if achou:
            return achou
    if not isinstance(base, str) or isinstance(interno, str) or not isinstance(equipe, str):
        return []
    return [k for k, d, nomes in zip(resp.index, resp["_doc"], resp["_nomes"])
            if _nomes_compativeis(base, nomes) and not (isinstance(d, str) and docs_conflitam(doc, d))]


def responsaveis_de_polo(df, cad, polos):
    """Responsáveis de polo que venderam no mês e não estão no cadastro, no formato de D_Parceiros_Ajustes.csv
    (entram no fim da PLANILHA GERAL). A venda no próprio polo continua no repasse do polo (ver classificar)."""
    resp = polos[polos["_status"].isin({norm(s) for s in config.STATUS_RESPONSAVEL_PARCEIRO})]
    colunas = ["nome_planilha", "apelidos_pincel", "cpf_cnpj_extra", "data_inicio_parceria", "situacao",
               "cpf_cnpj", "pf_pj", "email", "observacao", "_origem"]   # as de D_Parceiros_Ajustes.csv
    b = df[df["NOME PARCEIRO"].map(lambda v: isinstance(v, str))]
    if resp.empty or b.empty:
        return pd.DataFrame(columns=colunas + ["_inicio"])
    b = b.assign(_cpf=b["CPF PARCEIRO"].map(lambda v: "" if vazio(v) else txt_int(v)))
    grupos = b.groupby(["NOME PARCEIRO", "_cpf"], sort=False).agg(
        n=("NOME PARCEIRO", "size"), base=("_vend_base", "first"), doc=("_vend_doc", "first"),
        interno=("_motivo_interno", "first"), equipe=("EQUIPE", "first"),
        fora=("CLASSIFICAÇÃO VENDEDOR", lambda s: s.ne("RESPONSÁVEL DO POLO").any()))
    cruz, pessoas = Cruzador(cad), {}
    for (nome, cpf), g in grupos.iterrows():
        if cruz.cruzar(nome, cpf or None)[0] is not None:
            continue
        chaves = _polos_do_responsavel(resp, g["base"], g["doc"], g["interno"], g["equipe"])
        if not chaves:
            continue
        p = pessoas.setdefault(g["doc"] if isinstance(g["doc"], str) else g["base"],
                               {"nomes": {}, "cpf": "", "polos": [], "fora": False})
        p["nomes"][nome] = p["nomes"].get(nome, 0) + int(g["n"])
        p["cpf"] = p["cpf"] or cpf
        p["polos"] += [k for k in chaves if k not in p["polos"]]
        p["fora"] = p["fora"] or bool(g["fora"])
    regs = []
    for p in pessoas.values():
        if not p["fora"]:   # só vendeu no próprio polo: nada a receber como parceiro
            continue
        nomes = sorted(p["nomes"], key=lambda n: -p["nomes"][n])   # nome mais usado no Pincel
        linhas = resp.loc[p["polos"]]
        regs.append({"nome_planilha": nomes[0], "apelidos_pincel": "|".join(nomes[1:]) or None,
                     "cpf_cnpj": p["cpf"] or None, "pf_pj": tipo_pessoa(p["cpf"]) or None,
                     "email": next((e for e in linhas["email"] if not vazio(e)), None),
                     "observacao": "Responsável do polo " + " | ".join(
                         f"{k} ({s})" for k, s in zip(linhas["polo_key"], linhas["status_polo"])),
                     "_origem": ORIGEM_RESPONSAVEL})
    r = pd.DataFrame(regs, columns=colunas)
    r["_inicio"] = pd.NaT
    return r


# ================================ 3. CÁLCULO POR LINHA ================================
def ids_da_excecao(e, cad):
    ids = []
    for i, nomes, docs in zip(cad.index, cad["_todos_nomes"], cad["_docs"]):
        if (e["_nome"] and (e["_nome"] in nomes or _nomes_compativeis(e["_nome"], nomes))) or \
                (e["_doc"] and any(docs_compativeis(e["_doc"], d) for d in docs)):
            ids.append(i)
    return ids


class IndiceResponsaveis:
    """Responsáveis de polos (D_Polos, qualquer situação): para saber se quem vendeu é um 'parceiro-polo'."""

    def __init__(self, polos):
        self.por_doc, self.por_primeiro = {}, {}
        for chave, doc, nomes, st in zip(polos["_chave"], polos["_doc"], polos["_nomes"], polos["status_polo"]):
            if doc:
                self.por_doc.setdefault(doc[:11], []).append((doc, chave, st))
            for n in nomes:
                p = palavras(n)
                if len(p) >= 2:
                    self.por_primeiro.setdefault(p[0], []).append((n, chave, st))
        self._cache = {}

    def polo_de(self, base, doc):
        k = (base, doc)
        if k not in self._cache:
            achou = None
            if isinstance(doc, str):
                achou = next(((c, s) for d, c, s in self.por_doc.get(doc[:11], []) if docs_compativeis(doc, d)), None)
            p = palavras(base) if isinstance(base, str) else []
            if not achou and len(p) >= 2:
                achou = next(((c, s) for n, c, s in self.por_primeiro.get(p[0], []) if _nomes_compativeis(base, [n])),
                             None)
            self._cache[k] = achou
        return self._cache[k]


def calcular_parceiros(df, cad, regras, excecoes, polos):
    """Acrescenta ao df: PARCEIRO DIGITAL, CRUZAMENTO, CLASSIFICAÇÃO PARCEIRO, MOTIVO PARCEIRO, % PARCEIRO,
    VALOR DO PARCEIRO CALCULADO (sem arredondar, como a fórmula do Excel) e a memória da regra."""
    cruz = Cruzador(cad)
    pares = {}
    for n, c in zip(df["NOME PARCEIRO"], df["CPF PARCEIRO"]):
        if isinstance(n, str) and (n, str(c)) not in pares:
            pares[(n, str(c))] = cruz.cruzar(n, c)
    res = [pares.get((n, str(c)), (None, None)) if isinstance(n, str) else (None, None)
           for n, c in zip(df["NOME PARCEIRO"], df["CPF PARCEIRO"])]
    df["_parceiro_id"] = pd.Series([r[0] for r in res], index=df.index, dtype=object)
    df["CRUZAMENTO"] = pd.Series([r[1] for r in res], index=df.index, dtype=object)
    df["PARCEIRO DIGITAL"] = ap(df["_parceiro_id"], lambda i: cad.at[i, "NOME"] if isinstance(i, str) else None)

    # percentual pela regra (D_Comissao, papel PARCEIRO_COMERCIAL)
    p, vig = percentual(df, regras, "PARCEIRO_COMERCIAL")
    df["% PARCEIRO (REGRA)"] = p.astype(float)
    df["VIGÊNCIA PARCEIRO"] = vig
    tem_nome = df["NOME PARCEIRO"].map(lambda v: isinstance(v, str))
    df["_sem_regra_parceiro"] = df["% PARCEIRO (REGRA)"].isna() & tem_nome
    zerar = df["PARCELA"].map(lambda v: isinstance(v, (int, np.integer)) and v in config.PARCELAS_SEM_REPASSE_PARCEIRO)
    df.loc[zerar, "% PARCEIRO (REGRA)"] = 0.0
    df["% PARCEIRO (REGRA)"] = df["% PARCEIRO (REGRA)"].fillna(0.0)
    df["EXCEÇÃO PARCEIRO"] = None
    if excecoes is not None and len(excecoes):
        for _, e in excecoes[excecoes["_papel"] == "PARCEIRO"].iterrows():
            ids = ids_da_excecao(e, cad)
            if not ids:
                print(f"[Exceções] '{e['parceiro']}' não está na PLANILHA GERAL - exceção não aplicada.")
                continue
            m = filtro_excecao(df, e, df["_parceiro_id"].isin(ids))
            df.loc[m, "% PARCEIRO (REGRA)"] = e["_pct"]
            df.loc[m, "EXCEÇÃO PARCEIRO"] = f"{e['parceiro']}: {e['_pct']:.0%} ({e['aplica_em']})"

    # classificação de cada linha
    resp = IndiceResponsaveis(polos)

    def classificar(i, nome, motivo_interno, polo_key, data_mat, cls_vendedor, base, doc, equipe):
        if not isinstance(nome, str):
            return C_SEM, ""
        if not isinstance(i, str):   # fora da PLANILHA GERAL: só informativo (aba Vendedores Não Cadastrados)
            if cls_vendedor == "RESPONSÁVEL DO POLO":
                return C_RESPONSAVEL, f"venda no próprio polo ({polo_key}) - entra no repasse do polo"
            outro = resp.polo_de(base, doc)
            if outro:
                return C_PARCEIRO_POLO, f"responsável pelo polo {outro[0]} ({outro[1]})"
            if isinstance(motivo_interno, str):
                return C_INTERNO, motivo_interno
            if not isinstance(equipe, str):
                return C_SEM_EQUIPE, "sem EQUIPE no Pincel (manual, pág. 7: parceiros são 'Carteira' e 'Parceiros')"
            return C_NAO_CAD, f"EQUIPE '{equipe}'"
        sit = cad.at[i, "SITUAÇÃO"]
        if sit == NAO_RECEBE:
            return C_NAO_RECEBE, cad.at[i, "MOTIVO SITUAÇÃO"]
        proprio = isinstance(polo_key, str) and polo_key in cad.at[i, "POLOS PRÓPRIOS"]
        # responsável de polo incluído pelo programa: venda no polo dele (qualquer status) fica com o polo
        proprio = proprio or (cad.at[i, "ORIGEM"] == ORIGEM_RESPONSAVEL and cls_vendedor == "RESPONSÁVEL DO POLO")
        if config.EXCLUIR_VENDA_NO_PROPRIO_POLO and proprio:
            return C_PROPRIO_POLO, f"responsável pelo polo {polo_key}"
        inicio = cad.at[i, "INÍCIO DA PARCERIA"]
        if pd.notna(inicio) and pd.notna(data_mat) and pd.Timestamp(data_mat) < inicio:
            return C_ANTES, f"parceria a partir de {inicio:%d/%m/%Y}"
        return (C_RPA if sit == RPA else C_DIGITAL), ""
    cl = [classificar(*t) for t in zip(df["_parceiro_id"], df["NOME PARCEIRO"], df["_motivo_interno"],
                                       df["_polo_key"], df["DATA DE MATRÍCULA"], df["CLASSIFICAÇÃO VENDEDOR"],
                                       df["_vend_base"], df["_vend_doc"], df["EQUIPE"])]
    df["CLASSIFICAÇÃO PARCEIRO"] = [c[0] for c in cl]
    df["MOTIVO PARCEIRO"] = [c[1] for c in cl]
    df["PARCEIRO PAGO"] = df["CLASSIFICAÇÃO PARCEIRO"].isin(PAGOS)
    df["VALOR PARCEIRO (REGRA)"] = (df["VALOR PAGTO"] * df["% PARCEIRO (REGRA)"]).where(tem_nome, 0.0)
    df["% PARCEIRO"] = df["% PARCEIRO (REGRA)"].where(df["PARCEIRO PAGO"], 0.0)
    df["VALOR DO PARCEIRO CALCULADO"] = df["VALOR PARCEIRO (REGRA)"].where(df["PARCEIRO PAGO"], 0.0)
    df["DIF. PARCEIRO (CALC - PINCEL)"] = arred(arred(df["VALOR DO PARCEIRO CALCULADO"]) -
                                                df["VALOR DO PARCEIRO"].where(df["PARCEIRO PAGO"], 0.0)).values
    return cruz


# ===================================== 4. RESUMOS =====================================
def _nome_exibicao(nomes):
    v = nomes.dropna()
    return v.mode().iloc[0] if len(v) else ""


def _ordem(s):
    return s.map(lambda n: sem_acento(str(n)).upper())


def nome_relatorio(df_parceiro):
    """Nome do arquivo do parceiro = nome mais usado no Pincel (como nos relatórios feitos à mão)."""
    return _nome_exibicao(df_parceiro["NOME PARCEIRO"])


def resumo_parceiros(df, cad):
    """Uma linha por Parceiro Digital que recebe no mês (a 'tabela dinâmica' do manual)."""
    b = df[df["PARCEIRO PAGO"]].copy()
    if b.empty:
        return pd.DataFrame()
    b["_mat"] = np.where(b["TIPO PARCELA"].eq("MAT"), b["VALOR DO PARCEIRO CALCULADO"], 0.0)
    g = b.groupby("_parceiro_id", sort=False)
    r = pd.DataFrame({
        "NOME NO PINCEL": g["NOME PARCEIRO"].agg(_nome_exibicao),
        "QTD PAGAMENTOS": g.size(), "REPASSE MATRÍCULAS": g["_mat"].sum(),
        "TOTAL SEM ARREDONDAR": g["VALOR DO PARCEIRO CALCULADO"].sum(),
        "VALOR DO PARCEIRO (PINCEL)": g["VALOR DO PARCEIRO"].sum(),
        "EXCEÇÃO": g["EXCEÇÃO PARCEIRO"].agg(lambda s: " | ".join(sorted(set(s.dropna()))))})
    info = cad.loc[r.index]
    r.insert(0, "Nº PLANILHA", info["Nº"].values)
    r.insert(1, "NOME PARCEIRO (PLANILHA GERAL)", info["NOME"].values)
    r.insert(3, "CPF/CNPJ", info["CPF"].values)
    r.insert(4, "PF / PJ", info["PF / PJ"].values)
    r.insert(5, "SITUAÇÃO", info["SITUAÇÃO"].values)
    r["TOTAL REPASSE"] = [arred1(v) for v in r["TOTAL SEM ARREDONDAR"]]
    r["REPASSE MATRÍCULAS"] = [arred1(v) for v in r["REPASSE MATRÍCULAS"]]
    r["REPASSE MENSALIDADES"] = (r["TOTAL REPASSE"] - r["REPASSE MATRÍCULAS"]).round(2)
    r["VALOR DO PARCEIRO (PINCEL)"] = [arred1(v) for v in r["VALOR DO PARCEIRO (PINCEL)"]]
    r["DIFERENÇA (CALC - PINCEL)"] = (r["TOTAL REPASSE"] - r["VALOR DO PARCEIRO (PINCEL)"]).round(2)
    r["INÍCIO DA PARCERIA"] = info["INÍCIO DA PARCERIA"].values
    r["OBSERVAÇÃO (PLANILHA)"] = info["OBSERVAÇÃO"].values
    r["NOVO NA PLANILHA"] = np.where(info["NOVO"].values, "SIM - incluído no fim da PLANILHA GERAL", "")
    r["_id"] = r.index
    r = r.assign(_o=_ordem(r["NOME PARCEIRO (PLANILHA GERAL)"])).sort_values("_o").drop(columns="_o")
    return r[["Nº PLANILHA", "NOME PARCEIRO (PLANILHA GERAL)", "NOME NO PINCEL", "CPF/CNPJ", "PF / PJ", "SITUAÇÃO",
              "QTD PAGAMENTOS", "REPASSE MATRÍCULAS", "REPASSE MENSALIDADES", "TOTAL REPASSE",
              "VALOR DO PARCEIRO (PINCEL)", "DIFERENÇA (CALC - PINCEL)", "EXCEÇÃO", "INÍCIO DA PARCERIA",
              "OBSERVAÇÃO (PLANILHA)", "NOVO NA PLANILHA", "TOTAL SEM ARREDONDAR", "_id"]].reset_index(drop=True)


def tabela_rpa(resumo):
    """Manual, pág. 10: planilha para o RH com nome, CPF, valor bruto e 20% de INSS patronal."""
    if resumo.empty:
        return pd.DataFrame()
    r = resumo[resumo["SITUAÇÃO"].eq(RPA) & (resumo["TOTAL REPASSE"] > 0)]
    return pd.DataFrame({"NOME DO PARCEIRO": r["NOME PARCEIRO (PLANILHA GERAL)"].values, "CPF": r["CPF/CNPJ"].values,
                         "REPASSE PINCEL": r["TOTAL REPASSE"].values,
                         "SALDO CARTEIRA PROMINAS ACADÊMICO": np.nan})


def nao_recebem(df, cad):
    """Parceiros cadastrados com vendas no mês que não recebem (POLO, DISTRATO, venda no próprio polo...)."""
    b = df[df["_parceiro_id"].notna() & ~df["PARCEIRO PAGO"]].copy()
    if b.empty:
        return pd.DataFrame()
    b["_motivo"] = np.where(b["MOTIVO PARCEIRO"].astype(str) != "", b["CLASSIFICAÇÃO PARCEIRO"] + ": " +
                            b["MOTIVO PARCEIRO"].astype(str), b["CLASSIFICAÇÃO PARCEIRO"])
    g = b.groupby(["_parceiro_id", "_motivo"], sort=False)
    r = g.agg(**{"NOME NO PINCEL": ("NOME PARCEIRO", _nome_exibicao), "QTD PAGAMENTOS": ("VALOR PAGTO", "size"),
                 "VALOR PELA REGRA": ("VALOR PARCEIRO (REGRA)", "sum"),
                 "VALOR DO PARCEIRO (PINCEL)": ("VALOR DO PARCEIRO", "sum")}).reset_index()
    r.insert(0, "NOME PARCEIRO (PLANILHA GERAL)", r["_parceiro_id"].map(cad["NOME"]))
    r.insert(2, "PF / PJ", r["_parceiro_id"].map(cad["PF / PJ"]))
    r = r.rename(columns={"_motivo": "MOTIVO"})
    for c in ["VALOR PELA REGRA", "VALOR DO PARCEIRO (PINCEL)"]:
        r[c] = [arred1(v) for v in r[c]]
    r = r.assign(_o=_ordem(r["NOME PARCEIRO (PLANILHA GERAL)"])).sort_values(["MOTIVO", "_o"])
    return r[["NOME PARCEIRO (PLANILHA GERAL)", "NOME NO PINCEL", "PF / PJ", "MOTIVO", "QTD PAGAMENTOS",
              "VALOR PELA REGRA", "VALOR DO PARCEIRO (PINCEL)"]].reset_index(drop=True)


def nao_cadastrados(df, cruz):
    """Quem vendeu e não está na PLANILHA GERAL. Só 'NÃO CADASTRADO NA PLANILHA GERAL' precisa de análise
    (pode ser parceiro novo); os demais grupos são vendedores internos, responsáveis de polo etc."""
    b = df[df["CLASSIFICAÇÃO PARCEIRO"].isin(FORA_DO_CADASTRO)].copy()
    if b.empty:
        return pd.DataFrame()
    b["_doc_txt"] = b["CPF PARCEIRO"].map(lambda v: "" if vazio(v) else str(v))
    b["_chave"] = b["_vend_base"].fillna("") + "|" + b["_vend_doc"].fillna("")
    g = b.groupby(["CLASSIFICAÇÃO PARCEIRO", "_chave"], sort=False)
    r = g.agg(**{"NOME PARCEIRO": ("NOME PARCEIRO", _nome_exibicao), "CPF PARCEIRO": ("_doc_txt", "first"),
                 "EQUIPE": ("EQUIPE", lambda s: " / ".join(sorted(set(s.dropna())))),
                 "MOTIVO": ("MOTIVO PARCEIRO", "first"), "QTD PAGAMENTOS": ("VALOR PAGTO", "size"),
                 "VALOR PELA REGRA": ("VALOR PARCEIRO (REGRA)", "sum"),
                 "VALOR DO PARCEIRO (PINCEL)": ("VALOR DO PARCEIRO", "sum")}).reset_index()
    r = r.rename(columns={"CLASSIFICAÇÃO PARCEIRO": "CLASSIFICAÇÃO"})
    for c in ["VALOR PELA REGRA", "VALOR DO PARCEIRO (PINCEL)"]:
        r[c] = [arred1(v) for v in r[c]]
    r["PARECIDO COM (PLANILHA GERAL)"] = [cruz.sugerir(n) if cl == C_NAO_CAD else ""
                                         for n, cl in zip(r["NOME PARCEIRO"], r["CLASSIFICAÇÃO"])]
    r["_cl"] = r["CLASSIFICAÇÃO"].map({c: k for k, c in enumerate(FORA_DO_CADASTRO)})
    r = r.sort_values(["_cl", "VALOR PELA REGRA"], ascending=[True, False])
    return r[["CLASSIFICAÇÃO", "NOME PARCEIRO", "CPF PARCEIRO", "EQUIPE", "MOTIVO", "QTD PAGAMENTOS",
              "VALOR PELA REGRA", "VALOR DO PARCEIRO (PINCEL)", "PARECIDO COM (PLANILHA GERAL)"]].reset_index(drop=True)


def cruzamentos(df, cad):
    """Nomes do Pincel ligados ao cadastro por CPF/CNPJ, nome contido, nome parecido ou apelido."""
    b = df[df["_parceiro_id"].notna() & df["CRUZAMENTO"].ne("NOME IGUAL")].copy()
    if b.empty:
        return pd.DataFrame()
    b["_doc_txt"] = b["CPF PARCEIRO"].map(lambda v: "" if vazio(v) else str(v))
    g = b.groupby(["NOME PARCEIRO", "_doc_txt", "_parceiro_id", "CRUZAMENTO"], sort=False)
    r = g.agg(**{"QTD PAGAMENTOS": ("VALOR PAGTO", "size"),
                 "VALOR DO PARCEIRO CALCULADO": ("VALOR DO PARCEIRO CALCULADO", "sum")}).reset_index()
    r["NOME NA PLANILHA GERAL"] = r["_parceiro_id"].map(cad["NOME"])
    r["CPF/CNPJ NA PLANILHA GERAL"] = r["_parceiro_id"].map(cad["CPF"])
    r["VALOR DO PARCEIRO CALCULADO"] = [arred1(v) for v in r["VALOR DO PARCEIRO CALCULADO"]]
    r = r.rename(columns={"NOME PARCEIRO": "NOME NO PINCEL", "_doc_txt": "CPF/CNPJ NO PINCEL",
                          "CRUZAMENTO": "COMO FOI LIGADO"})
    return r[["NOME NO PINCEL", "CPF/CNPJ NO PINCEL", "NOME NA PLANILHA GERAL", "CPF/CNPJ NA PLANILHA GERAL",
              "COMO FOI LIGADO", "QTD PAGAMENTOS", "VALOR DO PARCEIRO CALCULADO"]].sort_values(
        ["COMO FOI LIGADO", "NOME NA PLANILHA GERAL"]).reset_index(drop=True)
