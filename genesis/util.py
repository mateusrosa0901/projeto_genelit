"""Funções de apoio: texto, nomes, CPF/CNPJ, datas, valores, arredondamento e leitura de CSV."""
import calendar
import re
import unicodedata
from difflib import SequenceMatcher

import numpy as np
import pandas as pd

MESES = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO",
         "NOVEMBRO", "DEZEMBRO"]
PALAVRAS_VAZIAS = {"DE", "DA", "DO", "DAS", "DOS", "E", "LTDA", "ME", "EIRELI", "EPP", "MEI", "SA"}


# ------------------------------------- texto -------------------------------------
def vazio(v):
    """None, NaN, NaT ou texto em branco."""
    if v is None:
        return True
    try:
        if pd.isna(v):
            return True
    except (TypeError, ValueError):
        return False
    return str(v).strip() == ""


def ap(serie, f):
    """Aplica f item a item mantendo objetos Python (mesmo resultado no pandas 2 e 3)."""
    return pd.Series([f(v) for v in serie], index=serie.index, dtype=object)


def sem_acento(t):
    return unicodedata.normalize("NFKD", str(t)).encode("ascii", "ignore").decode("ascii")


def norm(v):
    """MAIÚSCULAS, sem acento e sem espaços duplicados (vazio -> None)."""
    return None if vazio(v) else (re.sub(r"\s+", " ", sem_acento(str(v)).upper()).strip() or None)


def tokens(v):
    t = norm(v)
    return frozenset(re.findall(r"[A-Z0-9]+", t)) if t else frozenset()


def nome_base(v):
    """Nome só com letras: '64.329.952 BRUNA APARECIDA SANTOS' -> 'BRUNA APARECIDA SANTOS',
    'Brisa Santana Pires (COMERCIAL SÃO JOSÉ)' -> 'BRISA SANTANA PIRES'."""
    t = norm(v)
    if not t:
        return None
    t = re.sub(r"[^A-Z ]", " ", re.sub(r"\(.*?\)|\(.*$", " ", t))
    return re.sub(r"\s+", " ", t).strip() or None


def palavras(nome):
    """Palavras que identificam o nome (sem DE, DA, DOS, LTDA...)."""
    return [p for p in nome.split() if p not in PALAVRAS_VAZIAS and len(p) > 1] if isinstance(nome, str) else []


def mesmo_nome(a, b):
    """'IEDA ALVES PAIVA' x 'IEDA ALVES DE PAIVA': mesmo 1º nome e um contido no outro (2+ palavras)."""
    if not isinstance(a, str) or not isinstance(b, str):
        return False
    if a == b:
        return True
    pa, pb = palavras(a), palavras(b)
    return min(len(pa), len(pb)) >= 2 and pa[0] == pb[0] and (set(pa) <= set(pb) or set(pb) <= set(pa))


def semelhanca(a, b):
    return SequenceMatcher(None, a, b).ratio() if isinstance(a, str) and isinstance(b, str) else 0.0


def txt_int(v):
    """123.0 -> '123' (o Excel transforma números em float)."""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    t = str(v).strip()
    return t.split(".")[0] if re.fullmatch(r"\d+\.0+", t) else t


# ---------------------------------- CPF / CNPJ ----------------------------------
def documento(v):
    """CPF/CNPJ só com dígitos e sem zeros à esquerda; documentos fictícios (999..., 000...) -> None."""
    if vazio(v):
        return None
    d = re.sub(r"\D", "", txt_int(v)).lstrip("0")
    return None if len(d) < 6 or len(set(d)) == 1 else d


def docs_compativeis(a, b):
    """Mesmo documento, aceitando CNPJ sem os dígitos finais (o Pincel às vezes corta: '220175350001')."""
    if not a or not b:
        return False
    if a == b:
        return True
    curto, longo = sorted((a, b), key=len)
    return len(curto) >= 11 and longo.startswith(curto)


def docs_conflitam(a, b):
    """Dois documentos do mesmo tipo (CPF x CPF ou CNPJ x CNPJ) e diferentes = pessoas diferentes."""
    if not a or not b or docs_compativeis(a, b):
        return False
    return (len(a) <= 11) == (len(b) <= 11)


def tipo_pessoa(v):
    if vazio(v):
        return ""
    t = txt_int(v)
    d = re.sub(r"\D", "", t).lstrip("0")
    return "PJ" if "/" in t or len(d) > 11 else ("PF" if d else "")


# ------------------------------- números e datas -------------------------------
def arred(v):
    """Centavos 'meio para cima', igual ao ARRED do Excel (o round() do Python arredonda 0,5 para o par)."""
    v = pd.to_numeric(pd.Series(v), errors="coerce").astype(float)
    return np.sign(v) * np.floor(np.round(v.abs() * 100, 6) + 0.5) / 100


def arred1(x):
    """arred() para um único número."""
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return 0.0
    return float(np.sign(x) * np.floor(np.round(abs(x) * 100, 6) + 0.5) / 100)


def datas(s):
    """Datas do Excel, 'dd/mm/aaaa' ou 'aaaa-mm-dd'; dia inexistente (29/02/2025) vira o último dia do mês."""
    if pd.api.types.is_datetime64_any_dtype(s):
        return pd.to_datetime(s).dt.normalize()

    def uma(v):
        if vazio(v):
            return pd.NaT
        if hasattr(v, "year"):
            return pd.Timestamp(v).normalize()
        t = str(v).strip()[:10]
        m = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{4})", t)
        if m:
            d, mes, a = map(int, m.groups())
        else:
            m = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})", t)
            if not m:
                return pd.to_datetime(t, dayfirst=True, errors="coerce")
            a, mes, d = map(int, m.groups())
        return pd.Timestamp(a, mes, min(d, calendar.monthrange(a, mes)[1])) if 1 <= mes <= 12 else pd.NaT
    return pd.to_datetime(pd.Series([uma(v) for v in s], index=s.index))


def uma_data(v):
    """Uma data (texto ou data) -> Timestamp ou NaT."""
    return datas(pd.Series([v], dtype=object)).iloc[0]


def valores(s):
    """Número do Excel ou texto brasileiro ('1.260,00', 'R$ 265,50')."""
    if pd.api.types.is_numeric_dtype(s):
        return s.astype(float)

    def um(v):
        if vazio(v):
            return np.nan
        if isinstance(v, (int, float, np.number)):
            return float(v)
        t = str(v).replace("R$", "").replace(" ", "").strip()
        try:
            return float(t.replace(".", "").replace(",", ".") if "," in t else t)
        except ValueError:
            return np.nan
    return pd.Series([um(v) for v in s], index=s.index, dtype=float)


def pct(v):
    """'20%' | '0,2' | '0.2' | '20' -> 0.2"""
    if vazio(v):
        return np.nan
    t = str(v).strip().replace(",", ".")
    n = float(t.rstrip("%").strip())
    return n / 100 if t.endswith("%") or n > 1 else n


def moeda(v):
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def texto_pct(p):
    """0.2 -> '0.2' (sem zeros sobrando) para montar fórmulas do Excel."""
    return f"{float(p):.6f}".rstrip("0").rstrip(".")


# ------------------------------------ arquivos ------------------------------------
def ler_csv(caminho):
    """CSV como texto, separado por ',' ou ';' (Excel em português), em UTF-8 ou Latin-1."""
    df = None
    for cod in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            with open(caminho, encoding=cod) as f:
                cab = f.readline()
            df = pd.read_csv(caminho, dtype=str, encoding=cod, keep_default_na=False,
                             sep=";" if cab.count(";") > cab.count(",") else ",")
            break
        except UnicodeDecodeError:
            continue
    if df is None:
        raise ValueError(f"Não foi possível ler {caminho}")
    df.columns = [str(c).strip() for c in df.columns]
    df = df.astype(object)
    for c in df.columns:
        df[c] = ap(df[c], lambda v: None if vazio(v) else str(v).strip())
    return df


def exigir(df, colunas, onde):
    falta = [c for c in colunas if c not in df.columns]
    if falta:
        raise ValueError(f"Colunas ausentes {onde}: {falta}")


def nome_arquivo_seguro(texto, limite=120):
    t = re.sub(r'[\\/:*?"<>|\r\n\t]', "-", str(texto)).strip().strip(".")
    return re.sub(r"\s+", " ", t)[:limite] or "SEM NOME"
