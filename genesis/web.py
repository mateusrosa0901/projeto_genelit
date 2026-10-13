"""Execução do cálculo para a versão web (Streamlit).

Cada execução usa uma pasta temporária no lugar das pastas do projeto (entrada, dados, referencias, saida):
os arquivos recebidos são gravados nela, o cálculo roda como no main.py e a pasta do mês vira um .zip.
A pasta é apagada no fim, então nada dos dados fica no servidor.

O config é global (as pastas ficam nele), por isso só uma execução roda por vez.
"""
import contextlib
import io
import shutil
import tempfile
import threading
import zipfile
from pathlib import Path

import config
from .processo import gerar_saidas, imprimir_resumo, processar
from .util import MESES

# Arquivos da pasta "dados" que não vão para o GitHub (vêm por upload ou do Google Drive): nome -> obrigatório
ARQUIVOS_DADOS = {config.ARQ_PLANILHA_GERAL: True, config.ARQ_POLOS: True, config.ARQ_EXCECOES: True,
                  config.ARQ_AJUSTES_PARCEIROS: False}
# Arquivos de "dados" que estão no repositório e são copiados para a pasta temporária
ARQUIVOS_DADOS_REPOSITORIO = [config.ARQ_COMISSAO]

_TRAVA = threading.Lock()
_PASTAS = ("PASTA_PROJETO", "PASTA_ENTRADA", "PASTA_DADOS", "PASTA_REFERENCIAS", "PASTA_SAIDA")


@contextlib.contextmanager
def _pastas_temporarias(raiz):
    originais = {n: getattr(config, n) for n in _PASTAS}
    try:
        config.PASTA_PROJETO = raiz
        for n, sub in (("PASTA_ENTRADA", "entrada"), ("PASTA_DADOS", "dados"), ("PASTA_REFERENCIAS", "referencias"),
                       ("PASTA_SAIDA", "saida")):
            setattr(config, n, raiz / sub)
            (raiz / sub).mkdir()
        yield originais["PASTA_DADOS"]
    finally:
        for n, v in originais.items():
            setattr(config, n, v)


def _zip(pasta):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for arq in sorted(pasta.rglob("*")):
            if arq.is_file():
                z.write(arq, Path(pasta.name) / arq.relative_to(pasta))
    return buf.getvalue()


def executar(receita, dados, referencias=()):
    """receita: (nome do arquivo, bytes) do relatório do Pincel; dados: {nome em ARQUIVOS_DADOS: bytes};
    referencias: [(nome, bytes)] das planilhas feitas à mão (opcional, só conferência).
    Devolve {"zip", "nome_zip", "referencia", "registro"}. Erros do cálculo sobem com o registro em .registro."""
    falta = [n for n, obrigatorio in ARQUIVOS_DADOS.items() if obrigatorio and n not in dados]
    if falta:
        raise FileNotFoundError(f"Arquivos de dados ausentes: {', '.join(falta)}")
    registro = io.StringIO()
    with _TRAVA, tempfile.TemporaryDirectory(prefix="genesis_") as tmp, _pastas_temporarias(Path(tmp)) as dados_repo:
        nome, conteudo = receita
        sufixo = Path(nome).suffix.lower() or ".xlsx"
        (config.PASTA_ENTRADA / f"relatorioReceitaPolo_convertido{sufixo}").write_bytes(conteudo)
        for n, conteudo in dados.items():
            (config.PASTA_DADOS / n).write_bytes(conteudo)
        for n in ARQUIVOS_DADOS_REPOSITORIO:
            shutil.copy2(dados_repo / n, config.PASTA_DADOS / n)
        for n, conteudo in referencias:
            (config.PASTA_REFERENCIAS / Path(n).name).write_bytes(conteudo)
        try:
            with contextlib.redirect_stdout(registro):
                r = processar()
                pasta, _, _ = gerar_saidas(r, caminho_completo=False)
                imprimir_resumo(r)
        except Exception as erro:
            erro.registro = registro.getvalue()
            raise
        ref = r["referencia"]
        return {"zip": _zip(pasta), "nome_zip": f"{pasta.name} - {ref.year}.zip",
                "referencia": f"{MESES[ref.month - 1]}/{ref.year}", "registro": registro.getvalue()}
