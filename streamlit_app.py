"""
PROJETO GENESIS - versão web (Streamlit)
========================================
Upload do relatório "Receita Polo" do Pincel -> download da pasta do mês em .zip (a mesma pasta que o main.py
gera em saida/<ANO>/<MM>. REF. <MÊS>/).

Arquivos de dados (PLANILHA GERAL, D_Polos, D_Excecoes e, opcional, D_Parceiros_Ajustes):
  * sem [drive] no .streamlit/secrets.toml: upload na própria página;
  * com [drive] e [gcp_service_account]: lidos da pasta do Google Drive configurada.
Com "senha_app" no secrets.toml a página pede senha.

Rodar no computador:  streamlit run streamlit_app.py
"""
import hmac
from pathlib import Path

import pandas as pd
import streamlit as st

from genesis import __version__
from genesis.drive import baixar_arquivos
from genesis.web import ARQUIVOS_DADOS, executar

st.set_page_config(page_title="Genesis - Repasse", page_icon="📊")


def segredo(chave):
    try:
        return st.secrets.get(chave)
    except FileNotFoundError:  # sem secrets.toml
        return None


def autorizado():
    senha = segredo("senha_app")
    if not senha or st.session_state.get("autorizado"):
        return True
    with st.form("senha"):
        digitada = st.text_input("Senha de acesso", type="password")
        if st.form_submit_button("Entrar"):
            if hmac.compare_digest(digitada.encode(), str(senha).encode()):
                st.session_state["autorizado"] = True
                st.rerun()
            st.error("Senha incorreta.")
    return False


def dados_do_drive(drive):
    credenciais = segredo("gcp_service_account")
    if not credenciais or not drive.get("pasta_id"):
        raise ValueError("Configuração do Google Drive incompleta no secrets.toml: faltam [gcp_service_account] "
                         "ou pasta_id em [drive].")
    baixados = baixar_arquivos(credenciais, drive["pasta_id"], list(ARQUIVOS_DADOS))
    falta = [n for n, obrigatorio in ARQUIVOS_DADOS.items() if obrigatorio and n not in baixados]
    if falta:
        raise FileNotFoundError(f"Não encontrei na pasta do Google Drive: {', '.join(falta)}. Confira os nomes e se "
                                f"a pasta está compartilhada com {credenciais.get('client_email')}.")
    for nome, (_, alterado) in baixados.items():
        quando = pd.Timestamp(alterado).tz_convert("America/Sao_Paulo").strftime("%d/%m/%Y %H:%M")
        st.write(f"Google Drive: {nome} (alterado em {quando})")
    return {nome: conteudo for nome, (conteudo, _) in baixados.items()}


st.title("Genesis — Repasse de Parceiros Digitais e Polos")
st.caption(f"Versão {__version__}")
if not autorizado():
    st.stop()

st.markdown("No Pincel: **Gestão > Receita Polo > Geral > Mês/Ano > Gerar Relatório** e envie o arquivo abaixo.")
receita = st.file_uploader("Relatório Receita Polo (relatorioReceitaPolo_convertido.xlsx)", type=["xlsx", "xls", "csv"])

drive = segredo("drive")
enviados = {}
st.subheader("Cadastros")
if drive:
    st.caption("PLANILHA GERAL, D_Polos, D_Excecoes e D_Parceiros_Ajustes são lidos da pasta do Google Drive.")
else:
    st.caption("Arquivos da pasta `dados`. Ficam só durante o cálculo e são apagados em seguida.")
    for nome, obrigatorio in ARQUIVOS_DADOS.items():
        arq = st.file_uploader(nome + ("" if obrigatorio else " (opcional)"), type=[Path(nome).suffix[1:]], key=nome)
        if arq:
            enviados[nome] = arq.getvalue()

with st.expander("Conferência com planilhas feitas à mão (opcional)"):
    referencias = st.file_uploader("Relatórios de parceiro, PLANILHA GERAL e relatórios de polo do mesmo mês. Não "
                                   "entram no cálculo: aparecem nas abas de conferência do TOTAL REPASSE.",
                                   type=["xlsx", "xls"], accept_multiple_files=True)

pronto = receita is not None and (drive or all(n in enviados for n, o in ARQUIVOS_DADOS.items() if o))
if st.button("Calcular repasse", type="primary", disabled=not pronto):
    st.session_state.pop("resultado", None)
    with st.status("Calculando o repasse...", expanded=True) as status:
        try:
            dados = dados_do_drive(drive) if drive else enviados
            st.write("Calculando e gerando os relatórios (leva cerca de 1 minuto)...")
            st.session_state["resultado"] = executar((receita.name, receita.getvalue()), dados,
                                                     [(f.name, f.getvalue()) for f in referencias or []])
            status.update(label="Repasse calculado.", state="complete", expanded=False)
        except (FileNotFoundError, ValueError, KeyError) as erro:
            status.update(label="Não foi possível calcular.", state="error")
            st.error(f"ERRO: {erro}")
            if getattr(erro, "registro", ""):
                st.code(erro.registro, language=None)
        except Exception as erro:
            status.update(label="Erro inesperado.", state="error")
            st.exception(erro)

resultado = st.session_state.get("resultado")
if resultado:
    st.success(f"Referência: {resultado['referencia']}")
    st.download_button(f"Baixar {resultado['nome_zip']}", resultado["zip"], file_name=resultado["nome_zip"],
                       mime="application/zip", type="primary", on_click="ignore")
    st.code(resultado["registro"].strip(), language=None)
