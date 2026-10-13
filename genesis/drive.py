"""Leitura dos arquivos de dados numa pasta do Google Drive (versão web).

A pasta é compartilhada (como Leitor) com a conta de serviço do Google Cloud. As credenciais da conta e o ID
da pasta ficam no .streamlit/secrets.toml (modelo em .streamlit/secrets.exemplo.toml).
"""
from pathlib import Path

from .util import norm

ESCOPO = ["https://www.googleapis.com/auth/drive.readonly"]
PLANILHA_GOOGLE = "application/vnd.google-apps.spreadsheet"
EXPORTAR = {".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".csv": "text/csv"}


def baixar_arquivos(credenciais, pasta_id, nomes):
    """{nome: (bytes, data de alteração no Drive)} dos arquivos da pasta com um dos `nomes` (com ou sem a
    extensão, sem diferença de maiúsculas/acentos). Com nomes repetidos vale o alterado por último.
    Arquivo convertido para Planilhas Google é exportado no formato da extensão do nome (.xlsx ou .csv)."""
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    cred = service_account.Credentials.from_service_account_info(dict(credenciais), scopes=ESCOPO)
    drive = build("drive", "v3", credentials=cred, cache_discovery=False)
    arquivos, pagina = [], None
    while True:
        resp = drive.files().list(q=f"'{pasta_id}' in parents and trashed = false", orderBy="modifiedTime desc",
                                  fields="nextPageToken, files(id, name, mimeType, modifiedTime)", pageToken=pagina,
                                  supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
        arquivos += resp.get("files", [])
        pagina = resp.get("nextPageToken")
        if not pagina:
            break
    baixados = {}
    for nome in nomes:
        aceitos = {norm(nome), norm(Path(nome).stem)}
        f = next((f for f in arquivos if norm(f["name"]) in aceitos), None)
        if f is None:
            continue
        if f["mimeType"] == PLANILHA_GOOGLE:
            pedido = drive.files().export_media(fileId=f["id"], mimeType=EXPORTAR[Path(nome).suffix.lower()])
        else:
            pedido = drive.files().get_media(fileId=f["id"], supportsAllDrives=True)
        baixados[nome] = (pedido.execute(), f["modifiedTime"])
    return baixados
