"""
Baixa um recurso (planilha/CSV) do portal de dados abertos da ANTT
e salva na pasta ./dados do repositório.

Executado automaticamente pelo GitHub Actions (ver .github/workflows/download-semanal.yml)
ou manualmente:  python baixar_dados_antt.py
"""

import os
import sys
import shutil
from datetime import date

import requests

# ---------------------------------------------------------------------------
# CONFIGURAÇÃO — ajuste apenas esta seção
# ---------------------------------------------------------------------------

# ID do recurso (o trecho após /resource/ na URL do portal)
RESOURCE_ID = "f90fb6c6-9ecf-4b9d-86d7-153bdf0c1fd1"

# Nome "estável" do arquivo. É este que o Power Automate vai consumir sempre
# na mesma URL, sem precisar saber a data.
NOME_ESTAVEL = "dados_antt_ultimo.csv"

# Guardar também uma cópia com a data no nome (histórico)?
GUARDAR_HISTORICO = True

PASTA_SAIDA = "dados"

API_BASE = "https://dados.antt.gov.br/api/3/action"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Accept": "*/*",
}

TIMEOUT = 120


# ---------------------------------------------------------------------------


def obter_url_do_recurso(resource_id: str) -> tuple[str, str]:
    """Consulta a API CKAN da ANTT e devolve (url_download, formato)."""
    url = f"{API_BASE}/resource_show?id={resource_id}"
    print(f"[1/3] Consultando metadados: {url}")

    resp = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
    resp.raise_for_status()

    payload = resp.json()
    if not payload.get("success"):
        raise RuntimeError(f"API retornou success=false: {payload}")

    recurso = payload["result"]
    url_download = recurso.get("url")
    formato = (recurso.get("format") or "").upper()

    if not url_download:
        raise RuntimeError("Recurso não possui campo 'url'.")

    print(f"      Nome:    {recurso.get('name')}")
    print(f"      Formato: {formato}")
    print(f"      URL:     {url_download}")
    return url_download, formato


def baixar(url: str, destino: str) -> int:
    """Baixa o arquivo em streaming. Devolve o tamanho em bytes."""
    print(f"[2/3] Baixando para {destino} ...")

    with requests.get(url, headers=HEADERS, timeout=TIMEOUT, stream=True) as r:
        r.raise_for_status()
        with open(destino, "wb") as f:
            for bloco in r.iter_content(chunk_size=65536):
                if bloco:
                    f.write(bloco)

    tamanho = os.path.getsize(destino)
    if tamanho == 0:
        raise RuntimeError("Arquivo baixado está vazio.")

    print(f"      OK — {tamanho:,} bytes".replace(",", "."))
    return tamanho


def main() -> int:
    os.makedirs(PASTA_SAIDA, exist_ok=True)

    try:
        url_download, formato = obter_url_do_recurso(RESOURCE_ID)
    except Exception as e:
        print(f"ERRO ao obter metadados do recurso: {e}", file=sys.stderr)
        return 1

    # Mantém a extensão original se ela for diferente de .csv
    extensao = os.path.splitext(url_download.split("?")[0])[1] or ".csv"
    nome_estavel = os.path.splitext(NOME_ESTAVEL)[0] + extensao
    caminho_estavel = os.path.join(PASTA_SAIDA, nome_estavel)

    try:
        baixar(url_download, caminho_estavel)
    except Exception as e:
        print(f"ERRO ao baixar o arquivo: {e}", file=sys.stderr)
        return 1

    if GUARDAR_HISTORICO:
        nome_datado = (
            f"{os.path.splitext(nome_estavel)[0]}-{date.today().isoformat()}{extensao}"
        )
        caminho_datado = os.path.join(PASTA_SAIDA, nome_datado)
        shutil.copyfile(caminho_estavel, caminho_datado)
        print(f"[3/3] Cópia histórica criada: {caminho_datado}")
    else:
        print("[3/3] Histórico desativado.")

    print("\nConcluído com sucesso.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
