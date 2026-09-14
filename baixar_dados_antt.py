"""
Baixa o recurso de municípios do portal de dados abertos da ANTT,
limpa o conteúdo (mantém apenas CONCESSIONARIA + MUNICIPIO, sem repetições)
e salva na pasta ./dados do repositório.
"""

import os
import sys
import shutil
from datetime import date

import requests
import pandas as pd

# ---------------------------------------------------------------------------
# CONFIGURAÇÃO — ajuste apenas esta seção
# ---------------------------------------------------------------------------

RESOURCE_ID = "f90fb6c6-9ecf-4b9d-86d7-153bdf0c1fd1"
COLUNAS_DESEJADAS = ["concessionaria", "municipio"]
NOME_ESTAVEL = "concessionaria_municipio.csv"
GUARDAR_HISTORICO = False
GUARDAR_BRUTO = False
SEP_SAIDA = ";"
PASTA_SAIDA = "dados"

API_BASE = "https://dados.antt.gov.br/api/3/action"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Accept": "*/*",
}

TIMEOUT = 180
ENCODINGS = ["utf-8-sig", "utf-8", "latin-1"]


# ---------------------------------------------------------------------------


def obter_url_do_recurso(resource_id: str) -> str:
    """Consulta a API CKAN da ANTT e devolve a URL de download do recurso."""
    url = f"{API_BASE}/resource_show?id={resource_id}"
    print(f"[1/4] Consultando metadados: {url}")

    resp = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
    resp.raise_for_status()

    payload = resp.json()
    if not payload.get("success"):
        raise RuntimeError(f"API retornou success=false: {payload}")

    recurso = payload["result"]
    url_download = recurso.get("url")
    if not url_download:
        raise RuntimeError("Recurso não possui campo 'url'.")

    print(f"      Nome:    {recurso.get('name')}")
    print(f"      Formato: {(recurso.get('format') or '').upper()}")
    print(f"      URL:     {url_download}")
    return url_download


def baixar(url: str, destino: str) -> int:
    """Baixa o arquivo em streaming. Devolve o tamanho em bytes."""
    print("[2/4] Baixando arquivo bruto ...")

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


def ler_csv(caminho: str) -> pd.DataFrame:
    """Lê o CSV testando codificações e separadores comuns."""
    ultimo_erro = None

    for enc in ENCODINGS:
        for sep in [";", ","]:
            try:
                df = pd.read_csv(
                    caminho,
                    sep=sep,
                    encoding=enc,
                    dtype=str,
                    engine="python",
                    on_bad_lines="skip",
                )
                if df.shape[1] < 2:
                    continue
                print(f"      Lido com encoding='{enc}', sep='{sep}'")
                return df
            except Exception as e:  # noqa: BLE001
                ultimo_erro = e
                continue

    raise RuntimeError(f"Não foi possível ler o CSV. Último erro: {ultimo_erro}")


def normalizar_nome_coluna(nome: str) -> str:
    """Deixa o nome da coluna em minúsculas, sem espaços/BOM nas pontas."""
    return str(nome).strip().lstrip("\ufeff").lower()


def limpar(caminho_bruto: str) -> pd.DataFrame:
    """Mantém apenas as colunas desejadas, remove vazios e duplicidades."""
    print("[3/4] Limpando os dados ...")

    df = ler_csv(caminho_bruto)
    print(f"      Linhas no arquivo bruto: {len(df):,}".replace(",", "."))

    mapa = {normalizar_nome_coluna(c): c for c in df.columns}

    faltantes = [c for c in COLUNAS_DESEJADAS if c not in mapa]
    if faltantes:
        raise RuntimeError(
            f"Colunas não encontradas no arquivo: {faltantes}. "
            f"Colunas disponíveis: {list(df.columns)}"
        )

    df = df[[mapa[c] for c in COLUNAS_DESEJADAS]]
    df.columns = COLUNAS_DESEJADAS

    for col in COLUNAS_DESEJADAS:
        df[col] = (
            df[col]
            .fillna("")
            .astype(str)
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
        )

    antes = len(df)
    df = df[(df[COLUNAS_DESEJADAS] != "").all(axis=1)]
    vazias = antes - len(df)
    if vazias:
        print(f"      Linhas descartadas por campo vazio: {vazias:,}".replace(",", "."))

    antes = len(df)
    df = df.drop_duplicates(subset=COLUNAS_DESEJADAS, keep="first")
    print(f"      Duplicidades removidas: {antes - len(df):,}".replace(",", "."))

    df = df.sort_values(by=COLUNAS_DESEJADAS, kind="stable").reset_index(drop=True)

    print(f"      Linhas no arquivo final: {len(df):,}".replace(",", "."))
    print(
        f"      Concessionárias distintas: {df[COLUNAS_DESEJADAS[0]].nunique()} | "
        f"Municípios distintos: {df[COLUNAS_DESEJADAS[1]].nunique()}"
    )

    if df.empty:
        raise RuntimeError("Resultado da limpeza ficou vazio — verifique a origem.")

    return df


def main() -> int:
    os.makedirs(PASTA_SAIDA, exist_ok=True)
    caminho_bruto = os.path.join(PASTA_SAIDA, "_bruto_temp.csv")

    try:
        url_download = obter_url_do_recurso(RESOURCE_ID)
        baixar(url_download, caminho_bruto)
        df = limpar(caminho_bruto)
    except Exception as e:  # noqa: BLE001
        print(f"ERRO: {e}", file=sys.stderr)
        if os.path.exists(caminho_bruto):
            os.remove(caminho_bruto)
        return 1

    caminho_estavel = os.path.join(PASTA_SAIDA, NOME_ESTAVEL)
    df.to_csv(caminho_estavel, sep=SEP_SAIDA, index=False, encoding="utf-8-sig")
    print(f"[4/4] Arquivo gerado: {caminho_estavel}")

    if GUARDAR_HISTORICO:
        base, ext = os.path.splitext(NOME_ESTAVEL)
        caminho_datado = os.path.join(
            PASTA_SAIDA, f"{base}-{date.today().isoformat()}{ext}"
        )
        shutil.copyfile(caminho_estavel, caminho_datado)
        print(f"      Cópia histórica: {caminho_datado}")

    if GUARDAR_BRUTO:
        destino_bruto = os.path.join(PASTA_SAIDA, "bruto_completo.csv")
        shutil.move(caminho_bruto, destino_bruto)
        print(f"      Arquivo bruto mantido: {destino_bruto}")
    else:
        os.remove(caminho_bruto)

    print("\nConcluído com sucesso.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
