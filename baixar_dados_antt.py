"""
Baixa o recurso de municípios do portal de dados abertos da ANTT,
limpa o conteúdo (mantém apenas Concessionária + Município, sem repetições)
e gera uma planilha .xlsx formatada na pasta ./dados do repositório.

Executado automaticamente pelo GitHub Actions
(ver .github/workflows/download-semanal.yml) ou manualmente:

    python baixar_dados_antt.py

Dependências: requests, openpyxl
"""

import os
import sys
import shutil
from datetime import date, datetime, timezone, timedelta

import requests
import pandas as pd

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# CONFIGURAÇÃO — ajuste apenas esta seção
# ---------------------------------------------------------------------------

# ID do recurso (o trecho após /resource/ na URL do portal)
RESOURCE_ID = "f90fb6c6-9ecf-4b9d-86d7-153bdf0c1fd1"

# Colunas de origem (como vêm no CSV da ANTT) e os títulos que terão na planilha
COLUNAS_ORIGEM = ["concessionaria", "municipio"]
TITULOS_PLANILHA = ["Concessionária", "Município"]

# Nome "estável" do arquivo (URL fixa para o Power Automate)
NOME_ESTAVEL = "concessionaria_municipio.xlsx"

# Nome das abas da planilha
NOME_ABA = "Municípios"
NOME_ABA_METADADOS = "Metadados"

# Gerar a aba de metadados?
GERAR_ABA_METADADOS = True

# Texto fixo identificando a origem
ORIGEM_DOS_DADOS = "ANTT - Agência Nacional de Transportes Terrestres"

# Gerar também a versão em CSV, além do XLSX?
GERAR_CSV = False

# Guardar também uma cópia com a data no nome (histórico)?
GUARDAR_HISTORICO = False

PASTA_SAIDA = "dados"

# --- Aparência da planilha ---------------------------------------------------
FONTE = "Arial"
TAMANHO_FONTE = 11

COR_CABECALHO_FUNDO = "1F3864"   # azul-escuro institucional
COR_CABECALHO_TEXTO = "FFFFFF"   # branco
COR_LINHA_ALTERNADA = "EDF2F9"   # azul muito claro (zebra)
COR_BORDA = "B4C6E7"

LARGURA_MINIMA = 14
LARGURA_MAXIMA = 55
# -----------------------------------------------------------------------------

API_BASE = "https://dados.antt.gov.br/api/3/action"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Accept": "*/*",
}

TIMEOUT = 180

ENCODINGS = ["utf-8-sig", "utf-8", "cp1252", "latin-1"]

# Fuso de Brasília (UTC-3) — o runner do GitHub roda em UTC
FUSO_BR = timezone(timedelta(hours=-3))


# ---------------------------------------------------------------------------
# Download
# ---------------------------------------------------------------------------


def obter_recurso(resource_id: str) -> dict:
    """Consulta a API CKAN da ANTT e devolve os metadados do recurso."""
    url = f"{API_BASE}/resource_show?id={resource_id}"
    print(f"[1/5] Consultando metadados: {url}")

    resp = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
    resp.raise_for_status()

    payload = resp.json()
    if not payload.get("success"):
        raise RuntimeError(f"API retornou success=false: {payload}")

    recurso = payload["result"]
    if not recurso.get("url"):
        raise RuntimeError("Recurso não possui campo 'url'.")

    print(f"      Nome:    {recurso.get('name')}")
    print(f"      Formato: {(recurso.get('format') or '').upper()}")
    return recurso


def baixar(url: str, destino: str) -> int:
    """Baixa o arquivo em streaming. Devolve o tamanho em bytes."""
    print("[2/5] Baixando arquivo bruto ...")

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


# ---------------------------------------------------------------------------
# Leitura e limpeza
# ---------------------------------------------------------------------------


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


def corrigir_mojibake(texto: str) -> str:
    """Conserta acentuação corrompida (ex.: 'BraganÃ§a' -> 'Bragança').

    Ocorre quando bytes UTF-8 foram interpretados como latin-1/cp1252.
    A correção só é aplicada se o resultado realmente melhorar o texto.
    """
    if not isinstance(texto, str):
        return texto

    # Marcadores típicos de texto UTF-8 lido como latin-1
    if not any(m in texto for m in ("Ã", "Â", "â€", "Ãƒ")):
        return texto

    try:
        corrigido = texto.encode("latin-1", errors="strict").decode(
            "utf-8", errors="strict"
        )
    except (UnicodeEncodeError, UnicodeDecodeError):
        return texto

    # Só aceita se sumiram os marcadores e não apareceu caractere de substituição
    if "\ufffd" in corrigido:
        return texto
    return corrigido


def normalizar_nome_coluna(nome: str) -> str:
    """Deixa o nome da coluna em minúsculas, sem espaços/BOM nas pontas."""
    return str(nome).strip().lstrip("\ufeff").lower()


def limpar(caminho_bruto: str) -> pd.DataFrame:
    """Mantém apenas as colunas desejadas, corrige acentos, remove duplicidades."""
    print("[3/5] Limpando os dados ...")

    df = ler_csv(caminho_bruto)
    print(f"      Linhas no arquivo bruto: {len(df):,}".replace(",", "."))

    mapa = {normalizar_nome_coluna(c): c for c in df.columns}

    faltantes = [c for c in COLUNAS_ORIGEM if c not in mapa]
    if faltantes:
        raise RuntimeError(
            f"Colunas não encontradas no arquivo: {faltantes}. "
            f"Colunas disponíveis: {list(df.columns)}"
        )

    df = df[[mapa[c] for c in COLUNAS_ORIGEM]]
    df.columns = COLUNAS_ORIGEM

    corrigidos = 0
    for col in COLUNAS_ORIGEM:
        serie = df[col].fillna("").astype(str)
        antes_txt = serie.copy()
        serie = serie.map(corrigir_mojibake)
        corrigidos += int((serie != antes_txt).sum())
        df[col] = serie.str.replace(r"\s+", " ", regex=True).str.strip()

    if corrigidos:
        print(f"      Acentuação corrigida em {corrigidos:,} células".replace(",", "."))

    antes = len(df)
    df = df[(df[COLUNAS_ORIGEM] != "").all(axis=1)]
    vazias = antes - len(df)
    if vazias:
        print(f"      Linhas descartadas por campo vazio: {vazias:,}".replace(",", "."))

    antes = len(df)
    df = df.drop_duplicates(subset=COLUNAS_ORIGEM, keep="first")
    print(f"      Duplicidades removidas: {antes - len(df):,}".replace(",", "."))

    df = df.sort_values(by=COLUNAS_ORIGEM, kind="stable").reset_index(drop=True)

    print(f"      Linhas no arquivo final: {len(df):,}".replace(",", "."))
    print(
        f"      Concessionárias distintas: {df[COLUNAS_ORIGEM[0]].nunique()} | "
        f"Municípios distintos: {df[COLUNAS_ORIGEM[1]].nunique()}"
    )

    if df.empty:
        raise RuntimeError("Resultado da limpeza ficou vazio — verifique a origem.")

    return df


# ---------------------------------------------------------------------------
# Geração da planilha formatada
# ---------------------------------------------------------------------------


def montar_metadados(df: pd.DataFrame, recurso: dict, momento: datetime) -> list:
    """Monta a lista de pares (Propriedade, Valor) da aba de metadados."""
    return [
        ("Origem dos Dados", ORIGEM_DOS_DADOS),
        ("Recurso Selecionado", recurso.get("name") or "(sem nome)"),
        ("URL do Recurso", recurso.get("url")),
        ("Data de Extração", momento.strftime("%Y-%m-%d %H:%M:%S")),
        ("Total de Registros Únicos", len(df)),
        ("Concessionárias Distintas", int(df[COLUNAS_ORIGEM[0]].nunique())),
        ("Municípios Distintos", int(df[COLUNAS_ORIGEM[1]].nunique())),
        ("Última Atualização na ANTT", recurso.get("last_modified")
         or recurso.get("created") or "(não informado)"),
        ("ID do Recurso", RESOURCE_ID),
        ("Gerado por", "GitHub Actions — repositório dados-antt"),
    ]


def escrever_aba_metadados(wb: Workbook, metadados: list) -> None:
    """Cria e formata a aba de metadados no padrão Propriedade | Valor."""
    ws = wb.create_sheet(NOME_ABA_METADADOS)

    fonte_cabecalho = Font(
        name=FONTE, size=TAMANHO_FONTE, bold=True, color=COR_CABECALHO_TEXTO
    )
    fundo_cabecalho = PatternFill("solid", start_color=COR_CABECALHO_FUNDO)
    fonte_propriedade = Font(name=FONTE, size=TAMANHO_FONTE, bold=True)
    fonte_valor = Font(name=FONTE, size=TAMANHO_FONTE)
    alinh = Alignment(horizontal="left", vertical="center")

    ws.append(["Propriedade", "Valor"])
    for c in (1, 2):
        cel = ws.cell(row=1, column=c)
        cel.font = fonte_cabecalho
        cel.fill = fundo_cabecalho
        cel.alignment = alinh
    ws.row_dimensions[1].height = 22

    for i, (propriedade, valor) in enumerate(metadados, start=2):
        ws.cell(row=i, column=1, value=propriedade).font = fonte_propriedade
        cel_valor = ws.cell(row=i, column=2, value=valor)
        cel_valor.font = fonte_valor
        # Números gravados como número (evita o aviso verde do Excel)
        if isinstance(valor, int):
            cel_valor.number_format = "#,##0"
        ws.cell(row=i, column=1).alignment = alinh
        cel_valor.alignment = alinh
        ws.row_dimensions[i].height = 18

    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 95
    ws.sheet_view.showGridLines = False


def gerar_xlsx(df: pd.DataFrame, destino: str, metadados: list | None = None) -> None:
    """Grava o DataFrame em .xlsx com formatação de tabela."""
    print("[4/5] Gerando planilha formatada ...")

    wb = Workbook()
    ws = wb.active
    ws.title = NOME_ABA

    n_col = len(TITULOS_PLANILHA)

    fonte_cabecalho = Font(
        name=FONTE, size=TAMANHO_FONTE, bold=True, color=COR_CABECALHO_TEXTO
    )
    fundo_cabecalho = PatternFill("solid", start_color=COR_CABECALHO_FUNDO)
    fonte_corpo = Font(name=FONTE, size=TAMANHO_FONTE)
    fundo_zebra = PatternFill("solid", start_color=COR_LINHA_ALTERNADA)

    lado = Side(style="thin", color=COR_BORDA)
    borda = Border(left=lado, right=lado, top=lado, bottom=lado)

    alinh_cabecalho = Alignment(horizontal="left", vertical="center")
    alinh_corpo = Alignment(horizontal="left", vertical="center")

    # Cabeçalho
    ws.append(TITULOS_PLANILHA)
    for c in range(1, n_col + 1):
        cel = ws.cell(row=1, column=c)
        cel.font = fonte_cabecalho
        cel.fill = fundo_cabecalho
        cel.alignment = alinh_cabecalho
        cel.border = borda
    ws.row_dimensions[1].height = 24

    # Corpo
    for i, registro in enumerate(df.itertuples(index=False), start=2):
        ws.append(list(registro))
        zebra = i % 2 == 0
        for c in range(1, n_col + 1):
            cel = ws.cell(row=i, column=c)
            cel.font = fonte_corpo
            cel.alignment = alinh_corpo
            cel.border = borda
            if zebra:
                cel.fill = fundo_zebra
        ws.row_dimensions[i].height = 18

    ultima_linha = ws.max_row

    # Filtro, congelamento de painel e largura das colunas
    ws.auto_filter.ref = f"A1:{get_column_letter(n_col)}{ultima_linha}"
    ws.freeze_panes = "A2"

    for c in range(1, n_col + 1):
        letra = get_column_letter(c)
        maior = max(
            [len(str(TITULOS_PLANILHA[c - 1]))]
            + [len(str(v)) for v in df.iloc[:, c - 1].tolist()]
        )
        largura = min(max(maior + 7, LARGURA_MINIMA), LARGURA_MAXIMA)
        ws.column_dimensions[letra].width = largura

    ws.sheet_view.showGridLines = False

    if metadados:
        escrever_aba_metadados(wb, metadados)
        print(f"      Aba '{NOME_ABA_METADADOS}' criada com {len(metadados)} itens")

    wb.active = 0  # abre sempre na aba de dados
    wb.save(destino)
    print(f"      Planilha salva: {destino} ({ultima_linha - 1} linhas de dados)")


# ---------------------------------------------------------------------------


def main() -> int:
    os.makedirs(PASTA_SAIDA, exist_ok=True)
    caminho_bruto = os.path.join(PASTA_SAIDA, "_bruto_temp.csv")

    momento = datetime.now(timezone.utc).astimezone(FUSO_BR)

    try:
        recurso = obter_recurso(RESOURCE_ID)
        baixar(recurso["url"], caminho_bruto)
        df = limpar(caminho_bruto)
    except Exception as e:  # noqa: BLE001
        print(f"ERRO: {e}", file=sys.stderr)
        if os.path.exists(caminho_bruto):
            os.remove(caminho_bruto)
        return 1

    caminho_estavel = os.path.join(PASTA_SAIDA, NOME_ESTAVEL)
    metadados = (
        montar_metadados(df, recurso, momento) if GERAR_ABA_METADADOS else None
    )

    try:
        gerar_xlsx(df, caminho_estavel, metadados)
    except Exception as e:  # noqa: BLE001
        print(f"ERRO ao gerar a planilha: {e}", file=sys.stderr)
        os.remove(caminho_bruto)
        return 1

    if GERAR_CSV:
        caminho_csv = os.path.join(
            PASTA_SAIDA, os.path.splitext(NOME_ESTAVEL)[0] + ".csv"
        )
        df_csv = df.copy()
        df_csv.columns = TITULOS_PLANILHA
        df_csv.to_csv(caminho_csv, sep=";", index=False, encoding="utf-8-sig")
        print(f"      CSV adicional: {caminho_csv}")

    if GUARDAR_HISTORICO:
        base, ext = os.path.splitext(NOME_ESTAVEL)
        caminho_datado = os.path.join(
            PASTA_SAIDA, f"{base}-{date.today().isoformat()}{ext}"
        )
        shutil.copyfile(caminho_estavel, caminho_datado)
        print(f"[5/5] Cópia histórica: {caminho_datado}")
    else:
        print("[5/5] Histórico desativado.")

    os.remove(caminho_bruto)

    print("\nConcluído com sucesso.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
