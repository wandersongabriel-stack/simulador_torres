# app.py
# =============================
# PAINEL DE TORRES (Streamlit)
# - Card do topo = RESULTADO REAL (algoritmo)
# - Simulador de compra + nova coluna "Modelos (SKUs) a comprar" (ceil(faltante / N))
# - Tooltips nas colunas (hover) via column_config
# - Admin: pode subir base (na sessão); demais usuários não
# - Botão "Gerar kits agora" + cache (não recomputa toda hora)
# - Abas: simulador + relatórios (kits_resumo, kits_itens, estoque_restante, falha_proximo_kit)
# =============================

import time
import re
import io
import os
import random
from collections import defaultdict, deque

import pandas as pd
import numpy as np
import streamlit as st
import requests
import hmac


# =============================
# CONFIG GERAL
# =============================
st.set_page_config(page_title="Painel de Torres", layout="wide")

def check_login():
    if st.session_state.get("logged_in", False):
        return True

    # Exibe apenas a tela de login até autenticar
    st.title("Login")
    st.caption("Entre para acessar o simulador.")

    usuario = st.text_input("Usuário", key="login_user")
    senha = st.text_input("Senha", type="password", key="login_pass")

    if st.button("Entrar", type="primary"):
        user_ok = hmac.compare_digest(
            str(usuario),
            str(st.secrets["app_auth"]["user"])
        )
        pass_ok = hmac.compare_digest(
            str(senha),
            str(st.secrets["app_auth"]["password"])
        )

        if user_ok and pass_ok:
            st.session_state["logged_in"] = True
            st.session_state.pop("login_pass", None)
            st.rerun()
        else:
            st.error("Usuário ou senha inválidos.")

    st.stop()


if not check_login():
    st.stop()


TARGET_MIN_DEFAULT = 10000
TARGET_MAX_DEFAULT = 10090

RULES = {
    "CJ": (22, 25),
    "CK": (8, 12),
    "CO": (20, 28),
    "ES": (2, 3),
    "PF": (10, 15),
    "PR": (2, 3),
    "SEM": (1, 2),
    "PM": (2, 4),
    "C_FEMININO": (10, 15),
    "C_MASCULINO": (2, 4),
    "BR_TRIO": (8, 12),
    "BR_GRANDE": (8, 10),
    "BR_DEMAIS": (60, None),
}

PREFIX_DIRECT = ["CJ", "CK", "CO", "ES", "PF", "SEM", "PM", "PR"]
ADJUST_CATS = {"BR_DEMAIS", "CO"}  # categorias usadas como "ajuste fino"

DISPLAY_NAME = {
    "CJ": "CJ",
    "CK": "CK",
    "CO": "CO",
    "ES": "ES",
    "PF": "PF",
    "PR": "PR",
    "SEM": "SEM",
    "PM": "PM",
    "C_FEMININO": "C - FEMININO",
    "C_MASCULINO": "C - MASCULINO",
    "BR_TRIO": "BR - TRIO",
    "BR_GRANDE": "BR - GRANDE",
    "BR_DEMAIS": "BR - OUTROS",
}

# Gerador
DEFAULT_MAX_KITS = 200
ATTEMPTS_PER_KIT = 40
TABU_ITERS = 250
TABU_TENURE = 25
NEIGHBORHOOD_SAMPLES = 250
SEED = 7
random.seed(SEED)
np.random.seed(SEED)

# Base default no repo
DEFAULT_API_URL = "http://177.39.19.116/WebAPIFeniciaOCA/table/List2"
DEFAULT_API_TABLE = "SELECT grupo, referencia, qtdreal, prc_venda FROM CADMAT"


# =============================
# CSS (PowerBI-like + inputs legíveis)
# =============================
st.markdown(
    """
    <style>
    :root{
      --bg:#070a0f;
      --panel:#0f1420;
      --panel2:#0b111b;
      --border:#202a3a;
      --border2:#2a3750;
      --text:#e9eef8;
      --muted:#a7b4c7;
      --accent:#5aa9ff;
      --accent2:#7c5cff;
    }

    .stApp { background: radial-gradient(1200px 600px at 20% 0%, #0d1630 0%, var(--bg) 55%); }
    html, body, [class*="css"]  { color: var(--text); }

    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}
    [data-testid="stToolbar"] {visibility: hidden; height: 0px;}
    [data-testid="stDecoration"] {display: none;}

    .block-container { padding-top: 0.85rem; padding-bottom: 2rem; }

    section[data-testid="stSidebar"]{
      background: linear-gradient(180deg, #0b1220 0%, #070a0f 100%);
      border-right: 1px solid var(--border);
    }

    .topbar{
      background: linear-gradient(90deg, rgba(90,169,255,0.18) 0%, rgba(124,92,255,0.12) 45%, rgba(0,0,0,0) 100%);
      border: 1px solid var(--border);
      border-radius: 16px;
      padding: 14px 16px;
      box-shadow: 0 10px 30px rgba(0,0,0,0.35);
    }
    .topbar-title{
      font-size: 28px;
      font-weight: 800;
      letter-spacing: 0.8px;
      margin: 0;
      line-height: 1.15;
    }
    .topbar-sub{
      color: var(--muted);
      margin-top: 6px;
      font-size: 13px;
    }

    div[data-testid="stMetric"]{
      background: linear-gradient(180deg, var(--panel) 0%, var(--panel2) 100%);
      border: 1px solid var(--border);
      padding: 14px 14px 10px 14px;
      border-radius: 16px;
      box-shadow: 0 10px 30px rgba(0,0,0,0.35);
    }
    div[data-testid="stMetric"] label { color: var(--muted) !important; }
    div[data-testid="stMetric"] [data-testid="stMetricValue"]{
      font-size: 34px;
      font-weight: 800;
    }

    .stTabs [data-baseweb="tab-list"]{
      gap: 8px;
      border-bottom: 1px solid var(--border);
      padding-bottom: 8px;
    }
    .stTabs [data-baseweb="tab"]{
      background: rgba(255,255,255,0.03);
      border: 1px solid var(--border);
      border-radius: 999px;
      padding: 8px 14px;
      color: var(--muted);
    }
    .stTabs [aria-selected="true"]{
      background: linear-gradient(90deg, rgba(90,169,255,0.18), rgba(124,92,255,0.15));
      border: 1px solid var(--border2);
      color: var(--text);
    }

    .stButton>button, .stDownloadButton>button{
      border-radius: 12px !important;
      border: 1px solid var(--border2) !important;
      background: linear-gradient(180deg, rgba(90,169,255,0.16), rgba(90,169,255,0.06)) !important;
      color: var(--text) !important;
      box-shadow: 0 8px 24px rgba(0,0,0,0.35);
    }
    .stButton>button:hover, .stDownloadButton>button:hover{
      border-color: rgba(90,169,255,0.55) !important;
      transform: translateY(-1px);
    }

    div[data-baseweb="input"] input, textarea{
      background: rgba(255,255,255,0.03) !important;
      border: 1px solid var(--border) !important;
      color: var(--text) !important;
      border-radius: 12px !important;
    }

    /* Sidebar inputs com texto ESCURO (legível) */
    section[data-testid="stSidebar"] div[data-baseweb="input"] input{
      color: #0b0f18 !important;
      background: rgba(255,255,255,0.90) !important;
      border: 1px solid rgba(255,255,255,0.18) !important;
    }
    section[data-testid="stSidebar"] div[data-baseweb="input"] input::placeholder{
      color: rgba(11,15,24,0.55) !important;
    }
    section[data-testid="stSidebar"] label{
      color: rgba(233,238,248,0.92) !important;
    }
    section[data-testid="stSidebar"] button[aria-label="Increment"],
    section[data-testid="stSidebar"] button[aria-label="Decrement"]{
      background: rgba(255,255,255,0.85) !important;
      border: 1px solid rgba(255,255,255,0.18) !important;
    }
    section[data-testid="stSidebar"] button[aria-label="Increment"] svg,
    section[data-testid="stSidebar"] button[aria-label="Decrement"] svg{
      fill: #0b0f18 !important;
    }

    .dataframe-shell{
      background: linear-gradient(180deg, var(--panel) 0%, var(--panel2) 100%);
      border: 1px solid var(--border);
      border-radius: 16px;
      padding: 12px;
      box-shadow: 0 10px 30px rgba(0,0,0,0.35);
    }

    [data-testid="stDataFrame"]{
      border-radius: 12px;
      overflow: hidden;
      border: 1px solid rgba(255,255,255,0.06);
    }
    [data-testid="stDataFrame"] tbody tr:hover{
      background: rgba(90,169,255,0.08) !important;
    }

    hr { border: 0; height: 1px; background: var(--border); margin: 14px 0; }
    .muted { color: var(--muted); font-size: 13px; }
    </style>
    """,
    unsafe_allow_html=True
)


# =============================
# BASE PREP
# =============================
def norm_sku(s: str) -> str:
    return re.sub(r"\s+", "", str(s).strip()).upper()

def assign_category(row) -> str:
    sku = row["Sku_norm"]

    if int(row.get("BASE_Corrente_Feminina", 0) or 0) == 1:
        return "C_FEMININO"
    if int(row.get("BASE_Corrente_Masculina", 0) or 0) == 1:
        return "C_MASCULINO"

    if sku.startswith("BR"):
        if int(row.get("BASE_Trio", 0) or 0) == 1:
            return "BR_TRIO"
        if int(row.get("TIPO_Brinco_Grande", 0) or 0) == 1:
            return "BR_GRANDE"
        return "BR_DEMAIS"

    for p in PREFIX_DIRECT:
        if sku.startswith(p):
            return p

    return "OUTROS"

def preparar_base_from_df(df: pd.DataFrame) -> pd.DataFrame:
    for col in ["Sku", "Estoque", "Preco"]:
        if col not in df.columns:
            raise ValueError(f"A planilha precisa ter a coluna '{col}'.")

    df = df.copy()
    df["Sku"] = df["Sku"].astype(str)
    df["Sku_norm"] = df["Sku"].apply(norm_sku)

    df["Estoque"] = pd.to_numeric(df["Estoque"], errors="coerce").fillna(0).astype(int)
    df["Preco"] = pd.to_numeric(df["Preco"], errors="coerce")

    df["categoria"] = df.apply(assign_category, axis=1)

    allowed = set(RULES.keys())
    df = df[(df["Estoque"] > 0) & (df["Preco"].notna()) & (df["categoria"].isin(allowed))].copy()
    return df


# =============================
# API CADMAT + CLASSIFICAÇÃO (sem Excel no dia-a-dia)
# - Busca CADMAT via API (grupo, referencia, qtdreal, prc_venda)
# - Cria codigo = GRUPO+REFERENCIA (normalizado)
# - Faz JOIN com:
#   * categoria_produto.xlsx (BR: TRIO/GRANDE + validação de banho)
#   * genero_produto.xlsx (C: FEM/MASC)
# - Gera as flags esperadas pelo assign_category:
#   BASE_Corrente_Feminina, BASE_Corrente_Masculina, BASE_Trio, TIPO_Brinco_Grande
# - Remove itens sem cadastro ou com banho inválido
# =============================

def _to_str_clean(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if re.match(r"^\d+\.0$", s):
        s = s[:-2]
    return s

def make_codigo(grupo, referencia) -> str:
    g = _to_str_clean(grupo).upper().replace(" ", "")
    r = _to_str_clean(referencia).upper().replace(" ", "")
    return f"{g}{r}"

def norm_codigo(s) -> str:
    """Normaliza um código removendo tudo que não é letra/número e padronizando maiúsculo."""
    if pd.isna(s):
        return ""
    s = str(s).strip().upper()
    if s.endswith(".0"):
        s = s[:-2]
    s = re.sub(r"[^A-Z0-9]", "", s)
    return s

def parse_pt_decimal(x) -> float:
    """Converte '1.234,56' / '0,000' / '123' em float."""
    if pd.isna(x):
        return np.nan
    s = str(x).strip()
    if not s:
        return np.nan
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return np.nan

@st.cache_data(show_spinner=False)
def load_classificacao_from_bytes(cat_bytes: bytes, gen_bytes: bytes) -> pd.DataFrame:
    """
    Consolida as duas planilhas em um lookup por 'codigo'.

    Regras novas:
    - BR com banho vazio/nulo é inválido
    - BR com banho contendo 'Ródio' ou 'Rodio' é inválido
    - BR com banho 'Não definido' / 'Nao definido' é inválido
    - a invalidação não depende da API deixar de trazer o item;
      ela só serve para barrar o item na base tratada após o merge.
    """
    df_cat = pd.read_excel(io.BytesIO(cat_bytes)).copy()
    df_gen = pd.read_excel(io.BytesIO(gen_bytes)).copy()

    for d in (df_cat, df_gen):
        if "codigo" not in d.columns:
            raise ValueError("Planilhas de classificação precisam ter a coluna 'codigo'.")
        d["codigo"] = d["codigo"].apply(norm_codigo)

    # =============================
    # categoria_produto: trio / grande / validação de banho
    # =============================
    base = df_cat.get("base", pd.Series([""] * len(df_cat))).astype(str).str.strip().str.upper()
    modelo = df_cat.get("modelo", pd.Series([""] * len(df_cat))).astype(str).str.strip().str.upper()

    df_cat["is_trio"] = base.eq("TRIO") | modelo.str.contains("TRIO", na=False)
    df_cat["is_grande"] = modelo.str.contains("GRANDE", na=False)

    # marca que o código existe na categoria_produto
    df_cat["tem_cadastro_cat"] = True

    # regra do banho
    if "banho" in df_cat.columns:
        banho_raw = df_cat["banho"]

        banho_txt = banho_raw.fillna("").astype(str).str.strip().str.lower()

        tem_banho = banho_txt.ne("")
        banho_e_rodio = banho_txt.str.contains(r"rodio|ródio", case=False, regex=True, na=False)
        banho_nao_definido = banho_txt.str.contains(r"nao definido|não definido", case=False, regex=True, na=False)

        df_cat["banho_valido"] = tem_banho & (~banho_e_rodio) & (~banho_nao_definido)
    else:
        # se a coluna não existir, considera inválido para não deixar passar sem controle
        df_cat["banho_valido"] = False

    # se houver múltiplas linhas por código, consolida:
    # - trio/grande: se alguma linha marcar, mantém True
    # - tem_cadastro_cat: True
    # - banho_valido: se alguma linha for válida, considera válido
    df_cat_lookup = (
        df_cat.groupby("codigo", as_index=False)
        .agg(
            is_trio=("is_trio", "max"),
            is_grande=("is_grande", "max"),
            tem_cadastro_cat=("tem_cadastro_cat", "max"),
            banho_valido=("banho_valido", "max"),
        )
    )

    # =============================
    # genero_produto: feminino / masculino
    # =============================
    genero = df_gen.get("genero", pd.Series([""] * len(df_gen))).astype(str).str.strip().str.upper()

    df_gen["is_fem"] = genero.eq("MULHER") | genero.eq("FEM") | genero.eq("FEMININO")
    df_gen["is_masc"] = genero.eq("HOMEM") | genero.eq("MASC") | genero.eq("MASCULINO")

    df_gen_lookup = (
        df_gen[["codigo", "is_fem", "is_masc"]]
        .drop_duplicates("codigo")
        .copy()
    )

    cls = (
        df_cat_lookup
        .merge(df_gen_lookup, on="codigo", how="outer")
        .fillna({
            "is_trio": False,
            "is_grande": False,
            "tem_cadastro_cat": False,
            "banho_valido": False,
            "is_fem": False,
            "is_masc": False,
        })
    )

    return cls

def _api_get_table(
    base_url: str,
    usuario: str,
    senha: str,
    tabela_sql: str,
    timeout: int = 60,
    raise_on_error: bool = True,
) -> list[dict]:
    """Chama a API e devolve a lista de registros (Data['table'])."""
    headers = {"Usuario": usuario, "Senha": senha, "Tabela": tabela_sql}
    try:
        r = requests.get(base_url, headers=headers, timeout=timeout)
        if r.status_code == 404:
            return []
        r.raise_for_status()
    except requests.RequestException:
        if raise_on_error:
            raise
        return []

    try:
        data = r.json()
    except Exception:
        if raise_on_error:
            raise
        return []

    if isinstance(data, str):
        return []
    table = data.get("table") if isinstance(data, dict) else None
    if not table:
        return []
    if isinstance(table, list):
        return table
    return []


@st.cache_data(show_spinner=False, ttl=300)
def fetch_lookup_table(base_url: str, usuario: str, senha: str, table_name: str) -> pd.DataFrame:
    sql = f"SELECT codigo, nome FROM {table_name}"
    rows = _api_get_table(base_url, usuario, senha, sql, timeout=90, raise_on_error=False)
    if not rows:
        rows = _api_get_table(base_url, usuario, senha, table_name, timeout=90, raise_on_error=False)

    df = pd.DataFrame(rows)
    if df.empty:
        return pd.DataFrame(columns=["codigo", "nome"])

    df.columns = [str(c).strip().lower() for c in df.columns]
    if "codigo" not in df.columns:
        df["codigo"] = ""
    if "nome" not in df.columns:
        df["nome"] = ""

    df["codigo"] = df["codigo"].astype(str).str.strip().str.upper().str.zfill(2)
    df["nome"] = df["nome"].astype(str).str.strip().str.upper()
    return df[["codigo", "nome"]].drop_duplicates().copy()

def fetch_cadmat_paginado(
    base_url: str,
    usuario: str,
    senha: str,
    grupos: list[str],
    colunas: list[str] | None = None,
    only_stock_gt0: bool = True,
    max_pages_per_group: int = 5000,
    timeout: int = 60,
) -> pd.DataFrame:
    """Replica a paginação do Power BI: para cada grupo, pagina por referencia (referencia > last_ref)."""
    if colunas is None:
        colunas = ["grupo", "referencia", "qtdreal", "prc_venda"]
    cols_sql = ", ".join(colunas)

    all_rows: list[dict] = []
    for g in grupos:
        g_clean = _to_str_clean(g).strip()
        if not g_clean:
            continue

        last_ref = ""
        pages = 0

        while pages < max_pages_per_group:
            where_parts = [f"grupo = '{g_clean}'", f"referencia > '{last_ref}'"]
            if only_stock_gt0:
                where_parts.insert(1, "qtdreal > 0")
            where_sql = " AND ".join(where_parts)

            sql = f"SELECT {cols_sql} FROM CADMAT WHERE {where_sql} ORDER BY referencia"
            rows = _api_get_table(base_url, usuario, senha, sql, timeout=timeout)

            if not rows:
                break

            all_rows.extend(rows)

            try:
                last_ref_new = str(rows[-1].get("referencia", "")).strip()
            except Exception:
                last_ref_new = ""

            if not last_ref_new or last_ref_new == last_ref:
                break

            last_ref = last_ref_new
            pages += 1

    return pd.DataFrame(all_rows)

@st.cache_data(show_spinner=False, ttl=300)
def fetch_cadmat_full_bruto(
    base_url: str,
    usuario: str,
    senha: str,
    timeout: int = 90,
) -> pd.DataFrame:
    """
    Baixa a CADMAT bruta igual ao teste.py:
    - pagina por referencia > last_ref
    - sem merge com categoria/genero
    - sem filtros de banho
    - sem classificação do app
    """
    cols = (
        "grupo,referencia,descricao,caracter,qtdreal,pu_mat,prc_venda,"
        "unid_emb,qtd_emb,unid_mat,prc_venda2,prc_venda3,peso_unit,marca,"
        "prc_venda4,barra,val_larg,val_comp,barra_emb,ncm"
    )

    all_rows = []
    last_ref = ""

    while True:
        sql = f"SELECT {cols} FROM CADMAT WHERE referencia > '{last_ref}' ORDER BY referencia"
        rows = _api_get_table(base_url, usuario, senha, sql, timeout=timeout)

        if not rows:
            break

        all_rows.extend(rows)

        last_ref_new = str(rows[-1].get("referencia", "")).strip()
        if not last_ref_new or last_ref_new == last_ref:
            break

        last_ref = last_ref_new
        time.sleep(0.10)

    return pd.DataFrame(all_rows)


def df_single_excel_bytes(df: pd.DataFrame, sheet_name: str) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    bio.seek(0)
    return bio.getvalue()

def build_raw_base_from_api(
    cadmat_df: pd.DataFrame,
    tabcol_df: pd.DataFrame | None = None,
    tablin_df: pd.DataFrame | None = None,
    tabgrp_df: pd.DataFrame | None = None,
) -> tuple[pd.DataFrame, dict]:

    cad = cadmat_df.copy()
    cad.columns = [str(c).strip().lower() for c in cad.columns]

    needed = ["grupo", "referencia", "qtdreal", "prc_venda"]
    missing = [c for c in needed if c not in cad.columns]
    if missing:
        raise ValueError(f"CADMAT não trouxe colunas esperadas: {missing}")

    for opt in ["gradecol", "gradelin", "gradegrp"]:
        if opt not in cad.columns:
            cad[opt] = ""

    cad["Sku"] = cad.apply(
        lambda r: norm_codigo(make_codigo(r["grupo"], r["referencia"])),
        axis=1,
    )
    cad["Estoque"] = cad["qtdreal"].apply(parse_pt_decimal)
    cad["Preco"] = cad["prc_venda"].apply(parse_pt_decimal)
    cad["Estoque"] = np.floor(cad["Estoque"].fillna(0)).astype(int)
    cad = cad[(cad["Estoque"] > 0) & (cad["Preco"].notna()) & (cad["Preco"] > 0)].copy()

    cad["Grupo_norm"] = cad["grupo"].astype(str).str.strip().str.upper()
    cad["gradecol"] = cad["gradecol"].fillna("").astype(str).str.strip().str.upper().str.zfill(2)
    cad["gradelin"] = cad["gradelin"].fillna("").astype(str).str.strip().str.upper().str.zfill(2)
    cad["gradegrp"] = cad["gradegrp"].fillna("").astype(str).str.strip().str.upper().str.zfill(2)

    base = cad.copy()

    if tabcol_df is not None and not tabcol_df.empty:
        tc = tabcol_df.copy()
        tc["codigo"] = tc["codigo"].astype(str).str.strip().str.upper().str.zfill(2)
        tc["nome"] = tc["nome"].astype(str).str.strip().str.upper()
        base = base.merge(tc.rename(columns={"codigo": "gradecol", "nome": "tabcol_nome"}), on="gradecol", how="left")
    else:
        base["tabcol_nome"] = ""

    if tablin_df is not None and not tablin_df.empty:
        tl = tablin_df.copy()
        tl["codigo"] = tl["codigo"].astype(str).str.strip().str.upper().str.zfill(2)
        tl["nome"] = tl["nome"].astype(str).str.strip().str.upper()
        base = base.merge(tl.rename(columns={"codigo": "gradelin", "nome": "tablin_nome"}), on="gradelin", how="left")
    else:
        base["tablin_nome"] = ""

    if tabgrp_df is not None and not tabgrp_df.empty:
        tg = tabgrp_df.copy()
        tg["codigo"] = tg["codigo"].astype(str).str.strip().str.upper().str.zfill(2)
        tg["nome"] = tg["nome"].astype(str).str.strip().str.upper()
        base = base.merge(tg.rename(columns={"codigo": "gradegrp", "nome": "tabgrp_nome"}), on="gradegrp", how="left")
    else:
        base["tabgrp_nome"] = ""

    for c in ["tabcol_nome", "tablin_nome", "tabgrp_nome"]:
        base[c] = base[c].fillna("").astype(str).str.strip().str.upper()

    base["BASE_Corrente_Feminina"] = (
        (base["Grupo_norm"] == "C") & (base["gradelin"] == "09")
    ).astype(int)
    base["BASE_Corrente_Masculina"] = (
        (base["Grupo_norm"] == "C") & (base["gradelin"] == "11")
    ).astype(int)
    base["BASE_Trio"] = (
        (base["Grupo_norm"] == "BR") & (base["gradelin"] == "19")
    ).astype(int)
    base["TIPO_Brinco_Grande"] = (
        (base["Grupo_norm"] == "BR") & (base["gradecol"] == "05")
    ).astype(int)
    base["eh_ouro"] = base["gradegrp"].eq("01")

    itens_nao_ouro = base.loc[
        ~base["eh_ouro"],
        ["Sku", "Grupo_norm", "Estoque", "Preco", "gradegrp", "tabgrp_nome"],
    ].copy()
    itens_nao_ouro["motivo"] = "Não é ouro"

    base = base.loc[base["eh_ouro"]].copy()
    base = base.loc[
        (base["Grupo_norm"] != "C")
        | (base["BASE_Corrente_Feminina"] == 1)
        | (base["BASE_Corrente_Masculina"] == 1)
    ].copy()

    diag = {}
    c_items = base[base["Grupo_norm"] == "C"].copy()
    diag["c_classificados"] = c_items[["Sku", "gradelin", "tablin_nome", "Estoque", "Preco"]].head(300)
    diag["c_fora_regra_genero"] = cad.loc[
        (cad["Grupo_norm"] == "C") & (cad["gradegrp"].eq("01")) & (~cad["gradelin"].isin(["09", "11"])),
        ["Sku", "gradelin", "gradegrp", "Estoque", "Preco"],
    ].head(200)

    br_items = base[base["Grupo_norm"] == "BR"].copy()
    diag["br_classificados"] = br_items[["Sku", "gradecol", "tabcol_nome", "gradelin", "tablin_nome", "Estoque", "Preco"]].head(300)
    diag["br_sem_tipo"] = br_items.loc[
        (br_items["BASE_Trio"] == 0) & (br_items["TIPO_Brinco_Grande"] == 0),
        ["Sku", "gradecol", "tabcol_nome", "gradelin", "tablin_nome", "Estoque", "Preco"],
    ].head(200)

    diag["itens_nao_ouro"] = itens_nao_ouro.head(300)

    df_raw = base[[
        "Sku",
        "Estoque",
        "Preco",
        "BASE_Corrente_Feminina",
        "BASE_Corrente_Masculina",
        "BASE_Trio",
        "TIPO_Brinco_Grande",
    ]].copy()

    return df_raw, diag

def bytes_from_local_file(path: str) -> bytes:
    if not os.path.exists(path):
        raise FileNotFoundError(f"Não encontrei o arquivo '{path}'.")
    with open(path, "rb") as f:
        return f.read()

@st.cache_data(show_spinner=False)
def load_base_from_bytes(xlsx_bytes: bytes) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(xlsx_bytes))
    return preparar_base_from_df(df)

def df_to_excel_bytes(sheets: dict) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=name[:31])
    bio.seek(0)
    return bio.getvalue()

def get_active_base() -> tuple[pd.DataFrame, str, bytes]:
    # 1) Fonte API (CADMAT)
    if st.session_state.get("use_api", False):
        api_url = st.secrets["api"]["url"]
        api_user = st.secrets["api"]["user"]
        api_pass = st.secrets["api"]["password"]
        api_table = st.session_state.get("api_table", DEFAULT_API_TABLE)

        try:
            tabcol = fetch_lookup_table(api_url, api_user, api_pass, "TABCOL")
        except Exception as e:
            st.warning(f"Falha ao consultar TABCOL: {e}")
            tabcol = pd.DataFrame(columns=["codigo", "nome"])

        try:
            tablin = fetch_lookup_table(api_url, api_user, api_pass, "TABLIN")
        except Exception as e:
            st.warning(f"Falha ao consultar TABLIN: {e}")
            tablin = pd.DataFrame(columns=["codigo", "nome"])

        try:
            tabgrp = fetch_lookup_table(api_url, api_user, api_pass, "TABGRP")
        except Exception as e:
            st.warning(f"Falha ao consultar TABGRP: {e}")
            tabgrp = pd.DataFrame(columns=["codigo", "nome"])

        cad = fetch_cadmat_paginado(
            api_url,
            api_user,
            api_pass,
            grupos=st.session_state.get("lista_grupos", ["BR", "C", "CJ", "CK", "CO", "ES", "PF", "PR", "SEM", "PM"]),
            colunas=["grupo", "referencia", "descricao", "caracter", "qtdreal", "prc_venda", "gradecol", "gradelin", "gradegrp"],
            only_stock_gt0=st.session_state.get("only_stock_gt0", True),
            max_pages_per_group=int(st.session_state.get("max_pages_per_group", 5000)),
            timeout=90,
        )

        df_raw, diag = build_raw_base_from_api(cad, tabcol, tablin, tabgrp)
        st.session_state["api_diag"] = diag

        b = df_to_excel_bytes({"base_api": df_raw})
        base = load_base_from_bytes(b)
        return base, f"API CADMAT ({api_table[:35]}...)", b

    # 2) Fonte Excel (upload admin em sessão)
    if "base_bytes" in st.session_state and st.session_state["base_bytes"]:
        b = st.session_state["base_bytes"]
        base = load_base_from_bytes(b)
        name = st.session_state.get("base_name", "upload_admin.xlsx")
        return base, name, b


# =============================
# CAPACIDADE TEÓRICA (diagnóstico)
# =============================
def by_sku_table(df: pd.DataFrame) -> pd.DataFrame:
    return df.groupby(["categoria", "Sku_norm"], as_index=False).agg(
        Estoque=("Estoque", "max"),
        Preco=("Preco", "min"),
        Sku=("Sku", "first"),
    )

def max_kits_category_from_stocks(stocks: np.ndarray, m: int) -> int:
    stocks = np.array(stocks, dtype=int)
    stocks = stocks[stocks > 0]
    if len(stocks) < m:
        return 0
    hi = int(stocks.sum() // m)
    lo = 0

    def feasible(k: int) -> bool:
        if k <= 0:
            return True
        return int(np.minimum(stocks, k).sum()) >= k * m

    while lo < hi:
        mid = (lo + hi + 1) // 2
        if feasible(mid):
            lo = mid
        else:
            hi = mid - 1
    return lo

def capacity_table_correct(df: pd.DataFrame) -> pd.DataFrame:
    bs = by_sku_table(df)
    rows = []
    for cat, (mn, mx) in RULES.items():
        sub = bs[bs["categoria"] == cat]
        stocks = sub["Estoque"].to_numpy(dtype=int)
        kits_cat = max_kits_category_from_stocks(stocks, mn)
        rows.append({
            "categoria": cat,
            "Grupo": DISPLAY_NAME.get(cat, cat),
            "min_por_kit": mn,
            "skus_unicos": int(sub["Sku_norm"].nunique()),
            "estoque_total": int(sub["Estoque"].sum()),
            "kits_max_cat": int(kits_cat),
        })
    out = pd.DataFrame(rows)
    out["gargalo"] = out["kits_max_cat"] == out["kits_max_cat"].min()
    return out.sort_values(["kits_max_cat", "Grupo"], ascending=[True, True])

def kits_possible_overall_correct(df: pd.DataFrame) -> tuple[int, str, pd.DataFrame]:
    t = capacity_table_correct(df)
    kits_max = int(t["kits_max_cat"].min()) if len(t) else 0
    gargalos = t.loc[t["kits_max_cat"] == kits_max, "Grupo"].tolist()
    gargalo_str = ", ".join(gargalos) if gargalos else "-"
    return kits_max, gargalo_str, t


# =============================
# SIMULADOR DE COMPRA
# =============================
def summarize_category(df: pd.DataFrame):
    bs = by_sku_table(df)

    cat_raw = bs.groupby("categoria", as_index=False).agg(
        skus_unicos=("Sku_norm", "nunique"),
        estoque_total=("Estoque", "sum"),
        preco_min=("Preco", "min"),
        preco_med=("Preco", "median"),
        preco_max=("Preco", "max"),
    )

    cat = pd.DataFrame({"categoria": list(RULES.keys())}).merge(
        cat_raw, on="categoria", how="left"
    )
    cat["skus_unicos"] = cat["skus_unicos"].fillna(0).astype(int)
    cat["estoque_total"] = cat["estoque_total"].fillna(0).astype(int)

    quantiles = [
        (0.05, "p05"), (0.10, "p10"), (0.25, "p25"), (0.35, "p35"),
        (0.50, "p50"), (0.60, "p60"), (0.75, "p75"), (0.85, "p85"),
        (0.90, "p90"), (0.95, "p95"),
    ]
    for q, name in quantiles:
        cat[f"preco_{name}"] = cat["categoria"].map(
            lambda c: float(np.quantile(bs.loc[bs["categoria"] == c, "Preco"], q))
            if (bs["categoria"] == c).any()
            else np.nan
        )

    cat["min_por_kit"] = cat["categoria"].map(lambda c: RULES[c][0])
    return cat, bs

def min_cost_theoretical(bs: pd.DataFrame) -> float:
    total = 0.0
    for cat, (mn, _) in RULES.items():
        prices = bs.loc[bs["categoria"] == cat, "Preco"].sort_values().to_numpy()
        if len(prices) < mn:
            return np.inf
        total += float(prices[:mn].sum())
    return float(total)

def choose_price_band(direction: str, weight: float, is_adjust: bool):
    if direction == "cheaper":
        if is_adjust:
            return ("p05", "p60", "ajuste-fino barato (P05–P60)")
        if weight >= 0.18:
            return ("p05", "p35", "peso alto: comprar bem barato (P05–P35)")
        if weight >= 0.10:
            return ("p10", "p50", "comprar barato (P10–P50)")
        return ("p10", "p60", "barato-médio (P10–P60)")

    if direction == "pricier":
        if is_adjust:
            return ("p60", "p95", "ajuste-fino mais caro (P60–P95)")
        if weight >= 0.18:
            return ("p50", "p85", "subir valor com controle (P50–P85)")
        return ("p60", "p90", "subir valor (P60–P90)")

    if is_adjust:
        return ("p10", "p90", "ajuste amplo (P10–P90)")
    return ("p25", "p75", "faixa padrão (P25–P75)")

def shortage_slots_for_target(stocks: np.ndarray, m: int, k: int) -> int:
    stocks = np.array(stocks, dtype=int)
    stocks = stocks[stocks > 0]
    have = int(np.minimum(stocks, k).sum())
    need = int(k * m)
    return max(0, need - have)

def simulator_purchase_table(df: pd.DataFrame, target_kits: int, target_min: float, target_max: float):
    cat, bs = summarize_category(df)

    min_cost = min_cost_theoretical(bs)
    if np.isinf(min_cost):
        direction = "neutral"
    elif min_cost > target_max:
        direction = "cheaper"
    elif min_cost < target_min:
        direction = "pricier"
    else:
        direction = "neutral"

    cat["contrib"] = cat["preco_med"] * cat["min_por_kit"]
    contrib_sum = float(cat["contrib"].sum()) if float(cat["contrib"].sum()) > 0 else 1.0
    cat["peso"] = cat["contrib"] / contrib_sum

    faltas = []
    lo_list, hi_list, label_list = [], [], []
    for _, r in cat.iterrows():
        c = r["categoria"]
        mn = int(r["min_por_kit"])
        stocks = bs.loc[bs["categoria"] == c, "Estoque"].to_numpy(dtype=int)
        falta_slots = shortage_slots_for_target(stocks, mn, int(target_kits))
        faltas.append(falta_slots)

        is_adjust = c in ADJUST_CATS
        lo, hi, label = choose_price_band(direction, float(r["peso"]), is_adjust)
        lo_list.append(lo)
        hi_list.append(hi)
        label_list.append(label)

    cat["falta_slots"] = pd.Series(faltas, index=cat.index).astype(int)
    cat["estrategia_preco"] = label_list
    cat["preco_sugerido_de"] = [round(float(r[f"preco_{lo}"]), 2) for r, lo in zip(cat.to_dict("records"), lo_list)]
    cat["preco_sugerido_ate"] = [round(float(r[f"preco_{hi}"]), 2) for r, hi in zip(cat.to_dict("records"), hi_list)]
    cat["preco_sugerido_medio"] = (cat["preco_sugerido_de"] + cat["preco_sugerido_ate"]) / 2.0

    cat["requerido_slots"] = int(target_kits) * cat["min_por_kit"]
    cat["indice_faltante"] = np.where(cat["requerido_slots"] > 0, cat["falta_slots"] / cat["requerido_slots"], 0.0)
    cat["custo_reposicao"] = cat["falta_slots"] * cat["preco_sugerido_medio"]

    N = max(int(target_kits), 1)
    cat["modelos_sku_a_comprar"] = np.ceil(cat["falta_slots"] / N).astype(int)

    cat["unid_sugeridas_por_sku"] = np.where(
        cat["modelos_sku_a_comprar"] > 0,
        np.ceil(cat["falta_slots"] / cat["modelos_sku_a_comprar"]),
        0
    ).astype(int)
    cat["unid_sugeridas_por_sku"] = np.minimum(cat["unid_sugeridas_por_sku"], N).astype(int)

    out = pd.DataFrame({
        "Grupo": cat["categoria"].map(lambda x: DISPLAY_NAME.get(x, x)),
        "Estoque": cat["estoque_total"].astype(int),
        "Faltante para a meta": cat["falta_slots"].astype(int),
        "Modelos (SKUs) a comprar": cat["modelos_sku_a_comprar"].astype(int),
        "Unid. sugeridas por SKU": cat["unid_sugeridas_por_sku"].astype(int),
        "Índice faltante por grupo": cat["indice_faltante"].astype(float),
        "Custo de reposição": cat["custo_reposicao"].astype(float),
        "Preço sugerido (de)": cat["preco_sugerido_de"].astype(float),
        "Preço sugerido (até)": cat["preco_sugerido_ate"].astype(float),
        "Estratégia de preço": cat["estrategia_preco"].astype(str),
    }).sort_values(["Faltante para a meta", "Grupo"], ascending=[False, True])

    total_row = pd.DataFrame([{
        "Grupo": "Total",
        "Estoque": int(out["Estoque"].sum()),
        "Faltante para a meta": int(out["Faltante para a meta"].sum()),
        "Modelos (SKUs) a comprar": int(out["Modelos (SKUs) a comprar"].sum()),
        "Unid. sugeridas por SKU": np.nan,
        "Índice faltante por grupo": float(out["Índice faltante por grupo"].mean()) if len(out) else 0.0,
        "Custo de reposição": float(out["Custo de reposição"].sum()),
        "Preço sugerido (de)": np.nan,
        "Preço sugerido (até)": np.nan,
        "Estratégia de preço": "",
    }])
    out = pd.concat([out, total_row], ignore_index=True)

    return out, direction, float(min_cost)

def df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")

def fmt_brl(x: float) -> str:
    return f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


# =============================
# KIT GENERATOR (VALUE-FIRST + TABU + FALHA)
# =============================
def build_structures(base: pd.DataFrame):
    stock0 = base.groupby("Sku_norm")["Estoque"].max().to_dict()
    price = base.groupby("Sku_norm")["Preco"].min().to_dict()
    cat_of = base.groupby("Sku_norm")["categoria"].first().to_dict()

    pools = {}
    for cat in RULES.keys():
        d = base[base["categoria"] == cat].drop_duplicates("Sku_norm").copy()
        d.sort_values(["Preco", "Estoque"], ascending=[True, True], inplace=True)
        pools[cat] = d["Sku_norm"].tolist()

    return stock0, price, cat_of, pools

def objective(total: float, tmin: float, tmax: float) -> float:
    mid = (tmin + tmax) / 2
    if tmin <= total <= tmax:
        return abs(total - mid)
    if total < tmin:
        return 10_000 + (tmin - total)
    return 10_000 + (total - tmax)

def can_add_cat(counts, cat):
    mn, mx = RULES[cat]
    return True if mx is None else counts[cat] < mx

def pick_k_skus(cat, k, used, stock, pools):
    cands = [s for s in pools[cat] if stock.get(s, 0) > 0 and s not in used]
    if len(cands) < k:
        return None
    return cands[:k]

def pick_best_fit(cat, used, stock, pools, price, current_total, tmin, tmax):
    max_add = tmax - current_total
    if max_add <= 0:
        return None

    gap = tmin - current_total
    cands = [
        s for s in pools[cat]
        if stock.get(s, 0) > 0 and s not in used and float(price[s]) <= max_add
    ]
    if not cands:
        return None

    if gap > 0:
        return min(cands, key=lambda s: abs(float(price[s]) - gap))
    return min(cands, key=lambda s: objective(current_total + float(price[s]), tmin, tmax))

def greedy_build(stock, pools, price, cat_of, tmin, tmax):
    used = set()
    counts = defaultdict(int)
    selected = []
    total = 0.0

    for cat, (mn, mx) in RULES.items():
        pick = pick_k_skus(cat, mn, used, stock, pools)
        if pick is None:
            return None
        for s in pick:
            used.add(s)
            selected.append(s)
            counts[cat] += 1
            total += float(price[s])

    if total > tmax:
        return None

    cats_order = ["BR_DEMAIS", "CO"] + [c for c in RULES.keys() if c not in ("BR_DEMAIS", "CO")]
    step = 0
    while total < tmin and step < 1200:
        step += 1
        chosen = None
        chosen_cat = None

        for cat in cats_order:
            if not can_add_cat(counts, cat):
                continue
            s = pick_best_fit(cat, used, stock, pools, price, total, tmin, tmax)
            if s is not None:
                chosen = s
                chosen_cat = cat
                break

        if chosen is None:
            return None

        used.add(chosen)
        selected.append(chosen)
        counts[chosen_cat] += 1
        total += float(price[chosen])
        if total > tmax:
            return None

    return {"skus": selected, "used": used, "counts": dict(counts), "total": total}

def tabu_improve(sol, stock, pools, price, cat_of, tmin, tmax,
                max_iters=TABU_ITERS, tenure=TABU_TENURE, samples=NEIGHBORHOOD_SAMPLES):
    skus = sol["skus"][:]
    used = set(skus)
    counts = defaultdict(int, sol["counts"])
    total = float(sol["total"])
    tabu = deque(maxlen=tenure)

    def is_tabu(m): return m in tabu
    def add_tabu(m): tabu.append(m)

    def kit_items(cat):
        return [s for s in skus if cat_of.get(s) == cat]

    def best_in_cat_for_swap(cat, out_sku, current_total):
        max_add = tmax - (current_total - float(price[out_sku]))
        cands = [
            s for s in pools[cat]
            if stock.get(s, 0) > 0 and s not in used and float(price[s]) <= max_add
        ]
        if not cands:
            return None
        return min(
            cands[:300],
            key=lambda s: objective(current_total - float(price[out_sku]) + float(price[s]), tmin, tmax)
        )

    best_total = total
    best_skus = skus[:]
    best_counts = counts.copy()
    best_obj = objective(total, tmin, tmax)

    for _ in range(max_iters):
        best_neighbor = None
        best_neighbor_obj = None
        best_move = None

        for _ in range(samples):
            r = random.random()
            if r < 0.65:
                move_type, cat = "swap", "BR_DEMAIS"
            elif r < 0.88:
                move_type, cat = "swap", "CO"
            elif r < 0.95:
                move_type, cat = "add", "BR_DEMAIS"
            else:
                move_type, cat = "remove", "BR_DEMAIS"

            if move_type == "swap":
                items = kit_items(cat)
                if not items:
                    continue
                out_sku = random.choice(items)
                in_sku = best_in_cat_for_swap(cat, out_sku, total)
                if in_sku is None:
                    continue
                move = (move_type, cat, out_sku, in_sku)
                if is_tabu(move):
                    continue
                new_total = total - float(price[out_sku]) + float(price[in_sku])
                new_obj = objective(new_total, tmin, tmax)
                if best_neighbor_obj is None or new_obj < best_neighbor_obj:
                    best_neighbor_obj = new_obj
                    best_neighbor = ("swap", out_sku, in_sku, new_total)
                    best_move = move

            elif move_type == "add":
                if not can_add_cat(counts, "BR_DEMAIS"):
                    continue
                s = pick_best_fit("BR_DEMAIS", used, stock, pools, price, total, tmin, tmax)
                if s is None:
                    continue
                move = (move_type, "BR_DEMAIS", None, s)
                if is_tabu(move):
                    continue
                new_total = total + float(price[s])
                if new_total > tmax:
                    continue
                new_obj = objective(new_total, tmin, tmax)
                if best_neighbor_obj is None or new_obj < best_neighbor_obj:
                    best_neighbor_obj = new_obj
                    best_neighbor = ("add", None, s, new_total)
                    best_move = move

            else:
                min_br = RULES["BR_DEMAIS"][0]
                if counts["BR_DEMAIS"] <= min_br:
                    continue
                items = kit_items("BR_DEMAIS")
                if not items:
                    continue
                out_sku = random.choice(items)
                move = (move_type, "BR_DEMAIS", out_sku, None)
                if is_tabu(move):
                    continue
                new_total = total - float(price[out_sku])
                new_obj = objective(new_total, tmin, tmax)
                if best_neighbor_obj is None or new_obj < best_neighbor_obj:
                    best_neighbor_obj = new_obj
                    best_neighbor = ("remove", out_sku, None, new_total)
                    best_move = move

        if best_neighbor is None:
            break

        kind, out_sku, in_sku, new_total = best_neighbor
        if kind == "swap":
            skus.remove(out_sku)
            used.remove(out_sku)
            skus.append(in_sku)
            used.add(in_sku)
            total = new_total
        elif kind == "add":
            skus.append(in_sku)
            used.add(in_sku)
            counts["BR_DEMAIS"] += 1
            total = new_total
        else:
            skus.remove(out_sku)
            used.remove(out_sku)
            counts["BR_DEMAIS"] -= 1
            total = new_total

        add_tabu(best_move)

        cur_obj = objective(total, tmin, tmax)
        if cur_obj < best_obj:
            best_obj = cur_obj
            best_total = total
            best_skus = skus[:]
            best_counts = counts.copy()

        if tmin <= total <= tmax and best_obj <= 1.0:
            break

    return {"skus": best_skus, "used": set(best_skus), "counts": dict(best_counts), "total": float(best_total)}

def try_build_one_with_reason(stock, pools, price, tmin, tmax):
    used = set()
    counts = defaultdict(int)
    total = 0.0

    falhas_minimos = []
    detalhes_ok = []

    # 1) Verifica TODOS os mínimos, sem parar no primeiro
    for cat, (mn, mx) in RULES.items():
        cands = [s for s in pools[cat] if stock.get(s, 0) > 0 and s not in used]
        disp = len(cands)

        if disp < mn:
            falhas_minimos.append(f"{cat} precisa {mn}, disponíveis {disp}")
            continue

        pick = cands[:mn]
        cat_sum = sum(float(price[s]) for s in pick)

        for s in pick:
            used.add(s)
            counts[cat] += 1
            total += float(price[s])

        detalhes_ok.append((cat, mn, cat_sum))

    # Se qualquer categoria falhou nos mínimos, retorna TODAS
    if falhas_minimos:
        return None, "Falhou nos mínimos: " + " | ".join(falhas_minimos)

    # 2) Se os mínimos já estourarem o teto, informa isso
    if total > tmax:
        return None, f"Mínimos estouram teto: total_minimos={total:.2f} > {tmax}"

    # 3) Tenta completar até o piso
    cats_order = ["BR_DEMAIS", "CO"] + [c for c in RULES.keys() if c not in ("BR_DEMAIS", "CO")]
    step = 0
    while total < tmin and step < 1200:
        step += 1
        chosen = None
        for cat in cats_order:
            if not can_add_cat(counts, cat):
                continue
            s = pick_best_fit(cat, used, stock, pools, price, total, tmin, tmax)
            if s is not None:
                chosen = s
                break

        if chosen is None:
            return None, f"Não conseguiu completar: total={total:.2f}, falta={tmin-total:.2f}, nenhum item cabe até {tmax}"

        used.add(chosen)
        counts[cat_of_from_pools(chosen, pools)] += 1
        total += float(price[chosen])

        if total > tmax:
            return None, f"Estourou teto ao completar: total={total:.2f} > {tmax}"

    if not (tmin <= total <= tmax):
        return None, f"Terminou fora da faixa: total={total:.2f}"

    return {"total": total, "counts": dict(counts), "skus": list(used)}, "OK"

def cat_of_from_pools(sku, pools):
    for cat, itens in pools.items():
        if sku in itens:
            return cat
    return None

def diagnose_next_kit(stock, pools, price):
    rows = []

    for cat, (mn, mx) in RULES.items():
        cands = [s for s in pools[cat] if stock.get(s, 0) > 0]
        skus_disp = len(cands)
        estoque_total = sum(stock.get(s, 0) for s in cands)
        status = "OK" if skus_disp >= mn else "FALTA_SKU"
        rows.append({
            "tipo": "minimos_viabilidade",
            "categoria": cat,
            "min": mn,
            "max": (mx if mx is not None else "∞"),
            "skus_disponiveis": skus_disp,
            "estoque_total_categoria": estoque_total,
            "status": status
        })

    used = set()
    total_min = 0.0
    min_break = None
    min_details = []

    for cat, (mn, mx) in RULES.items():
        cands = [s for s in pools[cat] if stock.get(s, 0) > 0 and s not in used]
        if len(cands) < mn:
            min_break = f"Quebrou em {cat}: precisa {mn}, tem {len(cands)} (sem repetir SKU)."
            break
        pick = cands[:mn]
        cat_sum = sum(float(price[s]) for s in pick)
        total_min += cat_sum
        used.update(pick)
        min_details.append((cat, mn, cat_sum, total_min))

    rows.append({
        "tipo": "custo_minimo_mins",
        "categoria": "-",
        "min": "-",
        "max": "-",
        "skus_disponiveis": "-",
        "estoque_total_categoria": "-",
        "status": f"TOTAL_MIN={total_min:.2f} | " + ("OK" if (min_break is None) else "IMPOSSIVEL") + (f" | {min_break}" if min_break else "")
    })

    for cat, mn, cat_sum, acc in min_details:
        rows.append({
            "tipo": "custo_minimo_detalhe",
            "categoria": cat,
            "min": mn,
            "max": "-",
            "skus_disponiveis": "-",
            "estoque_total_categoria": "-",
            "status": f"soma_cat={cat_sum:.2f} | acumulado={acc:.2f}"
        })

    return pd.DataFrame(rows)


# =============================
# CACHE DO GERADOR (relatórios)
# =============================
@st.cache_data(show_spinner=False)
def generate_kits_reports(base_bytes: bytes, tmin: float, tmax: float, max_kits: int) -> dict:
    base = load_base_from_bytes(base_bytes)
    stock0, price, cat_of, pools = build_structures(base)

    stock = dict(stock0)
    kits = []
    failure_info = None

    for kit_id in range(1, max_kits + 1):
        best = None
        for _ in range(ATTEMPTS_PER_KIT):
            sol = greedy_build(stock, pools, price, cat_of, tmin, tmax)
            if sol is None:
                continue
            sol2 = tabu_improve(sol, stock, pools, price, cat_of, tmin, tmax)
            if tmin <= sol2["total"] <= tmax:
                best = sol2
                break

        if best is None:
            _, reason = try_build_one_with_reason(stock, pools, price, tmin, tmax)
            failure_info = {"kit_que_falhou": kit_id, "motivo": reason}
            break

        for s in best["skus"]:
            stock[s] -= 1

        best["kit_id"] = kit_id
        kits.append(best)

    base_lookup = base.drop_duplicates("Sku_norm").set_index("Sku_norm")

    rows_items = []
    for k in kits:
        for s in k["skus"]:
            r = base_lookup.loc[s]
            rows_items.append({
                "kit_id": k["kit_id"],
                "Sku": r["Sku"],
                "Sku_norm": s,
                "categoria": r["categoria"],
                "Preco": float(r["Preco"]),
            })
    kits_itens = pd.DataFrame(rows_items)

    summary_rows = []
    for k in kits:
        row = {"kit_id": k["kit_id"], "total_preco": float(k["total"]), "qtd_itens": int(len(k["skus"]))}
        for cat in RULES:
            row[f"qtd_{cat}"] = int(k["counts"].get(cat, 0))
        summary_rows.append(row)
    kits_resumo = pd.DataFrame(summary_rows)

    if not stock:
        estoque_restante = pd.DataFrame(columns=["Sku_norm", "Estoque_restante", "Sku", "categoria", "Preco"])
    else:
        estoque_restante = (
            pd.DataFrame([{"Sku_norm": s, "Estoque_restante": q} for s, q in stock.items()])
            .merge(
                base[["Sku_norm", "Sku", "categoria", "Preco"]].drop_duplicates("Sku_norm"),
                on="Sku_norm",
                how="left"
            )
            .sort_values(["categoria", "Sku_norm"])
        )

    falha_df = diagnose_next_kit(stock, pools, price)
    if failure_info:
        header = pd.DataFrame([{
            "tipo": "resumo_falha",
            "categoria": "-",
            "min": "-",
            "max": "-",
            "skus_disponiveis": "-",
            "estoque_total_categoria": "-",
            "status": f"Kit {failure_info['kit_que_falhou']} falhou | {failure_info['motivo']}"
        }])
        falha_df = pd.concat([header, falha_df], ignore_index=True)

    return {
        "kits_resumo": kits_resumo,
        "kits_itens": kits_itens,
        "estoque_restante": estoque_restante,
        "falha_proximo_kit": falha_df,
        "qtd_kits": len(kits),
        "failure_info": failure_info,
    }

@st.cache_data(show_spinner=False)
def compute_real_kits_count(base_bytes: bytes, tmin: float, tmax: float, max_kits: int) -> int:
    reports = generate_kits_reports(base_bytes, tmin, tmax, max_kits)
    return int(reports.get("qtd_kits", 0))


# =============================
# UI - Sidebar
# =============================
with st.sidebar:
    st.header("Fonte de dados")

    # Sempre usar API
    st.session_state["use_api"] = True

    if st.session_state["use_api"]:
        st.subheader("Conexão API")
        st.caption("A conexão com a API está protegida e configurada via secrets.")
        st.code(st.secrets["api"]["url"], language=None)

        st.session_state["lista_grupos"] = ["BR", "C", "CJ", "CK", "CO", "ES", "PF", "PR", "SEM", "PM"]
        st.session_state["only_stock_gt0"] = True
        st.session_state["max_pages_per_group"] = 5000

        st.markdown("**Classificação (C gênero / BR tipo)**")
        st.caption("Usado para separar C Feminino/Masculino e BR Trio/Grande/Demais.")

        if st.button("Atualizar agora (limpar cache API)"):
            fetch_lookup_table.clear()
            fetch_cadmat_paginado.clear()
            st.session_state.pop("api_diag", None)
            st.info("Cache da API limpo. A base será recarregada na próxima atualização.")

        if st.button("Sair"):
            for k in ["logged_in", "login_user", "login_pass"]:
                st.session_state.pop(k, None)
            st.rerun()

        st.divider()

    st.header("Configurações")
    target_min = st.number_input("Preço mínimo do kit", value=TARGET_MIN_DEFAULT, step=10)
    target_max = st.number_input("Preço máximo do kit", value=TARGET_MAX_DEFAULT, step=10)

    st.divider()
    st.header("Simulador de compra")

    if "desired_kits" not in st.session_state:
        st.session_state["desired_kits"] = 100

    slider_value = st.slider(
        "Quantidade de torres",
        min_value=1,
        max_value=500,
        value=int(st.session_state["desired_kits"]),
        step=1,
        key="desired_kits_slider",
    )

    if slider_value != st.session_state["desired_kits"]:
        st.session_state["desired_kits"] = slider_value

    input_value = st.number_input(
        "Ou digite a quantidade de torres",
        min_value=1,
        max_value=500,
        value=int(st.session_state["desired_kits"]),
        step=1,
        key="desired_kits_input",
    )

    if input_value != st.session_state["desired_kits"]:
        st.session_state["desired_kits"] = int(input_value)

    desired_kits = int(st.session_state["desired_kits"])

    st.divider()
    st.header("Geração de kits")
    max_kits = st.number_input("Gerar até (máx kits)", min_value=1, max_value=500, value=DEFAULT_MAX_KITS, step=10)


# =============================
# MAIN
# =============================
try:
    base_df, base_name, base_bytes = get_active_base()
except Exception as e:
    st.error(str(e))
    st.stop()

kits_teorico, gargalo, _ = kits_possible_overall_correct(base_df)
kits_real = compute_real_kits_count(base_bytes, float(target_min), float(target_max), int(max_kits))

c1, c2 = st.columns([2.6, 1.0])
with c1:
    st.markdown(
        f"""
        <div class="topbar">
          <div class="topbar-title">PAINEL DE TORRES</div>
          <div class="topbar-sub">
            Base ativa: <b>{base_name}</b>
            &nbsp;|&nbsp; Faixa: <b>{fmt_brl(float(target_min))}</b> a <b>{fmt_brl(float(target_max))}</b>
            &nbsp;|&nbsp; Gargalo(s): <b>{gargalo}</b>
            &nbsp;|&nbsp; Teórico (estoque): <b>{kits_teorico}</b>
          </div>
        </div>
        """,
        unsafe_allow_html=True
    )
with c2:
    st.metric("Quantidade de torres atualmente", kits_real)

st.markdown("<hr/>", unsafe_allow_html=True)

if st.session_state.get("use_api", False):
    with st.expander("Diagnóstico da classificação (API)", expanded=False):
        diag = st.session_state.get("api_diag", {}) or {}
        c_sem = diag.get("c_fora_regra_genero")
        br_sem = diag.get("br_sem_tipo")
        br_bloq = diag.get("itens_bloqueados_banho")

        st.markdown("**C (corrente) sem gênero definido** — esses itens não viram C_FEMININO/C_MASCULINO e podem cair em OUTROS.")
        if isinstance(c_sem, pd.DataFrame) and len(c_sem):
            st.dataframe(c_sem, use_container_width=True, hide_index=True)
        else:
            st.caption("Nenhum item C sem gênero (ou ainda não carregou a base).")

        st.divider()
        st.markdown("**BR sem marcação TRIO/GRANDE** — esses itens entram como BR_DEMAIS (fallback).")
        if isinstance(br_sem, pd.DataFrame) and len(br_sem):
            st.dataframe(br_sem, use_container_width=True, hide_index=True)
        else:
            st.caption("Nenhum item BR sem tipo (ou ainda não carregou a base).")

        st.divider()
        st.markdown("**Itens bloqueados por cadastro/banho inválido** — esses itens foram removidos da base final e não entram nos kits.")
        if isinstance(br_bloq, pd.DataFrame) and len(br_bloq):
            st.dataframe(br_bloq, use_container_width=True, hide_index=True)
        else:
            st.caption("Nenhum item bloqueado por cadastro/banho inválido.")

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "Simulador de compra",
    "Kits resumo",
    "Kits itens",
    "Estoque restante",
    "Falha próximo kit"
])


# =============================
# TAB 1 - Simulador
# =============================
with tab1:
    left, right = st.columns([1.15, 2.85])

    sim_table, _, _ = simulator_purchase_table(base_df, int(desired_kits), float(target_min), float(target_max))

    with left:
        st.subheader("Ações")
        st.markdown('<div class="muted">Altere a quantidade de torres na sidebar para recalcular a tabela.</div>', unsafe_allow_html=True)

        st.download_button(
            "Baixar Simulador (Excel)",
            data=df_to_excel_bytes({"simulador_compra": sim_table}),
            file_name="simulador_compra.xlsx",
        )
        st.download_button(
            "Baixar Simulador (CSV)",
            data=df_to_csv_bytes(sim_table),
            file_name="simulador_compra.csv",
        )

        st.divider()
        st.subheader("Gerar kits")
        st.markdown('<div class="muted">Gera os relatórios abaixo usando o algoritmo (value-first + tabu).</div>', unsafe_allow_html=True)

        if st.button("Gerar kits agora"):
            st.session_state["last_gen"] = generate_kits_reports(
                base_bytes, float(target_min), float(target_max), int(max_kits)
            )
            st.success(f"Kits gerados: {st.session_state['last_gen']['qtd_kits']}")

        if "last_gen" in st.session_state:
            st.info(f"Kits gerados: {st.session_state['last_gen']['qtd_kits']}")

        if st.button("Limpar cache dos kits"):
            if "last_gen" in st.session_state:
                del st.session_state["last_gen"]
            generate_kits_reports.clear()
            compute_real_kits_count.clear()
            st.info("Cache limpo (kits + métrica do topo).")

    with right:
        st.subheader("Simulador de compra")

        display_df = sim_table.copy()
        display_df["Índice faltante por grupo"] = display_df["Índice faltante por grupo"].apply(
            lambda x: f"{x*100:.2f}%" if pd.notna(x) else ""
        )
        display_df["Custo de reposição"] = display_df["Custo de reposição"].apply(
            lambda x: fmt_brl(x) if pd.notna(x) else ""
        )
        display_df["Preço sugerido (de)"] = display_df["Preço sugerido (de)"].apply(
            lambda x: fmt_brl(x) if pd.notna(x) else ""
        )
        display_df["Preço sugerido (até)"] = display_df["Preço sugerido (até)"].apply(
            lambda x: fmt_brl(x) if pd.notna(x) else ""
        )

        st.markdown('<div class="dataframe-shell">', unsafe_allow_html=True)
        st.dataframe(
            display_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Grupo": st.column_config.TextColumn(
                    "Grupo",
                    help="Categoria do kit (ex.: CJ, CO, BR - OUTROS...)."
                ),
                "Estoque": st.column_config.NumberColumn(
                    "Estoque",
                    help="Estoque total disponível no grupo (somando todas as unidades do grupo)."
                ),
                "Faltante para a meta": st.column_config.NumberColumn(
                    "Faltante para a meta",
                    help=(
                        "Quantos 'slots' de itens faltam nesse grupo para montar N kits respeitando o mínimo por kit.\n\n"
                        "Cálculo: faltante = max(0, N*min_por_kit − Σ min(estoque_sku, N))."
                    )
                ),
                "Modelos (SKUs) a comprar": st.column_config.NumberColumn(
                    "Modelos (SKUs) a comprar",
                    help=(
                        "Estimativa mínima de SKUs/modelos diferentes a comprar para cobrir o faltante.\n\n"
                        "Como 1 SKU pode aparecer no máximo 1x por kit, ao longo de N kits cada SKU contribui no máximo N unidades úteis.\n"
                        "Cálculo (correto): ceil(faltante_slots / N)."
                    )
                ),
                "Unid. sugeridas por SKU": st.column_config.NumberColumn(
                    "Unid. sugeridas por SKU",
                    help=(
                        "Sugestão de unidades por SKU novo para cobrir o faltante com os modelos estimados.\n"
                        "Cálculo: ceil(faltante_slots / modelos), limitado a N (acima de N não ajuda a meta)."
                    )
                ),
                "Índice faltante por grupo": st.column_config.TextColumn(
                    "Índice faltante por grupo",
                    help="Percentual do mínimo necessário que está faltando: faltante / (N * min_por_kit)."
                ),
                "Custo de reposição": st.column_config.TextColumn(
                    "Custo de reposição",
                    help=(
                        "Estimativa de custo para cobrir o faltante do grupo.\n\n"
                        "Cálculo: custo = faltante * preço_médio_sugerido, onde preço_médio_sugerido = (de + até)/2."
                    )
                ),
                "Preço sugerido (de)": st.column_config.TextColumn(
                    "Preço sugerido (de)",
                    help="Faixa inferior sugerida para compra (percentil do preço do grupo)."
                ),
                "Preço sugerido (até)": st.column_config.TextColumn(
                    "Preço sugerido (até)",
                    help="Faixa superior sugerida para compra (percentil do preço do grupo)."
                ),
                "Estratégia de preço": st.column_config.TextColumn(
                    "Estratégia de preço",
                    help=(
                        "Explica por que a faixa de preço foi sugerida (baratear/encarecer/ajuste fino),\n"
                        "considerando o impacto do grupo no custo do kit e a faixa alvo (R$ min–max)."
                    )
                ),
            }
        )
        st.markdown("</div>", unsafe_allow_html=True)


# =============================
# RELATÓRIOS (gerados pelo botão)
# =============================
def render_report(tab, key: str, title: str):
    with tab:
        st.subheader(title)

        if "last_gen" not in st.session_state:
            st.info("Clique em **Gerar kits agora** na aba 'Simulador de compra' para montar os relatórios.")
            return

        df = st.session_state["last_gen"].get(key)
        if df is None:
            st.warning("Relatório não disponível.")
            return

        col_a, col_b = st.columns([1, 1])
        with col_a:
            st.download_button("Baixar Excel", data=df_to_excel_bytes({key: df}), file_name=f"{key}.xlsx")
        with col_b:
            st.download_button("Baixar CSV", data=df_to_csv_bytes(df), file_name=f"{key}.csv")

        st.dataframe(df, use_container_width=True, hide_index=True)

render_report(tab2, "kits_resumo", "Kits resumo")
render_report(tab3, "kits_itens", "Kits itens")
render_report(tab4, "estoque_restante", "Estoque restante")
render_report(tab5, "falha_proximo_kit", "Falha próximo kit")
