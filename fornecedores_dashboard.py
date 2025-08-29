import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import plotly.express as px

# Tema
# =========================
def tema_escuro_ativo():
    try:
        return st.runtime.scriptrunner.get_script_run_ctx().session.theme == "dark"
    except:
        return False

if tema_escuro_ativo():
    cor_fundo = "#2a2a2a"
    cor_texto = "#ffffff"
    cor_subtitulo = "#cccccc"
else:
    cor_fundo = "#f4f4f4"
    cor_texto = "#000000"
    cor_subtitulo = "#444444"

st.set_page_config(page_title="Fornecedores Ativos", page_icon="📦", layout="wide")
st.markdown(f"""
    <style>
    .metric-box {{
        background-color: {cor_fundo};
        color: {cor_texto};
        padding: 20px;
        border-radius: 12px;
        text-align: center;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.1);
        transition: all 0.3s ease;
    }}
    .metric-box h1 {{
        font-size: 36px;
        margin: 0;
    }}
    .metric-box small {{
        font-size: 14px;
        color: {cor_subtitulo};
    }}
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: center;'>📦 Painel de Fornecedores Ativos</h1>", unsafe_allow_html=True)
st.markdown("---")

# Carregar dados
# =========================
@st.cache_data(ttl=900)
def carregar_dados():
    df_forn = pd.read_excel("FornecedoresAtivos.xlsx", sheet_name=0)
    df_ped = pd.read_excel("UltForn.xlsx", sheet_name=0)
    return df_forn, df_ped

df, df_pedidos = carregar_dados()

# Tratamento / normalização
# =========================
# chaves/strings estáveis
df["FORN_RAZAO"] = df["FORN_RAZAO"].astype(str).str.strip()
df["FORN_FANTASIA"] = df["FORN_FANTASIA"].astype(str).str.strip()
df["CATEGORIAS"] = df["CATEGORIAS"].astype(str).str.upper().str.strip()
df_pedidos["PED_FORNECEDOR"] = df_pedidos["PED_FORNECEDOR"].astype(str)
df["FORN_DTCADASTRO"] = pd.to_datetime(df["FORN_DTCADASTRO"], errors="coerce")
df["FORN_CNPJ"] = df["FORN_CNPJ"].astype(str).str.replace(r"\D", "", regex=True).str.zfill(14)
df["CNPJ_FORMATADO"] = df["FORN_CNPJ"].str.replace(
    r"(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})", r"\1.\2.\3/\4-\5", regex=True
)
df["FORN_UF"] = df["FORN_UF"].astype(str).str.upper().str.strip()

df_pedidos["PED_DT"] = pd.to_datetime(df_pedidos["PED_DT"], errors="coerce")
df_pedidos["PED_FORNECEDOR"] = df_pedidos["PED_FORNECEDOR"].astype(str).str.replace(r"\D", "", regex=True).str.zfill(14)

# Último pedido por fornecedor (global, antes de filtros)
ult = (df_pedidos.dropna(subset=["PED_DT"])
       .groupby("PED_FORNECEDOR", as_index=False)["PED_DT"].max()
       .rename(columns={"PED_FORNECEDOR": "FORN_CNPJ", "PED_DT": "ULTIMO_PEDIDO"}))

df = df.merge(ult, how="left", on="FORN_CNPJ")

# Filtros
# =========================
col1, col2, col3 = st.columns(3)

ufs = col1.multiselect(
    "Filtrar por UF:",
    sorted(df["FORN_UF"].dropna().unique()),
    placeholder="Selecione o estado"
)

# categorias vêm em string separada por vírgulas (ex.: "Elétrico, Hidráulico")
todas_categorias = (
    df["CATEGORIAS"]
    .dropna()
    .astype(str)
    .str.split(",")
    .explode()
    .str.strip()
    .replace("", pd.NA)
    .dropna()
    .unique()
)
categorias = col2.multiselect(
    "Filtrar por Categoria:",
    sorted(todas_categorias),
    placeholder="Escolha uma categoria"
)

data_ini, data_fim = col3.date_input(
    "Filtrar por período de cadastro:",
    value=(df["FORN_DTCADASTRO"].min().date() if pd.notna(df["FORN_DTCADASTRO"].min()) else datetime.today().date(),
           df["FORN_DTCADASTRO"].max().date() if pd.notna(df["FORN_DTCADASTRO"].max()) else datetime.today().date())
)

# Aplicar filtros
df_filtrado = df.copy()

if ufs:
    df_filtrado = df_filtrado[df_filtrado["FORN_UF"].isin(ufs)]

if categorias:
    categorias_set = set([c.upper() for c in categorias])
    def tem_intersec(x):
        if pd.isna(x): 
            return False
        tokens = [t.strip().upper() for t in str(x).split(",") if t.strip()]
        return bool(set(tokens) & categorias_set)
    df_filtrado = df_filtrado[df_filtrado["CATEGORIAS"].apply(tem_intersec)]

if data_ini and data_fim:
    di = pd.to_datetime(data_ini)
    dt_fim = pd.to_datetime(data_fim)
    df_filtrado = df_filtrado[(df_filtrado["FORN_DTCADASTRO"] >= di) & (df_filtrado["FORN_DTCADASTRO"] <= dt_fim)]

if st.button("🔄 Resetar filtros"):
    st.rerun()

# Janelas de tempo (hoje, 12m, 30d, 90d)
# =========================
hoje = pd.Timestamp.today().normalize()
h12m = hoje - pd.DateOffset(months=12)
h30d = hoje - pd.Timedelta(days=30)
h90d = hoje - pd.Timedelta(days=90)

# CNPJ set após filtros (para cruzar com pedidos)
cnpjs_filtrados = set(df_filtrado["FORN_CNPJ"].astype(str))

# Pedidos filtrados pelo universo e por 12m (para métricas e Top 10)
ped_12m = df_pedidos[
    (df_pedidos["PED_FORNECEDOR"].isin(cnpjs_filtrados)) &
    (df_pedidos["PED_DT"] >= h12m)
].copy()

# Flag ativos 12m (com base em ULTIMO_PEDIDO >= h12m)
df_filtrado["ATIVO_12M"] = df_filtrado["ULTIMO_PEDIDO"].ge(h12m)

# Dias desde último pedido (NaT -> NaN)
df_filtrado["DIAS_DESDE_ULTIMO"] = (hoje - df_filtrado["ULTIMO_PEDIDO"]).dt.days

# Métricas
# =========================
total_forn = int(df_filtrado["FORN_CNPJ"].nunique())

cadastrados_30d = int(
    df_filtrado.loc[df_filtrado["FORN_DTCADASTRO"].ge(h30d), "FORN_CNPJ"].nunique()
)

# --- Base "ativa" (12m) para contagens baseadas em pedidos ---
ativos_set = set(df_filtrado.loc[df_filtrado["ATIVO_12M"], "FORN_CNPJ"].astype(str))

# pedidos nos 12m feitos para fornecedores ATIVOS
ped_12m_ativos = ped_12m[ped_12m["PED_FORNECEDOR"].isin(ativos_set)].copy()

# nº de fornecedores ATIVOS usados nos 12m (distintos)
usados_12m_ativos = int(ped_12m_ativos["PED_FORNECEDOR"].nunique())

# nº de fornecedores (ativos do sistema) usados nos 12m
usados_12m_ativos = int(ped_12m["PED_FORNECEDOR"].nunique())

# % dos ATIVOS (da planilha) que foram usados nos últimos 12 meses
pct_ativos_12m = (usados_12m_ativos / total_forn) if total_forn else 0.0

# Novos com uso (30 dias): cadastrados nos últimos 30 dias E com pelo menos 1 pedido em qualquer data
novos_com_uso_30d = int(
    df_filtrado.loc[
        df_filtrado["FORN_DTCADASTRO"].ge(h30d) & df_filtrado["ULTIMO_PEDIDO"].notna(),
        "FORN_CNPJ"
    ].nunique()
)

# Tempo desde último pedido (dias)
tempo_medio_sem_uso = float(df_filtrado["DIAS_DESDE_ULTIMO"].dropna().mean()) if df_filtrado["DIAS_DESDE_ULTIMO"].notna().any() else 0.0
dias_sem_uso_series = df_filtrado["DIAS_DESDE_ULTIMO"].dropna()
mediana_sem_uso = float(dias_sem_uso_series.median()) if not dias_sem_uso_series.empty else 0.0
p90_sem_uso     = float(np.percentile(dias_sem_uso_series, 90)) if not dias_sem_uso_series.empty else 0.0

# Risco de inatividade: sem uso há ≥ 90 dias (NaT conta como risco)
risco_inatividade_90d = int(
    df_filtrado.loc[
        (df_filtrado["ULTIMO_PEDIDO"].isna()) | (df_filtrado["ULTIMO_PEDIDO"] < h90d),
        "FORN_CNPJ"
    ].nunique()
)

# Concentração 80/20 (Pareto por quantidade de pedidos nos últimos 12 meses)
contagens = ped_12m.groupby("PED_FORNECEDOR").size().sort_values(ascending=False)
total_ped_12m = int(contagens.sum())
if total_ped_12m > 0:
    acum = contagens.cumsum() / total_ped_12m
    n_fornecedores_para_80 = int(np.searchsorted(acum.values, 0.80, side="left") + 1)
    fornecedores_usados_12m = int(contagens.shape[0])
else:
    n_fornecedores_para_80 = 0
    fornecedores_usados_12m = 0

# KPIs – Fileira 1 (5 cards)
# =========================
f1 = st.columns(6)
with f1[0]:
    st.markdown(f"""<div class="metric-box"><h1>{total_forn}</h1><small>Total de Fornecedores Ativos</small></div>""", unsafe_allow_html=True)
with f1[1]:
    st.markdown(f"""<div class="metric-box"><h1>{cadastrados_30d}</h1><small>Cadastrados nos últimos 30 dias</small></div>""", unsafe_allow_html=True)
with f1[2]:
    st.markdown(f"""<div class="metric-box"><h1>{usados_12m_ativos}</h1><small>Fornecedores Utilizados (12m)</small></div>""", unsafe_allow_html=True)
with f1[3]:
    st.markdown(f"""<div class="metric-box"><h1>{pct_ativos_12m:.0%}</h1><small>% Fornecedores Utilizados (12m)</small></div>""", unsafe_allow_html=True)
with f1[4]:
    st.markdown(f"""<div class="metric-box"><h1>{novos_com_uso_30d}</h1><small>Fornecedores Novos Utilizados</small></div>""", unsafe_allow_html=True)
with f1[5]:
    cap_80 = f"{n_fornecedores_para_80}/{fornecedores_usados_12m}" if fornecedores_usados_12m else "0/0"
    st.markdown(f"""<div class="metric-box"><h1>{cap_80}</h1><small>Concentração 80% dos pedidos (12m)</small></div>""", unsafe_allow_html=True)

# Alerta simples (opcional): muitos inativos
if total_forn:
    perc_risco = risco_inatividade_90d / total_forn
    if perc_risco >= 0.30:
        st.warning(f"⚠️ {risco_inatividade_90d} fornecedores (≈{perc_risco:.0%}) sem uso há ≥ 90 dias.")
    elif perc_risco >= 0.15:
        st.info(f"ℹ️ {risco_inatividade_90d} fornecedores (≈{perc_risco:.0%}) sem uso há ≥ 90 dias.")

st.divider()

# Tabela principal
# =========================
df_filtrado = df_filtrado.sort_values(by="FORN_DTCADASTRO", ascending=False)

tabela = df_filtrado[[
    "FORN_RAZAO", "FORN_FANTASIA", "FORN_UF", "FORN_DTCADASTRO", "ULTIMO_PEDIDO", "DIAS_DESDE_ULTIMO"
]].rename(columns={
    "FORN_RAZAO": "Razão Social",
    "FORN_FANTASIA": "Nome Fantasia",
    "FORN_UF": "UF",
    "FORN_DTCADASTRO": "Data de Cadastro",
    "ULTIMO_PEDIDO": "Data Último Pedido",
    "DIAS_DESDE_ULTIMO": "Dias desde o Último Pedido"
}).copy()

# Formata cadastro
tabela["Data de Cadastro"] = pd.to_datetime(tabela["Data de Cadastro"], errors="coerce").dt.strftime("%d/%m/%Y")

# Regra de "não utilizado nos últimos 12 meses"
mask_sem_uso = df_filtrado["ULTIMO_PEDIDO"].isna() | (df_filtrado["ULTIMO_PEDIDO"] < h12m)

# Formata "Último Pedido": data se usou em 12m, texto se não usou
ult = pd.to_datetime(tabela["Data Último Pedido"], errors="coerce")
tabela["Data Último Pedido"] = np.where(
    mask_sem_uso.values,
    "Não utilizado nos últimos 12 meses",
    ult.dt.strftime("%d/%m/%Y")
)

# "Dias desde o Último Pedido": número ou "—" quando não tem registro
dias = pd.to_numeric(tabela["Dias desde o Último Pedido"], errors="coerce").astype("Int64")
tabela["Dias desde o Último Pedido"] = dias.astype(str).replace({"<NA>": "—"})

st.subheader("Fornecedores (cadastro + último uso)")
tabela = tabela[["Razão Social","Nome Fantasia","UF","Data de Cadastro","Data Último Pedido","Dias desde o Último Pedido"]]
st.dataframe(tabela, use_container_width=True)
st.markdown("---")

# Top 10 fornecedores (12m) – respeita filtros
# =========================

# Pedidos 12m já filtrados por CNPJs da visão
df_top = ped_12m.merge(
    df[["FORN_CNPJ", "FORN_FANTASIA"]],
    left_on="PED_FORNECEDOR",
    right_on="FORN_CNPJ",
    how="left"
)

top10 = (
    df_top.groupby("FORN_FANTASIA", dropna=False)
    .size()
    .reset_index(name="Quantidade de Pedidos")
    .sort_values(by="Quantidade de Pedidos", ascending=False)
    .head(10)
)

st.markdown("### 🏆 Top 10 Fornecedores dos Últimos 12 Meses")

if top10.empty:
    st.info("Sem pedidos nos últimos 12 meses para o conjunto filtrado.")
else:
    # 1) Ordena do MAIOR -> MENOR e reseta índice
    top10 = top10.sort_values("Quantidade de Pedidos", ascending=False).reset_index(drop=True)

    # 2) Eixo Y como CATEGÓRICO ORDENADO na MESMA ordem do DataFrame
    ordem_y = top10["FORN_FANTASIA"].tolist()
    top10["FORN_FANTASIA"] = pd.Categorical(top10["FORN_FANTASIA"], categories=ordem_y, ordered=True)

    fig = px.bar(
        top10,
        x="Quantidade de Pedidos",
        y="FORN_FANTASIA",
        orientation="h",
        text="Quantidade de Pedidos",
        color="Quantidade de Pedidos",
        color_continuous_scale=["#7FC7FF", "#0066CC"],
    )

    fig.update_traces(
        hovertemplate="<b>%{y}</b><br>Pedidos: %{x}<extra></extra>",
        textposition="outside"
    )

    # 3) Força o Plotly a usar a ordem da categoria (sem reverse, sem alphabetical)
    fig.update_yaxes(categoryorder="total ascending")

    fig.update_layout(
        yaxis_title="Fornecedor",
        xaxis_title="Quantidade de Pedidos",
        showlegend=False,
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(color=cor_texto)
    )
    st.plotly_chart(fig, use_container_width=True)

# 📍 Distribuição por UF (após filtros) — vertical + desc + "Outras"
# =========================
st.markdown("### 📍 Distribuição por UF")
alvo = {"RJ", "SC", "SP"}
contagem = df_filtrado["FORN_UF"].value_counts(dropna=False)

rj = int(contagem.get("RJ", 0))
sc = int(contagem.get("SC", 0))
sp = int(contagem.get("SP", 0))
outras = int(contagem.drop(list(alvo), errors="ignore").sum())

df_uf_plot = pd.DataFrame(
    {"UF": ["RJ", "SC", "SP", "Outras"], "Fornecedores": [rj, sc, sp, outras]}
).sort_values("Fornecedores", ascending=False)

fig_uf = px.bar(
    df_uf_plot,
    x="UF",                    
    y="Fornecedores",
    text="Fornecedores",
    color="Fornecedores",
    color_continuous_scale=["#7FC7FF", "#0066CC"],
)
fig_uf.update_traces(textposition="outside")
fig_uf.update_layout(
    xaxis_title="UF",
    yaxis_title="Total de Fornecedores",
    showlegend=False,
    plot_bgcolor="rgba(0,0,0,0)",
    paper_bgcolor="rgba(0,0,0,0)",
    font=dict(color=cor_texto),
)
st.plotly_chart(fig_uf, use_container_width=True)

# 🧩 Distribuição por Categoria (após filtros) — remove NaN + DESC
# =========================
st.markdown("### 🧩 Distribuição por Categoria")
cats = (
    df_filtrado["CATEGORIAS"]
    .dropna()
    .astype(str)
    .str.split(",")
    .explode()
    .str.strip()
)

# remove vazios e rótulos "NaN"/"nan" que vieram como string
cats = cats[cats.ne("") & ~cats.str.match(r"(?i)^nan$")]

# conta e nomeia colunas de forma explícita
dist_cat = cats.value_counts().reset_index()
dist_cat.columns = ["Categoria", "Fornecedores"]

# top 15
dist_cat = dist_cat.sort_values("Fornecedores", ascending=False).head(15)

if dist_cat.empty:
    st.info("Sem categorias para exibir com os filtros atuais.")
else:
    # gráfico horizontal
    fig_cat = px.bar(
        dist_cat,
        x="Fornecedores",
        y="Categoria",
        orientation="h",
        text="Fornecedores",
        color="Fornecedores",
        color_continuous_scale=["#7FC7FF", "#0066CC"],
    )
    fig_cat.update_traces(textposition="outside")
    fig_cat.update_yaxes(
        categoryorder="total ascending"  # garante maior em cima
    )
    fig_cat.update_layout(
        xaxis_title="Total de Fornecedores",
        yaxis_title="Categoria",
        showlegend=False,
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(color=cor_texto),
    )
    st.plotly_chart(fig_cat, use_container_width=True)

# Exportar Excel
# =========================
def converter_excel(df_export):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Fornecedores')
    return output.getvalue()

tabela_export = tabela.copy()
excel_bytes = converter_excel(tabela_export)

st.download_button(
    label="📥 Baixar tabela filtrada (.xlsx)",
    data=excel_bytes,
    file_name="fornecedores_filtrados.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
