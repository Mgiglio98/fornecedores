import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import plotly.express as px

# =========================
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
st.markdown(
    "<p style='text-align: center; font-size:18px; color:gray;'>Visão consolidada dos fornecedores cadastrados e ativos no sistema e sua utilização com base nas requisições previamente criadas</p>",
    unsafe_allow_html=True
)
st.markdown("---")

# =========================
# Carregar dados
# =========================
@st.cache_data(ttl=900)
def carregar_dados():
    return pd.read_excel("AnaliseFornecedores.xlsx", sheet_name=0)

df = carregar_dados()

# =========================
# Janela de tempo (mantendo comportamento de 12 meses)
# =========================
hoje = pd.Timestamp.today().normalize()
h12m = hoje - pd.DateOffset(months=12)
h30d = hoje - pd.Timedelta(days=30)
h90d = hoje - pd.Timedelta(days=90)

# =========================
# Tratamento / normalização
# =========================
df["FORN_RAZAO"] = df["FORN_RAZAO"].astype(str).str.strip()
df["FORN_FANTASIA"] = df["FORN_FANTASIA"].astype(str).str.strip()
df["CATEGORIAS"] = df["CATEGORIAS"].astype(str).str.upper().str.strip()
df["FORN_DTCADASTRO"] = pd.to_datetime(df["FORN_DTCADASTRO"], errors="coerce")
df["FORN_CNPJ"] = df["FORN_CNPJ"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
df["FORN_UF"] = df["FORN_UF"].astype(str).str.upper().str.strip()
df["ULTIMO_PEDIDO"] = pd.to_datetime(df["ULTIMO_PEDIDO"], errors="coerce")
df["QTD_OFS_12M"] = pd.to_numeric(df["QTD_OFS_12M"], errors="coerce").fillna(0).astype(int)

# =========================
# Filtros
# =========================
col1, col2, col3 = st.columns(3)

ufs = col1.multiselect(
    "Filtrar por UF:",
    sorted(df["FORN_UF"].dropna().unique()),
    placeholder="Selecione o estado"
)

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
    value=(
        df["FORN_DTCADASTRO"].min().date() if pd.notna(df["FORN_DTCADASTRO"].min()) else datetime.today().date(),
        df["FORN_DTCADASTRO"].max().date() if pd.notna(df["FORN_DTCADASTRO"].max()) else datetime.today().date()
    )
)

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

# Dias desde último pedido (NaT -> NaN)
df_filtrado["DIAS_DESDE_ULTIMO"] = (hoje - df_filtrado["ULTIMO_PEDIDO"]).dt.days

# =========================
# Métricas
# =========================
total_forn = int(df_filtrado["FORN_CNPJ"].nunique())

cadastrados_30d = int(
    df_filtrado.loc[df_filtrado["FORN_DTCADASTRO"].ge(h30d), "FORN_CNPJ"].nunique()
)

# nº de fornecedores usados nos últimos 12 meses
usados_12m_ativos = int(
    df_filtrado.loc[df_filtrado["QTD_OFS_12M"].gt(0), "FORN_CNPJ"].nunique()
)

# % dos ATIVOS (da planilha) que foram usados nos últimos 12 meses
pct_ativos_12m = (usados_12m_ativos / total_forn) if total_forn else 0.0

# Novos com uso (30 dias): cadastrados nos últimos 30 dias E com pelo menos 1 OF nos últimos 12 meses
novos_com_uso_30d = int(
    df_filtrado.loc[
        df_filtrado["FORN_DTCADASTRO"].ge(h30d) &
        df_filtrado["QTD_OFS_12M"].gt(0),
        "FORN_CNPJ"
    ].nunique()
)

# Tempo desde último pedido (dias)
dias_sem_uso_series = df_filtrado["DIAS_DESDE_ULTIMO"].dropna()
tempo_medio_sem_uso = float(dias_sem_uso_series.mean()) if not dias_sem_uso_series.empty else 0.0
mediana_sem_uso = float(dias_sem_uso_series.median()) if not dias_sem_uso_series.empty else 0.0
p90_sem_uso     = float(np.percentile(dias_sem_uso_series, 90)) if not dias_sem_uso_series.empty else 0.0

# Risco de inatividade: sem uso há ≥ 90 dias (NaT conta como risco)
risco_inatividade_90d = int(
    df_filtrado.loc[
        (df_filtrado["ULTIMO_PEDIDO"].isna()) | (df_filtrado["ULTIMO_PEDIDO"] < h90d),
        "FORN_CNPJ"
    ].nunique()
)

# Concentração 80/20 (Pareto por nº de OFs distintas nos últimos 12 meses)
contagens = (
    df_filtrado.loc[df_filtrado["QTD_OFS_12M"].gt(0)]
               .set_index("FORN_CNPJ")["QTD_OFS_12M"]
               .sort_values(ascending=False)
)
total_of_12m = int(contagens.sum())
if total_of_12m > 0:
    acum = contagens.cumsum() / total_of_12m
    n_fornecedores_para_80 = int(np.searchsorted(acum.values, 0.80, side="left") + 1)
    fornecedores_usados_12m = int(contagens.shape[0])
else:
    n_fornecedores_para_80 = 0
    fornecedores_usados_12m = 0

# =========================
# KPIs – Fileira 1
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
    st.markdown(f"""<div class="metric-box"><h1>{cap_80}</h1><small>Concentração 80% das OFs (12m)</small></div>""", unsafe_allow_html=True)

# Alerta simples (opcional)
if total_forn:
    perc_risco = risco_inatividade_90d / total_forn
    if perc_risco >= 0.30:
        st.warning(f"⚠️ {risco_inatividade_90d} fornecedores (≈{perc_risco:.0%}) sem uso há ≥ 90 dias.")
    elif perc_risco >= 0.15:
        st.info(f"ℹ️ {risco_inatividade_90d} fornecedores (≈{perc_risco:.0%}) sem uso há ≥ 90 dias.")

st.divider()

# =========================
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
ult_datas = pd.to_datetime(tabela["Data Último Pedido"], errors="coerce")
tabela["Data Último Pedido"] = np.where(
    mask_sem_uso.values,
    "Não utilizado nos últimos 12 meses",
    ult_datas.dt.strftime("%d/%m/%Y")
)

# "Dias desde o Último Pedido": número ou "—" quando não tem registro
dias = pd.to_numeric(tabela["Dias desde o Último Pedido"], errors="coerce").astype("Int64")
tabela["Dias desde o Último Pedido"] = dias.astype(str).replace({"<NA>": "—"})

st.subheader("Fornecedores - Visão Geral")
tabela = tabela[["Razão Social","Nome Fantasia","UF","Data de Cadastro","Data Último Pedido","Dias desde o Último Pedido"]]
st.dataframe(tabela, use_container_width=True)
st.markdown("---")

# =========================
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
    label="📥 Baixar tabela em Excel",
    data=excel_bytes,
    file_name="fornecedores_filtrados.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# =========================
# Top 10 Fornecedores (12m) – por nº de OFs distintas
# =========================
top10 = (
    df_filtrado.loc[df_filtrado["QTD_OFS_12M"].gt(0),
                    ["FORN_CNPJ", "FORN_FANTASIA", "QTD_OFS_12M"]]
               .rename(columns={"QTD_OFS_12M": "Quantidade de OFs"})
               .nlargest(10, "Quantidade de OFs")
               .copy()
)

st.markdown("### 🏆 Top 10 Fornecedores dos Últimos 12 Meses (por OFs)")
if top10.empty:
    st.info("Sem OFs nos últimos 12 meses para o conjunto filtrado.")
else:
    top10 = top10.sort_values("Quantidade de OFs", ascending=False).reset_index(drop=True)
    nomes_duplicados = top10["FORN_FANTASIA"].duplicated(keep=False)
    top10["FORNECEDOR_EXIBICAO"] = top10["FORN_FANTASIA"]
    top10.loc[nomes_duplicados, "FORNECEDOR_EXIBICAO"] = (
        top10.loc[nomes_duplicados, "FORN_FANTASIA"]
        + " (" + top10.loc[nomes_duplicados, "FORN_CNPJ"] + ")"
    )
    ordem_y = top10["FORNECEDOR_EXIBICAO"].tolist()
    top10["FORNECEDOR_EXIBICAO"] = pd.Categorical(
        top10["FORNECEDOR_EXIBICAO"], categories=ordem_y, ordered=True
    )

    fig = px.bar(
        top10,
        x="Quantidade de OFs",
        y="FORNECEDOR_EXIBICAO",
        orientation="h",
        text="Quantidade de OFs",
        color="Quantidade de OFs",
        color_continuous_scale=["#7FC7FF", "#0066CC"],
    )
    fig.update_traces(hovertemplate="<b>%{y}</b><br>OFs: %{x}<extra></extra>", textposition="outside")
    fig.update_yaxes(categoryorder="total ascending")
    fig.update_layout(
        yaxis_title="Fornecedor",
        xaxis_title="Quantidade de OFs",
        showlegend=False,
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(color=cor_texto)
    )
    st.plotly_chart(fig, use_container_width=True)

# =========================
# 📍 Distribuição por UF (após filtros)
# =========================
st.markdown("### 📍 Distribuição por UF")
alvo = {"RJ", "SC", "SP"}
contagem = df_filtrado["FORN_UF"].value_counts(dropna=False)

rj = int(contagem.get("RJ", 0))
sc = int(contagem.get("SC", 0))
sp = int(contagem.get("SP", 0))
outras = int(contagem.drop(list(alvo), errors="ignore").sum())

df_uf_plot = pd.DataFrame({"UF": ["RJ", "SC", "SP", "Outras"], "Fornecedores": [rj, sc, sp, outras]}).sort_values("Fornecedores", ascending=False)

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

# =========================
# 🧩 Distribuição por Categoria (após filtros)
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
    fig_cat.update_yaxes(categoryorder="total ascending")
    fig_cat.update_layout(
        xaxis_title="Total de Fornecedores",
        yaxis_title="Categoria",
        showlegend=False,
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(color=cor_texto),
    )
    st.plotly_chart(fig_cat, use_container_width=True)
