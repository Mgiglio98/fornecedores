import streamlit as st
import pandas as pd
from io import BytesIO
import plotly.express as px

# --- Estilo customizado ---
st.markdown("""
    <style>
    .main {
        background-color: #111111;
        color: white;
    }
    .metric-box {
        background-color: #262730;
        padding: 20px;
        border-radius: 12px;
        text-align: center;
    }
    .metric-box h1 {
        font-size: 36px;
        margin: 0;
    }
    .metric-box small {
        font-size: 14px;
        color: #AAAAAA;
    }
    </style>
""", unsafe_allow_html=True)

# --- Configuração da Página ---
st.set_page_config(page_title="Fornecedores Ativos", layout="wide")
st.markdown("<h1 style='text-align: center;'>📦 Painel de Fornecedores Ativos</h1>", unsafe_allow_html=True)
st.markdown("---")

# --- Carregar os dados ---
@st.cache_data
def carregar_dados():
    df_forn = pd.read_excel("FornecedoresAtivos.xlsx", sheet_name=0)
    df_ped = pd.read_excel("UltForn.xlsx", sheet_name=0)
    return df_forn, df_ped

df, df_pedidos = carregar_dados()

# --- Tratamento das colunas ---
df["FORN_DTCADASTRO"] = pd.to_datetime(df["FORN_DTCADASTRO"], errors="coerce")
df["CNPJ_FORMATADO"] = df["FORN_CNPJ"].astype(str).str.zfill(14).str.replace(
    r"(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})", r"\1.\2.\3/\4-\5", regex=True
)
df_pedidos["PED_DT"] = pd.to_datetime(df_pedidos["PED_DT"], errors="coerce")

# --- Cruzamento: Último Pedido por fornecedor ---
df_ultimos = df_pedidos.groupby("PED_FORNECEDOR")["PED_DT"].max().reset_index()
df_ultimos.rename(columns={"PED_DT": "ULTIMO_PEDIDO", "PED_FORNECEDOR": "FORN_CNPJ"}, inplace=True)

# Garantir que CNPJ seja string
df["FORN_CNPJ"] = df["FORN_CNPJ"].astype(str)
df_ultimos["FORN_CNPJ"] = df_ultimos["FORN_CNPJ"].astype(str)

# Merge para trazer a info de último pedido
df = df.merge(df_ultimos, how="left", on="FORN_CNPJ")

# --- Filtros ---
col1, col2, col3 = st.columns(3)

ufs = col1.multiselect("Filtrar por UF:", sorted(df["FORN_UF"].dropna().unique()), placeholder="Selecione o estado")

todas_categorias = df["CATEGORIAS"].dropna().str.split(",").explode().str.strip().unique()
categorias = col2.multiselect("Filtrar por Categoria:", sorted(todas_categorias), placeholder="Escolha uma categoria")

data_ini, data_fim = col3.date_input(
    "Filtrar por período de cadastro:",
    value=(df["FORN_DTCADASTRO"].min(), df["FORN_DTCADASTRO"].max())
)

# --- Aplicar filtros ---
df_filtrado = df.copy()

if ufs:
    df_filtrado = df_filtrado[df_filtrado["FORN_UF"].isin(ufs)]

if categorias:
    df_filtrado = df_filtrado[df_filtrado["CATEGORIAS"]
        .apply(lambda x: any(cat in x for cat in categorias) if isinstance(x, str) else False)
    ]

if data_ini and data_fim:
    df_filtrado = df_filtrado[
        (df_filtrado["FORN_DTCADASTRO"] >= pd.to_datetime(data_ini)) &
        (df_filtrado["FORN_DTCADASTRO"] <= pd.to_datetime(data_fim))
    ]

# --- Métricas ---
col4, col5, col6 = st.columns(3)

with col4:
    st.markdown(f"""
        <div class="metric-box">
            <h1>{len(df_filtrado)}</h1>
            <small>Total de Fornecedores (após filtro)</small>
        </div>
    """, unsafe_allow_html=True)

with col5:
    recentes = df_filtrado[df_filtrado["FORN_DTCADASTRO"] >= pd.Timestamp.now() - pd.DateOffset(days=30)]
    st.markdown(f"""
        <div class="metric-box">
            <h1>{len(recentes)}</h1>
            <small>Cadastrados nos últimos 30 dias</small>
        </div>
    """, unsafe_allow_html=True)

with col6:
    usados = df_filtrado[df_filtrado["ULTIMO_PEDIDO"].notnull()]
    st.markdown(f"""
        <div class="metric-box">
            <h1>{len(usados)}</h1>
            <small>Fornecedores usados nos últimos 12 meses</small>
        </div>
    """, unsafe_allow_html=True)

# --- Tabela ---
df_filtrado = df_filtrado.sort_values(by="FORN_DTCADASTRO", ascending=False)

tabela = df_filtrado[[ 
    "FORN_RAZAO", "FORN_FANTASIA", "FORN_UF", "FORN_DTCADASTRO"
]].rename(columns={
    "FORN_RAZAO": "Razão Social",
    "FORN_FANTASIA": "Nome Fantasia",
    "FORN_UF": "UF",
    "FORN_DTCADASTRO": "Data de Cadastro"
})

tabela["Último Pedido"] = df_filtrado["ULTIMO_PEDIDO"].dt.strftime('%d/%m/%Y')
tabela["Último Pedido"] = tabela["Último Pedido"].where(df_filtrado["ULTIMO_PEDIDO"].notna(), "")
tabela["Data de Cadastro"] = pd.to_datetime(tabela["Data de Cadastro"]).dt.strftime('%d/%m/%Y')

st.markdown("---")
st.dataframe(tabela, use_container_width=True)

# --- Top 10 Fornecedores ---
st.markdown("### 🏆 Top 10 Fornecedores Mais Utilizados nos Últimos 12 Meses")

df_top = df_pedidos.copy()
df_top["PED_FORNECEDOR"] = df_top["PED_FORNECEDOR"].astype(str)
df["FORN_CNPJ"] = df["FORN_CNPJ"].astype(str)

df_top = df_top.merge(df[["FORN_CNPJ", "FORN_FANTASIA"]], left_on="PED_FORNECEDOR", right_on="FORN_CNPJ", how="inner")

top10 = (
    df_top.groupby("FORN_FANTASIA")
    .size()
    .reset_index(name="Quantidade de Pedidos")
    .sort_values(by="Quantidade de Pedidos", ascending=False)
    .head(10)
)

fig = px.bar(
    top10,
    x="Quantidade de Pedidos",
    y="FORN_FANTASIA",
    orientation="h",
    text="Quantidade de Pedidos",
    color="Quantidade de Pedidos",
    color_continuous_scale=["#7FC7FF", "#0066CC"],
    category_orders={"FORN_FANTASIA": top10["FORN_FANTASIA"].tolist()}
)

fig.update_layout(
    yaxis_title="Fornecedor",
    xaxis_title="Quantidade de Pedidos",
    showlegend=False,
    plot_bgcolor="rgba(0,0,0,0)",
    paper_bgcolor="rgba(0,0,0,0)",
    font=dict(color="white")
)

fig.update_traces(textposition="outside")
st.plotly_chart(fig, use_container_width=True)

# --- Exportar para Excel ---
def converter_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Fornecedores')
    return output.getvalue()

excel_bytes = converter_excel(tabela)

st.download_button(
    label="📥 Baixar tabela filtrada (.xlsx)",
    data=excel_bytes,
    file_name="fornecedores_filtrados.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
