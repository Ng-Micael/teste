"""Dashboard: Educação e Economia na América Latina.

Dados: API pública do Banco Mundial (sem necessidade de chave).
Rodar com: streamlit run app.py
"""

import pandas as pd
import plotly.express as px
import requests
import streamlit as st

st.set_page_config(
    page_title="Educação e Economia na América Latina",
    page_icon="📊",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Configurações
# ---------------------------------------------------------------------------
PAISES = {
    "Brasil": "BRA",
    "Argentina": "ARG",
    "Chile": "CHL",
    "Colômbia": "COL",
    "México": "MEX",
    "Peru": "PER",
    "Uruguai": "URY",
    "Equador": "ECU",
    "Bolívia": "BOL",
    "Paraguai": "PRY",
}
CODIGO_PARA_PAIS = {codigo: nome for nome, codigo in PAISES.items()}

INDICADORES = {
    "PIB per capita (US$)": "NY.GDP.PCAP.CD",
    "Gasto público em educação (% do PIB)": "SE.XPD.TOTL.GD.ZS",
    "Matrícula no ensino médio (% bruta)": "SE.SEC.ENRR",
}

ANO_INICIAL, ANO_FINAL = 2000, 2024


# ---------------------------------------------------------------------------
# Dados
# ---------------------------------------------------------------------------
@st.cache_data(ttl=60 * 60 * 24, show_spinner="Baixando dados do Banco Mundial...")
def baixar_indicador(codigo_indicador: str) -> pd.DataFrame:
    """Baixa um indicador do Banco Mundial para todos os países da lista."""
    paises = ";".join(PAISES.values())
    url = f"https://api.worldbank.org/v2/country/{paises}/indicator/{codigo_indicador}"
    resposta = requests.get(
        url,
        params={
            "format": "json",
            "per_page": 2000,
            "date": f"{ANO_INICIAL}:{ANO_FINAL}",
        },
        timeout=30,
    )
    resposta.raise_for_status()
    registros = resposta.json()[1] or []

    linhas = [
        {
            "pais": CODIGO_PARA_PAIS.get(r["countryiso3code"], r["country"]["value"]),
            "ano": int(r["date"]),
            "valor": r["value"],
        }
        for r in registros
    ]
    df = pd.DataFrame(linhas, columns=["pais", "ano", "valor"])
    return df.dropna(subset=["valor"])


@st.cache_data(ttl=60 * 60 * 24)
def carregar_dados() -> pd.DataFrame:
    """Junta todos os indicadores em uma tabela: uma linha por país e ano."""
    tabelas = []
    for nome, codigo in INDICADORES.items():
        df = baixar_indicador(codigo)
        df["indicador"] = nome
        tabelas.append(df)

    longo = pd.concat(tabelas, ignore_index=True)
    largo = longo.pivot_table(
        index=["pais", "ano"], columns="indicador", values="valor"
    ).reset_index()
    return largo


# ---------------------------------------------------------------------------
# Interface
# ---------------------------------------------------------------------------
st.title("📊 Educação e Economia na América Latina")
st.caption(
    "Como o desenvolvimento econômico e os investimentos em educação "
    "se relacionam nos países latino-americanos? Fonte: Banco Mundial."
)

try:
    dados = carregar_dados()
except requests.RequestException:
    st.error("Não foi possível baixar os dados agora. Tente novamente em instantes.")
    st.stop()

# Filtros na barra lateral
st.sidebar.header("Filtros")
paises_escolhidos = st.sidebar.multiselect(
    "Países", options=list(PAISES.keys()), default=list(PAISES.keys())
)
indicador = st.sidebar.selectbox("Indicador para a evolução", list(INDICADORES.keys()))
anos = st.sidebar.slider(
    "Período", ANO_INICIAL, ANO_FINAL, (ANO_INICIAL, ANO_FINAL)
)

if not paises_escolhidos:
    st.warning("Selecione pelo menos um país na barra lateral.")
    st.stop()

filtrado = dados[
    dados["pais"].isin(paises_escolhidos) & dados["ano"].between(anos[0], anos[1])
]

# Gráfico 1: evolução ao longo do tempo
st.subheader(f"Evolução: {indicador}")
serie = filtrado.dropna(subset=[indicador])
if serie.empty:
    st.info("Sem dados para essa combinação de filtros.")
else:
    fig_linha = px.line(
        serie,
        x="ano",
        y=indicador,
        color="pais",
        markers=True,
        labels={"ano": "Ano", "pais": "País"},
    )
    st.plotly_chart(fig_linha, use_container_width=True)

# Gráfico 2: relação entre PIB per capita e gasto em educação
st.subheader("Países mais ricos investem mais em educação?")
x_col = "PIB per capita (US$)"
y_col = "Gasto público em educação (% do PIB)"

anos_com_dados = sorted(filtrado.dropna(subset=[x_col, y_col])["ano"].unique())
if not anos_com_dados:
    st.info("Sem dados suficientes para comparar PIB e gasto em educação.")
else:
    ano_escolhido = st.select_slider(
        "Ano da comparação", options=anos_com_dados, value=anos_com_dados[-1]
    )
    recorte = filtrado[
        (filtrado["ano"] == ano_escolhido)
    ].dropna(subset=[x_col, y_col])

    fig_disp = px.scatter(
        recorte,
        x=x_col,
        y=y_col,
        color="pais",
        text="pais",
        size_max=20,
    )
    fig_disp.update_traces(marker={"size": 14}, textposition="top center")
    st.plotly_chart(fig_disp, use_container_width=True)

# Tabela com os dados
with st.expander("Ver tabela de dados"):
    st.dataframe(filtrado.sort_values(["pais", "ano"]), use_container_width=True)

# Espaço para as SUAS conclusões (isso é o que mais impressiona no LinkedIn!)
st.subheader("O que os dados mostram")
st.markdown(
    """
    - *Escreva aqui a sua 1ª conclusão (ex.: tendência do Brasil vs. Chile).*
    - *Escreva aqui a 2ª conclusão.*
    - *Escreva aqui as limitações da análise (ex.: dados ausentes em alguns anos).*
    """
)
