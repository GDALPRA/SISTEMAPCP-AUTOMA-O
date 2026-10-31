import streamlit as st

st.set_page_config(
    page_title="Portal de Automação - PCP & Produção",
    page_icon="⚙️",
    layout="wide"
)

# Definindo as páginas do site
pg = st.navigation([
    st.Page("views/home.py", title="Início / Visão Geral", icon="🏠"),
    st.Page("views/bandejas_pead.py", title="Cálculo de Bandejas & PEAD", icon="📦"),
    # Quando tiver novas automações, basta adicioná-las aqui embaixo:
    # st.Page("views/futuro_sistema.py", title="Novo Sistema", icon="🔧"),
])

# Executa a navegação
pg.run()