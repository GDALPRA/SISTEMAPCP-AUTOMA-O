import streamlit as st

st.set_page_config(
    page_title="Portal de Automação - PCP & Produção",
    page_icon="⚙️",
    layout="wide"
)

# Executa diretamente o sistema de bandejas e PEAD
pg = st.Page("views/bandejas_pead.py", title="Cálculo de Bandejas & PEAD", icon="📦")
pg.run()