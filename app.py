import streamlit as st

st.set_page_config(
    page_title="Portal de Automação - PCP & Produção",
    page_icon="⚙️",
    layout="wide"
)

# Executa o ficheiro da vista diretamente
with open("views/bandejas_pead.py", encoding="utf-8") as f:
    exec(f.read())