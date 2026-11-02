import streamlit as st

st.set_page_config(
    page_title="Portal de Automação Industrial | PCP",
    page_icon="⚙️",
    layout="wide"
)

# Definindo as páginas do portal no menu lateral
pg = st.navigation([
    st.Page("views/bandejas_pead.py", title="Bandejas ABS & PEAD", icon="📦"),
    st.Page("views/calculo_tubos.py", title="Cálculo de Tubos", icon="🧪"),
])

# Executa a navegação selecionada
pg.run()