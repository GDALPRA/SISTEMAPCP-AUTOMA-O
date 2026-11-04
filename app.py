import streamlit as st

st.set_page_config(
    page_title="Portal PCP & Engenharia",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Definição das páginas do portal com os nomes corretos dos arquivos
bandejas_page = st.Page(
    "views/bandejas_pead.py",
    title="Bandejas ABS & PEAD",
    icon="📦",
    default=True,
)

tubos_page = st.Page(
    "views/calculo_tubos.py",
    title="Cálculo de Tubos (Chassis)",
    icon="🧪",
)

portas_page = st.Page(
    "views/calculo_portas.py",
    title="Portas & Frames (Vertical Alto)",
    icon="🚪",
)

# Configuração da Navegação Lateral
pg = st.navigation(
    {
        "Módulos Operacionais PCP": [
            bandejas_page,
            tubos_page,
            portas_page,
        ]
    }
)

pg.run()