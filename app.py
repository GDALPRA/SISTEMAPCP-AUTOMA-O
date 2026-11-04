import streamlit as st

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="PCP Enterprise | Gestão Executiva",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- CSS GLOBAL PARA ESTILIZAÇÃO EXECUTIVA (ENTERPRISE UI) ---
st.markdown(
    """
    <style>
        /* Estilização geral do fundo e fontes */
        .main {
            background-color: #F4F6F9;
        }
        
        /* Barra lateral corporativa */
        [data-testid="stSidebar"] {
            background-color: #1A2530;
            color: #FFFFFF;
        }
        
        [data-testid="stSidebar"] .stMarkdown h1, 
        [data-testid="stSidebar"] .stMarkdown h2, 
        [data-testid="stSidebar"] .stMarkdown h3,
        [data-testid="stSidebar"] span {
            color: #ECF0F1 !important;
        }

        /* Cartões de Métricas / KPIs */
        div[data-testid="metric-container"] {
            background-color: #FFFFFF;
            border: 1px solid #E2E8F0;
            padding: 15px 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.02);
            transition: all 0.3s ease;
        }
        div[data-testid="metric-container"]:hover {
            box-shadow: 0 4px 12px rgba(31,73,125,0.15);
            border-color: #1F497D;
        }

        /* Tabelas e Dataframes */
        dataframe {
            border-radius: 6px;
        }
        
        /* Botões de Download Executivos */
        .stDownloadButton button {
            background-color: #1F497D !important;
            color: white !important;
            font-weight: 600;
            border-radius: 6px;
            border: none;
            padding: 10px 20px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            transition: background-color 0.2s;
        }
        .stDownloadButton button:hover {
            background-color: #16335B !important;
        }
        
        /* Ocultar elementos padrão do Streamlit para visual limpo */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
    </style>
""",
    unsafe_allow_html=True,
)

# --- SIDEBAR EXECUTIVA ---
with st.sidebar:
  st.markdown(
      """
        <div style="text-align: center; padding: 10px 0 20px 0;">
            <h2 style="color: #FFFFFF; margin: 0; font-size: 20px; font-weight: 700;">INDÚSTRIA DE EXPOSITORES</h2>
            <p style="color: #94A3B8; font-size: 12px; margin: 5px 0 0 0;">Portal Corporativo de PCP & Engenharia</p>
        </div>
    """,
      unsafe_allow_html=True,
  )

  st.markdown("---")
  st.markdown(
      "<p style='color: #94A3B8; font-size: 11px; font-weight: 600; text-transform:"
      " uppercase; letter-spacing: 1px;'>Módulos Operacionais</p>",
      unsafe_allow_html=True,
  )

# --- DEFINIÇÃO DAS PÁGINAS DO SISTEMA ---
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

# --- NAVEGAÇÃO CORPORATIVA ---
pg = st.navigation(
    {
        "PLANEJAMENTO E CONTROLE (PCP)": [
            bandejas_page,
            tubos_page,
            portas_page,
        ]
    }
)

# Rodapé discreto na barra lateral
with st.sidebar:
  st.markdown("---")
  st.markdown(
      """
        <div style="text-align: center; color: #64748B; font-size: 11px;">
            <p style="margin: 0;">Versão 3.5 Enterprise</p>
            <p style="margin: 2px 0 0 0;">Ambiente Integrado de Fábrica</p>
        </div>
    """,
      unsafe_allow_html=True,
  )

pg.run()