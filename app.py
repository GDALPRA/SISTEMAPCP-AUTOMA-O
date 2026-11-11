import streamlit as st

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="NSF Tecnologia | Portal PCP & Engenharia",
    page_icon="🕊️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- CSS GLOBAL COM A NOVA IDENTIDADE VISUAL NSF ENTERPRISE UI ---
st.markdown(
    """
    <style>
        /* Fundo principal e tipografia */
        .main {
            background-color: #F8FAFC;
            font-family: 'Inter', 'Segoe UI', sans-serif;
        }
        
        /* Sidebar executiva com a cor de marca NSF (Deep Royal Navy) */
        [data-testid="stSidebar"] {
            background-color: #0A2540 !important;
            border-right: 1px solid #1E3A8A;
        }
        
        /* Cor de textos no Menu Lateral */
        [data-testid="stSidebar"] .stMarkdown h1, 
        [data-testid="stSidebar"] .stMarkdown h2, 
        [data-testid="stSidebar"] .stMarkdown h3,
        [data-testid="stSidebar"] span,
        [data-testid="stSidebar"] p {
            color: #F8FAFC !important;
        }
        
        /* Estilização da seleção de navegadores (Radio e st.Page Navigation) */
        [data-testid="stSidebar"] div[role="radiogroup"] > label {
            background-color: transparent;
            border-radius: 6px;
            padding: 8px 12px;
            margin-bottom: 4px;
            transition: all 0.2s ease;
        }
        
        [data-testid="stSidebar"] div[role="radiogroup"] > label:hover {
            background-color: rgba(255, 255, 255, 0.08);
        }
        
        /* Cartões de Métricas Executivas */
        div[data-testid="metric-container"] {
            background-color: #FFFFFF;
            border: 1px solid #E2E8F0;
            border-left: 4px solid #0052CC;
            padding: 18px 22px;
            border-radius: 8px;
            box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);
        }
        
        /* Botão Primário e de Download com padrão NSF */
        .stDownloadButton button, .stButton button[kind="primary"] {
            background: linear-gradient(135deg, #0052CC 0%, #0A2540 100%) !important;
            color: white !important;
            font-weight: 600;
            border-radius: 6px;
            border: none;
            padding: 10px 24px;
            box-shadow: 0 2px 4px rgba(0, 82, 204, 0.25);
            transition: all 0.3s ease;
        }
        
        .stDownloadButton button:hover, .stButton button[kind="primary"]:hover {
            background: linear-gradient(135deg, #003D99 0%, #061729 100%) !important;
            box-shadow: 0 4px 8px rgba(0, 82, 204, 0.4);
            transform: translateY(-1px);
        }

        /* Ocultar elementos padrão da interface Streamlit */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
    </style>
""",
    unsafe_allow_html=True,
)

# --- SIDEBAR EXECUTIVA DA NSF TECNOLOGIA ---
with st.sidebar:
    # Logotipo vetorial da NSF com o símbolo da pomba estilizada e tipografia oficial
    st.markdown(
        """
        <div style="text-align: center; padding: 15px 0 10px 0;">
            <svg width="180" height="50" viewBox="0 0 350 90" fill="none" xmlns="http://www.w3.org/2000/svg">
                <!-- Asa Azul da Pomba NSF -->
                <path d="M 20 20 Q 90 0, 110 50 Q 70 45, 20 20 Z" fill="#0052CC"/>
                <!-- Asa Verde da Pomba NSF -->
                <path d="M 110 50 Q 115 15, 125 10 Q 115 70, 75 75 Q 95 65, 110 50 Z" fill="#70C04B"/>
                <!-- Corpo Amarelo/Laranja NSF -->
                <path d="M 20 60 Q 60 90, 125 45 Q 80 50, 20 60 Z" fill="#FFC700"/>
                <!-- Texto NSF -->
                <text x="145" y="65" font-family="'Segoe UI', Roboto, sans-serif" font-weight="900" font-size="64" fill="#FFFFFF" letter-spacing="2">NSF</text>
            </svg>
            <p style="color: #70C04B; font-size: 11px; font-weight: 700; margin: 4px 0 0 0; letter-spacing: 1.5px; text-transform: uppercase;">
                Tecnologia em Instalações Comerciais
            </p>
            <p style="color: #94A3B8; font-size: 10px; margin: 2px 0 0 0;">
                Desenvolvendo o varejo, transformando vidas.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("---")
    st.markdown(
        "<p style='color: #70C04B; font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 1.2px; margin-bottom: 8px;'>Módulos de Engenharia & PCP</p>",
        unsafe_allow_html=True,
    )

# --- DEFINIÇÃO DAS PÁGINAS DO SISTEMA COM ÍCONES EXECUTIVOS ---
bandejas_page = st.Page(
    "views/bandejas_pead.py",
    title="Bandejas ABS & PEAD",
    icon=":material/layers:",
    default=True,
)

tubos_page = st.Page(
    "views/calculo_tubos.py",
    title="Cálculo de Tubos (Chassis)",
    icon=":material/grid_on:",
)

portas_page = st.Page(
    "views/calculo_portas.py",
    title="Portas & Frames (VA & VAP)",
    icon=":material/door_front:",
)

chapas_page = st.Page(
    "views/calculo_chapas.py",
    title="Chapas Cristal VA (Bojo/Costas/Teto)",
    icon=":material/description:",
)

inox_page = st.Page(
    "views/calculo_inox.py",
    title="Chapas de Aço Inox (Expositores VA)",
    icon=":material/shield:",
)

gerencial_page = st.Page(
    "views/processar_of.py",
    title="Relatórios Gerenciais (ORFAs)",
    icon=":material/analytics:",
)

# --- NAVEGAÇÃO CORPORATIVA DA PLATAFORMA ---
pg = st.navigation(
    {
        "PLANEJAMENTO E CONTROLE (PCP)": [
            bandejas_page,
            tubos_page,
            portas_page,
            chapas_page,
            inox_page,
            gerencial_page,
        ]
    }
)

# Rodapé institucional na barra lateral
with st.sidebar:
    st.markdown("---")
    st.markdown(
        """
        <div style="text-align: center; color: #64748B; font-size: 11px;">
            <p style="margin: 0; font-weight: 600; color: #94A3B8;">NSF Platform v4.0 Enterprise</p>
            <p style="margin: 3px 0 0 0; font-size: 10px;">Ambiente Integrado de Fábrica</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

# Executa o roteamento das páginas
pg.run()