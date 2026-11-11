import streamlit as st

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="NSF Tecnologia | Portal PCP Enterprise",
    page_icon="🕊️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- CSS GLOBAL DARK EXECUTIVO SLIM ---
st.markdown(
    """
    <style>
        /* 1. FORÇAR MODO ESCURO ENTERPRISE NA PÁGINA INTEIRA */
        html, body, [data-testid="stAppViewContainer"], .main {
            background-color: #0E131F !important;
            color: #E2E8F0 !important;
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
        }

        /* 2. SIDEBAR ESCURA E SLIM */
        [data-testid="stSidebar"] {
            background-color: #080C14 !important;
            border-right: 1px solid #1E293B !important;
        }

        [data-testid="stSidebar"] * {
            color: #94A3B8 !important;
        }

        /* Título dos Grupos na Sidebar */
        [data-testid="stSidebar"] div[data-testid="stMarkdownContainer"] p strong,
        [data-testid="stSidebar"] span[data-testid="stHeader"] {
            color: #64748B !important;
            font-size: 11px !important;
            letter-spacing: 1.2px !important;
            text-transform: uppercase !important;
        }

        /* 3. ESTILIZAÇÃO DAS ABAS DE NAVEGAÇÃO */
        [data-testid="stSidebar"] div[role="radiogroup"] > label,
        [data-testid="stSidebarNav"] li a {
            border-radius: 6px !important;
            padding: 8px 12px !important;
            margin-bottom: 2px !important;
            transition: all 0.15s ease-in-out !important;
        }

        [data-testid="stSidebar"] div[role="radiogroup"] > label:hover,
        [data-testid="stSidebarNav"] li a:hover {
            background-color: #172033 !important;
        }

        /* Item Selecionado na Sidebar */
        [data-testid="stSidebarNav"] li a[aria-current="page"],
        [data-testid="stSidebar"] div[role="radiogroup"] > label[aria-checked="true"] {
            background-color: #172033 !important;
            border-left: 3px solid #0052CC !important;
        }

        [data-testid="stSidebarNav"] li a[aria-current="page"] span,
        [data-testid="stSidebar"] div[role="radiogroup"] > label[aria-checked="true"] span {
            color: #FFFFFF !important;
            font-weight: 600 !important;
        }

        /* 4. DRAG & DROP E FILE UPLOADER DARK SLIM */
        [data-testid="stFileUploader"] {
            background-color: #131927 !important;
            border: 1px dashed #334155 !important;
            border-radius: 8px !important;
            padding: 12px !important;
        }

        [data-testid="stFileUploader"] section {
            background-color: #131927 !important;
            color: #94A3B8 !important;
        }

        /* Botões Dark Executivos */
        .stDownloadButton button, .stButton button[kind="primary"] {
            background: linear-gradient(135deg, #0052CC 0%, #0A2540 100%) !important;
            color: #FFFFFF !important;
            font-weight: 500 !important;
            font-size: 13px !important;
            border-radius: 6px !important;
            border: 1px solid #1E40AF !important;
            padding: 8px 18px !important;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.2) !important;
        }

        .stDownloadButton button:hover, .stButton button[kind="primary"]:hover {
            background: linear-gradient(135deg, #0041A3 0%, #061729 100%) !important;
            border-color: #2563EB !important;
        }

        /* Ocultar cabeçalhos desnecessários */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        .stHeader, header[data-testid="stHeader"] {
            background-color: transparent !important;
        }
    </style>
""",
    unsafe_allow_html=True,
)

# --- SIDEBAR COM O LOGO OFICIAL VETORIZADO NSF ---
with st.sidebar:
    st.markdown(
        """
        <div style="text-align: center; padding: 20px 0 15px 0;">
            <svg width="190" height="60" viewBox="0 0 280 85" fill="none" xmlns="http://www.w3.org/2000/svg">
                <!-- Asa Superior Azul NSF -->
                <path d="M 25 22 C 45 8, 85 10, 108 36 C 82 32, 50 30, 25 22 Z" fill="#0052CC"/>
                <!-- Asa Verde Natureza NSF -->
                <path d="M 108 36 C 118 18, 122 12, 128 8 C 118 55, 78 68, 55 64 C 80 58, 98 48, 108 36 Z" fill="#70C04B"/>
                <!-- Base Laranja / Amarela NSF -->
                <path d="M 20 54 C 42 74, 75 70, 110 40 C 72 48, 40 50, 20 54 Z" fill="#FF8A00"/>
                <path d="M 20 54 C 32 66, 52 68, 75 62 C 50 60, 32 58, 20 54 Z" fill="#FFC700"/>
                <!-- Tipografia NSF -->
                <text x="140" y="58" font-family="'Segoe UI', Roboto, Helvetica, sans-serif" font-weight="900" font-size="56" fill="#FFFFFF" letter-spacing="1.5">NSF</text>
            </svg>
            <div style="margin-top: 6px;">
                <p style="color: #FFFFFF !important; font-size: 11px; font-weight: 700; margin: 0; letter-spacing: 0.8px; text-transform: uppercase;">
                    TECNOLOGIA EM INSTALAÇÕES COMERCIAIS
                </p>
                <p style="color: #64748B !important; font-size: 10px; margin: 3px 0 0 0;">
                    Desenvolvendo o varejo, transformando vidas.
                </p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("<hr style='border-color: #1E293B; margin: 10px 0;'>", unsafe_allow_html=True)
    st.markdown(
        "<p style='color: #64748B !important; font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 1.2px;'>Módulos Operacionais</p>",
        unsafe_allow_html=True,
    )

# --- DEFINIÇÃO DAS PÁGINAS DO SISTEMA COM ÍCONES SLIM ---
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

# Rodapé minimalista na barra lateral
with st.sidebar:
    st.markdown("<hr style='border-color: #1E293B; margin: 15px 0 10px 0;'>", unsafe_allow_html=True)
    st.markdown(
        """
        <div style="text-align: center; color: #475569 !important; font-size: 10px;">
            <p style="margin: 0; font-weight: 600; color: #64748B !important;">NSF Enterprise v4.0</p>
            <p style="margin: 2px 0 0 0;">Ambiente Integrado de Fábrica</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

# Executa o roteador
pg.run()