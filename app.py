import streamlit as st

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="NSF Tecnologia | Portal PCP Enterprise",
    page_icon="🕊️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- CSS GLOBAL EXECUTIVO SLIM (OVERRIDE TOTAL DE TEMA) ---
st.markdown(
    """
    <style>
        /* 1. FORÇAR COR DE FUNDO DA ÁREA DE TRABALHO (SLIM OFF-WHITE) */
        html, body, [data-testid="stAppViewContainer"], .main {
            background-color: #F8FAFC !important;
            color: #0F172A !important;
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
        }

        /* 2. SIDEBAR EXECUTIVA SLIM (DARK NAVY NSF) */
        [data-testid="stSidebar"] {
            background-color: #0F172A !important;
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

        /* 3. ESTILIZAÇÃO DAS ABAS DE NAVEGAÇÃO (SLIM & MODERNO) */
        [data-testid="stSidebar"] div[role="radiogroup"] > label,
        [data-testid="stSidebarNav"] li a {
            border-radius: 6px !important;
            padding: 8px 12px !important;
            margin-bottom: 2px !important;
            transition: all 0.15s ease-in-out !important;
        }

        [data-testid="stSidebar"] div[role="radiogroup"] > label:hover,
        [data-testid="stSidebarNav"] li a:hover {
            background-color: #1E293B !important;
        }

        /* Item Selecionado na Sidebar */
        [data-testid="stSidebarNav"] li a[aria-current="page"],
        [data-testid="stSidebar"] div[role="radiogroup"] > label[aria-checked="true"] {
            background-color: #1E293B !important;
            border-left: 3px solid #0052CC !important;
        }

        [data-testid="stSidebarNav"] li a[aria-current="page"] span,
        [data-testid="stSidebar"] div[role="radiogroup"] > label[aria-checked="true"] span {
            color: #FFFFFF !important;
            font-weight: 600 !important;
        }

        /* 4. CABEÇALHO SLIM (SEM CAIXAS AZUIS GIGANTES) */
        .stHeader, header[data-testid="stHeader"] {
            background-color: transparent !important;
        }

        /* Cartões de Entrada / Drag & Drop Slim */
        [data-testid="stFileUploader"] {
            background-color: #FFFFFF !important;
            border: 1px dashed #CBD5E1 !important;
            border-radius: 8px !important;
            padding: 10px !important;
            box-shadow: 0 1px 3px rgba(0,0,0,0.02) !important;
        }

        [data-testid="stFileUploader"] section {
            background-color: #FFFFFF !important;
        }

        /* Botões Slim e Executivos */
        .stDownloadButton button, .stButton button[kind="primary"] {
            background-color: #0052CC !important;
            color: #FFFFFF !important;
            font-weight: 500 !important;
            font-size: 13px !important;
            border-radius: 6px !important;
            border: none !important;
            padding: 8px 18px !important;
            box-shadow: 0 1px 2px rgba(0, 0, 0, 0.05) !important;
        }

        .stDownloadButton button:hover, .stButton button[kind="primary"]:hover {
            background-color: #0041A3 !important;
        }

        /* Ocultar elementos desnecessários da barra superior */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
    </style>
""",
    unsafe_allow_html=True,
)

# --- SIDEBAR COM O LOGO OFICIAL VETORIZADO NSF ---
with st.sidebar:
    # Logo oficial vetorizado fiel ao manual de marca NSF (Pássaro estilizado + Tipografia)
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

# --- CONFIGURAÇÃO DAS PÁGINAS COM ÍCONES SLIM (MATERIAL DESIGN) ---
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

# --- ESTRUTURA DE NAVEGAÇÃO DA PLATAFORMA ---
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

# Executa o roteador do Streamlit
pg.run()