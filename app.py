import streamlit as st

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="NSF Tecnologia | Portal PCP Enterprise",
    page_icon="🕊️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- GERENCIAMENTO DE TEMA (CLARO / ESCURO) ---
with st.sidebar:
    st.markdown(
        """
        <div style="text-align: center; padding: 10px 0 5px 0;">
            <p style="font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 1.2px; margin-bottom: 6px;">
                ⚙️ Configuração de Tema
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    tema_selecionado = st.selectbox(
        "Modo de Exibição",
        options=["Escuro (Dark)", "Claro (Light)"],
        index=0,
        label_visibility="collapsed",
    )

is_dark = tema_selecionado == "Escuro (Dark)"

# --- PALETA DE CORES DINÂMICA (ALTO CONTRASTE GARANTIDO) ---
if is_dark:
    BG_APP = "#0E131F"
    BG_SIDEBAR = "#080C14"
    TEXT_MAIN = "#E2E8F0"
    TEXT_MUTED = "#94A3B8"
    BORDER_COLOR = "#1E293B"
    HOVER_COLOR = "#172033"
    UPLOADER_BG = "#131927"
    LOGO_TEXT = "#FFFFFF"
    SLOGAN_TEXT = "#64748B"
else:
    BG_APP = "#F8FAFC"
    BG_SIDEBAR = "#FFFFFF"
    TEXT_MAIN = "#1E293B"
    TEXT_MUTED = "#334155"
    BORDER_COLOR = "#CBD5E1"
    HOVER_COLOR = "#F1F5F9"
    UPLOADER_BG = "#F1F5F9"
    LOGO_TEXT = "#0A2540"
    SLOGAN_TEXT = "#475569"

# --- CSS GLOBAL DINÂMICO ---
st.markdown(
    f"""
    <style>
        /* 1. FUNDO E TEXTO GERAL */
        html, body, [data-testid="stAppViewContainer"], .main {{
            background-color: {BG_APP} !important;
            color: {TEXT_MAIN} !important;
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
        }}

        /* 2. SIDEBAR */
        [data-testid="stSidebar"] {{
            background-color: {BG_SIDEBAR} !important;
            border-right: 1px solid {BORDER_COLOR} !important;
        }}

        [data-testid="stSidebar"] * {{
            color: {TEXT_MUTED} !important;
        }}

        /* Título dos Grupos na Sidebar */
        [data-testid="stSidebar"] div[data-testid="stMarkdownContainer"] p strong,
        [data-testid="stSidebar"] span[data-testid="stHeader"] {{
            color: {TEXT_MUTED} !important;
            font-size: 11px !important;
            letter-spacing: 1.2px !important;
            text-transform: uppercase !important;
        }}

        /* 3. ABAS DE NAVEGAÇÃO */
        [data-testid="stSidebar"] div[role="radiogroup"] > label,
        [data-testid="stSidebarNav"] li a {{
            border-radius: 6px !important;
            padding: 8px 12px !important;
            margin-bottom: 2px !important;
            transition: all 0.15s ease-in-out !important;
        }}

        [data-testid="stSidebar"] div[role="radiogroup"] > label:hover,
        [data-testid="stSidebarNav"] li a:hover {{
            background-color: {HOVER_COLOR} !important;
        }}

        /* Item Selecionado na Sidebar */
        [data-testid="stSidebarNav"] li a[aria-current="page"],
        [data-testid="stSidebar"] div[role="radiogroup"] > label[aria-checked="true"] {{
            background-color: {HOVER_COLOR} !important;
            border-left: 3px solid #0052CC !important;
        }}

        [data-testid="stSidebarNav"] li a[aria-current="page"] span,
        [data-testid="stSidebar"] div[role="radiogroup"] > label[aria-checked="true"] span {{
            color: {LOGO_TEXT} !important;
            font-weight: 600 !important;
        }}

        /* 4. DRAG & DROP E FILE UPLOADER */
        [data-testid="stFileUploader"] {{
            background-color: {UPLOADER_BG} !important;
            border: 1px dashed {BORDER_COLOR} !important;
            border-radius: 8px !important;
            padding: 12px !important;
        }}

        [data-testid="stFileUploader"] section {{
            background-color: {UPLOADER_BG} !important;
            color: {TEXT_MUTED} !important;
        }}

        /* Botões Executivos */
        .stDownloadButton button, .stButton button[kind="primary"] {{
            background: linear-gradient(135deg, #0052CC 0%, #0A2540 100%) !important;
            color: #FFFFFF !important;
            font-weight: 500 !important;
            font-size: 13px !important;
            border-radius: 6px !important;
            border: 1px solid #1E40AF !important;
            padding: 8px 18px !important;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.2) !important;
        }}

        .stDownloadButton button:hover, .stButton button[kind="primary"]:hover {{
            background: linear-gradient(135deg, #0041A3 0%, #061729 100%) !important;
            border-color: #2563EB !important;
        }}

        /* Ocultar elementos padrão desnecessários */
        #MainMenu {{visibility: hidden;}}
        footer {{visibility: hidden;}}
        .stHeader, header[data-testid="stHeader"] {{
            background-color: transparent !important;
        }}
    </style>
""",
    unsafe_allow_html=True,
)

# --- SIDEBAR COM A LOGO OFICIAL 100% FIEL VETORIZADA ---
with st.sidebar:
    st.markdown(
        f"""
        <div style="text-align: center; padding: 15px 5px 10px 5px;">
            <svg width="220" height="75" viewBox="0 0 310 95" fill="none" xmlns="http://www.w3.org/2000/svg">
                <!-- Símbolo da Ave NSF (Fiel ao Original) -->
                <!-- Asa Superior Azul -->
                <path d="M 28 24 C 52 8, 94 11, 120 40 C 90 35, 56 33, 28 24 Z" fill="#2563EB"/>
                <!-- Cabeça e Bico Verde -->
                <path d="M 120 40 C 131 20, 136 13, 142 8 C 131 61, 87 75, 61 70 C 90 64, 110 52, 120 40 Z" fill="#22C55E"/>
                <!-- Base Laranja / Amarela (Cauda) -->
                <path d="M 22 59 C 47 82, 84 77, 123 44 C 80 53, 44 55, 22 59 Z" fill="#F97316"/>
                <path d="M 22 59 C 35 73, 58 75, 83 68 C 55 66, 35 64, 22 59 Z" fill="#EAB308"/>
                
                <!-- Tipografia NSF Perfeita -->
                <text x="146" y="63" font-family="'Segoe UI', Roboto, Helvetica, sans-serif" font-weight="900" font-size="58" fill="{LOGO_TEXT}" letter-spacing="2">NSF</text>
            </svg>
            
            <div style="margin-top: 4px; padding: 0 5px;">
                <p style="color: {LOGO_TEXT} !important; font-size: 10px; font-weight: 800; margin: 0; letter-spacing: 0.9px; text-transform: uppercase; line-height: 1.3;">
                    TECNOLOGIA EM INSTALAÇÕES COMERCIAIS
                </p>
                <p style="color: {SLOGAN_TEXT} !important; font-size: 9.5px; font-weight: 500; margin: 4px 0 0 0; line-height: 1.2;">
                    Desenvolvendo o varejo, transformando vidas.
                </p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(f"<hr style='border-color: {BORDER_COLOR}; margin: 10px 0;'>", unsafe_allow_html=True)
    st.markdown(
        f"<p style='color: {SLOGAN_TEXT} !important; font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 1.2px;'>Módulos Operacionais</p>",
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

materiais_page = st.Page(
    "views/materiais_padrao.py",
    title="Materiais Padrão (Refrigeração)",
    icon=":material/inventory_2:",
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
            materiais_page,
        ]
    }
)

# Rodapé minimalista na barra lateral
with st.sidebar:
    st.markdown(f"<hr style='border-color: {BORDER_COLOR}; margin: 15px 0 10px 0;'>", unsafe_allow_html=True)
    st.markdown(
        f"""
        <div style="text-align: center; color: {SLOGAN_TEXT} !important; font-size: 10px;">
            <p style="margin: 0; font-weight: 600; color: {LOGO_TEXT} !important;">NSF Enterprise v4.0</p>
            <p style="margin: 2px 0 0 0;">Ambiente Integrado de Fábrica</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

# Executa o roteador
pg.run()