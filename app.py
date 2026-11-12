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
        <div style="text-align: center; padding: 5px 0 2px 0;">
            <p style="font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 1.2px; margin-bottom: 4px;">
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

# --- PALETA DE CORES DINÂMICA (PADRÃO MERCADO ERP) ---
if is_dark:
    BG_APP = "#0B0F19"
    BG_SIDEBAR = "#05070D"
    TEXT_MAIN = "#F1F5F9"
    TEXT_MUTED = "#94A3B8"
    BORDER_COLOR = "#1E293B"
    HOVER_COLOR = "#172033"
    UPLOADER_BG = "#131927"
    LOGO_TEXT = "#FFFFFF"
    SLOGAN_TEXT = "#64748B"
    CARD_BG = "#111827"
    KPI_BG = "#131B2E"
else:
    BG_APP = "#F8FAFC"
    BG_SIDEBAR = "#FFFFFF"
    TEXT_MAIN = "#0F172A"
    TEXT_MUTED = "#475569"
    BORDER_COLOR = "#E2E8F0"
    HOVER_COLOR = "#F1F5F9"
    UPLOADER_BG = "#F1F5F9"
    LOGO_TEXT = "#1E3A8A"
    SLOGAN_TEXT = "#64748B"
    CARD_BG = "#FFFFFF"
    KPI_BG = "#FFFFFF"

# --- CSS GLOBAL DINÂMICO & ESTILO ERP PROFISSIONAL ---
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
            border-left: 3px solid #2563EB !important;
        }}

        [data-testid="stSidebarNav"] li a[aria-current="page"] span,
        [data-testid="stSidebar"] div[role="radiogroup"] > label[aria-checked="true"] span {{
            color: {LOGO_TEXT} !important;
            font-weight: 600 !important;
        }}

        /* 4. CARDS ESTILO ERP CORPORATIVO */
        .erp-kpi-card {{
            background-color: {KPI_BG};
            border: 1px solid {BORDER_COLOR};
            border-radius: 10px;
            padding: 18px 20px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        }}

        /* Estilização Perfeita, Uniforme e Alinhada dos Botões de Acesso Rápido */
        .stButton button {{
            background-color: {CARD_BG} !important;
            color: {TEXT_MAIN} !important;
            border: 1px solid {BORDER_COLOR} !important;
            border-radius: 10px !important;
            padding: 15px 20px !important;
            text-align: center !important;
            font-size: 14px !important;
            font-weight: 700 !important;
            width: 100% !important;
            height: 70px !important;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05) !important;
            transition: all 0.2s ease-in-out !important;
        }}

        .stButton button:hover {{
            border-color: #2563EB !important;
            box-shadow: 0 10px 15px -3px rgba(37, 99, 235, 0.15) !important;
            transform: translateY(-2px);
        }}

        /* Ocultar elementos padrão desnecessários do Streamlit */
        #MainMenu {{visibility: hidden;}}
        footer {{visibility: hidden;}}
        .stHeader, header[data-testid="stHeader"] {{
            background-color: transparent !important;
        }}
    </style>
""",
    unsafe_allow_html=True,
)

# --- SIDEBAR COM O LOGOTIPO OFICIAL NSF ---
with st.sidebar:
    st.markdown(
        f"""
        <div style="text-align: center; padding: 15px 5px 10px 5px;">
            <svg width="210" height="65" viewBox="0 0 320 95" fill="none" xmlns="http://www.w3.org/2000/svg">
                <!-- Asa Superior Azul -->
                <path d="M 15 28 C 35 10, 75 12, 102 42 C 78 38, 45 35, 15 28 Z" fill="#2563EB"/>
                <!-- Cabeça e Bico Verde -->
                <path d="M 102 42 C 112 22, 116 15, 122 10 C 112 58, 70 72, 46 66 C 70 60, 88 50, 102 42 Z" fill="#22C55E"/>
                <!-- Base Laranja / Amarela (Cauda) -->
                <path d="M 10 56 C 32 76, 65 72, 102 44 C 68 52, 36 54, 10 56 Z" fill="#F97316"/>
                <path d="M 10 56 C 22 66, 45 68, 68 60 C 42 58, 26 56, 10 56 Z" fill="#EAB308"/>
                <!-- Tipografia NSF -->
                <text x="135" y="62" font-family="'Segoe UI', Roboto, Helvetica, sans-serif" font-weight="900" font-size="56" fill="{LOGO_TEXT}" letter-spacing="2">NSF</text>
            </svg>
            <div style="margin-top: 4px; padding: 0 5px;">
                <p style="color: {LOGO_TEXT} !important; font-size: 9.5px; font-weight: 800; margin: 0; letter-spacing: 0.9px; text-transform: uppercase; line-height: 1.3;">
                    TECNOLOGIA EM INSTALAÇÕES COMERCIAIS
                </p>
                <p style="color: {SLOGAN_TEXT} !important; font-size: 9px; font-weight: 500; margin: 3px 0 0 0; line-height: 1.2;">
                    Desenvolvendo o varejo, transformando vidas.
                </p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(f"<hr style='border-color: {BORDER_COLOR}; margin: 10px 0;'>", unsafe_allow_html=True)
    st.markdown(
        f"<p style='color: {SLOGAN_TEXT} !important; font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 1.2px;'>Navegação do Sistema</p>",
        unsafe_allow_html=True,
    )

# --- DEFINIÇÃO DOS MÓDULOS E PÁGINAS DO SISTEMA ---
bandejas_page = st.Page(
    "views/bandejas_pead.py",
    title="Bandejas ABS & PEAD",
    icon=":material/layers:",
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

# --- TELA INICIAL (DASHBOARD ERP PROFISSIONAL) ---
def render_home():
    st.markdown(
        """
        <div style="padding: 10px 0 20px 0;">
            <h1 style="font-size: 26px; font-weight: 800; margin-bottom: 5px; display: flex; align-items: center; gap: 10px;">
                🕊️ Portal PCP Enterprise — Painel de Controle Executivo
            </h1>
            <p style="font-size: 14px; color: #94A3B8; margin: 0;">
                Sistema de Gestão de Produção, Planejamento de Insumos e Automação de Engenharia para Refrigeração Comercial.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Cards de Indicadores Estilo ERP
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    with kpi1:
        st.markdown(
            f"""
            <div class="erp-kpi-card">
                <span style="font-size: 11px; font-weight: 700; color: #64748B; text-transform: uppercase; letter-spacing: 0.8px;">Módulos Ativos</span>
                <h3 style="margin: 5px 0 0 0; font-size: 22px; font-weight: 700; color: {TEXT_MAIN};">7 Subsistemas</h3>
                <span style="font-size: 11px; color: #22C55E; font-weight: 600;">● Operacional & Estável</span>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with kpi2:
        st.markdown(
            f"""
            <div class="erp-kpi-card">
                <span style="font-size: 11px; font-weight: 700; color: #64748B; text-transform: uppercase; letter-spacing: 0.8px;">Linhas de Fábrica</span>
                <h3 style="margin: 5px 0 0 0; font-size: 22px; font-weight: 700; color: {TEXT_MAIN};">VA, VCA & VRA</h3>
                <span style="font-size: 11px; color: #3B82F6; font-weight: 600;">◆ Integrado ao PCP</span>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with kpi3:
        st.markdown(
            f"""
            <div class="erp-kpi-card">
                <span style="font-size: 11px; font-weight: 700; color: #64748B; text-transform: uppercase; letter-spacing: 0.8px;">Insumos Gerenciados</span>
                <h3 style="margin: 5px 0 0 0; font-size: 22px; font-weight: 700; color: {TEXT_MAIN};">Aço, PEAD & ABS</h3>
                <span style="font-size: 11px; color: #22C55E; font-weight: 600;">⚡ Cálculo Dinâmico</span>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with kpi4:
        st.markdown(
            f"""
            <div class="erp-kpi-card">
                <span style="font-size: 11px; font-weight: 700; color: #64748B; text-transform: uppercase; letter-spacing: 0.8px;">Ambiente</span>
                <h3 style="margin: 5px 0 0 0; font-size: 22px; font-weight: 700; color: {TEXT_MAIN};">Enterprise v4.0</h3>
                <span style="font-size: 11px; color: #F59E0B; font-weight: 600;">★ Alta Performance</span>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        f"""
        <h3 style="font-size: 18px; font-weight: 700; color: {TEXT_MAIN}; margin-bottom: 15px;">
            🚀 Acesso Rápido aos Módulos Operacionais
        </h3>
        """,
        unsafe_allow_html=True,
    )

    # Grid de Acesso Rápido com 3 Colunas (Alinhadas, Tamanhos Iguais e Simétricas para as 7 páginas)
    col1, col2, col3 = st.columns(3)

    with col1:
        if st.button("📦 Bandejas ABS & PEAD", key="btn_bandejas"):
            st.switch_page(bandejas_page)
        st.markdown("<div style='margin-bottom: 10px;'></div>", unsafe_allow_html=True)
        if st.button("🚪 Portas & Frames (VA & VAP)", key="btn_portas"):
            st.switch_page(portas_page)

    with col2:
        if st.button("📐 Cálculo de Tubos (Chassis)", key="btn_tubos"):
            st.switch_page(tubos_page)
        st.markdown("<div style='margin-bottom: 10px;'></div>", unsafe_allow_html=True)
        if st.button("🛡️ Chapas de Aço Inox", key="btn_inox"):
            st.switch_page(inox_page)

    with col3:
        if st.button("📄 Chapas Cristal VA", key="btn_chapas"):
            st.switch_page(chapas_page)
        st.markdown("<div style='margin-bottom: 10px;'></div>", unsafe_allow_html=True)
        if st.button("📊 Relatórios Gerenciais (ORFAs)", key="btn_gerencial"):
            st.switch_page(gerencial_page)

    st.markdown("<div style='margin-bottom: 10px;'></div>", unsafe_allow_html=True)
    
    # 7º Botão centralizado na grade, mantendo exatamente o mesmo tamanho e padrão
    c_left, c_center, c_right = st.columns(3)
    with c_center:
        if st.button("⚙️ Materiais Padrão", key="btn_materiais"):
            st.switch_page(materiais_page)

home_page = st.Page(
    render_home,
    title="Início (Portal ERP)",
    icon=":material/dashboard:",
    default=True,
)

# --- NAVEGAÇÃO CORPORATIVA DA PLATAFORMA ---
pg = st.navigation(
    {
        "GERENCIAMENTO GERAL": [
            home_page,
        ],
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