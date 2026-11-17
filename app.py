import streamlit as st

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="NSF Tecnologia | Portal PCP Enterprise",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- GERENCIAMENTO DE TEMA (CLARO / ESCURO) ---
with st.sidebar:
    st.markdown(
        """
        <div style="text-align: center; padding: 2px 0 2px 0;">
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

# --- PALETA DE CORES INSTITUCIONAL (PADRÃO ERP CORPORATIVO) ---
if is_dark:
    BG_APP = "#0B0F19"
    BG_SIDEBAR = "#05070D"
    TEXT_MAIN = "#F1F5F9"
    TEXT_MUTED = "#94A3B8"
    BORDER_COLOR = "#1E293B"
    HOVER_COLOR = "#172033"
    LOGO_TEXT = "#FFFFFF"
    SLOGAN_TEXT = "#94A3B8"
    CARD_BG = "#111827"
    KPI_BG = "#131B2E"
else:
    BG_APP = "#F8FAFC"
    BG_SIDEBAR = "#FFFFFF"
    TEXT_MAIN = "#0F172A"
    TEXT_MUTED = "#475569"
    BORDER_COLOR = "#E2E8F0"
    HOVER_COLOR = "#F1F5F9"
    LOGO_TEXT = "#1E3A8A"
    SLOGAN_TEXT = "#64748B"
    CARD_BG = "#FFFFFF"
    KPI_BG = "#FFFFFF"

# --- CSS GLOBAL CORPORATIVO (SEM ÍCONES GENÉRICOS) ---
st.markdown(
    f"""
    <style>
        html, body, [data-testid="stAppViewContainer"], .main {{
            background-color: {BG_APP} !important;
            color: {TEXT_MAIN} !important;
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
        }}

        [data-testid="stSidebar"] {{
            background-color: {BG_SIDEBAR} !important;
            border-right: 1px solid {BORDER_COLOR} !important;
        }}

        [data-testid="stSidebar"] * {{
            color: {TEXT_MUTED} !important;
        }}

        /* Links de Navegação do Sidebar Limpos */
        [data-testid="stSidebarNav"] li a {{
            border-radius: 6px !important;
            padding: 9px 14px !important;
            margin-bottom: 3px !important;
            transition: all 0.15s ease-in-out !important;
        }}

        [data-testid="stSidebarNav"] li a:hover {{
            background-color: {HOVER_COLOR} !important;
        }}

        [data-testid="stSidebarNav"] li a[aria-current="page"] {{
            background-color: {HOVER_COLOR} !important;
            border-left: 3px solid #2563EB !important;
        }}

        [data-testid="stSidebarNav"] li a[aria-current="page"] span {{
            color: {LOGO_TEXT} !important;
            font-weight: 600 !important;
        }}

        /* Cards de KPIs */
        .erp-kpi-card {{
            background-color: {KPI_BG};
            border: 1px solid {BORDER_COLOR};
            border-radius: 10px;
            padding: 18px 20px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        }}

        /* Botões de Acesso Rápido Estilo ERP (Moderne & Profissional) */
        .stButton button {{
            background-color: {CARD_BG} !important;
            color: {TEXT_MAIN} !important;
            border: 1px solid {BORDER_COLOR} !important;
            border-radius: 8px !important;
            padding: 15px 20px !important;
            text-align: left !important;
            font-size: 14px !important;
            font-weight: 600 !important;
            width: 100% !important;
            height: 65px !important;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.02) !important;
            transition: all 0.2s ease-in-out !important;
        }}

        .stButton button:hover {{
            border-color: #2563EB !important;
            background-color: {HOVER_COLOR} !important;
            transform: translateY(-1px);
        }}

        #MainMenu {{visibility: hidden;}}
        footer {{visibility: hidden;}}
        .stHeader, header[data-testid="stHeader"] {{
            background-color: transparent !important;
        }}
    </style>
""",
    unsafe_allow_html=True,
)

# --- LOGOTIPO CORPORATIVO DA EMPRESA ---
with st.sidebar:
    st.markdown(
        f"""
        <div style="text-align: center; padding: 15px 5px 12px 5px;">
            <!-- Logotipo Corporativo Minimalista e Sóbrio -->
            <div style="display: flex; align-items: center; justify-content: center; gap: 8px; margin-bottom: 6px;">
                <div style="width: 10px; height: 32px; background-color: #2563EB; border-radius: 2px;"></div>
                <h2 style="font-family: 'Segoe UI', sans-serif; font-weight: 900; font-size: 32px; color: {LOGO_TEXT}; margin: 0; letter-spacing: 3px;">NSF</h2>
            </div>
            <div style="border-top: 1px solid {BORDER_COLOR}; padding-top: 8px; margin-top: 4px;">
                <p style="color: {LOGO_TEXT} !important; font-size: 9px; font-weight: 700; margin: 0; letter-spacing: 1px; text-transform: uppercase;">
                    TECNOLOGIA EM INSTALAÇÕES COMERCIAIS
                </p>
                <p style="color: {SLOGAN_TEXT} !important; font-size: 8.5px; font-weight: 400; margin: 3px 0 0 0; font-style: italic;">
                    "Desenvolvendo o varejo, transformando vidas."
                </p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(f"<hr style='border-color: {BORDER_COLOR}; margin: 8px 0;'>", unsafe_allow_html=True)

# --- DEFINIÇÃO DOS MÓDULOS E PÁGINAS DO SISTEMA (SEM ÍCONES POLUENTES) ---
bandejas_page = st.Page("views/bandejas_pead.py", title="Bandejas ABS & PEAD")
tubos_page = st.Page("views/calculo_tubos.py", title="Cálculo de Tubos (Chassis)")
portas_page = st.Page("views/calculo_portas.py", title="Portas & Frames (VA & VAP)")
chapas_page = st.Page("views/calculo_chapas.py", title="Chapas Cristal VA (Bojo/Costas/Teto)")
inox_page = st.Page("views/calculo_inox.py", title="Chapas de Aço Inox (Expositores VA)")
gerencial_page = st.Page("views/processar_of.py", title="Relatórios Gerenciais (ORFAs)")
materiais_page = st.Page("views/materiais_padrao.py", title="Materiais Padrão (Refrigeração)")
injecao_page = st.Page("views/op_injeção.py", title="OP de Injeção de PU")

# --- TELA INICIAL (DASHBOARD EXECUTIVO) ---
def render_home():
    st.markdown(
        """
        <div style="padding: 10px 0 20px 0;">
            <h1 style="font-size: 24px; font-weight: 800; margin-bottom: 5px;">
                Portal PCP Enterprise — Painel de Controle Executivo
            </h1>
            <p style="font-size: 14px; color: #94A3B8; margin: 0;">
                Sistema corporativo de gestão de produção, planejamento de insumos e automação de engenharia.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    with kpi1:
        st.markdown(
            f"""
            <div class="erp-kpi-card">
                <span style="font-size: 11px; font-weight: 700; color: #64748B; text-transform: uppercase;">Módulos Ativos</span>
                <h3 style="margin: 5px 0 0 0; font-size: 20px; font-weight: 700; color: {TEXT_MAIN};">8 Subsistemas</h3>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with kpi2:
        st.markdown(
            f"""
            <div class="erp-kpi-card">
                <span style="font-size: 11px; font-weight: 700; color: #64748B; text-transform: uppercase;">Linhas de Fábrica</span>
                <h3 style="margin: 5px 0 0 0; font-size: 20px; font-weight: 700; color: {TEXT_MAIN};">VA, VCA, VRA & VIL</h3>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with kpi3:
        st.markdown(
            f"""
            <div class="erp-kpi-card">
                <span style="font-size: 11px; font-weight: 700; color: #64748B; text-transform: uppercase;">Insumos</span>
                <h3 style="margin: 5px 0 0 0; font-size: 20px; font-weight: 700; color: {TEXT_MAIN};">Aço, PU, PEAD & ABS</h3>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with kpi4:
        st.markdown(
            f"""
            <div class="erp-kpi-card">
                <span style="font-size: 11px; font-weight: 700; color: #64748B; text-transform: uppercase;">Ambiente</span>
                <h3 style="margin: 5px 0 0 0; font-size: 20px; font-weight: 700; color: {TEXT_MAIN};">Enterprise v4.0</h3>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        f"""
        <h3 style="font-size: 16px; font-weight: 700; color: {TEXT_MAIN}; margin-bottom: 12px;">
            Acesso Rápido aos Módulos Operacionais
        </h3>
        """,
        unsafe_allow_html=True,
    )

    col1, col2, col3 = st.columns(3)

    with col1:
        if st.button("Bandejas ABS & PEAD", key="btn_bandejas"):
            st.switch_page(bandejas_page)
        st.markdown("<div style='margin-bottom: 8px;'></div>", unsafe_allow_html=True)
        if st.button("Portas & Frames (VA & VAP)", key="btn_portas"):
            st.switch_page(portas_page)
        st.markdown("<div style='margin-bottom: 8px;'></div>", unsafe_allow_html=True)
        if st.button("Chapas Cristal VA", key="btn_chapas"):
            st.switch_page(chapas_page)

    with col2:
        if st.button("Cálculo de Tubos (Chassis)", key="btn_tubos"):
            st.switch_page(tubos_page)
        st.markdown("<div style='margin-bottom: 8px;'></div>", unsafe_allow_html=True)
        if st.button("Chapas de Aço Inox", key="btn_inox"):
            st.switch_page(inox_page)
        st.markdown("<div style='margin-bottom: 8px;'></div>", unsafe_allow_html=True)
        if st.button("Materiais Padrão", key="btn_materiais"):
            st.switch_page(materiais_page)

    with col3:
        if st.button("OP de Injeção de PU", key="btn_injecao"):
            st.switch_page(injecao_page)
        st.markdown("<div style='margin-bottom: 8px;'></div>", unsafe_allow_html=True)
        if st.button("Relatórios Gerenciais (ORFAs)", key="btn_gerencial"):
            st.switch_page(gerencial_page)

home_page = st.Page(render_home, title="Início (Portal ERP)", default=True)

# --- NAVEGAÇÃO DA PLATAFORMA ---
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
            injecao_page,
            gerencial_page,
            materiais_page,
        ]
    }
)

# Rodapé lateral discreto
with st.sidebar:
    st.markdown(f"<hr style='border-color: {BORDER_COLOR}; margin: 12px 0 8px 0;'>", unsafe_allow_html=True)
    st.markdown(
        f"""
        <div style="text-align: center; color: {SLOGAN_TEXT} !important; font-size: 9px;">
            <p style="margin: 0; font-weight: 600; color: {LOGO_TEXT} !important;">NSF Enterprise v4.0</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

# Executa o roteador
pg.run()