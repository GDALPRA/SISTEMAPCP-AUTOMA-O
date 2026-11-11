from datetime import datetime
import io
import math
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd

# ReportLab para geração do PDF de Compras
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, portrait
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    HRFlowable,
    KeepTogether,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)
import streamlit as st

# --- CABEÇALHO EXECUTIVO SLIM ---
st.markdown(
    """
    <div style="margin-bottom: 20px;">
        <h2 style="margin: 0; color: #FFFFFF; font-size: 22px; font-weight: 700; display: flex; align-items: center; gap: 10px;">
            📦 MATERIAIS PADRÃO — REFRIGERAÇÃO
        </h2>
        <p style="margin: 4px 0 0 0; color: #94A3B8; font-size: 13px;">
            Banco de Insumos, Cálculo de Chapas por Blank, Regras de Perfis Tubulares (6m/m/kg) & Solicitação de Compras
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

# --- BASE DE DADOS DOS MATERIAIS PADRÃO ---
@st.cache_data
def carregar_base_materiais():
    return [
        # CHAPAS EM INOX PARA REFRIGERAÇÃO
        {"cod": "03583", "nome": "CH 0.60 X 1240 X 2560 INOX 304 ESC PVC LASER", "tipo": "Chapas em Inox", "esp": 0.60, "larg": 1240, "comp": 2560, "dens": 7.85},
        {"cod": "104818", "nome": "CH 0.60 X 1240 X 2000 INOX 304 ESC PVC LASER", "tipo": "Chapas em Inox", "esp": 0.60, "larg": 1240, "comp": 2000, "dens": 7.85},
        {"cod": "157228", "nome": "CH 0.60 X 1240 X 3840 INOX 304 ESC PVC LASER", "tipo": "Chapas em Inox", "esp": 0.60, "larg": 1240, "comp": 3840, "dens": 7.85},
        {"cod": "157219", "nome": "CH 0.80 X 1240 X 3900 INOX 304 ESC PVC LASER", "tipo": "Chapas em Inox", "esp": 0.80, "larg": 1240, "comp": 3900, "dens": 7.85},
        {"cod": "107358", "nome": "CH 0.80 X 1240 X 2660 INOX 304 ESC PVC LASER", "tipo": "Chapas em Inox", "esp": 0.80, "larg": 1240, "comp": 2660, "dens": 7.85},
        {"cod": "153549", "nome": "CH 0.40 X 1240 X 2540 INOX 304 2B FOSCO", "tipo": "Chapas em Inox", "esp": 0.40, "larg": 1240, "comp": 2540, "dens": 7.85},
        {"cod": "156082", "nome": "CH 0.40 X 1240 X 3780 INOX 304 2B FOSCO", "tipo": "Chapas em Inox", "esp": 0.40, "larg": 1240, "comp": 3780, "dens": 7.85},
        {"cod": "157101", "nome": "CH 0.40 X 1040 X 2060 INOX 430 PVC LASER", "tipo": "Chapas em Inox", "esp": 0.40, "larg": 1040, "comp": 2060, "dens": 7.85},
        {"cod": "151947", "nome": "CH 0.40 X 1240 X 2550 INOX 430 PVC LASER", "tipo": "Chapas em Inox", "esp": 0.40, "larg": 1240, "comp": 2550, "dens": 7.85},
        {"cod": "157226", "nome": "CH 0.40 X 1240 X 3780 INOX 430 PVC LASER", "tipo": "Chapas em Inox", "esp": 0.40, "larg": 1240, "comp": 3780, "dens": 7.85},
        {"cod": "150723", "nome": "CH 1.00 X 1250 X 3780 INOX 304 ESC PVC LASER", "tipo": "Chapas em Inox", "esp": 1.00, "larg": 1250, "comp": 3780, "dens": 7.85},
        {"cod": "103938", "nome": "CH 2.00 X 1040 X 2640 INOX 304 2B FOSCO", "tipo": "Chapas em Inox", "esp": 2.00, "larg": 1040, "comp": 2640, "dens": 7.85},
        {"cod": "156102", "nome": "CH 1.20 X 1240 X 2960 INOX 304 ESC PVC LASER", "tipo": "Chapas em Inox", "esp": 1.20, "larg": 1240, "comp": 2960, "dens": 7.85},

        # CHAPAS ZN MN PARA REFRIGERAÇÃO
        {"cod": "X1", "nome": "CH 1.20 X 1200 X 3780 NBR 7008 Z100 ZN MN", "tipo": "Chapas Zn Mn", "esp": 1.20, "larg": 1200, "comp": 3780, "dens": 7.85},
        {"cod": "X2", "nome": "CH 1.20 X 1200 X 2550 NBR 7008 Z100 ZN MN", "tipo": "Chapas Zn Mn", "esp": 1.20, "larg": 1200, "comp": 2550, "dens": 7.85},
        {"cod": "108691", "nome": "CH 0.65 X 1200 X 1950 NBR 7008 Z100 ZN MN", "tipo": "Chapas Zn Mn", "esp": 0.65, "larg": 1200, "comp": 1950, "dens": 7.85},
        {"cod": "101652", "nome": "CH 0.65 X 1200 X 2550 NBR 7008 Z100 ZN MN", "tipo": "Chapas Zn Mn", "esp": 0.65, "larg": 1200, "comp": 2550, "dens": 7.85},
        {"cod": "108688", "nome": "CH 0.65 X 1200 X 2220 NBR 7008 Z100 ZN MN", "tipo": "Chapas Zn Mn", "esp": 0.65, "larg": 1200, "comp": 2220, "dens": 7.85},
        {"cod": "108690", "nome": "CH 0.65 X 1200 X 2060 NBR 7008 Z100 ZN MN", "tipo": "Chapas Zn Mn", "esp": 0.65, "larg": 1200, "comp": 2060, "dens": 7.85},
        {"cod": "108689", "nome": "CH 0.65 X 1200 X 2640 NBR 7008 Z100 ZN MN", "tipo": "Chapas Zn Mn", "esp": 0.65, "larg": 1200, "comp": 2640, "dens": 7.85},
        {"cod": "108749", "nome": "CH 0.95 X 1200 X 2300 NBR 7008 ZC Z100 MINS/OL", "tipo": "Chapas Zn Mn", "esp": 0.95, "larg": 1200, "comp": 2300, "dens": 7.85},
        {"cod": "119723", "nome": "CH 0.95 X 1200 X 2550 NBR 7008 ZC Z100 MINS/OL", "tipo": "Chapas Zn Mn", "esp": 0.95, "larg": 1200, "comp": 2550, "dens": 7.85},
        {"cod": "153962", "nome": "CH 0.95 X 1200 X 3780 NBR 7008 ZC Z100 MINS/OL", "tipo": "Chapas Zn Mn", "esp": 0.95, "larg": 1200, "comp": 3780, "dens": 7.85},
        {"cod": "149233", "nome": "CH 1.80 X 1200 X 1700 GALV MINIMIZ Z100 ZAR 280 GI", "tipo": "Chapas Zn Mn", "esp": 1.80, "larg": 1200, "comp": 1700, "dens": 7.85},
        {"cod": "156965", "nome": "CH 1.95 X 1200 X 1950 GALV MINI Z100 ZAR 280 GI", "tipo": "Chapas Zn Mn", "esp": 1.95, "larg": 1200, "comp": 1950, "dens": 7.85},
        {"cod": "54357", "nome": "CH 1.95 X 1200 X 2100 GALV Z100 ZAR 280 GI", "tipo": "Chapas Zn Mn", "esp": 1.95, "larg": 1200, "comp": 2100, "dens": 7.85},
        {"cod": "136309", "nome": "CH 1.95 X 1200 X 2550 GALV Z100 ZAR 280 GI", "tipo": "Chapas Zn Mn", "esp": 1.95, "larg": 1200, "comp": 2550, "dens": 7.85},

        # CHAPAS PRÉ PINTADA
        {"cod": "153381", "nome": "CH 0.43 X 1200 X 4000 NBR 7008 ZC Z275 RAL 9003", "tipo": "Chapas Pré Pintada", "esp": 0.43, "larg": 1200, "comp": 4000, "dens": 7.85},
        {"cod": "153210", "nome": "CH 0.43 X 1200 X 3780 NBR 7008 ZC Z275 RAL 9003", "tipo": "Chapas Pré Pintada", "esp": 0.43, "larg": 1200, "comp": 3780, "dens": 7.85},
        {"cod": "153400", "nome": "CH 0.43 X 1200 X 2550 NBR 7008 ZC Z275 RAL 9003", "tipo": "Chapas Pré Pintada", "esp": 0.43, "larg": 1200, "comp": 2550, "dens": 7.85},

        # CHAPAS CN ZN
        {"cod": "153401", "nome": "CH 0.43 X 1200 X 2550 NBR 7008 ZC Z275 NL S/OL", "tipo": "Chapas Cn Zn", "esp": 0.43, "larg": 1200, "comp": 2550, "dens": 7.85},
        {"cod": "153209", "nome": "CH 0.43 X 1200 X 3780 NBR 7008 ZC Z275 NL S/OL", "tipo": "Chapas Cn Zn", "esp": 0.43, "larg": 1200, "comp": 3780, "dens": 7.85},
        {"cod": "153380", "nome": "CH 0.43 X 1200 X 4000 NBR 7008 ZC Z275 NL S/OL", "tipo": "Chapas Cn Zn", "esp": 0.43, "larg": 1200, "comp": 4000, "dens": 7.85},

        # PERFIS DE INOX
        {"cod": "158212", "nome": "TUBO INOX 304 NORM A 554 25X25 X 1.00 X 6000 POL", "tipo": "Perfis de Inox", "esp": 1.00, "larg": 25, "comp": 6000, "dens": 7.85, "dim2": 25, "formato": "quadrado"},
        {"cod": "158217", "nome": "TUBO INOX 304 NORM A 554 15 X 15 X 1,00 X 6000 POL", "tipo": "Perfis de Inox", "esp": 1.00, "larg": 15, "comp": 6000, "dens": 7.85, "dim2": 15, "formato": "quadrado"},
        {"cod": "174028", "nome": "TUBO INOX 304 NORM A 554 20 X 20 X 1,20 X 6000 POL", "tipo": "Perfis de Inox", "esp": 1.20, "larg": 20, "comp": 6000, "dens": 7.85, "dim2": 20, "formato": "quadrado"},
        {"cod": "158214", "nome": "TUBO INOX 304 NORM A 554 20 X 10 X 1,00 X 6000 POL", "tipo": "Perfis de Inox", "esp": 1.00, "larg": 20, "comp": 6000, "dens": 7.85, "dim2": 10, "formato": "retangular"},
        {"cod": "158216", "nome": "TUBO INOX 304 NORM A 554 40 X 20 X 1,00 X 6000 POL", "tipo": "Perfis de Inox", "esp": 1.00, "larg": 40, "comp": 6000, "dens": 7.85, "dim2": 20, "formato": "retangular"},
        {"cod": "158215", "nome": "TUBO INOX 304 NORM A 554 30 X 30 X 1,00 X 6000 POL", "tipo": "Perfis de Inox", "esp": 1.00, "larg": 30, "comp": 6000, "dens": 7.85, "dim2": 30, "formato": "quadrado"},
        {"cod": "82856", "nome": "TUBO INOX REDONDO 304 NORM A 554 1 1/4 POL", "tipo": "Perfis de Inox", "esp": 1.20, "larg": 31.75, "comp": 6000, "dens": 7.85, "formato": "redondo"},
        {"cod": "164009", "nome": "TUBO INOX REDONDO 304 NORM A 554 1 POL", "tipo": "Perfis de Inox", "esp": 1.20, "larg": 25.4, "comp": 6000, "dens": 7.85, "formato": "redondo"},
        {"cod": "74202", "nome": "TUBO INOX REDONDO 304 NORM A 554 2 POL", "tipo": "Perfis de Inox", "esp": 1.50, "larg": 50.8, "comp": 6000, "dens": 7.85, "formato": "redondo"},
        
        # PERFIS MINIMIZADOS / CARBONO
        {"cod": "165413", "nome": "TUBO CARB RET NBR 6591 25X25X1,11 X 6000 MINI GALV", "tipo": "Perfis Minimizados", "esp": 1.11, "larg": 25, "comp": 6000, "dens": 7.85, "dim2": 25, "formato": "quadrado"},
    ]

materiais = carregar_base_materiais()

# Função para calcular o peso por metro (kg/m) de qualquer perfil tubular (quadrado, retangular ou redondo)
def calcular_peso_metro(mat):
    esp_mm = mat["esp"]
    dens = mat["dens"]
    formato = mat.get("formato", "quadrado")
    
    if formato in ["quadrado", "retangular"]:
        d1_m = mat["larg"] / 1000.0
        d2_m = mat.get("dim2", mat["larg"]) / 1000.0
        esp_m = esp_mm / 1000.0
        # Perímetro da seção * espessura * 1 metro * densidade
        perimetro = 2 * (d1_m + d2_m)
        return round(perimetro * esp_m * 1.0 * dens, 3)
    elif formato == "redondo":
        # Diâmetro externo em metros
        diam_m = mat["larg"] / 1000.0
        esp_m = esp_mm / 1000.0
        # Diâmetro interno
        diam_int_m = diam_m - (2 * esp_m)
        # Área da seção transversal = pi/4 * (D^2 - d^2)
        import math as m_geom
        area = (m_geom.pi / 4.0) * ((diam_m ** 2) - (diam_int_m ** 2))
        # Peso por metro = Área (m²) * 1m * Densidade (kg/m³)
        return round(area * 1.0 * dens, 3)
    else:
        return 0.5

for m in materiais:
    m["is_tubular"] = "TUBO" in m["nome"].upper() or m["tipo"] in ["Perfis de Inox", "Perfis Minimizados"]
    if m["is_tubular"]:
        m["peso_metro"] = calcular_peso_metro(m)
        # Cada peça é uma barra padrão de 6 metros
        m["peso_unit"] = round(m["peso_metro"] * 6.0, 2)
    else:
        # Chapa plana
        esp = m["esp"]
        larg_m = m["larg"] / 1000.0
        comp_m = m["comp"] / 1000.0
        m["peso_unit"] = round(esp * larg_m * comp_m * m["dens"], 2)
        m["peso_metro"] = 0.0

# --- FILTROS E BARRA DE PESQUISA SUPERIOR ---
col_filtro1, col_filtro2 = st.columns([2, 3])

with col_filtro1:
    categorias = ["Todas"] + list(set([m["tipo"] for m in materiais]))
    cat_selecionada = st.selectbox("Filtrar por Categoria de Insumo:", categorias)

with col_filtro2:
    busca_termo = st.text_input("🔍 Pesquisar por Código ou Descrição:", placeholder="Ex: 158212, Inox 304, Redondo...")

# Filtragem combinada
materiais_filtrados = [
    m for m in materiais 
    if (cat_selecionada == "Todas" or m["tipo"] == cat_selecionada) and
       (not busca_termo or busca_termo.lower() in m["cod"].lower() or busca_termo.lower() in m["nome"].lower())
]

st.markdown("---")
st.markdown("<p style='color: #94A3B8; font-size: 13px;'>Para <b>Perfis Tubulares</b>: <b>Pcs</b> considera cada peça uma barra de 6m; <b>Metros (m)</b> calcula pelo comprimento exato informado; <b>Kg</b> insere o peso direto sem conversões.</p>", unsafe_allow_html=True)

solicitacao_itens = []

for idx, mat in enumerate(materiais_filtrados):
    cols = st.columns([1.2, 4.2, 1.8, 2.0, 2.2])
    
    with cols[0]:
        st.markdown(f"<p style='color: #38BDF8; font-weight:600; font-size:12px; margin-top:8px;'>#{mat['cod']}</p>", unsafe_allow_html=True)
    with cols[1]:
        st.markdown(f"<p style='color: #E2E8F0; font-size:12px; margin-top:8px;'>{mat['nome']}</p>", unsafe_allow_html=True)
    with cols[2]:
        if mat['is_tubular']:
            peso_exib = f"{mat['peso_metro']} kg/m"
        else:
            peso_exib = f"{mat['peso_unit']} kg/un"
        st.markdown(f"<p style='color: #94A3B8; font-size:12px; margin-top:8px;'>{peso_exib}</p>", unsafe_allow_html=True)
    with cols[3]:
        opcoes_modo = ["Pcs", "Metros (m)", "Kg"] if mat["is_tubular"] else ["Pcs", "Kg"]
        modo = st.selectbox("Modo", opcoes_modo, key=f"modo_{mat['cod']}_{idx}", label_visibility="collapsed")
    with cols[4]:
        val = st.number_input("Qtd", min_value=0.0, step=1.0, key=f"val_{mat['cod']}_{idx}", label_visibility="collapsed")
        
    if val > 0:
        if mat['is_tubular']:
            if modo == "Pcs":
                qtd_pecas = int(val)
                # Cada peça é uma barra de 6 metros
                peso_total = round(qtd_pecas * mat['peso_metro'] * 6.0, 2)
                detalhe_qtd = f"{qtd_pecas} Pcs ({qtd_pecas * 6} metros em barras de 6m)"
            elif modo == "Metros (m)":
                metros_sel = round(val, 2)
                peso_total = round(metros_sel * mat['peso_metro'], 2)
                barras_aprox = math.ceil(metros_sel / 6.0)
                detalhe_qtd = f"{metros_sel} m (~{barras_aprox} barras de 6m)"
            else:  # Kg direto
                peso_total = round(val, 2)
                detalhe_qtd = f"{peso_total} Kg (Informado Direto)"
        else:
            # Chapas planas
            if modo == "Pcs":
                qtd_pecas = int(val)
                peso_total = round(qtd_pecas * mat['peso_unit'], 2)
                detalhe_qtd = f"{qtd_pecas} Pcs"
            else:  # Kg direto
                peso_total = round(val, 2)
                detalhe_qtd = f"{peso_total} Kg"
            
        solicitacao_itens.append({
            "cod": mat["cod"],
            "nome": mat["nome"],
            "tipo": mat["tipo"],
            "detalhe_qtd": detalhe_qtd,
            "peso_total": peso_total
        })

st.markdown("---")

# --- RESUMO E EXPORTAÇÃO PARA COMPRAS ---
if solicitacao_itens:
    st.markdown(f"### 📋 Resumo da Solicitação para Compras ({len(solicitacao_itens)} itens selecionados)")
    
    df_resumo = pd.DataFrame(solicitacao_itens)
    df_exib = df_resumo[["cod", "nome", "tipo", "detalhe_qtd", "peso_total"]]
    df_exib.columns = ["Código", "Descrição do Material", "Categoria", "Quantidade Solicitada", "Peso Total (Kg)"]
    st.dataframe(df_exib, use_container_width=True, hide_index=True)
    
    peso_geral = df_resumo["peso_total"].sum()
    st.metric(label="Peso Total Consolidado da Solicitação", value=f"{peso_geral:,.2f} kg")
    
    st.write("")
    
    # --- GERADOR DE PDF PARA COMPRAS ---
    def gerar_pdf_compras(itens, peso_total_geral):
        pdf_out = io.BytesIO()
        doc = SimpleDocTemplate(pdf_out, pagesize=portrait(A4), leftMargin=1.5*cm, rightMargin=1.5*cm, topMargin=1.5*cm, bottomMargin=1.5*cm)
        styles = getSampleStyleSheet()
        
        style_title = ParagraphStyle('T', fontName='Helvetica-Bold', fontSize=14, textColor=colors.white, alignment=1)
        style_sub = ParagraphStyle('S', fontName='Helvetica-Bold', fontSize=10, textColor=colors.HexColor("#0052CC"))
        style_h = ParagraphStyle('H', fontName='Helvetica-Bold', fontSize=8, textColor=colors.white, alignment=1)
        style_c = ParagraphStyle('C', fontName='Helvetica', fontSize=8, leading=10)
        style_cc = ParagraphStyle('CC', fontName='Helvetica', fontSize=8, leading=10, alignment=1)
        
        story = []
        
        t_top = Table([[Paragraph("SOLICITAÇÃO DE COMPRAS DE MATERIAIS — REFRIGERAÇÃO", style_title)]], colWidths=[18*cm])
        t_top.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#0A2540")),
            ('TOPPADDING', (0,0), (-1,-1), 10),
            ('BOTTOMPADDING', (0,0), (-1,-1), 10),
        ]))
        story.append(t_top)
        story.append(Spacer(1, 10))
        
        data_atual = datetime.now().strftime("%d/%m/%Y %H:%M")
        info_texto = f"<b>Data de Emissão:</b> {data_atual} &nbsp;&nbsp;|&nbsp;&nbsp; <b>Destino:</b> Departamento de Suprimentos & Compras"
        t_info = Table([[Paragraph(info_texto, style_sub)]], colWidths=[18*cm])
        t_info.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#F1F5F9")),
            ('TOPPADDING', (0,0), (-1,-1), 6),
            ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ]))
        story.append(t_info)
        story.append(Spacer(1, 10))
        
        rows = [[
            Paragraph("Cód.", style_h),
            Paragraph("Descrição do Material", style_h),
            Paragraph("Qtd Solicitada", style_h),
            Paragraph("Peso Total", style_h)
        ]]
        
        for item in itens:
            rows.append([
                Paragraph(str(item["cod"]), style_cc),
                Paragraph(str(item["nome"]), style_c),
                Paragraph(str(item["detalhe_qtd"]), style_cc),
                Paragraph(f"{item['peso_total']} kg", style_cc),
            ])
            
        t_table = Table(rows, colWidths=[2*cm, 9.5*cm, 3.5*cm, 3*cm])
        t_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#0052CC")),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor("#F8FAFC")]),
            ('TOPPADDING', (0,0), (-1,-1), 5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ]))
        story.append(t_table)
        story.append(Spacer(1, 15))
        
        t_total = Table([[Paragraph(f"<b>PESO TOTAL DA SOLICITAÇÃO: {peso_total_geral:,.2f} KG</b>", ParagraphStyle('Tot', fontName='Helvetica-Bold', fontSize=10, textColor=colors.HexColor("#0A2540"), alignment=2))]], colWidths=[18*cm])
        t_total.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#E2E8F0")),
            ('TOPPADDING', (0,0), (-1,-1), 8),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ]))
        story.append(t_total)
        
        doc.build(story)
        pdf_out.seek(0)
        return pdf_out.getvalue()

    st.write("")
    if st.button("📥 Gerar PDF de Solicitação para Compras", type="primary"):
        pdf_bytes = gerar_pdf_compras(solicitacao_itens, peso_geral)
        st.download_button(
            label="💾 Baixar PDF de Solicitação (.pdf)",
            data=pdf_bytes,
            file_name=f"Solicitacao_Compras_Refrigeracao_{datetime.now().strftime('%Y%m%d')}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
else:
    st.info("ℹ️ Pesquise ou selecione quantidades em pelo menos um dos materiais acima para habilitar o resumo e a exportação para compras.")