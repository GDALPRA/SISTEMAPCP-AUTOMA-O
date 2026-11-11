from datetime import datetime
import io
import re
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
            Banco de Insumos, Cálculo de Pesos por Blank & Solicitação de Compras
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

# --- BASE DE DADOS DOS MATERIAIS PADRÃO (EXTRAÍDA DO CATÁLOGO) ---
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
        {"cod": "158212", "nome": "TUBO INOX 304 NORM A 554 25X25 X 1.00 X 6000 POL", "tipo": "Perfis de Inox", "esp": 1.00, "larg": 25, "comp": 6000, "dens": 7.85},
        {"cod": "158217", "nome": "TUBO INOX 304 NORM A 554 15 X 15 X 1,00 X 6000 POL", "tipo": "Perfis de Inox", "esp": 1.00, "larg": 15, "comp": 6000, "dens": 7.85},
        {"cod": "174028", "nome": "TUBO INOX 304 NORM A 554 20 X 20 X 1,20 X 6000 POL", "tipo": "Perfis de Inox", "esp": 1.20, "larg": 20, "comp": 6000, "dens": 7.85},
        {"cod": "158214", "nome": "TUBO INOX 304 NORM A 554 20 X 10 X 1,00 X 6000 POL", "tipo": "Perfis de Inox", "esp": 1.00, "larg": 20, "comp": 6000, "dens": 7.85},
        {"cod": "158216", "nome": "TUBO INOX 304 NORM A 554 40 X 20 X 1,00 X 6000 POL", "tipo": "Perfis de Inox", "esp": 1.00, "larg": 40, "comp": 6000, "dens": 7.85},
        {"cod": "158215", "nome": "TUBO INOX 304 NORM A 554 30 X 30 X 1,00 X 6000 POL", "tipo": "Perfis de Inox", "esp": 1.00, "larg": 30, "comp": 6000, "dens": 7.85},
    ]

materiais = carregar_base_materiais()

# Função matemática para calcular o peso unitário da chapa/tubo com base no blank
def calcular_peso_unitario(mat):
    esp = mat["esp"]
    larg_m = mat["larg"] / 1000.0
    comp_m = mat["comp"] / 1000.0
    dens = mat["dens"]
    
    if "TUBO" in mat["nome"].upper() or mat["tipo"] == "Perfis de Inox":
        # Estimativa de perfil linear em barras de 6m
        peso_metro = esp * larg_m * dens * 0.05 # aproximação para tubos estruturais
        return round(peso_metro * (mat["comp"] / 1000.0), 2)
    else:
        # Fórmula padrão de chapa: Espessura (mm) x Largura (m) x Comprimento (m) x Densidade (7.85)
        peso = esp * larg_m * comp_m * dens
        return round(peso, 2)

# Adiciona o peso unitário calculado a cada item
for m in materiais:
    m["peso_unit"] = calcular_peso_unitario(m)

# --- INTERFACE DE SELEÇÃO E QUANTIFICAÇÃO ---
st.markdown("<p style='color: #94A3B8; font-size: 13px;'>Selecione os materiais abaixo, defina a quantidade por <b>Peças (Unidades)</b> ou diretamente em <b>Quilogramas (Kg)</b> e gere a solicitação de compras.</p>", unsafe_allow_html=True)

# Filtro por categoria
categorias = ["Todas"] + list(set([m["tipo"] for m in materiais]))
cat_selecionada = st.selectbox("Filtrar por Categoria de Insumo:", categorias)

materiais_filtrados = [m for m in materiais if cat_selecionada == "Todas" or m["tipo"] == cat_selecionada]

# Tabela interativa de seleção
st.markdown("---")
solicitacao_itens = []

for idx, mat in enumerate(materiais_filtrados):
    cols = st.columns([1.2, 4.5, 1.8, 1.8, 2.2])
    
    with cols[0]:
        st.markdown(f"<p style='color: #38BDF8; font-weight:600; font-size:12px; margin-top:8px;'>#{mat['cod']}</p>", unsafe_allow_html=True)
    with cols[1]:
        st.markdown(f"<p style='color: #E2E8F0; font-size:12px; margin-top:8px;'>{mat['nome']}</p>", unsafe_allow_html=True)
    with cols[2]:
        st.markdown(f"<p style='color: #94A3B8; font-size:12px; margin-top:8px;'>{mat['peso_unit']} kg/un</p>", unsafe_allow_html=True)
    with cols[3]:
        # Opção de selecionar modo de entrada: Peças ou Kgs
        modo = st.selectbox("Modo", ["Pcs", "Kg"], key=f"modo_{mat['cod']}_{idx}", label_visibility="collapsed")
    with cols[4]:
        val = st.number_input("Qtd", min_value=0.0, step=1.0, key=f"val_{mat['cod']}_{idx}", label_visibility="collapsed")
        
    if val > 0:
        if modo == "Pcs":
            qtd_pecas = int(val)
            peso_total = round(qtd_pecas * mat['peso_unit'], 2)
        else:
            peso_total = round(val, 2)
            qtd_pecas = max(1, int(peso_total / mat['peso_unit'])) if mat['peso_unit'] > 0 else 1
            
        solicitacao_itens.append({
            "cod": mat["cod"],
            "nome": mat["nome"],
            "tipo": mat["tipo"],
            "peso_unit": mat["peso_unit"],
            "qtd_pecas": qtd_pecas,
            "peso_total": peso_total
        })

st.markdown("---")

# --- RESUMO E EXPORTAÇÃO PARA COMPRAS ---
if solicitacao_itens:
    st.markdown(f"### 📋 Resumo da Solicitação para Compras ({len(solicitacao_itens)} itens selecionados)")
    
    df_resumo = pd.DataFrame(solicitacao_itens)
    df_exib = df_resumo[["cod", "nome", "tipo", "qtd_pecas", "peso_total"]]
    df_exib.columns = ["Código", "Descrição do Material", "Categoria", "Qtd (Pcs)", "Peso Total (Kg)"]
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
        
        # Cabeçalho Executivo do PDF
        t_top = Table([[Paragraph("SOLICITAÇÃO DE COMPRAS DE MATERIAIS — REFRIGERAÇÃO", style_title)]], colWidths=[18*cm])
        t_top.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#0A2540")),
            ('TOPPADDING', (0,0), (-1,-1), 10),
            ('BOTTOMPADDING', (0,0), (-1,-1), 10),
        ]))
        story.append(t_top)
        story.append(Spacer(1, 10))
        
        # Informações de Data e Emissor
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
        
        # Tabela de Itens
        rows = [[
            Paragraph("Cód.", style_h),
            Paragraph("Descrição do Material (Blank / Insumo)", style_h),
            Paragraph("Qtd (Pcs)", style_h),
            Paragraph("Peso Unit.", style_h),
            Paragraph("Peso Total", style_h)
        ]]
        
        for item in itens:
            rows.append([
                Paragraph(str(item["cod"]), style_cc),
                Paragraph(str(item["nome"]), style_c),
                Paragraph(str(item["qtd_pecas"]), style_cc),
                Paragraph(f"{item['peso_unit']} kg", style_cc),
                Paragraph(f"{item['peso_total']} kg", style_cc),
            ])
            
        t_table = Table(rows, colWidths=[2*cm, 8.5*cm, 2*cm, 2.5*cm, 3*cm])
        t_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#0052CC")),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor("#F8FAFC")]),
            ('TOPPADDING', (0,0), (-1,-1), 5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ]))
        story.append(t_table)
        story.append(Spacer(1, 15))
        
        # Totalizador
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
    st.info("ℹ️ Insira quantidades (por Peças ou Kg) em pelo menos um dos materiais acima para habilitar o resumo e a exportação para compras.")