from datetime import datetime
import io
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd

# ReportLab para geração do PDF Executivo
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
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

# --- CONFIGURAÇÃO E CABEÇALHO DA INTERFACE ---
st.markdown(
    """
    <div style="margin-bottom: 25px;">
        <h2 style="margin: 0; color: #FFFFFF; font-size: 24px; font-weight: 700; display: flex; align-items: center; gap: 10px;">
            🚪 RELATORIO GERENCIAL OFs PLANEJADAS
        </h2>
        <p style="margin: 6px 0 0 0; color: #94A3B8; font-size: 14px;">
            Painel de Engenharia & PCP — Relatório Gerencial Executivo com Dashboard de Consumo
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.write("")

# --- FUNÇÃO AUXILIAR PARA IDENTIFICAR FAMÍLIA E DIMENSÃO DO EXPOSITOR ---
def extrair_familia_e_dimensao(modelo: str):
    m = modelo.strip().upper()
    
    nums = re.findall(r"\d{3,5}", m)
    dimensao = nums[0] if nums else "GERAL"

    if m.startswith("VAP"):
        familia = "VAP (CONGELADOS)"
    elif m.startswith("VRA"):
        familia = "VRA"
    elif m.startswith("VCA"):
        familia = "VCA"
    elif m.startswith("VIL"):
        familia = "VIL"
    elif m.startswith("VA"):
        familia = "VA (RESFRIADOS)"
    elif m.startswith("VR"):
        familia = "VR"
    elif m.startswith("IR"):
        familia = "IR"
    else:
        match = re.match(r"^([A-Z]+)", m)
        familia = match.group(1) if match else "OUTROS"

    return familia, dimensao

# --- ÁREA DE UPLOAD ---
uploaded_files = st.file_uploader(
    "Arraste ou selecione os arquivos PDF das ORFAs de origem (.pdf)",
    type=["pdf"],
    accept_multiple_files=True,
    key="pdf_uploader_orfas",
)

if uploaded_files:
    st.info(f"📁 {len(uploaded_files)} arquivo(s) PDF carregado(s) com sucesso.")

    if st.button("🚀 Processar ORFAs e Gerar Relatório Consolidado", type="primary"):
        with st.spinner("Analisando e isolando dados de cada ORFA..."):
            lotes_orfas = []

            for uploaded_file in uploaded_files:
                try:
                    import pdfplumber

                    with pdfplumber.open(uploaded_file) as pdf:
                        paginas_texto = [page.extract_text() or "" for page in pdf.pages]
                        texto_completo = "\n".join(paginas_texto)
                        linhas_todas = [l.strip() for p in paginas_texto for l in p.split("\n") if l.strip()]

                    num_orfa = ""
                    entrega = ""
                    pedido = ""
                    cliente = "NÃO INFORMADO"

                    for idx, linha in enumerate(linhas_todas[:35]):
                        linha_upper = linha.upper()
                        
                        if "ORFA" in linha_upper or "N°" in linha or "Nº" in linha:
                            m_orfa_direto = re.search(r"(\d{2,4}\s*/\s*\d{4}\s*-?\s*\([A-Za-z0-9\-/]+\)|\d{2,4}\s*/\s*\d{4})", linha)
                            if m_orfa_direto:
                                num_orfa = m_orfa_direto.group(1).strip()
                            elif idx + 1 < len(linhas_todas):
                                prox = linhas_todas[idx + 1]
                                m_orfa_prox = re.search(r"(\d{2,4}\s*/\s*\d{4}\s*-?\s*\([A-Za-z0-9\-/]+\)|\d{2,4}\s*/\s*\d{4})", prox)
                                if m_orfa_prox:
                                    num_orfa = m_orfa_prox.group(1).strip()

                        if "DATA ENTREGA" in linha_upper or ("ENTREGA" in linha_upper and idx < 15):
                            m_val = re.search(r"(\d{2}/\d{2}/\d{4})", linha)
                            if m_val:
                                entrega = m_val.group(1).strip()
                            elif idx + 1 < len(linhas_todas):
                                prox = linhas_todas[idx + 1]
                                m_val2 = re.search(r"(\d{2}/\d{2}/\d{4})", prox)
                                if m_val2:
                                    entrega = m_val2.group(1).strip()

                        if "PEDIDO" in linha_upper and "REPRES" not in linha_upper:
                            m_val = re.search(r"(P\d{8,10}[A-Za-z0-9\-]*)", linha)
                            if m_val:
                                pedido = m_val.group(1).strip()
                            elif idx + 1 < len(linhas_todas):
                                prox = linhas_todas[idx + 1]
                                m_val2 = re.search(r"(P\d{8,10}[A-Za-z0-9\-]*)", prox)
                                if m_val2:
                                    pedido = m_val2.group(1).strip()

                        if "CLIENTE:" in linha_upper:
                            m_cli = re.search(r"CLIENTE[:\s]*(.+)", linha, re.IGNORECASE)
                            if m_cli and len(m_cli.group(1).strip()) > 2:
                                cliente = m_cli.group(1).strip()
                            elif idx + 1 < len(linhas_todas):
                                cliente = linhas_todas[idx + 1].strip()

                    if not num_orfa:
                        orfa_fb = re.search(r"(\d{2,4}\s*/\s*\d{4}\s*-?\s*\([A-Za-z0-9\-/]+\))", texto_completo)
                        num_orfa = orfa_fb.group(1).strip() if orfa_fb else uploaded_file.name.replace(".pdf", "")
                    
                    if not entrega:
                        dt_fb = re.search(r"(\d{2}/\d{2}/\d{4})", texto_completo)
                        entrega = dt_fb.group(1).strip() if dt_fb else "-"

                    if not pedido:
                        ped_fb = re.search(r"(P\d{8,10}[A-Za-z0-9\-]*)", texto_completo)
                        pedido = ped_fb.group(1).strip() if ped_fb else "N/A"

                    itens_orfa = []
                    todas_iluminacoes = re.findall(r"(3000K|4000K)", texto_completo, re.IGNORECASE)
                    ilum_idx = 0

                    for pagina in paginas_texto:
                        if "Resumo de Balcões por Endereço" in pagina:
                            continue

                        linhas = pagina.split("\n")
                        for i, linha in enumerate(linhas):
                            linha_clean = linha.strip()
                            linha_upper = linha_clean.upper()
                            
                            if any(termo in linha_upper for termo in ["DIVISORI", "CANTONEIRA", "ESCANINHO"]):
                                continue

                            m_item = re.search(r"^(\d{2})\s+([A-Z0-9\s\-\/+]+?)(?:\s*\(([\d\.\-\s]+)\)|$)", linha_clean)
                            
                            if m_item:
                                q_str, modelo_raw, end_opt = m_item.groups()
                                modelo_limpo = re.sub(r"\s*\([\d\.\-\s]+\)", "", modelo_raw).strip()
                                modelo_upper = modelo_limpo.upper()
                                
                                if any(termo in modelo_upper for termo in ["DIVISORI", "CANTONEIRA"]):
                                    continue

                                endereco = "1.1"
                                if end_opt and re.search(r"^\d+[\d\.\-\s]*$", end_opt.strip()):
                                    endereco = end_opt.strip()
                                else:
                                    end_match = re.search(r"\(([\d\.\-]+)\)", linha_clean)
                                    if end_match:
                                        endereco = end_match.group(1).strip()

                                bloco = "\n".join(linhas[max(0, i-6):min(len(linhas), i+12)])
                                acb_tt = re.search(r"Acb\.\s*TT/PNL:\s*([^\n|]+)", bloco, re.IGNORECASE)
                                caren = re.search(r"Caren\.\s*Frontal:\s*([^\n|]+)", bloco, re.IGNORECASE)
                                acol_ext = re.search(r"Acb\.\s*Externo:\s*([^\n|]+)", bloco, re.IGNORECASE)
                                gas = re.search(r"Fluido/Gás:\s*([^\n|]+)", bloco, re.IGNORECASE)
                                
                                val_ilu = "-"
                                if ilum_idx < len(todas_iluminacoes):
                                    val_ilu = todas_iluminacoes[ilum_idx].upper()
                                    ilum_idx += 1

                                val_tt = acb_tt.group(1).strip() if acb_tt else "-"
                                val_car = caren.group(1).strip() if caren else "-"
                                val_ext = acol_ext.group(1).strip() if acol_ext else "-"
                                val_gas = gas.group(1).strip() if gas else "-"

                                val_tt = re.split(r"Ach\.|Acb\.|Bojo", val_tt)[0].strip()
                                val_car = re.split(r"Porta|Serigrafia", val_car)[0].strip()
                                val_ext = re.split(r"Fluido|Gás|Porta", val_ext)[0].strip()

                                itens_orfa.append({
                                    "qtd": q_str.strip(),
                                    "modelo": modelo_limpo,
                                    "end": endereco,
                                    "acb_tt_pnl": val_tt if val_tt else "-",
                                    "carenagem": val_car if val_car else "-",
                                    "acb_ext": val_ext if val_ext else "-",
                                    "gas": val_gas if val_gas else "-",
                                    "ilum": val_ilu
                                })

                    lotes_orfas.append({
                        "orfa": num_orfa,
                        "pedido": pedido,
                        "entrega": entrega,
                        "cliente": cliente,
                        "itens": itens_orfa
                    })

                except Exception as e:
                    st.error(f"⚠️ Erro ao processar {uploaded_file.name}: {e}")

            if lotes_orfas:
                st.success(f"🎉 Processamento concluído! Total de ORFAs no lote: {len(lotes_orfas)}")

                # --- PROCESSAMENTO DOS DADOS PARA O RESUMO E DASHBOARD DE CONSUMO ---
                registros_cores = []
                for lote in lotes_orfas:
                    for item in lote['itens']:
                        try:
                            qtd_num = int(item["qtd"])
                        except ValueError:
                            qtd_num = 1
                        
                        familia, dimensao = extrair_familia_e_dimensao(item["modelo"])
                        
                        acb_tt = item["acb_tt_pnl"].strip().upper() if item["acb_tt_pnl"].strip() != "-" else "N/A"
                        caren_frontal = item["carenagem"].strip().upper() if item["carenagem"].strip() != "-" else "N/A"
                        acb_externo = item["acb_ext"].strip().upper() if item["acb_ext"].strip() != "-" else "N/A"

                        registros_cores.append({
                            "Família": familia,
                            "Dimensão": dimensao,
                            "ACB. TT/PNL": acb_tt,
                            "CAREN. FRONTAL": caren_frontal,
                            "ACB. EXTERNO": acb_externo,
                            "Quantidade": qtd_num
                        })

                if registros_cores:
                    df_resumo_cores = (
                        pd.DataFrame(registros_cores)
                        .groupby(["Família", "Dimensão", "ACB. TT/PNL", "CAREN. FRONTAL", "ACB. EXTERNO"], as_index=False)["Quantidade"]
                        .sum()
                        .sort_values(by=["Família", "Dimensão", "Quantidade"], ascending=[True, True, False])
                    )
                else:
                    df_resumo_cores = pd.DataFrame(columns=["Família", "Dimensão", "ACB. TT/PNL", "CAREN. FRONTAL", "ACB. EXTERNO", "Quantidade"])

                # --- CÁLCULO DO DASHBOARD DE CONSUMO DE MATERIAIS ---
                total_geral_expositores = df_resumo_cores["Quantidade"].sum() if not df_resumo_cores.empty else 1

                def calcular_dashboard_material(coluna_material):
                    if df_resumo_cores.empty:
                        return pd.DataFrame(columns=["Material / Acabamento", "Qtd Expositores", "Porcentagem (%)"])
                    
                    df_dash = (
                        df_resumo_cores.groupby(coluna_material, as_index=False)["Quantidade"]
                        .sum()
                        .rename(columns={coluna_material: "Material / Acabamento", "Quantidade": "Qtd Expositores"})
                    )
                    df_dash["Porcentagem (%)"] = (df_dash["Qtd Expositores"] / total_geral_expositores * 100).round(2)
                    return df_dash.sort_values(by="Qtd Expositores", ascending=False)

                df_dash_tt = calcular_dashboard_material("ACB. TT/PNL")
                df_dash_caren = calcular_dashboard_material("CAREN. FRONTAL")
                df_dash_ext = calcular_dashboard_material("ACB. EXTERNO")

                st.markdown("---")
                for lote in lotes_orfas:
                    st.markdown(f"#### 📄 ORFA: {lote['orfa']} | Pedido: {lote['pedido']} | Entrega: {lote['entrega']}")
                    st.markdown(f"**Cliente:** {lote['cliente']}")

                    df_exib = pd.DataFrame(
                        lote['itens'] if lote['itens'] else [{"qtd": "-", "modelo": "Nenhum item detectado", "end": "-", "acb_tt_pnl": "-", "carenagem": "-", "acb_ext": "-", "gas": "-", "ilum": "-"}],
                        columns=["qtd", "modelo", "end", "acb_tt_pnl", "carenagem", "acb_ext", "gas", "ilum"]
                    )
                    df_exib.columns = ["Qtd", "Item/Modelo", "End.", "ACB. TT/PNL", "CAREN. FRONTAL", "ACB. EXTERNO", "GÁS", "ILUM."]
                    st.dataframe(df_exib, use_container_width=True, hide_index=True)
                    st.markdown("---")

                # --- EXIBIÇÃO DO RESUMO E DASHBOARD NA TELA ---
                st.markdown("### 🎨 Resumo Geral Detalhado por Família, Dimensão e Acabamentos")
                if not df_resumo_cores.empty:
                    st.dataframe(df_resumo_cores, use_container_width=True, hide_index=True)
                
                st.markdown("---")
                st.markdown("### 📊 Dashboard de Consumo de Materiais e Acabamentos")
                
                col_b1, col_b2, col_b3 = st.columns(3)
                with col_b1:
                    st.markdown("##### 🔷 ACB. TT/PNL")
                    st.dataframe(df_dash_tt, use_container_width=True, hide_index=True)
                with col_b2:
                    st.markdown("##### 🔷 CAREN. FRONTAL")
                    st.dataframe(df_dash_caren, use_container_width=True, hide_index=True)
                with col_b3:
                    st.markdown("##### 🔷 ACB. EXTERNO")
                    st.dataframe(df_dash_ext, use_container_width=True, hide_index=True)
                st.markdown("---")

                # --- GERAÇÃO DO EXCEL ---
                def gerar_excel(lotes, df_cores, df_d_tt, df_d_caren, df_d_ext):
                    out = io.BytesIO()
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = "Consolidado ORFAs"
                    ws.views.sheetView[0].showGridLines = True

                    NAVY = "1F497D"
                    HEADER_BG = "2C5282"
                    SUB_BG = "E2E8F0"
                    ZEBRA = "F7FAFC"
                    BORDER_COLOR = "CBD5E0"

                    f_title = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
                    f_sub = Font(name="Calibri", size=10, bold=True, color=NAVY)
                    f_head = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
                    f_body = Font(name="Calibri", size=9)
                    f_bold_body = Font(name="Calibri", size=9, bold=True)

                    fill_title = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
                    fill_head = PatternFill(start_color=HEADER_BG, end_color=HEADER_BG, fill_type="solid")
                    fill_sub = PatternFill(start_color=SUB_BG, end_color=SUB_BG, fill_type="solid")
                    fill_zebra = PatternFill(start_color=ZEBRA, end_color=ZEBRA, fill_type="solid")
                    border = Border(left=Side(style='thin', color=BORDER_COLOR), right=Side(style='thin', color=BORDER_COLOR),
                                    top=Side(style='thin', color=BORDER_COLOR), bottom=Side(style='thin', color=BORDER_COLOR))

                    ws.merge_cells("A1:H1")
                    t1 = ws["A1"]
                    t1.value = f"RELATÓRIO CONSOLIDADO DE SEPARAÇÃO DE PEÇAS — Total de ORFAs: {len(lotes)}"
                    t1.font = f_title; t1.fill = fill_title; t1.alignment = Alignment(horizontal="center", vertical="center")
                    ws.row_dimensions[1].height = 25

                    curr_row = 3
                    for lote in lotes:
                        ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=8)
                        c_sub = ws.cell(row=curr_row, column=1, value=f"ORFA: {lote['orfa']}  |  Pedido: {lote['pedido']}  |  Entrega: {lote['entrega']}  |  Cliente: {lote['cliente']}")
                        c_sub.font = f_sub; c_sub.fill = fill_sub; c_sub.alignment = Alignment(horizontal="left", vertical="center", indent=1)
                        ws.row_dimensions[curr_row].height = 20
                        curr_row += 1

                        headers = ["Qtd", "Item/Modelo", "End.", "ACB. TT/PNL", "CAREN. FRONTAL", "ACB. EXTERNO", "GÁS", "ILUM."]
                        for c_idx, h in enumerate(headers, 1):
                            c = ws.cell(row=curr_row, column=c_idx, value=h)
                            c.font = f_head; c.fill = fill_head; c.alignment = Alignment(horizontal="center", vertical="center"); c.border = border
                        ws.row_dimensions[curr_row].height = 18
                        curr_row += 1

                        for idx, item in enumerate(lote['itens']):
                            vals = [item["qtd"], item["modelo"], item["end"], item["acb_tt_pnl"], item["carenagem"], item["acb_ext"], item["gas"], item["ilum"]]
                            for c_idx, val in enumerate(vals, 1):
                                c = ws.cell(row=curr_row, column=c_idx, value=val)
                                c.font = f_body; c.border = border
                                c.alignment = Alignment(horizontal="center" if c_idx in [1,3,7,8] else "left", vertical="center")
                                if idx % 2 == 1:
                                    c.fill = fill_zebra
                            curr_row += 1
                        curr_row += 1

                    # --- RESUMO DETALHADO ---
                    if not df_cores.empty:
                        curr_row += 1
                        ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=6)
                        c_res_title = ws.cell(row=curr_row, column=1, value="RESUMO GERAL DETALHADO POR FAMÍLIA, DIMENSÃO E ACABAMENTOS")
                        c_res_title.font = f_title; c_res_title.fill = fill_title; c_res_title.alignment = Alignment(horizontal="center", vertical="center")
                        ws.row_dimensions[curr_row].height = 22
                        curr_row += 1

                        res_headers = ["Família", "Dimensão", "ACB. TT/PNL", "CAREN. FRONTAL", "ACB. EXTERNO", "Qtd Expositores"]
                        for c_idx, h in enumerate(res_headers, 1):
                            c = ws.cell(row=curr_row, column=c_idx, value=h)
                            c.font = f_head; c.fill = fill_head; c.alignment = Alignment(horizontal="center", vertical="center"); c.border = border
                        ws.row_dimensions[curr_row].height = 18
                        curr_row += 1

                        for idx_c, row_c in df_cores.iterrows():
                            vals_c = [row_c["Família"], row_c["Dimensão"], row_c["ACB. TT/PNL"], row_c["CAREN. FRONTAL"], row_c["ACB. EXTERNO"], int(row_c["Quantidade"])]
                            for c_idx, val in enumerate(vals_c, 1):
                                c = ws.cell(row=curr_row, column=c_idx, value=val)
                                c.font = f_bold_body if c_idx in [1, 2, 6] else f_body
                                c.border = border
                                c.alignment = Alignment(horizontal="center" if c_idx in [1, 2, 6] else "left", vertical="center")
                                if idx_c % 2 == 1:
                                    c.fill = fill_zebra
                            curr_row += 1
                        curr_row += 2

                    # --- DASHBOARD DE CONSUMO NO EXCEL ---
                    def escrever_tabela_dashboard(df_d, nome_tabela):
                        nonlocal curr_row
                        ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=3)
                        c_dt = ws.cell(row=curr_row, column=1, value=f"DASHBOARD DE CONSUMO — {nome_tabela}")
                        c_dt.font = f_title; c_dt.fill = fill_title; c_dt.alignment = Alignment(horizontal="center", vertical="center")
                        ws.row_dimensions[curr_row].height = 20
                        curr_row += 1

                        d_headers = ["Material / Acabamento", "Qtd Expositores", "Porcentagem (%)"]
                        for c_idx, h in enumerate(d_headers, 1):
                            c = ws.cell(row=curr_row, column=c_idx, value=h)
                            c.font = f_head; c.fill = fill_head; c.alignment = Alignment(horizontal="center", vertical="center"); c.border = border
                        ws.row_dimensions[curr_row].height = 18
                        curr_row += 1

                        for idx_d, row_d in df_d.iterrows():
                            c1 = ws.cell(row=curr_row, column=1, value=row_d["Material / Acabamento"])
                            c1.font = f_body; c1.alignment = Alignment(horizontal="left", vertical="center"); c1.border = border

                            c2 = ws.cell(row=curr_row, column=2, value=int(row_d["Qtd Expositores"]))
                            c2.font = f_bold_body; c2.alignment = Alignment(horizontal="center", vertical="center"); c2.border = border

                            c3 = ws.cell(row=curr_row, column=3, value=f"{row_d['Porcentagem (%)']:.2f}%")
                            c3.font = f_bold_body; c3.alignment = Alignment(horizontal="center", vertical="center"); c3.border = border

                            if idx_d % 2 == 1:
                                c1.fill = fill_zebra; c2.fill = fill_zebra; c3.fill = fill_zebra
                            curr_row += 1
                        curr_row += 2

                    escrever_tabela_dashboard(df_d_tt, "ACB. TT/PNL")
                    escrever_tabela_dashboard(df_d_caren, "CARENAGEM FRONTAL")
                    escrever_tabela_dashboard(df_d_ext, "ACB. EXTERNO")

                    for col_letter, width in {'A': 25, 'B': 18, 'C': 18, 'D': 22, 'E': 22, 'F': 15}.items():
                        ws.column_dimensions[col_letter].width = width

                    wb.save(out)
                    out.seek(0)
                    return out.getvalue()

                # --- GERAÇÃO DO PDF CONSOLIDADO ---
                def gerar_pdf(lotes, df_cores, df_d_tt, df_d_caren, df_d_ext):
                    pdf_out = io.BytesIO()
                    doc = SimpleDocTemplate(pdf_out, pagesize=landscape(A4), leftMargin=1*cm, rightMargin=1*cm, topMargin=1*cm, bottomMargin=1*cm)
                    styles = getSampleStyleSheet()
                    
                    style_title = ParagraphStyle('T', fontName='Helvetica-Bold', fontSize=13, textColor=colors.white, alignment=1)
                    style_sub = ParagraphStyle('S', fontName='Helvetica-Bold', fontSize=9, textColor=colors.HexColor("#1F497D"))
                    style_h = ParagraphStyle('H', fontName='Helvetica-Bold', fontSize=8, textColor=colors.white, alignment=1)
                    style_c = ParagraphStyle('C', fontName='Helvetica', fontSize=7, leading=9)
                    style_cc = ParagraphStyle('CC', fontName='Helvetica', fontSize=7, leading=9, alignment=1)
                    style_ccb = ParagraphStyle('CCB', fontName='Helvetica-Bold', fontSize=7.5, leading=9, alignment=1)

                    story = []
                    
                    pdf_title_text = f"RELATÓRIO CONSOLIDADO DE SEPARAÇÃO DE PEÇAS — Total de ORFAs: {len(lotes)}"
                    t_top = Table([[Paragraph(pdf_title_text, style_title)]], colWidths=[27.7*cm])
                    t_top.setStyle(TableStyle([
                        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#1F497D")),
                        ('TOPPADDING', (0,0), (-1,-1), 8),
                        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
                    ]))
                    story.append(t_top)
                    story.append(Spacer(1, 10))

                    for lote in lotes:
                        bloqueio = []
                        txt_lote = f"■ ORFA: {lote['orfa']} | Pedido: {lote['pedido']} | Entrega: {lote['entrega']} | Cliente: {lote['cliente']}"
                        t_lote = Table([[Paragraph(txt_lote, style_sub)]], colWidths=[27.7*cm])
                        t_lote.setStyle(TableStyle([
                            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#E2E8F0")),
                            ('TOPPADDING', (0,0), (-1,-1), 4),
                            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
                        ]))
                        bloqueio.append(t_lote)

                        rows = [[
                            Paragraph("Qtd", style_h), Paragraph("Item/Modelo", style_h), Paragraph("End.", style_h),
                            Paragraph("■ ACB. TT/PNL", style_h), Paragraph("■ CAREN. FRONTAL", style_h),
                            Paragraph("■ ACB. EXTERNO", style_h), Paragraph("GÁS", style_h), Paragraph("■ ILUM.", style_h)
                        ]]

                        for item in lote['itens']:
                            rows.append([
                                Paragraph(item["qtd"], style_cc),
                                Paragraph(item["modelo"], style_c),
                                Paragraph(item["end"], style_cc),
                                Paragraph(item["acb_tt_pnl"], style_c),
                                Paragraph(item["carenagem"], style_c),
                                Paragraph(item["acb_ext"], style_c),
                                Paragraph(item["gas"], style_cc),
                                Paragraph(item["ilum"], style_cc),
                            ])

                        t_table = Table(rows, colWidths=[1.2*cm, 5.5*cm, 2.5*cm, 4.8*cm, 4.8*cm, 4.8*cm, 2*cm, 2*cm])
                        t_table.setStyle(TableStyle([
                            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#2C5282")),
                            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
                            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor("#F7FAFC")]),
                            ('TOPPADDING', (0,0), (-1,-1), 3),
                            ('BOTTOMPADDING', (0,0), (-1,-1), 3),
                        ]))
                        bloqueio.append(t_table)
                        bloqueio.append(Spacer(1, 10))
                        story.append(KeepTogether(bloqueio))

                    # --- RESUMO DETALHADO NO PDF ---
                    if not df_cores.empty:
                        story.append(Spacer(1, 8))
                        bloco_resumo = []
                        t_res_banner = Table([[Paragraph("RESUMO GERAL DETALHADO POR FAMÍLIA, DIMENSÃO E ACABAMENTOS", style_sub)]], colWidths=[27.7*cm])
                        t_res_banner.setStyle(TableStyle([
                            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#E2E8F0")),
                            ('TOPPADDING', (0,0), (-1,-1), 5),
                            ('BOTTOMPADDING', (0,0), (-1,-1), 5),
                        ]))
                        bloco_resumo.append(t_res_banner)

                        rows_cores = [[
                            Paragraph("Família", style_h), Paragraph("Dimensão", style_h), Paragraph("ACB. TT/PNL", style_h),
                            Paragraph("CAREN. FRONTAL", style_h), Paragraph("ACB. EXTERNO", style_h), Paragraph("Qtd Expositores", style_h)
                        ]]

                        for _, r_cor in df_cores.iterrows():
                            rows_cores.append([
                                Paragraph(str(r_cor["Família"]), style_ccb),
                                Paragraph(str(r_cor["Dimensão"]), style_cc),
                                Paragraph(str(r_cor["ACB. TT/PNL"]), style_c),
                                Paragraph(str(r_cor["CAREN. FRONTAL"]), style_c),
                                Paragraph(str(r_cor["ACB. EXTERNO"]), style_c),
                                Paragraph(str(int(r_cor["Quantidade"])), style_ccb),
                            ])

                        t_cores = Table(rows_cores, colWidths=[4.5*cm, 2.5*cm, 6.0*cm, 6.0*cm, 6.0*cm, 2.7*cm])
                        t_cores.setStyle(TableStyle([
                            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#2C5282")),
                            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
                            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor("#F7FAFC")]),
                            ('TOPPADDING', (0,0), (-1,-1), 4),
                            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
                        ]))
                        bloco_resumo.append(t_cores)
                        story.append(KeepTogether(bloco_resumo))

                    # --- DASHBOARD DE CONSUMO NO PDF ---
                    def criar_tabela_pdf_dashboard(df_d, titulo_dash):
                        bloco_d = []
                        bloco_d.append(Spacer(1, 10))
                        t_b = Table([[Paragraph(f"DASHBOARD DE CONSUMO — {titulo_dash}", style_sub)]], colWidths=[27.7*cm])
                        t_b.setStyle(TableStyle([
                            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#E2E8F0")),
                            ('TOPPADDING', (0,0), (-1,-1), 4),
                            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
                        ]))
                        bloco_d.append(t_b)

                        r_d = [[Paragraph("Material / Acabamento", style_h), Paragraph("Qtd Expositores", style_h), Paragraph("Porcentagem (%)", style_h)]]
                        for _, row_d in df_d.iterrows():
                            r_d.append([
                                Paragraph(str(row_d["Material / Acabamento"]), style_c),
                                Paragraph(str(int(row_d["Qtd Expositores"])), style_ccb),
                                Paragraph(f"{row_d['Porcentagem (%)']:.2f}%", style_ccb),
                            ])

                        t_dt = Table(r_d, colWidths=[17.7*cm, 5.0*cm, 5.0*cm])
                        t_dt.setStyle(TableStyle([
                            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#2C5282")),
                            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
                            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor("#F7FAFC")]),
                            ('TOPPADDING', (0,0), (-1,-1), 3),
                            ('BOTTOMPADDING', (0,0), (-1,-1), 3),
                        ]))
                        bloco_d.append(t_dt)
                        return KeepTogether(bloco_d)

                    if not df_d_tt.empty:
                        story.append(criar_tabela_pdf_dashboard(df_d_tt, "ACB. TT/PNL"))
                    if not df_d_caren.empty:
                        story.append(criar_tabela_pdf_dashboard(df_d_caren, "CARENAGEM FRONTAL"))
                    if not df_d_ext.empty:
                        story.append(criar_tabela_pdf_dashboard(df_d_ext, "ACB. EXTERNO"))

                    doc.build(story)
                    pdf_out.seek(0)
                    return pdf_out.getvalue()

                col_d1, col_d2 = st.columns(2)
                with col_d1:
                    st.download_button(
                        label="📥 Baixar Relatório Consolidado em Excel (.xlsx)",
                        data=gerar_excel(lotes_orfas, df_resumo_cores, df_dash_tt, df_dash_caren, df_dash_ext),
                        file_name="Relatorio_Consolidado_ORFAS.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                with col_d2:
                    st.download_button(
                        label="📥 Baixar Relatório Consolidado em PDF (.pdf)",
                        data=gerar_pdf(lotes_orfas, df_resumo_cores, df_dash_tt, df_dash_caren, df_dash_ext),
                        file_name="Relatorio_Consolidado_ORFAS.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                    )