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
    <div style="background-color: #1F497D; color: white; padding: 20px; border-radius: 8px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
        <h2 style="margin: 0; font-weight: 700;">📋 RELATÓRIO CONSOLIDADO DE SEPARAÇÃO DE PEÇAS (ORFAS)</h2>
        <p style="margin: 5px 0 0 0; opacity: 0.9; font-size: 14px;">Portal Corporativo de PCP & Engenharia - Processamento de ORFAs</p>
    </div>
""",
    unsafe_allow_html=True,
)

st.write("")

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
        with st.spinner("Analisando metadados e blocos das ORFAs..."):
            lotes_orfas = []

            for uploaded_file in uploaded_files:
                try:
                    import pdfplumber

                    with pdfplumber.open(uploaded_file) as pdf:
                        paginas_texto = [page.extract_text() or "" for page in pdf.pages]
                        texto_completo = "\n".join(paginas_texto)

                    # --- EXTRAÇÃO ROBUSTA DE METADADOS DO CABEÇALHO DA ORFA ---
                    orfa_m = re.search(r"N[º°]?\s*ORFA[:\s]*\n?([A-Za-z0-9\-/()]+)", texto_completo, re.IGNORECASE)
                    ped_m = re.search(r"N[º°]?\s*PEDIDO[:\s]*\n?([A-Za-z0-9\-/]+(?:\s*Rev[:\s]*[A-Za-z0-9]*)?)", texto_completo, re.IGNORECASE)
                    dt_entrega_m = re.search(r"DATA\s*ENTREGA[:\s]*\n?([\d/]+)", texto_completo, re.IGNORECASE)
                    cli_m = re.search(r"CLIENTE[:\s]*\n?([^\n]+)", texto_completo, re.IGNORECASE)

                    num_orfa = orfa_m.group(1).strip() if orfa_m else uploaded_file.name.replace(".pdf", "")
                    pedido = ped_m.group(1).strip() if ped_m else "N/A"
                    entrega = dt_entrega_m.group(1).strip() if dt_entrega_m else "-"
                    cliente = cli_m.group(1).strip() if cli_m else "NÃO INFORMADO"

                    itens_orfa = []

                    # Processamento por Páginas/Blocos da ORFA (evitando a página de resumo final para não duplicar)
                    for pagina in paginas_texto:
                        if "Resumo de Balcões por Endereço" in pagina:
                            continue

                        linhas = pagina.split("\n")
                        for i, linha in enumerate(linhas):
                            # Captura modelos com quantidade e o parênteses de endereço ex: "01 VABP PVT 2500 (1.1)"
                            m_item = re.search(r"^(\d{2})\s+([A-Z0-9\s\-/]+?(?:PVT|VAP|VAL|VAH|VRA|VRAG|VACP|VAHP|VIL|VCAG|ICFT|VRQU).*?)(?:\s*\(([\d\.\-\s]+)\)|$)", linha.strip())
                            
                            if m_item:
                                q_str, modelo_raw, end_opt = m_item.groups()
                                
                                modelo_limpo = re.sub(r"\s*\([\d\.\-\s]+\)", "", modelo_raw).strip()
                                
                                if not end_opt:
                                    end_match = re.search(r"\(([\d\.\-\s]+)\)", linha)
                                    endereco = end_match.group(1).strip() if end_match else "1.1"
                                else:
                                    endereco = end_opt.strip()

                                bloco = "\n".join(linhas[max(0, i-6):min(len(linhas), i+12)])
                                
                                acb_tt = re.search(r"Acb\.\s*TT/PNL:\s*([^\n|]+)", bloco, re.IGNORECASE)
                                caren = re.search(r"Caren\.\s*Frontal:\s*([^\n|]+)", bloco, re.IGNORECASE)
                                acb_ext = re.search(r"Acb\.\s*Externo:\s*([^\n|]+)", bloco, re.IGNORECASE)
                                gas = re.search(r"Fluido/Gás:\s*([^\n|]+)", bloco, re.IGNORECASE)
                                ilum = re.search(r"Iluminação:\s*([^\n|]+)", bloco, re.IGNORECASE)

                                val_tt = acb_tt.group(1).strip() if acb_tt else "INOX"
                                val_car = caren.group(1).strip() if caren else "PRETO (P1902)"
                                val_ext = acb_ext.group(1).strip() if acb_ext else "PRETO (P1902)"
                                val_gas = gas.group(1).strip() if gas else "GLICOL"
                                val_ilu = ilum.group(1).strip() if ilum else "4000K"

                                val_tt = re.split(r"Ach\.|Acb\.|Bojo", val_tt)[0].strip()
                                val_car = re.split(r"Porta|Serigrafia", val_car)[0].strip()
                                val_ext = re.split(r"Fluido|Gás|Porta", val_ext)[0].strip()

                                itens_orfa.append({
                                    "qtd": q_str.strip(),
                                    "modelo": modelo_limpo,
                                    "end": endereco,
                                    "acb_tt_pnl": val_tt if val_tt else "INOX",
                                    "carenagem": val_car if val_car else "PRETO (P1902)",
                                    "acb_ext": val_ext if val_ext else "PRETO (P1902)",
                                    "gas": val_gas if val_gas else "GLICOL",
                                    "ilum": val_ilu if val_ilu else "4000K"
                                })

                    if not itens_orfa and "Resumo de Balcões por Endereço" in texto_completo:
                        for pagina in paginas_texto:
                            if "Resumo de Balcões por Endereço" in pagina:
                                for linha in pagina.split("\n"):
                                    m_res = re.match(r"^(\d{2})\s+([A-Z0-9\s\+]+?)\s+(\d+)\s+(\d+)\s+([\d\.\-]+)", linha.strip())
                                    if m_res:
                                        q_str, modelo_base, larg, comp, endereco = m_res.groups()
                                        itens_orfa.append({
                                            "qtd": q_str.strip(),
                                            "modelo": f"{modelo_base.strip()} {comp.strip()}",
                                            "end": endereco.strip(),
                                            "acb_tt_pnl": "INOX",
                                            "carenagem": "VINHO (H1007)",
                                            "acb_ext": "VINHO (H1007)",
                                            "gas": "GLICOL",
                                            "ilum": "4000K"
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

                st.markdown("---")
                for lote in lotes_orfas:
                    st.markdown(f"#### 📄 ORFA: {lote['orfa']} | Pedido: {lote['pedido']} | Entrega: {lote['entrega']}")
                    st.markdown(f"**Cliente:** {lote['cliente']}")

                    df_exib = pd.DataFrame(lote['itens'])
                    df_exib.columns = ["Qtd", "Item/Modelo", "End.", "ACB. TT/PNL", "CAREN. FRONTAL", "ACB. EXTERNO", "GÁS", "ILUM."]
                    st.dataframe(df_exib, use_container_width=True, hide_index=True)
                    st.markdown("---")

                # --- GERAÇÃO DO EXCEL ---
                def gerar_excel(lotes):
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

                    for col_letter, width in {'A': 6, 'B': 25, 'C': 15, 'D': 20, 'E': 20, 'F': 20, 'G': 12, 'H': 10}.items():
                        ws.column_dimensions[col_letter].width = width

                    wb.save(out)
                    out.seek(0)
                    return out.getvalue()

                # --- GERAÇÃO DO PDF CONSOLIDADO ---
                def gerar_pdf(lotes):
                    pdf_out = io.BytesIO()
                    doc = SimpleDocTemplate(pdf_out, pagesize=landscape(A4), leftMargin=1*cm, rightMargin=1*cm, topMargin=1*cm, bottomMargin=1*cm)
                    styles = getSampleStyleSheet()
                    
                    style_title = ParagraphStyle('T', fontName='Helvetica-Bold', fontSize=13, textColor=colors.white, alignment=1)
                    style_sub = ParagraphStyle('S', fontName='Helvetica-Bold', fontSize=9, textColor=colors.HexColor("#1F497D"))
                    style_h = ParagraphStyle('H', fontName='Helvetica-Bold', fontSize=8, textColor=colors.white, alignment=1)
                    style_c = ParagraphStyle('C', fontName='Helvetica', fontSize=7, leading=9)
                    style_cc = ParagraphStyle('CC', fontName='Helvetica', fontSize=7, leading=9, alignment=1)

                    story = []
                    
                    t_top = Table([[Paragraph(f"RELATÓRIO CONSOLIDADO DE SEPARAÇÃO DE PEÇAS — Total de ORFAs: {len(lotes)}", style_title)]], colWidths=[27.7*cm])
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

                    doc.build(story)
                    pdf_out.seek(0)
                    return pdf_out.getvalue()

                col_d1, col_d2 = st.columns(2)
                with col_d1:
                    st.download_button(
                        label="📥 Baixar Relatório Consolidado em Excel (.xlsx)",
                        data=gerar_excel(lotes_orfas),
                        file_name="Relatorio_Consolidado_ORFAS.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                with col_d2:
                    st.download_button(
                        label="📥 Baixar Relatório Consolidado em PDF (.pdf)",
                        data=gerar_pdf(lotes_orfas),
                        file_name="Relatorio_Consolidado_ORFAS.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                    )