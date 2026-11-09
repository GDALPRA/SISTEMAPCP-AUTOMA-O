from datetime import datetime
import io
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd  # <-- Biblioteca pandas importada corretamente

# Bibliotecas para o PDF Executivo (ReportLab)
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
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

# --- ESTILIZAÇÃO DO CABEÇALHO DA PÁGINA ---
st.markdown(
    """
    <div style="background-color: #1F497D; color: white; padding: 20px; border-radius: 8px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
        <h2 style="margin: 0; font-weight: 700;">📋 RELATÓRIO CONSOLIDADO DE SEPARAÇÃO DE PEÇAS (ORFAS)</h2>
        <p style="margin: 5px 0 0 0; opacity: 0.9; font-size: 14px;">Portal Corporativo de PCP & Engenharia - Extração por Lote de ORFAs</p>
    </div>
""",
    unsafe_allow_html=True,
)

st.write("")

# --- ÁREA DE UPLOAD DE MÚLTIPLOS PDFS ---
uploaded_files = st.file_uploader(
    "Arraste ou selecione os arquivos PDF das ORFAs de origem (.pdf)",
    type=["pdf"],
    accept_multiple_files=True,
    key="pdf_uploader_orfas",
)

if uploaded_files:
  st.info(f"📁 {len(uploaded_files)} arquivo(s) PDF carregado(s) com sucesso na memória.")

  if st.button("🚀 Gerar Relatório Consolidado de ORFAs", type="primary"):
    with st.spinner("Analisando as ORFAs e estruturando os grupos de separação..."):
      lotes_orfas = []

      for uploaded_file in uploaded_files:
        try:
          import pdfplumber

          with pdfplumber.open(uploaded_file) as pdf:
            texto_completo = "\n".join([page.extract_text() or "" for page in pdf.pages])

          # Extração de Metadados da ORFA
          orfa_m = re.search(r"N[º°]\s*ORFA:\s*([A-Za-z0-9\-/]+)", texto_completo, re.IGNORECASE)
          ped_m = re.search(r"N[º°]\s*PEDIDO:\s*([A-Za-z0-9\-/]+)", texto_completo, re.IGNORECASE)
          dt_entrega_m = re.search(r"DATA ENTREGA:\s*([\d/]+)", texto_completo, re.IGNORECASE)
          cli_m = re.search(r"CLIENTE:\s*(.+)", texto_completo, re.IGNORECASE)

          num_orfa = orfa_m.group(1).strip() if orfa_m else uploaded_file.name
          pedido = ped_m.group(1).strip() if ped_m else "N/A"
          entrega = dt_entrega_m.group(1).strip() if dt_entrega_m else "-"
          cliente = cli_m.group(1).strip() if cli_m else "CLIENTE NÃO INFORMADO"

          # Características globais de acabamento extraídas do documento
          acb_tcn_m = re.search(r"Acb\.\s*TT/PNL:\s*([^\n]+)", texto_completo, re.IGNORECASE)
          carenagem_m = re.search(r"Caren\.\s*Frontal:\s*([^\n]+)", texto_completo, re.IGNORECASE)
          acb_ext_m = re.search(r"Acb\.\s*Externo:\s*([^\n]+)", texto_completo, re.IGNORECASE)
          gas_m = re.search(r"Fluido/Gás:\s*([^\n]+)", texto_completo, re.IGNORECASE)
          ilum_m = re.search(r"Iluminação:\s*([^\n]+)", texto_completo, re.IGNORECASE)

          acb_tcn_global = acb_tcn_m.group(1).strip() if acb_tcn_m else "INOX"
          carenagem_global = carenagem_m.group(1).strip() if carenagem_m else "PRETO (P1902)"
          acb_ext_global = acb_ext_m.group(1).strip() if acb_ext_m else "PRETO (P1902)"
          gas_global = gas_m.group(1).strip() if gas_m else "GLICOL"
          ilum_global = ilum_m.group(1).strip() if ilum_m else "4000K"

          itens_orfa = []

          # Varre a tabela de resumo final da ORFA (se presente no documento)
          if "Resumo de Balcões por Endereço" in texto_completo or "Balcão" in texto_completo:
            linhas = texto_completo.split("\n")
            capturando = False
            for linha in linhas:
              if "Resumo de Balcões por Endereço" in linha:
                capturando = True
                continue
              if capturando:
                # Padrão da tabela de resumo: Qtd (ex: 01, 03) | Balcão (ex: VABP PVT) | Largura | Comprimento | Endereço (ex: 1.2-1.3-1.4)
                m_res = re.match(r"^(\d{2})\s+([A-Z0-9\s\+]+?)\s+(\d+)\s+(\d+)\s+([\d\.\-]+)", linha.strip())
                if m_res:
                  q_str, modelo_base, largura, comprimento, endereco = m_res.groups()
                  modelo_completo = f"{modelo_base.strip()} {comprimento.strip()}"
                  itens_orfa.append({
                      "qtd": q_str.strip(),
                      "modelo": modelo_completo,
                      "end": endereco.strip(),
                      "acb_tt_pnl": acb_tcn_global,
                      "carenagem": carenagem_global,
                      "acb_ext": acb_ext_global,
                      "gas": gas_global,
                      "ilum": ilum_global
                  })

          # Fallback: Se não pegou pela tabela de resumo, busca linhas comuns de modelos no texto
          if not itens_orfa:
            linhas = texto_completo.split("\n")
            for linha in linhas:
              m_item = re.match(r"^(\s*0?[1-9])\s+([A-Z0-9\s\-]+?(?:PVT|VAP|VAL|VAH|VRA|VRAG|VACP|VAHP|VIL|VCAG).*?)(?:\s*\(|$)", linha.strip())
              if m_item:
                q_str, modelo = m_item.groups()
                try:
                  q_val = int(q_str.strip())
                except:
                  q_val = 1
                end_m = re.search(r"\(([\d\.\-]+)\)", linha)
                endereco = end_m.group(1) if end_m else "1.1"

                itens_orfa.append({
                    "qtd": f"{q_val:02d}",
                    "modelo": modelo.strip(),
                    "end": endereco,
                    "acb_tt_pnl": acb_tcn_global,
                    "carenagem": carenagem_global,
                    "acb_ext": acb_ext_global,
                    "gas": gas_global,
                    "ilum": ilum_global
                })

          if not itens_orfa:
            itens_orfa.append({
                "qtd": "01",
                "modelo": "EXPOSITOR GERAL ORFA",
                "end": "1.1",
                "acb_tt_pnl": acb_tcn_global,
                "carenagem": carenagem_global,
                "acb_ext": acb_ext_global,
                "gas": gas_global,
                "ilum": ilum_global
            })

          lotes_orfas.append({
              "orfa": num_orfa,
              "pedido": pedido,
              "entrega": entrega,
              "cliente": cliente,
              "itens": itens_orfa
          })

        except Exception as e:
          st.error(f"⚠️ Erro ao processar o arquivo {uploaded_file.name}: {e}")

      if not lotes_orfas:
        st.warning("❌ Nenhuma ORFA válida foi processada.")
      else:
        st.success(f"🎉 Processamento concluído! Total de ORFAs no lote: {len(lotes_orfas)}")

        # --- EXIBIÇÃO VISUAL DOS GRUPOS NA TELA ---
        st.markdown("---")
        for lote in lotes_orfas:
          st.markdown(f"#### 📄 ORFA: {lote['orfa']} | Pedido: {lote['pedido']} | Entrega: {lote['entrega']}")
          st.markdown(f"**Cliente:** {lote['cliente']}")

          df_exibicao = pd.DataFrame(lote['itens'])
          df_exibicao.columns = ["Qtd", "Item/Modelo", "End.", "ACB. TT/PNL", "CAREN. FRONTAL", "ACB. EXTERNO", "GÁS", "ILUM."]
          st.dataframe(df_exibicao, use_container_width=True, hide_index=True)
          st.markdown("---")

        # --- GERAÇÃO DO EXCEL CONSOLIDADO EXECUTIVO ---
        def gerar_excel_consolidado(lotes):
          output_bytes = io.BytesIO()
          wb = openpyxl.Workbook()
          ws = wb.active
          ws.title = "Consolidado ORFAs"
          ws.views.sheetView[0].showGridLines = True

          NAVY = "1F497D"
          WHITE = "FFFFFF"
          HEADER_BG = "2C5282"
          ZEBRA = "F7FAFC"
          BORDER_COLOR = "CBD5E0"

          font_title = Font(name="Calibri", size=13, bold=True, color=WHITE)
          font_subtitle = Font(name="Calibri", size=10, bold=True, color=NAVY)
          font_header = Font(name="Calibri", size=9, bold=True, color=WHITE)
          font_body = Font(name="Calibri", size=9)

          fill_title = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
          fill_header = PatternFill(start_color=HEADER_BG, end_color=HEADER_BG, fill_type="solid")
          fill_sub = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
          fill_zebra = PatternFill(start_color=ZEBRA, end_color=ZEBRA, fill_type="solid")
          thin_border = Border(left=Side(style='thin', color=BORDER_COLOR), right=Side(style='thin', color=BORDER_COLOR),
                               top=Side(style='thin', color=BORDER_COLOR), bottom=Side(style='thin', color=BORDER_COLOR))

          ws.merge_cells("A1:H1")
          t1 = ws["A1"]
          t1.value = f"RELATÓRIO CONSOLIDADO DE SEPARAÇÃO DE PEÇAS — Total de ORFAs: {len(lotes)}"
          t1.font = font_title; t1.fill = fill_title; t1.alignment = Alignment(horizontal="center", vertical="center")
          ws.row_dimensions[1].height = 25

          current_row = 3
          for lote in lotes:
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=8)
            cab_cell = ws.cell(row=current_row, column=1, value=f"ORFA: {lote['orfa']}  |  Pedido: {lote['pedido']}  |  Entrega: {lote['entrega']}  |  Cliente: {lote['cliente']}")
            cab_cell.font = font_subtitle; cab_cell.fill = fill_sub
            cab_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            ws.row_dimensions[current_row].height = 22
            current_row += 1

            headers = ["Qtd", "Item/Modelo", "End.", "ACB. TT/PNL", "CAREN. FRONTAL", "ACB. EXTERNO", "GÁS", "ILUM."]
            for c_idx, h in enumerate(headers, 1):
              c = ws.cell(row=current_row, column=c_idx, value=h)
              c.font = font_header; c.fill = fill_header
              c.alignment = Alignment(horizontal="center", vertical="center")
              c.border = thin_border
            ws.row_dimensions[current_row].height = 20
            current_row += 1

            for idx, item in enumerate(lote['itens']):
              vals = [item["qtd"], item["modelo"], item["end"], item["acb_tt_pnl"], item["carenagem"], item["acb_ext"], item["gas"], item["ilum"]]
              for c_idx, val in enumerate(vals, 1):
                c = ws.cell(row=current_row, column=c_idx, value=val)
                c.font = font_body; c.border = thin_border
                if c_idx in [1, 3, 7, 8]:
                  c.alignment = Alignment(horizontal="center", vertical="center")
                else:
                  c.alignment = Alignment(horizontal="left", vertical="center")
                if idx % 2 == 1:
                  c.fill = fill_zebra
              ws.row_dimensions[current_row].height = 18
              current_row += 1

            current_row += 2

          widths = {'A': 8, 'B': 30, 'C': 12, 'D': 18, 'E': 18, 'F': 18, 'G': 12, 'H': 12}
          for col_letter, width in widths.items():
            ws.column_dimensions[col_letter].width = width

          wb.save(output_bytes)
          output_bytes.seek(0)
          return output_bytes.getvalue()

        excel_data = gerar_excel_consolidado(lotes_orfas)

        st.write("---")
        st.download_button(
            label="📥 Baixar Relatório Consolidado de ORFAs em Excel (.xlsx)",
            data=excel_data,
            file_name="Relatorio_Consolidado_ORFAS.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )