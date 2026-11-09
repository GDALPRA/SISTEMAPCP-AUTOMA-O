from datetime import datetime
import io
import os
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

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
        <h2 style="margin: 0; font-weight: 700;">📊 PROCESSADOR DE PEDIDOS - EXCEL & PDF EXECUTIVO</h2>
        <p style="margin: 5px 0 0 0; opacity: 0.9; font-size: 14px;">Portal Corporativo de PCP & Engenharia - Automação de Relatórios</p>
    </div>
""",
    unsafe_allow_html=True,
)

st.write("")

# --- ÁREA DE UPLOAD DE MÚLTIPLOS PDFS ---
uploaded_files = st.file_uploader(
    "Arraste ou selecione os arquivos PDF dos pedidos de origem (.pdf)",
    type=["pdf"],
    accept_multiple_files=True,
    key="pdf_uploader_gerencial",
)

if uploaded_files:
  st.info(
      f"📁 {len(uploaded_files)} arquivo(s) PDF carregado(s) com sucesso na"
      " memória."
  )

  if st.button("🚀 Processar Dados e Gerar Relatórios", type="primary"):
    with st.spinner(
        "Extraindo dados dos PDFs e gerando relatórios executivos..."
    ):
      dados = []
      padrao_item = re.compile(
          r"^(\d{6})\s+(.+?)\s+(\d+,\d{2})$", re.MULTILINE
      )

      # Processamento dos arquivos enviados pelo Streamlit
      for uploaded_file in uploaded_files:
        try:
          import pdfplumber

          with pdfplumber.open(uploaded_file) as pdf:
            texto = "\n".join(
                [page.extract_text() or "" for page in pdf.pages]
            )

          ped_m = re.search(r"Pedido:\s*(\w+)", texto)
          dt_venda_m = re.search(r"Data de venda:\s*([\d/]+)", texto)
          prev_ent_m = re.search(r"Previsão de entrega:\_*\s*([\d/]+)", texto)
          cli_m = re.search(r"Cliente:\s*(.+)", texto)
          cid_m = re.search(r"LTDA\n\s*(.+- SP)", texto)
          frete_m = re.search(r"Tipo de Frete:\s*(\w+)", texto)

          pedido = ped_m.group(1) if ped_m else "N/A"
          data_venda = dt_venda_m.group(1) if dt_venda_m else "N/A"
          prev_entrega = prev_ent_m.group(1) if prev_ent_m else "N/A"
          cliente = cli_m.group(1).strip() if cli_m else "Cliente Desconhecido"
          cidade_uf = cid_m.group(1).strip() if cid_m else ""
          frete = frete_m.group(1) if frete_m else "N/A"

          for match in padrao_item.finditer(texto):
            cod_prod, descricao, qtd = match.groups()
            qtd_float = float(qtd.replace(".", "").replace(",", "."))

            dados.append({
                "arquivo": uploaded_file.name,
                "pedido": pedido,
                "data_venda": data_venda,
                "prev_entrega": prev_entrega,
                "cliente": cliente,
                "cidade_uf": cidade_uf,
                "frete": frete,
                "cod_prod": cod_prod,
                "descricao": descricao,
                "quantidade": qtd_float,
            })
        except Exception as e:
          st.error(f"⚠️ Erro ao ler {uploaded_file.name}: {e}")

      if not dados:
        st.warning(
            "❌ Nenhum dado compatível foi extraído dos PDFs enviados."
        )
      else:
        # =========================================================================
        # GERADOR 1: GERAR PLANILHA EXCEL EM MEMÓRIA (BYTESIO)
        # =========================================================================
        excel_buffer = io.BytesIO()
        wb = openpyxl.Workbook()

        # Aba 1 - Relatório Gerencial
        ws1 = wb.active
        ws1.title = "Relatório Gerencial"
        ws1.views.sheetView[0].showGridLines = True

        NAVY = "1B365D"
        HEADER_BLUE = "2C5282"
        ACCENT_BLUE = "3182CE"
        LIGHT_BLUE = "E8EEF5"
        ZEBRA = "F7FAFC"
        WHITE = "FFFFFF"
        BORDER_COLOR = "CBD5E0"

        font_title = Font(name="Calibri", size=15, bold=True, color=WHITE)
        font_header = Font(name="Calibri", size=10, bold=True, color=WHITE)
        font_data = Font(name="Calibri", size=10)
        font_total = Font(name="Calibri", size=10, bold=True, color=NAVY)

        fill_title = PatternFill(
            start_color=NAVY, end_color=NAVY, fill_type="solid"
        )
        fill_header = PatternFill(
            start_color=HEADER_BLUE, end_color=HEADER_BLUE, fill_type="solid"
        )
        fill_cli = PatternFill(
            start_color=ACCENT_BLUE, end_color=ACCENT_BLUE, fill_type="solid"
        )
        fill_zebra = PatternFill(
            start_color=ZEBRA, end_color=ZEBRA, fill_type="solid"
        )
        fill_summary = PatternFill(
            start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type="solid"
        )
        thin_border = Border(
            left=Side(style="thin", color=BORDER_COLOR),
            right=Side(style="thin", color=BORDER_COLOR),
            top=Side(style="thin", color=BORDER_COLOR),
            bottom=Side(style="thin", color=BORDER_COLOR),
        )

        ws1.merge_cells("A1:F1")
        t_cell = ws1["A1"]
        t_cell.value = "RELATÓRIO GERENCIAL DE ITENS EM ABERTO POR CLIENTE"
        t_cell.font = font_title
        t_cell.fill = fill_title
        t_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws1.row_dimensions[1].height = 30

        clientes_dict = {}
        for item in dados:
          cli = item["cliente"]
          if cli not in clientes_dict:
            clientes_dict[cli] = []
          clientes_dict[cli].append(item)

        row = 3
        for cli_nome, itens_cli in clientes_dict.items():
          cidade = itens_cli[0]["cidade_uf"]
          ws1.merge_cells(
              start_row=row, start_column=1, end_row=row, end_column=6
          )
          c_cli = ws1.cell(
              row=row, column=1, value=f"CLIENTE: {cli_nome}  |  {cidade}"
          )
          c_cli.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
          c_cli.fill = fill_cli
          c_cli.alignment = Alignment(
              horizontal="left", vertical="center", indent=1
          )
          ws1.row_dimensions[row].height = 24

          row += 1
          headers = [
              "Nº Pedido",
              "Cód. Produto",
              "Descrição do Produto",
              "Data Pedido",
              "Prev. Entrega",
              "Qtd. A Separar",
          ]
          for c_idx, h in enumerate(headers, 1):
            cell = ws1.cell(row=row, column=c_idx, value=h)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

          start_item = row + 1
          row += 1

          for idx, item in enumerate(itens_cli):
            ws1.cell(row=row, column=1, value=item["pedido"]).alignment = (
                Alignment(horizontal="center")
            )
            ws1.cell(row=row, column=2, value=item["cod_prod"]).alignment = (
                Alignment(horizontal="center")
            )
            ws1.cell(row=row, column=3, value=item["descricao"]).alignment = (
                Alignment(horizontal="left")
            )
            ws1.cell(row=row, column=4, value=item["data_venda"]).alignment = (
                Alignment(horizontal="center")
            )
            ws1.cell(row=row, column=5, value=item["prev_entrega"]).alignment = (
                Alignment(horizontal="center")
            )
            q = ws1.cell(row=row, column=6, value=item["quantidade"])
            q.alignment = Alignment(horizontal="right")
            q.number_format = "#,##0.00"

            for c in range(1, 7):
              cell = ws1.cell(row=row, column=c)
              cell.font = font_data
              cell.border = thin_border
              if idx % 2 == 1:
                cell.fill = fill_zebra
            row += 1

          # Subtotal
          ws1.merge_cells(
              start_row=row, start_column=1, end_row=row, end_column=5
          )
          tot_label = ws1.cell(
              row=row, column=1, value=f"TOTAL DE ITENS ({cli_nome}):"
          )
          tot_label.font = font_total
          tot_label.alignment = Alignment(horizontal="right", vertical="center")
          tot_val = ws1.cell(
              row=row, column=6, value=f"=SUM(F{start_item}:F{row-1})"
          )
          tot_val.font = font_total
          tot_val.alignment = Alignment(horizontal="right", vertical="center")
          tot_val.number_format = "#,##0.00"

          for c in range(1, 7):
            cell = ws1.cell(row=row, column=c)
            cell.fill = fill_summary
            cell.border = thin_border
          row += 2

        widths = {"A": 15, "B": 15, "C": 48, "D": 15, "E": 15, "F": 18}
        for col_letter, width in widths.items():
          ws1.column_dimensions[col_letter].width = width

        # Aba 2 - Base
        ws2 = wb.create_sheet(title="Itens Extraídos (Base)")
        ws2.append([
            "Arquivo PDF",
            "Pedido",
            "Data Venda",
            "Prev. Entrega",
            "Cliente",
            "Cidade/UF",
            "Frete",
            "Código Produto",
            "Descrição",
            "Qtd A Separar",
        ])
        for d in dados:
          ws2.append([
              d["arquivo"],
              d["pedido"],
              d["data_venda"],
              d["prev_entrega"],
              d["cliente"],
              d["cidade_uf"],
              d["frete"],
              d["cod_prod"],
              d["descricao"],
              d["quantidade"],
          ])

        # Aba 3 - Resumo
        ws3 = wb.create_sheet(title="Resumo Geral (Itens Totais)")
        produtos_dict = {}
        for item in dados:
          chave = (item["cod_prod"], item["descricao"])
          produtos_dict[chave] = produtos_dict.get(chave, 0.0) + item[
              "quantidade"
          ]

        produtos_ordenados = sorted(
            produtos_dict.items(), key=lambda x: x[1], reverse=True
        )
        total_geral_qtd = sum(produtos_dict.values())

        ws3.merge_cells("A1:E1")
        t_res = ws3["A1"]
        t_res.value = "DASHBOARD RESUMO DE PRODUTOS UNIFICADOS EM ABERTO"
        t_res.font = font_title
        t_res.fill = fill_title
        t_res.alignment = Alignment(horizontal="center", vertical="center")

        headers_resumo = [
            "Cód. Produto",
            "Descrição do Produto",
            "Qtd. Total A Separar",
            "% do Volume Total",
            "Impacto Operacional",
        ]
        for col_idx, h in enumerate(headers_resumo, 1):
          cell = ws3.cell(row=3, column=col_idx, value=h)
          cell.font = font_header
          cell.fill = fill_header
          cell.alignment = Alignment(horizontal="center", vertical="center")
          cell.border = thin_border

        row_r = 4
        for idx, ((cod_prod, desc), qtd_tot) in enumerate(produtos_ordenados):
          perc = (qtd_tot / total_geral_qtd) if total_geral_qtd > 0 else 0
          impacto = (
              "Alta Prioridade"
              if perc >= 0.15
              else (
                  "Média Prioridade" if perc >= 0.05 else "Baixa Prioridade"
              )
          )

          ws3.cell(row=row_r, column=1, value=cod_prod).alignment = Alignment(
              horizontal="center"
          )
          ws3.cell(row=row_r, column=2, value=desc).alignment = Alignment(
              horizontal="left"
          )
          ws3.cell(row=row_r, column=3, value=qtd_tot).number_format = (
              "#,##0.00"
          )
          ws3.cell(row=row_r, column=4, value=perc).number_format = "0.00%"
          ws3.cell(row=row_r, column=5, value=impacto).alignment = Alignment(
              horizontal="center"
          )

          for c in range(1, 6):
            cell = ws3.cell(row=row_r, column=c)
            cell.font = font_data
            cell.border = thin_border
            if idx % 2 == 1:
              cell.fill = fill_zebra
          row_r += 1

        wb.save(excel_buffer)
        excel_buffer.seek(0)

        # =========================================================================
        # GERADOR 2: GERAR PDF EXECUTIVO EM MEMÓRIA (BYTESIO)
        # =========================================================================
        pdf_buffer = io.BytesIO()
        doc = SimpleDocTemplate(
            pdf_buffer,
            pagesize=A4,
            leftMargin=1.5 * cm,
            rightMargin=1.5 * cm,
            topMargin=1.5 * cm,
            bottomMargin=1.5 * cm,
        )

        styles = getSampleStyleSheet()

        C_NAVY = colors.HexColor("#1B365D")
        C_SLATE = colors.HexColor("#2C5282")
        C_BLUE_ACCENT = colors.HexColor("#3182CE")
        C_BG_LIGHT = colors.HexColor("#F7FAFC")
        C_BORDER = colors.HexColor("#E2E8F0")

        style_title = ParagraphStyle(
            "DocTitle",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=18,
            textColor=colors.white,
            alignment=1,
        )
        style_subtitle = ParagraphStyle(
            "DocSubTitle",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=10,
            textColor=colors.HexColor("#E2E8F0"),
            alignment=1,
        )
        style_sec_heading = ParagraphStyle(
            "SecHeading",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=12,
            textColor=C_NAVY,
            spaceBefore=12,
            spaceAfter=6,
        )
        style_cell = ParagraphStyle(
            "Cell",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=8,
            leading=10,
            textColor=colors.HexColor("#2D3748"),
        )
        style_cell_center = ParagraphStyle("CellC", parent=style_cell, alignment=1)
        style_cell_right = ParagraphStyle("CellR", parent=style_cell, alignment=2)
        style_cell_header = ParagraphStyle(
            "CellH",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=8,
            leading=10,
            textColor=colors.white,
            alignment=1,
        )

        story = []

        header_data = [
            [
                Paragraph(
                    "RELATÓRIO EXECUTIVO DE PEDIDOS EM ABERTO", style_title
                )
            ],
            [
                Paragraph(
                    f"Consolidado Operacional de Separação • Gerado em:"
                    f" {datetime.now().strftime('%d/%m/%Y às %H:%M')}",
                    style_subtitle,
                )
            ],
        ]
        t_header = Table(header_data, colWidths=[18 * cm])
        t_header.setStyle(
            TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), C_NAVY),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, 0), 12),
                ("BOTTOMPADDING", (0, -1), (-1, -1), 12),
            ])
        )
        story.append(t_header)
        story.append(Spacer(1, 12))

        total_clientes = len(clientes_dict)
        total_pedidos = len(set(d["pedido"] for d in dados))

        kpi_data = [[
            Paragraph(
                f"<b>CLIENTES PENDENTES</b><br/><font size=16"
                f" color='{C_NAVY}'><b>{total_clientes}</b></font>",
                style_cell_center,
            ),
            Paragraph(
                f"<b>PEDIDOS EM ABERTO</b><br/><font size=16"
                f" color='{C_NAVY}'><b>{total_pedidos}</b></font>",
                style_cell_center,
            ),
            Paragraph(
                f"<b>TOTAL DE PEÇAS A SEPARAR</b><br/><font size=16"
                f" color='{C_NAVY}'><b>{total_geral_qtd:,.0f}</b></font>",
                style_cell_center,
            ),
        ]]
        t_kpis = Table(kpi_data, colWidths=[6 * cm, 6 * cm, 6 * cm])
        t_kpis.setStyle(
            TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), C_BG_LIGHT),
                ("BOX", (0, 0), (-1, -1), 1, C_BORDER),
                ("INNERGRID", (0, 0), (-1, -1), 0.5, C_BORDER),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ])
        )
        story.append(t_kpis)
        story.append(Spacer(1, 15))

        story.append(
            Paragraph(
                "1. RESUMO DE PRODUTOS UNIFICADOS (TOP IMPACTO)",
                style_sec_heading,
            )
        )
        story.append(
            HRFlowable(
                width="100%", thickness=1.5, color=C_NAVY, spaceAfter=8
            )
        )

        r_headers = [
            Paragraph("Cód.", style_cell_header),
            Paragraph("Descrição do Produto", style_cell_header),
            Paragraph("Qtd. Total", style_cell_header),
            Paragraph("% Vol.", style_cell_header),
            Paragraph("Prioridade", style_cell_header),
        ]
        r_rows = [r_headers]

        for (cod_p, desc), q_tot in produtos_ordenados:
          p_pct = (q_tot / total_geral_qtd) if total_geral_qtd > 0 else 0
          prio = (
              "Alta"
              if p_pct >= 0.15
              else ("Média" if p_pct >= 0.05 else "Baixa")
          )

          r_rows.append([
              Paragraph(cod_p, style_cell_center),
              Paragraph(desc, style_cell),
              Paragraph(f"{q_tot:,.2f}", style_cell_right),
              Paragraph(f"{p_pct*100:.1f}%", style_cell_right),
              Paragraph(f"<b>{prio}</b>", style_cell_center),
          ])

        t_resumo_pdf = Table(
            r_rows, colWidths=[2.2 * cm, 8.3 * cm, 2.5 * cm, 2 * cm, 3 * cm]
        )
        t_resumo_pdf.setStyle(
            TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), C_SLATE),
                ("GRID", (0, 0), (-1, -1), 0.5, C_BORDER),
                (
                    "ROWBACKGROUNDS",
                    (0, 1),
                    (-1, -1),
                    [colors.white, C_BG_LIGHT],
                ),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ])
        )
        story.append(t_resumo_pdf)
        story.append(Spacer(1, 15))

        story.append(
            Paragraph(
                "2. DETALHAMENTO DOS PEDIDOS EM ABERTO POR CLIENTE",
                style_sec_heading,
            )
        )
        story.append(
            HRFlowable(
                width="100%", thickness=1.5, color=C_NAVY, spaceAfter=8
            )
        )

        for cli_nome, itens_cli in clientes_dict.items():
          bloqueio_cliente = []
          cidade = itens_cli[0]["cidade_uf"]

          t_cli_bar = Table(
              [[Paragraph(f"<b>CLIENTE: {cli_nome}</b>  ({cidade})", style_cell_header)]],
              colWidths=[18 * cm],
          )
          t_cli_bar.setStyle(
              TableStyle([
                  ("BACKGROUND", (0, 0), (-1, -1), C_BLUE_ACCENT),
                  ("TOPPADDING", (0, 0), (-1, -1), 5),
                  ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
              ])
          )
          bloqueio_cliente.append(t_cli_bar)

          c_headers = [
              Paragraph("Pedido", style_cell_header),
              Paragraph("Cód.", style_cell_header),
              Paragraph("Descrição Item", style_cell_header),
              Paragraph("Data", style_cell_header),
              Paragraph("Prev.", style_cell_header),
              Paragraph("Qtd.", style_cell_header),
          ]
          c_rows = [c_headers]

          subtot = sum(i["quantidade"] for i in itens_cli)
          for i in itens_cli:
            c_rows.append([
                Paragraph(i["pedido"], style_cell_center),
                Paragraph(i["cod_prod"], style_cell_center),
                Paragraph(i["descricao"], style_cell),
                Paragraph(i["data_venda"], style_cell_center),
                Paragraph(i["prev_entrega"], style_cell_center),
                Paragraph(f"{i['quantidade']:,.2f}", style_cell_right),
            ])

          c_rows.append([
              Paragraph("<b>TOTAL CLIENTE:</b>", style_cell_right),
              "",
              "",
              "",
              "",
              Paragraph(f"<b>{subtot:,.2f}</b>", style_cell_right),
          ])

          t_cli_items = Table(
              c_rows,
              colWidths=[2.2 * cm, 2 * cm, 7.8 * cm, 2 * cm, 2 * cm, 2 * cm],
          )
          t_cli_items.setStyle(
              TableStyle([
                  ("BACKGROUND", (0, 0), (-1, 0), C_SLATE),
                  ("SPAN", (0, -1), (4, -1)),
                  ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#E2E8F0")),
                  ("GRID", (0, 0), (-1, -1), 0.5, C_BORDER),
                  (
                      "ROWBACKGROUNDS",
                      (0, 1),
                      (-1, -2),
                      [colors.white, C_BG_LIGHT],
                  ),
                  ("TOPPADDING", (0, 0), (-1, -1), 4),
                  ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
              ])
          )
          bloqueio_cliente.append(t_cli_items)
          bloqueio_cliente.append(Spacer(1, 10))

          story.append(KeepTogether(bloqueio_cliente))

        def add_footer(canvas, doc):
          canvas.saveState()
          canvas.setFont("Helvetica", 8)
          canvas.setFillColor(colors.HexColor("#718096"))
          canvas.drawString(
              1.5 * cm,
              1 * cm,
              "Relatório Gerencial de Separação - Uso Interno e Diretoria",
          )
          canvas.drawRightString(
              21 * cm - 1.5 * cm, 1 * cm, f"Página {doc.page}"
          )
          canvas.restoreState()

        doc.build(story, onFirstPage=add_footer, onLaterPages=add_footer)
        pdf_buffer.seek(0)

        st.success(
            "🎉 Relatórios gerados com sucesso a partir dos PDFs enviados!"
        )

        # Exibição de métricas visuais na tela
        col_m1, col_m2, col_m3 = st.columns(3)
        col_m1.metric("Clientes Pendentes", total_clientes)
        col_m2.metric("Pedidos em Aberto", total_pedidos)
        col_m3.metric("Total de Peças", f"{total_geral_qtd:,.0f}")

        st.write("---")

        # Botões de Download direto no navegador
        col_d1, col_d2 = st.columns(2)
        with col_d1:
          st.download_button(
              label="📥 Baixar Relatório Excel (.xlsx)",
              data=excel_buffer,
              file_name="Relatorio_Gerencial_Geral.xlsx",
              mime=(
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              ),
              use_container_width=True,
          )
        with col_d2:
          st.download_button(
              label="📥 Baixar Relatório PDF Executivo (.pdf)",
              data=pdf_buffer,
              file_name="Relatorio_Executivo_Diretoria.pdf",
              mime="application/pdf",
              use_container_width=True,
          )