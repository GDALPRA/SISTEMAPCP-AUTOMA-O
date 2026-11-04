import io
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import streamlit as st

# --- CONSTANTES TÉCNICAS E CÁLCULO DE PESO ---
# Densidade do Aço Cristal 0,43mm: ~3.3759 kg/m²
DENSIDADE_M2 = 3.3759

PESO_CHAPA_2550 = (1.20 * 2.55) * DENSIDADE_M2  # ~10.33 kg/chapa
PESO_CHAPA_3780 = (1.20 * 3.78) * DENSIDADE_M2  # ~15.31 kg/chapa
PESO_CHAPA_4000 = (1.20 * 4.00) * DENSIDADE_M2  # ~16.20 kg/chapa


def calcular_chapas_va(df_input):
  """Aplica as regras de consumo de chapas cristal (0,43mm) com base no layout real da planilha de Montagem."""
  total_ch_2550 = 0
  total_ch_3780 = 0
  total_ch_4000 = 0

  detalhamento = []

  for idx, row in df_input.iterrows():
    cliente = str(row.get("Name", row.get("Cliente / Projeto", "Cliente")))
    orfa = str(row.get("ORFA", row.get("OP / Pedido", f"OP-{idx+1}")))
    linha = str(row.get("LINHA", "VERTICAL ALTO"))
    alinhamento = str(row.get("ALINHAMENTO", "-"))

    dim_raw = row.get("DIMENSÃO", row.get("Modelo Expositor (VA)", "0"))
    dim_str = (
        str(dim_raw)
        .replace(",", ".")
        .strip()
        .rstrip("0")
        .rstrip(".")
        if "." in str(dim_raw)
        else str(dim_raw).strip()
    )

    qtd = int(row.get("Quantidade (Unid)", 1))

    ch_2550 = 0
    ch_3780 = 0
    ch_4000 = 0
    observacao = ""

    # REGRAS DE NEGÓCIO PCP (CHAPAS CRISTAL 0,43mm):
    if dim_str in ["1.25", "1.5", "2.25", "2.5"]:
      ch_2550 = 7 * qtd
      observacao = f"Uso de CH 2550 para {dim_str}m. Sobram retalhos no corte."
    elif dim_str in ["1.87", "3"]:
      ch_3780 = 7 * qtd
      observacao = f"Uso integral de CH 3780 para {dim_str}m."
    elif dim_str in ["3.75"]:
      ch_4000 = 1 * qtd
      ch_3780 = 6 * qtd
      observacao = "Mistura Padrão 3.75m: 1x CH 4000 (Bojo) + 6x CH 3780."
    else:
      observacao = f"⚠️ Dimensão '{dim_raw}' não reconhecida na regra VA."

    total_ch_2550 += ch_2550
    total_ch_3780 += ch_3780
    total_ch_4000 += ch_4000

    detalhamento.append({
        "ORFA / Pedido": orfa,
        "Cliente": cliente,
        "Linha": linha,
        "Dimensão (m)": dim_raw,
        "Alinhamento": alinhamento,
        "CH 2550 (0,43x1200x2550)": ch_2550,
        "CH 3780 (0,43x1200x3780)": ch_3780,
        "CH 4000 (0,43x1200x4000)": ch_4000,
        "Total Chapas": ch_2550 + ch_3780 + ch_4000,
        "Observação Técnica": observacao,
    })

  peso_2550 = total_ch_2550 * PESO_CHAPA_2550
  peso_3780 = total_ch_3780 * PESO_CHAPA_3780
  peso_4000 = total_ch_4000 * PESO_CHAPA_4000
  peso_total_geral = peso_2550 + peso_3780 + peso_4000

  resumo_chapas = {
      "total_ch_2550": total_ch_2550,
      "total_ch_3780": total_ch_3780,
      "total_ch_4000": total_ch_4000,
      "total_chapas_geral": total_ch_2550 + total_ch_3780 + total_ch_4000,
      "peso_2550_kg": peso_2550,
      "peso_3780_kg": peso_3780,
      "peso_4000_kg": peso_4000,
      "peso_total_kg": peso_total_geral,
  }

  return pd.DataFrame(detalhamento), resumo_chapas


def gerar_excel_formatado(df_detalhe, resumo):
  """Gera um arquivo Excel corporativo formatado com openpyxl."""
  wb = openpyxl.Workbook()

  # Cores e Estilos
  NAVY = "1F497D"
  WHITE = "FFFFFF"
  ZEBRA = "F2F5F9"
  BORDER_GRAY = "D9D9D9"

  font_header = Font(name="Calibri", size=11, bold=True, color=WHITE)
  font_body = Font(name="Calibri", size=11)
  font_title = Font(name="Calibri", size=13, bold=True, color=NAVY)

  fill_header = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
  fill_zebra = PatternFill(start_color=ZEBRA, end_color=ZEBRA, fill_type="solid")

  thin_border = Border(
      left=Side(style="thin", color=BORDER_GRAY),
      right=Side(style="thin", color=BORDER_GRAY),
      top=Side(style="thin", color=BORDER_GRAY),
      bottom=Side(style="thin", color=BORDER_GRAY),
  )

  # --- ABA 1: DETALHAMENTO ---
  ws1 = wb.active
  ws1.title = "Detalhamento_Chapas"
  ws1.views.sheetView[0].showGridLines = True

  # Título
  ws1.cell(
      row=1,
      column=1,
      value="RELATÓRIO DE CÁLCULO DE CHAPAS CRISTAL (EXPOSITORES VA)",
  ).font = font_title
  ws1.row_dimensions[1].height = 25

  # Cabeçalhos
  headers = list(df_detalhe.columns)
  for col_idx, h in enumerate(headers, 1):
    cell = ws1.cell(row=3, column=col_idx, value=h)
    cell.font = font_header
    cell.fill = fill_header
    cell.alignment = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )
    cell.border = thin_border
  ws1.row_dimensions[3].height = 28

  # Dados
  for row_idx, row_data in enumerate(df_detalhe.values, 4):
    ws1.row_dimensions[row_idx].height = 20
    for col_idx, val in enumerate(row_data, 1):
      cell = ws1.cell(row=row_idx, column=col_idx, value=val)
      cell.font = font_body
      cell.border = thin_border

      # Alinhamentos específicos
      if col_idx in [1, 3, 4, 5, 6, 7, 8, 9]:
        cell.alignment = Alignment(horizontal="center", vertical="center")
      else:
        cell.alignment = Alignment(horizontal="left", vertical="center")

      # Efeito Zebra
      if row_idx % 2 == 0:
        cell.fill = fill_zebra

  # Largura automática
  for col in ws1.columns:
    max_len = max(len(str(cell.value or "")) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws1.column_dimensions[col_letter].width = max(max_len + 4, 16)

  # --- ABA 2: RESUMO EXECUTIVO ---
  ws2 = wb.create_sheet(title="Resumo_Material_PCP")
  ws2.views.sheetView[0].showGridLines = True

  ws2.cell(
      row=1, column=1, value="RESUMO GERAL DE COMPRAS E ALMOXARIFADO"
  ).font = font_title
  ws2.row_dimensions[1].height = 25

  resumo_headers = [
      "Tipo de Chapa / Insumo",
      "Quantidade Total (Unid)",
      "Peso Estimado Total (Kg)",
  ]
  for col_idx, h in enumerate(resumo_headers, 1):
    cell = ws2.cell(row=3, column=col_idx, value=h)
    cell.font = font_header
    cell.fill = fill_header
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thin_border
  ws2.row_dimensions[3].height = 28

  resumo_rows = [
      [
          "Chapa 0,43 x 1200 x 2550 mm",
          resumo["total_ch_2550"],
          round(resumo["peso_2550_kg"], 2),
      ],
      [
          "Chapa 0,43 x 1200 x 3780 mm",
          resumo["total_ch_3780"],
          round(resumo["peso_3780_kg"], 2),
      ],
      [
          "Chapa 0,43 x 1200 x 4000 mm",
          resumo["total_ch_4000"],
          round(resumo["peso_4000_kg"], 2),
      ],
      [
          "TOTAL GERAL",
          resumo["total_chapas_geral"],
          round(resumo["peso_total_kg"], 2),
      ],
  ]

  for r_idx, r_data in enumerate(resumo_rows, 4):
    ws2.row_dimensions[r_idx].height = 22
    for c_idx, val in enumerate(r_data, 1):
      cell = ws2.cell(row=r_idx, column=c_idx, value=val)
      cell.border = thin_border
      if r_idx == 7:  # Linha do Total Geral em destaque
        cell.font = Font(name="Calibri", size=11, bold=True)
      else:
        cell.font = font_body

      if c_idx > 1:
        cell.alignment = Alignment(horizontal="center", vertical="center")
      else:
        cell.alignment = Alignment(horizontal="left", vertical="center")

  for col in ws2.columns:
    max_len = max(len(str(cell.value or "")) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws2.column_dimensions[col_letter].width = max(max_len + 6, 25)

  output = io.BytesIO()
  wb.save(output)
  return output.getvalue()


# --- INTERFACE EXECUTIVA STREAMLIT ---
st.title("📦 Módulo de Chapas Cristal - Expositores VA")
st.markdown(
    "Cálculo automatizado de blanks cristal 0,43mm (2550, 3780 e 4000) baseado"
    " na Planilha de Montagem Oficial."
)

st.markdown("---")

uploaded_file = st.file_uploader(
    "Envie a planilha de Montagem em Excel (.xlsx)", type=["xlsx"]
)

if uploaded_file:
  df_pedidos = pd.read_excel(uploaded_file)
  st.subheader("📋 Dados de Montagem Carregados")
  st.dataframe(df_pedidos.head(10), use_container_width=True)

  if st.button("🚀 Processar Cálculo de Chapas", type="primary"):
    df_resultado, resumo = calcular_chapas_va(df_pedidos)

    # --- METRICAS EXECUTIVAS ---
    st.markdown("---")
    st.subheader("📊 Indicadores Totais de Materiais Necessários")

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("CH 2550 mm", f"{resumo['total_ch_2550']} un")
    col2.metric("CH 3780 mm", f"{resumo['total_ch_3780']} un")
    col3.metric("CH 4000 mm", f"{resumo['total_ch_4000']} un")
    col4.metric(
        "Total de Chapas", f"{resumo['total_chapas_geral']} un", delta="Blanks"
    )
    col5.metric(
        "Peso Total Material",
        f"{resumo['peso_total_kg']:.2f} kg",
        delta="Aço 0,43mm",
    )

    st.markdown("---")
    st.subheader("📄 Detalhamento de Insumos por Item da Montagem")
    st.dataframe(df_resultado, use_container_width=True)

    # GERADOR DE EXCEL FORMATADO PROFISSIONAL
    excel_data = gerar_excel_formatado(df_resultado, resumo)

    st.download_button(
        label="📥 Baixar Relatório Executivo Formatado em Excel",
        data=excel_data,
        file_name="Relatorio_PCP_Chapas_Cristal_VA.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )