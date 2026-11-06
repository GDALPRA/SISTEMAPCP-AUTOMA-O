import io
import math
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import streamlit as st

# --- CONSTANTES TÉCNICAS E PESOS (INOX 304) ---
DENSIDADE_INOX = 7.93

# Pesos Unitários dos Blanks de Inox VA
PESO_CH_157219 = 0.80 * 1.24 * 3.90 * DENSIDADE_INOX  # CH 0.80 x 1240 x 3900 (~30.68 kg) - Prateleiras e Bandejas Baixas
PESO_CH_053583 = 0.60 * 1.24 * 2.56 * DENSIDADE_INOX  # CH 0.60 x 1240 x 2560 (~15.10 kg) - Painéis/Tetos
PESO_CH_104818 = 0.60 * 1.24 * 2.00 * DENSIDADE_INOX  # CH 0.60 x 1240 x 2000 (~11.80 kg) - Painéis/Tetos
PESO_CH_153549 = 0.40 * 1.24 * 2.54 * DENSIDADE_INOX  # CH 0.40 x 1240 x 2540 (~9.99 kg) - Bojos
PESO_CH_156082 = 0.40 * 1.24 * 3.78 * DENSIDADE_INOX  # CH 0.40 x 1240 x 3780 (~14.87 kg) - Bojos


def calcular_inox_va(df_input):
  """Aplica as regras completas de consumo de chapas de Aço Inox 304 para expositores VA."""
  detalhamento = []

  total_ch_157219 = 0.0  # 0.80 x 1240 x 3900 (Prateleiras & Bandejas Baixas)
  total_ch_053583 = 0.0  # 0.60 x 1240 x 2560
  total_ch_104818 = 0.0  # 0.60 x 1240 x 2000
  total_ch_153549 = 0.0  # 0.40 x 1240 x 2540 (Bojos)
  total_ch_156082 = 0.0  # 0.40 x 1240 x 3780 (Bojos)

  for idx, row in df_input.iterrows():
    cliente = str(row.get("Name", row.get("Cliente / Projeto", "Cliente")))
    orfa = str(row.get("ORFA", row.get("OP / Pedido", f"OP-{idx+1}")))
    linha = str(row.get("LINHA", "VERTICAL ALTO"))
    alinhamento = str(row.get("ALINHAMENTO", "-"))

    dim_raw = row.get("DIMENSÃO", row.get("Modelo Expositor (VA)", "0"))
    try:
      dim_float = float(str(dim_raw).replace(",", "."))
    except:
      dim_float = 0.0

    qtd = int(row.get("Quantidade (Unid)", 1))
    
    is_vap = "VAP" in alinhamento.upper()

    ch_157219_item = 0.0
    ch_053583_item = 0.0
    ch_104818_item = 0.0
    ch_153549_item = 0.0
    ch_156082_item = 0.0

    # ==========================================
    # 1. CÁLCULO DE PRATELEIRAS E BANDEJAS (Chapa 0.80mm - COD 157219)
    # ==========================================
    if dim_float == 1.87:
      # Prateleiras 1245 (10 un, cabem 12 na chapa 3900)
      prat_1245_tot = 10 * qtd
      ch_157219_item += prat_1245_tot / 12.0
      
      # Prateleiras 610 (10 un, cabem 22 na chapa 3900)
      prat_610_tot = 10 * qtd
      ch_157219_item += prat_610_tot / 22.0

      # Bandejas Baixas para 1.87m (1 unidade por módulo, equivalente a padrão 1.25/1.87 -> 1 bandeja, 6 por chapa)
      bandejas_187_tot = 1 * qtd
      ch_157219_item += bandejas_187_tot / 6.0

    elif is_vap:
      modulos_075 = round(dim_float / 0.75)
      prateleiras_vap_tot = (modulos_075 * 10) * qtd
      ch_157219_item += prateleiras_vap_tot / 18.0

      bandejas_vap_tot = (modulos_075 * 1) * qtd
      ch_157219_item += bandejas_vap_tot / 9.0

    else:
      modulos_125 = round(dim_float / 1.25, 2)
      if dim_float == 2.25:
        modulos_125 = 2.0
      elif dim_float == 3.0:
        modulos_125 = 2.5

      prateleiras_pad_tot = (modulos_125 * 10) * qtd
      ch_157219_item += prateleiras_pad_tot / 12.0

      bandejas_pad_tot = (modulos_125 * 1) * qtd
      ch_157219_item += bandejas_pad_tot / 6.0


    # ==========================================
    # 2. CÁLCULO DE BOJOS (Inox 0,40mm)
    # ==========================================
    if dim_float in [1.25, 1.5, 2.25, 2.5]:
      if dim_float == 1.25:
        ch_153549_item = math.ceil(qtd / 2.0)
      else:
        ch_153549_item = float(qtd)
    elif dim_float in [1.87, 3.0, 3.75]:
      if dim_float == 1.87:
        ch_156082_item = math.ceil(qtd / 2.0)
      else:
        ch_156082_item = float(qtd)


    # ==========================================
    # 3. CÁLCULO DE PAINÉIS DE FUNDO, TETOS E LATERIAIS (Inox 0,60mm)
    # ==========================================
    if dim_float == 1.25:
      paineis_fundo = (1 * 3) * qtd
      ch_053583_item += paineis_fundo / 4.0
      ch_053583_item += qtd / 4.0

    elif dim_float == 1.5:
      paineis_075 = (2 * 3) * qtd
      ch_053583_item += paineis_075 / 6.0
      ch_053583_item += qtd / 2.0

    elif dim_float == 1.87:
      ch_053583_item += (3 * qtd) / 4.0
      paineis_610 = 3 * qtd
      ch_104818_item += paineis_610 / 6.0
      tetos_187 = 1 * qtd
      ch_104818_item += tetos_187 / 2.0

    elif dim_float == 2.25:
      paineis_075 = (3 * 3) * qtd
      ch_053583_item += paineis_075 / 6.0
      ch_053583_item += qtd / 2.0

    elif dim_float == 2.5:
      paineis_fundo = (2 * 3) * qtd
      ch_053583_item += paineis_fundo / 4.0
      ch_053583_item += qtd / 2.0

    elif dim_float == 3.0:
      paineis_075 = (4 * 3) * qtd
      ch_053583_item += paineis_075 / 6.0
      tetos_15_para_3m = 2 * qtd
      ch_053583_item += tetos_15_para_3m / 2.0

    elif dim_float == 3.75:
      paineis_fundo = (3 * 3) * qtd
      ch_053583_item += paineis_fundo / 4.0
      tetos_375 = 2 * qtd
      ch_104818_item += tetos_375 / 2.0

    # Acumuladores Gerais
    total_ch_157219 += ch_157219_item
    total_ch_053583 += ch_053583_item
    total_ch_104818 += ch_104818_item
    total_ch_153549 += ch_153549_item
    total_ch_156082 += ch_156082_item

    peso_item = (
        (ch_157219_item * PESO_CH_157219) +
        (ch_053583_item * PESO_CH_053583) +
        (ch_104818_item * PESO_CH_104818) +
        (ch_153549_item * PESO_CH_153549) +
        (ch_156082_item * PESO_CH_156082)
    )

    detalhamento.append({
        "CLIENTE": cliente,
        "ORFA / OP": orfa,
        "LINHA": linha,
        "DIMENSÃO (m)": dim_raw,
        "ALINHAMENTO": alinhamento,
        "CH 3900 (0.8mm)": round(ch_157219_item, 2),
        "CH 2560 (0.6mm)": round(ch_053583_item, 2),
        "CH 2000 (0.6mm)": round(ch_104818_item, 2),
        "CH 2540 (0.4mm)": round(ch_153549_item, 2),
        "CH 3780 (0.4mm)": round(ch_156082_item, 2),
        "PESO TOTAL (kg)": round(peso_item, 2),
    })

  peso_total_geral = (
      (total_ch_157219 * PESO_CH_157219) +
      (total_ch_053583 * PESO_CH_053583) +
      (total_ch_104818 * PESO_CH_104818) +
      (total_ch_153549 * PESO_CH_153549) +
      (total_ch_156082 * PESO_CH_156082)
  )

  resumo = {
      "ch_3900_un": round(total_ch_157219, 2),
      "peso_3900_kg": total_ch_157219 * PESO_CH_157219,
      "ch_2560_un": round(total_ch_053583, 2),
      "peso_2560_kg": total_ch_053583 * PESO_CH_053583,
      "ch_2000_un": round(total_ch_104818, 2),
      "peso_2000_kg": total_ch_104818 * PESO_CH_104818,
      "ch_2540_un": round(total_ch_153549, 2),
      "peso_2540_kg": total_ch_153549 * PESO_CH_153549,
      "ch_3780_un": round(total_ch_156082, 2),
      "peso_3780_kg": total_ch_156082 * PESO_CH_156082,
      "total_chapas_geral": round(total_ch_157219 + total_ch_053583 + total_ch_104818 + total_ch_153549 + total_ch_156082, 2),
      "peso_total_kg": peso_total_geral,
  }

  return pd.DataFrame(detalhamento), resumo


def gerar_excel_estilo_tubos(df_detalhe, resumo):
  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = "Relatório Inox"
  ws.views.sheetView[0].showGridLines = True

  NAVY = "1F497D"
  WHITE = "FFFFFF"
  LIGHT_BLUE = "D9E1F2"
  ZEBRA = "F2F5F9"
  BORDER_GRAY = "D9D9D9"

  font_title = Font(name="Calibri", size=14, bold=True, color=NAVY)
  font_subtitle = Font(name="Calibri", size=10, italic=True, color="595959")
  font_header = Font(name="Calibri", size=10, bold=True, color=WHITE)
  font_section = Font(name="Calibri", size=11, bold=True, color=NAVY)
  font_body = Font(name="Calibri", size=10)
  font_bold = Font(name="Calibri", size=10, bold=True)

  fill_header = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
  fill_zebra = PatternFill(start_color=ZEBRA, end_color=ZEBRA, fill_type="solid")
  fill_kpi = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type="solid")

  thin_border = Border(
      left=Side(style="thin", color=BORDER_GRAY),
      right=Side(style="thin", color=BORDER_GRAY),
      top=Side(style="thin", color=BORDER_GRAY),
      bottom=Side(style="thin", color=BORDER_GRAY),
  )

  thick_border = Border(
      left=Side(style="thin", color=BORDER_GRAY),
      right=Side(style="thin", color=BORDER_GRAY),
      top=Side(style="thin", color=BORDER_GRAY),
      bottom=Side(style="medium", color=NAVY),
  )

  ws.cell(row=1, column=1, value="RELATÓRIO EXECUTIVO - PRODUÇÃO DE CHAPAS DE AÇO INOX (EXPOSITORES VA)").font = font_title
  ws.cell(row=2, column=1, value="PCP & Engenharia - Visão Estratégica para Diretoria e Compras").font = font_subtitle

  ws.cell(row=4, column=1, value="INDICADORES GERAIS DE INOX 304").font = font_section
  
  kpi_headers = ["TOTAL DE CHAPAS (UN)", "CH 3900 (0.8mm)", "CH 2560 (0.6mm)", "CH 2000 (0.6mm)", "CH 2540 (0.4mm)", "CH 3780 (0.4mm)", "PESO TOTAL (kg)"]
  kpi_values = [
      resumo["total_chapas_geral"],
      resumo["ch_3900_un"],
      resumo["ch_2560_un"],
      resumo["ch_2000_un"],
      resumo["ch_2540_un"],
      resumo["ch_3780_un"],
      round(resumo["peso_total_kg"], 2)
  ]

  for c_idx, h in enumerate(kpi_headers, 1):
    cell = ws.cell(row=5, column=c_idx, value=h)
    cell.font = font_header
    cell.fill = fill_header
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = thin_border

  for c_idx, val in enumerate(kpi_values, 1):
    cell = ws.cell(row=6, column=c_idx, value=val)
    cell.font = font_bold
    cell.fill = fill_kpi
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thin_border

  ws.cell(row=9, column=1, value="1. RESUMO DE CONSUMO POR BLANK DE INOX").font = font_section
  
  resumo_headers = ["CÓDIGO", "DESCRIÇÃO DO BLANK DE INOX", "QUANTIDADE TOTAL (UN)", "PESO TOTAL (kg)"]
  for c_idx, h in enumerate(resumo_headers, 1):
    cell = ws.cell(row=10, column=c_idx, value=h)
    cell.font = font_header
    cell.fill = fill_header
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thin_border

  resumo_data = [
      ["157219", "CH 0.80 x 1240 x 3900 INOX 304 ESC PVC LASER (Prateleiras/Bandejas)", resumo["ch_3900_un"], round(resumo["peso_3900_kg"], 2)],
      ["053583", "CH 0.60 x 1240 x 2560 INOX 304 ESC PVC LASER", resumo["ch_2560_un"], round(resumo["peso_2560_kg"], 2)],
      ["104818", "CH 0.60 x 1240 x 2000 INOX 304 ESC PVC LASER", resumo["ch_2000_un"], round(resumo["peso_2000_kg"], 2)],
      ["153549", "CH 0.40 x 1240 x 2540 INOX 304 2B (Bojos)", resumo["ch_2540_un"], round(resumo["peso_2540_kg"], 2)],
      ["156082", "CH 0.40 x 1240 x 3780 INOX 304 2B FOSCO (Bojos)", resumo["ch_3780_un"], round(resumo["peso_3780_kg"], 2)],
  ]

  for r_idx, row_vals in enumerate(resumo_data, 11):
    for c_idx, val in enumerate(row_vals, 1):
      cell = ws.cell(row=r_idx, column=c_idx, value=val)
      cell.font = font_body
      cell.border = thin_border
      if c_idx in [1, 3, 4]:
        cell.alignment = Alignment(horizontal="center", vertical="center")
      else:
        cell.alignment = Alignment(horizontal="left", vertical="center")

  start_row = 17
  ws.cell(row=start_row - 1, column=1, value="2. DEMANDA DETALHADA DE INOX POR CLIENTE E EXPOSITOR").font = font_section

  detalhe_headers = list(df_detalhe.columns)
  for c_idx, h in enumerate(detalhe_headers, 1):
    cell = ws.cell(row=start_row, column=c_idx, value=h)
    cell.font = font_header
    cell.fill = fill_header
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = thin_border

  current_row = start_row + 1
  for row_data in df_detalhe.values:
    for c_idx, val in enumerate(row_data, 1):
      cell = ws.cell(row=current_row, column=c_idx, value=val)
      cell.font = font_body
      cell.border = thin_border
      if c_idx in [2, 3, 4, 5, 6, 7, 8, 9, 10, 11]:
        cell.alignment = Alignment(horizontal="center", vertical="center")
      else:
        cell.alignment = Alignment(horizontal="left", vertical="center")
      
      if current_row % 2 == 0:
        cell.fill = fill_zebra
    current_row += 1

  for c_idx in range(1, len(detalhe_headers) + 1):
    cell = ws.cell(row=current_row, column=c_idx)
    cell.border = thick_border
    if c_idx == 1:
      cell.value = "TOTAL GERAL"
      cell.font = font_bold
      cell.alignment = Alignment(horizontal="left", vertical="center")
    elif c_idx in [6, 7, 8, 9, 10, 11]:
      col_let = get_column_letter(c_idx)
      cell.value = f"=SUM({col_let}{start_row+1}:{col_let}{current_row-1})"
      cell.font = font_bold
      cell.alignment = Alignment(horizontal="center", vertical="center")

  for col in ws.columns:
    max_len = max(len(str(cell.value or "")) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws.column_dimensions[col_letter].width = max(max_len + 5, 18)

  output = io.BytesIO()
  wb.save(output)
  return output.getvalue()


st.title("🛡️ Módulo de Chapas de Aço Inox - Expositores VA")
st.markdown("Cálculo automatizado de blanks de Inox 304 (0,40mm, 0,60mm e 0,80mm) incluindo Bojos, Painéis, Prateleiras (1245 e 610) e Bandejas Baixas.")

st.markdown("---")

uploaded_file = st.file_uploader("Envie a planilha de Montagem em Excel (.xlsx)", type=["xlsx"], key="inox_uploader")

if uploaded_file:
  df_pedidos = pd.read_excel(uploaded_file)
  st.subheader("📋 Dados de Montagem Carregados")
  st.dataframe(df_pedidos.head(10), use_container_width=True)

  if st.button("🚀 Processar Cálculo de Aço Inox Completo", type="primary"):
    df_resultado, resumo = calcular_inox_va(df_pedidos)

    st.markdown("---")
    st.subheader("📊 Indicadores Totais de Inox Necessários (Com Prateleiras de 1,87m)")

    col1, col2, col3, col4, col5, col6 = st.columns(6)
    col1.metric("CH 3900 (0.8mm)", f"{resumo['ch_3900_un']} un")
    col2.metric("CH 2560 (0.6mm)", f"{resumo['ch_2560_un']} un")
    col3.metric("CH 2000 (0.6mm)", f"{resumo['ch_2000_un']} un")
    col4.metric("CH 2540 (0.4mm)", f"{resumo['ch_2540_un']} un")
    col5.metric("CH 3780 (0.4mm)", f"{resumo['ch_3780_un']} un")
    col6.metric("Peso Total Inox", f"{resumo['peso_total_kg']:.2f} kg")

    st.markdown("---")
    st.subheader("📄 Detalhamento Completo por Item")
    st.dataframe(df_resultado, use_container_width=True)

    excel_data = gerar_excel_estilo_tubos(df_resultado, resumo)

    st.download_button(
        label="📥 Baixar Relatório Executivo de Aço Inox em Excel",
        data=excel_data,
        file_name="Relatorio_Executivo_Inox_VA.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )