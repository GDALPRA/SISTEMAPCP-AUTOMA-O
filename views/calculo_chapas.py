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
  detalhamento = []
  
  total_2550 = 0
  total_3780 = 0
  total_4000 = 0

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

    ch_2550 = 0
    ch_3780 = 0
    ch_4000 = 0

    # REGRAS DE NEGÓCIO PCP (CHAPAS CRISTAL 0,43mm):
    if dim_float in [1.25, 1.5, 2.25, 2.5]:
      ch_2550 = 7 * qtd
    elif dim_float in [1.87, 3.0]:
      ch_3780 = 7 * qtd
    elif dim_float in [3.75]:
      ch_4000 = 1 * qtd
      ch_3780 = 6 * qtd

    total_2550 += ch_2550
    total_3780 += ch_3780
    total_4000 += ch_4000

    detalhamento.append({
        "CLIENTE": cliente,
        "ORFA / OP": orfa,
        "LINHA": linha,
        "DIMENSÃO (m)": dim_raw,
        "ALINHAMENTO": alinhamento,
        "CH 2550 (UN)": ch_2550,
        "CH 3780 (UN)": ch_3780,
        "CH 4000 (UN)": ch_4000,
        "TOTAL CHAPAS": ch_2550 + ch_3780 + ch_4000,
        "PESO TOTAL (kg)": round((ch_2550 * PESO_CHAPA_2550) + (ch_3780 * PESO_CHAPA_3780) + (ch_4000 * PESO_CHAPA_4000), 2)
    })

  peso_2550_total = total_2550 * PESO_CHAPA_2550
  peso_3780_total = total_3780 * PESO_CHAPA_3780
  peso_4000_total = total_4000 * PESO_CHAPA_4000
  peso_geral = peso_2550_total + peso_3780_total + peso_4000_total
  total_geral_chapas = total_2550 + total_3780 + total_4000

  resumo = {
      "total_ch_2550": total_2550,
      "peso_2550_kg": peso_2550_total,
      "total_ch_3780": total_3780,
      "peso_3780_kg": peso_3780_total,
      "total_ch_4000": total_4000,
      "peso_peso_4000_kg": peso_4000_total,
      "total_chapas_geral": total_geral_chapas,
      "peso_total_kg": peso_geral
  }

  return pd.DataFrame(detalhamento), resumo


def gerar_excel_estilo_tubos(df_detalhe, resumo):
  """Gera o relatório em Excel seguindo exatamente o padrão visual do relatório de tubos."""
  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = "Relatório Chapas"
  ws.views.sheetView[0].showGridLines = True

  # Cores e Estilos Corporativos
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

  # DEFINIÇÃO CORRETA DA BORDA GROSSA QUE CAUSAVA O ERRO
  thick_border = Border(
      left=Side(style="thin", color=BORDER_GRAY),
      right=Side(style="thin", color=BORDER_GRAY),
      top=Side(style="thin", color=BORDER_GRAY),
      bottom=Side(style="medium", color=NAVY),
  )

  # --- CABEÇALHO DO RELATÓRIO ---
  ws.cell(row=1, column=1, value="RELATÓRIO EXECUTIVO - PRODUÇÃO DE CHAPAS CRISTAL (EXPOSITORES VA)").font = font_title
  ws.cell(row=2, column=1, value="PCP & Engenharia - Visão Estratégica para Diretoria e Compras").font = font_subtitle

  # --- BLOCO DE KPIS / TOTAIS GERAIS ---
  ws.cell(row=4, column=1, value="INDICADORES GERAIS DE MATERIAIS").font = font_section
  
  kpi_headers = ["TOTAL DE CHAPAS (UN)", "CH 2550 (UN)", "CH 3780 (UN)", "CH 4000 (UN)", "PESO TOTAL MATERIAL (kg)"]
  kpi_values = [
      resumo["total_chapas_geral"],
      resumo["total_ch_2550"],
      resumo["total_ch_3780"],
      resumo["total_ch_4000"],
      round(resumo["peso_total_kg"], 2)
  ]

  for c_idx, h in enumerate(kpi_headers, 1):
    cell = ws.cell(row=5, column=c_idx, value=h)
    cell.font = font_header
    cell.fill = fill_header
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thin_border

  for c_idx, val in enumerate(kpi_values, 1):
    cell = ws.cell(row=6, column=c_idx, value=val)
    cell.font = font_bold
    cell.fill = fill_kpi
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thin_border

  # --- SEÇÃO 1: RESUMO POR TIPO DE CHAPA ---
  ws.cell(row=9, column=1, value="1. RESUMO DE CONSUMO POR TIPO DE BLANK (0,43mm)").font = font_section
  
  resumo_headers = ["TIPO DE CHAPA", "DIMENSÕES (mm)", "QUANTIDADE TOTAL (UN)", "PESO TOTAL (kg)"]
  for c_idx, h in enumerate(resumo_headers, 1):
    cell = ws.cell(row=10, column=c_idx, value=h)
    cell.font = font_header
    cell.fill = fill_header
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thin_border

  resumo_data = [
      ["Chapa 2550", "0,43 x 1200 x 2550", resumo["total_ch_2550"], round(resumo["peso_2550_kg"], 2)],
      ["Chapa 3780", "0,43 x 1200 x 3780", resumo["total_ch_3780"], round(resumo["peso_3780_kg"], 2)],
      ["Chapa 4000", "0,43 x 1200 x 4000", resumo["total_ch_4000"], round(resumo["peso_peso_4000_kg"], 2)],
  ]

  for r_idx, row_vals in enumerate(resumo_data, 11):
    for c_idx, val in enumerate(row_vals, 1):
      cell = ws.cell(row=r_idx, column=c_idx, value=val)
      cell.font = font_body
      cell.border = thin_border
      if c_idx in [3, 4]:
        cell.alignment = Alignment(horizontal="center", vertical="center")
      else:
        cell.alignment = Alignment(horizontal="left", vertical="center")

  # --- SEÇÃO 2: DEMANDA DETALHADA POR CLIENTE / ORFA ---
  start_row = 16
  ws.cell(row=start_row - 1, column=1, value="2. DEMANDA DETALHADA DE CHAPAS POR CLIENTE E EXPOSITOR").font = font_section

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
      if c_idx in [2, 3, 4, 5, 6, 7, 8, 9, 10]:
        cell.alignment = Alignment(horizontal="center", vertical="center")
      else:
        cell.alignment = Alignment(horizontal="left", vertical="center")
      
      if current_row % 2 == 0:
        cell.fill = fill_zebra
    current_row += 1

  # Linha de Total Geral ao fim da tabela detalhada
  for c_idx in range(1, len(detalhe_headers) + 1):
    cell = ws.cell(row=current_row, column=c_idx)
    cell.border = thick_border
    if c_idx == 1:
      cell.value = "TOTAL GERAL"
      cell.font = font_bold
      cell.alignment = Alignment(horizontal="left", vertical="center")
    elif c_idx == 6:
      cell.value = f"=SUM(F{start_row+1}:F{current_row-1})"
      cell.font = font_bold
      cell.alignment = Alignment(horizontal="center", vertical="center")
    elif c_idx == 7:
      cell.value = f"=SUM(G{start_row+1}:G{current_row-1})"
      cell.font = font_bold
      cell.alignment = Alignment(horizontal="center", vertical="center")
    elif c_idx == 8:
      cell.value = f"=SUM(H{start_row+1}:H{current_row-1})"
      cell.font = font_bold
      cell.alignment = Alignment(horizontal="center", vertical="center")
    elif c_idx == 9:
      cell.value = f"=SUM(I{start_row+1}:I{current_row-1})"
      cell.font = font_bold
      cell.alignment = Alignment(horizontal="center", vertical="center")
    elif c_idx == 10:
      cell.value = f"=SUM(J{start_row+1}:J{current_row-1})"
      cell.font = font_bold
      cell.alignment = Alignment(horizontal="center", vertical="center")

  # Ajuste automático de largura das colunas
  for col in ws.columns:
    max_len = max(len(str(cell.value or "")) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws.column_dimensions[col_letter].width = max(max_len + 5, 18)

  output = io.BytesIO()
  wb.save(output)
  return output.getvalue()


# --- INTERFACE EXECUTIVA STREAMLIT ---
st.title("📦 Módulo de Chapas Cristal - Expositores VA")
st.markdown("Cálculo automatizado de blanks cristal 0,43mm baseado na Planilha de Montagem Oficial.")

st.markdown("---")

uploaded_file = st.file_uploader("Envie a planilha de Montagem em Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
  df_pedidos = pd.read_excel(uploaded_file)
  st.subheader("📋 Dados de Montagem Carregados")
  st.dataframe(df_pedidos.head(10), use_container_width=True)

  if st.button("🚀 Processar Cálculo de Chapas", type="primary"):
    df_resultado, resumo = calcular_chapas_va(df_pedidos)

    # --- MÉTRICAS EXECUTIVAS ---
    st.markdown("---")
    st.subheader("📊 Indicadores Totais de Materiais Necessários")

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("CH 2550 mm", f"{resumo['total_ch_2550']} un")
    col2.metric("CH 3780 mm", f"{resumo['total_ch_3780']} un")
    col3.metric("CH 4000 mm", f"{resumo['total_ch_4000']} un")
    col4.metric("Total de Chapas", f"{resumo['total_chapas_geral']} un", delta="Blanks")
    col5.metric("Peso Total Material", f"{resumo['peso_total_kg']:.2f} kg", delta="Aço 0,43mm")

    st.markdown("---")
    st.subheader("📄 Detalhamento de Insumos por Item da Montagem")
    st.dataframe(df_resultado, use_container_width=True)

    # GERADOR DE EXCEL FORMATADO ESTILO TUBOS
    excel_data = gerar_excel_estilo_tubos(df_resultado, resumo)

    st.download_button(
        label="📥 Baixar Relatório Executivo Formatado (Estilo Tubos) em Excel",
        data=excel_data,
        file_name="Relatorio_Executivo_Chapas_Cristal.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )