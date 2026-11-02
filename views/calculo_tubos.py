import io
import math
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

# --- ESTILIZAÇÃO DO CABEÇALHO DA PÁGINA ---
st.markdown(
    """
    <div style="background-color: #1F497D; color: white; padding: 20px; border-radius: 8px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
        <h2 style="margin: 0; font-weight: 700;">🧪 CÁLCULO DE TUBOS E PERFIS 60x30 (CHASSIS)</h2>
        <p style="margin: 5px 0 0 0; opacity: 0.9; font-size: 14px;">Painel de Engenharia & PCP - Relatório Gerencial Executivo</p>
    </div>
""",
    unsafe_allow_html=True,
)

st.write("")

uploaded_file = st.file_uploader(
    "Arraste ou selecione o seu ficheiro Excel de produção de chassis (.xlsx, .xls)",
    type=["xlsx", "xls", "csv"],
)

if uploaded_file is not None:
  try:
    if uploaded_file.name.endswith(".xls"):
      df = pd.read_excel(uploaded_file, engine="xlrd")
    else:
      df = pd.read_excel(uploaded_file, engine="openpyxl")
  except Exception:
    df = pd.read_excel(uploaded_file)

  df.columns = [str(col).strip().upper() for col in df.columns]

  # --- CONFIGURAÇÕES DE PESO E DIMENSÕES ---
  COMPRIMENTO_BARRA_PE = 5200.0
  QTD_PES_POR_BARRA = 4

  # Tabela de Pesos Reais Unitários por Viga (com o desconto de 10mm já aplicado nas medidas)
  # C1240=3kg, C1490=3kg, C1668=4kg, C1865=4kg, C2240=5kg, C2490=5kg, C2990=6kg, C3740=8kg
  # Para os pés (barra 5200mm / C5200), o peso unitário da barra é 10kg
  PESO_BARRA_PE_KG = 10.0

  # Regras de Pés, Vigas e Peso Unitário por Dimensão de Chassi (em milímetros)
  REGRAS_CHASSIS = {
      1250: {
          "viga_desconto": 10,
          "qtd_vigas": 2,
          "qtd_pes": 2,
          "peso_unit_viga": 3.0,
      },
      1500: {
          "viga_desconto": 10,
          "qtd_vigas": 2,
          "qtd_pes": 2,
          "peso_unit_viga": 3.0,
      },
      1680: {
          "viga_desconto": 10,
          "qtd_vigas": 2,
          "qtd_pes": 3,
          "peso_unit_viga": 4.0,
      },  # 1.76m -> viga 1670 (~1668) = 4kg
      1875: {
          "viga_desconto": 10,
          "qtd_vigas": 2,
          "qtd_pes": 3,
          "peso_unit_viga": 4.0,
      },  # 1.87m -> viga 1865 = 4kg
      2250: {
          "viga_desconto": 10,
          "qtd_vigas": 2,
          "qtd_pes": 3,
          "peso_unit_viga": 5.0,
      },
      2500: {
          "viga_desconto": 10,
          "qtd_vigas": 2,
          "qtd_pes": 3,
          "peso_unit_viga": 5.0,
      },
      3000: {
          "viga_desconto": 10,
          "qtd_vigas": 2,
          "qtd_pes": 4,
          "peso_unit_viga": 6.0,
      },
      3750: {
          "viga_desconto": 10,
          "qtd_vigas": 2,
          "qtd_pes": 4,
          "peso_unit_viga": 8.0,
      },
  }


  # Função de conversão exata das dimensões do Excel (em metros) para milímetros
  def converter_dimensao_para_mm(valor_dim):
    try:
      val = float(str(valor_dim).replace(",", "."))
    except:
      return 1250

    # Casos especiais informados
    if 1.74 <= val <= 1.78:
      return 1680  # 1.76 corresponde a 1680mm
    if 1.85 <= val <= 1.89:
      return 1875  # 1.87 corresponde a 1875mm

    if val <= 10.0:
      mm = round(val * 1000)
    else:
      mm = round(val)

    chaves_validas = list(REGRAS_CHASSIS.keys())
    return min(chaves_validas, key=lambda x: abs(x - mm))


  lista_processada = []

  for _, row in df.iterrows():
    cliente = str(row.get("NAME", row.get("CLIENTE", "N/A"))).strip()
    orfa = str(row.get("ORFA", row.get("ORFA / OP", "N/A"))).strip()
    qtd_chassis = int(row.get("QTD", row.get("QUANTIDADE", 1)))

    dim_val = row.get("DIMENSÃO", row.get("DIMENSAO", 1.25))
    dimensao_chassi = converter_dimensao_para_mm(dim_val)

    cfg = REGRAS_CHASSIS[dimensao_chassi]
    comprimento_viga_mm = dimensao_chassi - cfg["viga_desconto"]

    total_vigas_nec = cfg["qtd_vigas"] * qtd_chassis
    total_pes_nec = cfg["qtd_pes"] * qtd_chassis

    lista_processada.append({
        "CLIENTE": cliente,
        "ORFA / OP": orfa,
        "CHASSI (mm)": dimensao_chassi,
        "QTD CHASSIS": qtd_chassis,
        "COMPR. VIGA (mm)": comprimento_viga_mm,
        "PESO UNIT. VIGA (kg)": cfg["peso_unit_viga"],
        "QTD VIGAS": total_vigas_nec,
        "QTD PÉS": total_pes_nec,
    })

  df_itens = pd.DataFrame(lista_processada)

  # --- CONSOLIDAÇÃO PARA FABRICAÇÃO ---
  df_vigas_resumo = (
      df_itens.groupby(
          ["CHASSI (mm)", "COMPR. VIGA (mm)", "PESO UNIT. VIGA (kg)"],
          as_index=False,
      )
      .agg({"QTD VIGAS": "sum"})
      .rename(columns={"QTD VIGAS": "TOTAL DE VIGAS"})
  )
  df_vigas_resumo["METRAGEM TOTAL (m)"] = (
      df_vigas_resumo["COMPR. VIGA (mm)"] * df_vigas_resumo["TOTAL DE VIGAS"]
  ) / 1000.0
  # Peso total calculado usando o peso unitário real da viga
  df_vigas_resumo["PESO TOTAL (kg)"] = (
      df_vigas_resumo["TOTAL DE VIGAS"] * df_vigas_resumo["PESO UNIT. VIGA (kg)"]
  )

  total_pes_geral = df_itens["QTD PÉS"].sum()
  barras_pes_necessarias = math.ceil(total_pes_geral / QTD_PES_POR_BARRA)
  peso_total_pes = barras_pes_necessarias * PESO_BARRA_PE_KG

  pead_data = [{
      "DESCRIÇÃO DO ITEM": "Barras para Pés de Chassi (5200mm - rende 4 pés)",
      "COMPRIMENTO BARRA (mm)": COMPRIMENTO_BARRA_PE,
      "QTD BARRAS": int(barras_pes_necessarias),
      "PESO UNIT. BARRA (kg)": PESO_BARRA_PE_KG,
      "PESO TOTAL (kg)": round(peso_total_pes, 2),
  }]
  df_pes_resumo = pd.DataFrame(pead_data)

  total_barras_geral = int(
      barras_pes_necessarias
      + math.ceil(df_vigas_resumo["TOTAL DE VIGAS"].sum() / 4)
  )
  peso_geral_tubos = (
      df_vigas_resumo["PESO TOTAL (kg)"].sum() + peso_total_pes
  )

  # --- PAINEL VISUAL NA TELA ---
  st.write("---")
  col1, col2, col3, col4 = st.columns(4)
  col1.metric("TOTAL DE CHASSIS", int(df_itens["QTD CHASSIS"].sum()))
  col2.metric("TOTAL DE PÉS (UN)", total_pes_geral)
  col3.metric("BARRAS TOTAIS (EST.)", total_barras_geral)
  col4.metric(
      "PESO TOTAL TUBOS",
      f"{peso_geral_tubos:,.2f} kg"
      .replace(",", "X")
      .replace(".", ",")
      .replace("X", "."),
  )

  st.write("")
  st.write("### 1. Separação de Barras para Pés (5200 mm)")
  st.dataframe(df_pes_resumo, use_container_width=True, hide_index=True)

  st.write("### 2. Resumo de Vigas de Fixação por Dimensão de Chassi")
  st.dataframe(df_vigas_resumo, use_container_width=True, hide_index=True)

  st.write("### 3. Demanda Detalhada de Chassis por Cliente e ORFA")
  st.dataframe(df_itens, use_container_width=True, hide_index=True)


  # --- FUNÇÃO DE GERAÇÃO DO EXCEL EXECUTIVO ---
  def gerar_excel_tubos(
      tot_chassis,
      tot_pes_un,
      tot_barras_tot,
      peso_geral,
      df_pes_in,
      df_vigas_in,
      df_itens_in,
  ):
    output_bytes = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Relatório Tubos"
    ws.views.sheetView[0].showGridLines = True

    cor_header = "1F497D"
    cor_zebrado = "F8F9FA"

    font_main_title = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    font_subtitle = Font(name="Calibri", size=10, italic=True, color="FFFFFF")
    font_kpi_title = Font(name="Calibri", size=9, bold=True, color="555555")
    font_kpi_val = Font(name="Calibri", size=14, bold=True, color="1F497D")
    font_section = Font(name="Calibri", size=11, bold=True, color="1F497D")
    font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_body = Font(name="Calibri", size=10)
    font_total = Font(name="Calibri", size=10, bold=True)

    border_thin = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    fill_header = PatternFill(
        start_color=cor_header, end_color=cor_header, fill_type="solid"
    )

    # Cabeçalho Principal
    ws.merge_cells("A1:H1")
    ws.merge_cells("A2:H2")

    c1 = ws["A1"]
    c1.value = (
        "RELATÓRIO EXECUTIVO - PRODUÇÃO DE TUBOS E PERFIS 60x30 (CHASSIS)"
    )
    c1.font = font_main_title
    c1.fill = fill_header
    c1.alignment = Alignment(horizontal="center", vertical="center")

    c2 = ws["A2"]
    c2.value = (
        "Pesos Unitários Reais por Viga | PCP & Engenharia - Visão Estratégica"
        " para Diretoria"
    )
    c2.font = font_subtitle
    c2.fill = fill_header
    c2.alignment = Alignment(horizontal="center", vertical="center")

    ws.row_dimensions[1].height = 25
    ws.row_dimensions[2].height = 18

    # KPIs Superiores
    kpis = [
        ("TOTAL DE CHASSIS", tot_chassis, "#,##0", "A", "B"),
        ("TOTAL DE PÉS (UN)", tot_pes_un, "#,##0", "C", "C"),
        ("BARRAS TOTAIS (EST.)", tot_barras_tot, "#,##0", "D", "E"),
        ("PESO TOTAL (kg)", peso_geral, '#,##0.00" kg"', "F", "H"),
    ]

    for titulo, valor, formato, col_ini, col_fim in kpis:
      if col_ini != col_fim:
        ws.merge_cells(f"{col_ini}4:{col_fim}4")
        ws.merge_cells(f"{col_ini}5:{col_fim}5")
      ct = ws[f"{col_ini}4"]
      ct.value = titulo
      ct.font = font_kpi_title
      ct.alignment = Alignment(horizontal="center", vertical="center")

      cv = ws[f"{col_ini}5"]
      cv.value = valor
      cv.font = font_kpi_val
      cv.number_format = formato
      cv.alignment = Alignment(horizontal="center", vertical="center")

    ws.row_dimensions[4].height = 16
    ws.row_dimensions[5].height = 24

    current_row = 7
    blocos = [
        ("1. SEPARAÇÃO E BARRAS PARA PÉS DE CHASSIS", df_pes_in),
        ("2. RESUMO DE VIGAS DE FIXAÇÃO POR DIMENSÃO", df_vigas_in),
        ("3. DEMANDA DETALHADA DE CHASSIS POR CLIENTE", df_itens_in),
    ]

    for titulo_secao, df_data in blocos:
      ws.cell(row=current_row, column=1, value=titulo_secao).font = font_section
      ws.row_dimensions[current_row].height = 22
      current_row += 1

      for col_idx, col_name in enumerate(df_data.columns, 1):
        c = ws.cell(row=current_row, column=col_idx, value=col_name)
        c.font = font_header
        c.fill = fill_header
        c.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        c.border = border_thin
      ws.row_dimensions[current_row].height = 22
      start_data_row = current_row + 1
      current_row += 1

      for row_values in df_data.values:
        ws.row_dimensions[current_row].height = 18
        is_even = current_row % 2 == 0

        for col_idx, val in enumerate(row_values, 1):
          c = ws.cell(row=current_row, column=col_idx, value=val)
          c.font = font_body
          c.border = border_thin

          if is_even:
            c.fill = PatternFill(
                start_color=cor_zebrado,
                end_color=cor_zebrado,
                fill_type="solid",
            )

          if isinstance(val, float):
            c.number_format = "#,##0.00"
            c.alignment = Alignment(horizontal="right", vertical="center")
          elif isinstance(val, int):
            c.number_format = "#,##0"
            c.alignment = Alignment(horizontal="right", vertical="center")
          else:
            c.alignment = Alignment(horizontal="left", vertical="center")
        current_row += 1

      end_data_row = current_row - 1

      ws.row_dimensions[current_row].height = 20
      c_tot_lbl = ws.cell(row=current_row, column=1, value="TOTAL GERAL")
      c_tot_lbl.font = font_total
      c_tot_lbl.border = border_thin
      c_tot_lbl.alignment = Alignment(horizontal="left", vertical="center")

      for col_idx, col_name in enumerate(df_data.columns, 1):
        if col_idx > 1:
          c_tot = ws.cell(row=current_row, column=col_idx)
          c_tot.font = font_total
          c_tot.border = border_thin
          col_letter = get_column_letter(col_idx)

          if df_data[col_name].dtype in ["int64", "float64"]:
            c_tot.value = (
                f"=SUM({col_letter}{start_data_row}:{col_letter}{end_data_row})"
            )
            c_tot.alignment = Alignment(horizontal="right", vertical="center")
            if df_data[col_name].dtype == "float64":
              c_tot.number_format = "#,##0.00"
            else:
              c_tot.number_format = "#,##0"

      current_row += 3

    for col in ws.columns:
      max_len = 0
      col_letter = get_column_letter(col[0].column)
      for cell in col:
        v_str = str(cell.value or "")
        if len(v_str) > max_len:
          max_len = len(v_str)
      ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

    wb.save(output_bytes)
    output_bytes.seek(0)
    return output_bytes.getvalue()

  excel_tubos = gerar_excel_tubos(
      int(df_itens["QTD CHASSIS"].sum()),
      total_pes_geral,
      total_barras_geral,
      peso_geral_tubos,
      df_pes_resumo,
      df_vigas_resumo,
      df_itens,
  )

  st.write("")
  st.download_button(
      label="📥 Descarregar Relatório Executivo de Tubos (.xlsx)",
      data=excel_tubos,
      file_name="Relatorio_Executivo_Tubos_Chassis.xlsx",
      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      use_container_width=True,
  )