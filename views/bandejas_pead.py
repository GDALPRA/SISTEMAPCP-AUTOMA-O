import io
import math
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

# --- ESTILIZAÇÃO DO CABEÇALHO DA PÁGINA ---
st.markdown(
    """
    <div style="background-color: #1F497D; color: white; padding: 20px; border-radius: 8px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
        <h2 style="margin: 0; font-weight: 700;">📦 CÁLCULO DE BANDEJAS ABS E MATÉRIA-PRIMA PEAD</h2>
        <p style="margin: 5px 0 0 0; opacity: 0.9; font-size: 14px;">Painel de Engenharia & PCP - Relatório Gerencial Executivo</p>
    </div>
""",
    unsafe_allow_html=True,
)

st.write("")

uploaded_file = st.file_uploader(
    "Arraste ou selecione o seu ficheiro Excel de produção (.xlsx, .xls)",
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

  FATORES_BANDEJAS = {
      "VCA2P": {"C790": 4, "C530": 8, "C410": 0},
      "VRA2P": {"C790": 4, "C530": 0, "C410": 8},
      "VRA1P": {"C790": 4, "C530": 0, "C410": 4},
      "VRAG": {"C790": 4, "C530": 0, "C410": 0},
      "VCAG": {"C790": 4, "C530": 0, "C410": 0},
  }
  MODELOS_VALIDOS = list(FATORES_BANDEJAS.keys())

  def extrair_modulos_por_modelo(texto_alinhamento, dimensao_total):
    texto = str(texto_alinhamento).upper().replace(",", ".")
    modulos_por_modelo = {}
    encontrados = re.findall(r"([A-Z0-9]+)\s*\(([\d\.]+)\)", texto)

    if encontrados:
      for modelo, tamanho_str in encontrados:
        if modelo in MODELOS_VALIDOS:
          try:
            metragem = float(tamanho_str)
            mods = metragem / 1.25
            modulos_por_modelo[modelo] = (
                modulos_por_modelo.get(modelo, 0.0) + mods
            )
          except Exception:
            pass

    if not modulos_por_modelo:
      for modelo in MODELOS_VALIDOS:
        if modelo in texto:
          modulos_por_modelo[modelo] = (
              dimensao_total / 1.25 if dimensao_total > 0 else 0
          )
          break
    return modulos_por_modelo

  lista_detalhada = []
  for _, row in df.iterrows():
    cliente = str(row.get("NAME", row.get("CLIENTE", "N/A"))).strip()
    orfa = str(row.get("ORFA", row.get("ORFA / OP", "N/A"))).strip()
    linha = str(row.get("LINHA", "N/A")).strip()
    dimensao_val = row.get(
        "DIMENSÃO", row.get("DIMENSAO", row.get("TAMANHO", 0))
    )
    alinhamento = (
        str(row.get("ALINHAMENTO", row.get("MODELO", ""))).strip().upper()
    )

    try:
      dimensao = float(dimensao_val)
    except Exception:
      dimensao = 0.0

    modulos_dict = extrair_modulos_por_modelo(alinhamento, dimensao)

    for modelo, mod_qtd in modulos_dict.items():
      fatores = FATORES_BANDEJAS.get(
          modelo, {"C790": 0, "C530": 0, "C410": 0}
      )
      c790 = int(round(mod_qtd * fatores["C790"]))
      c530 = int(round(mod_qtd * fatores["C530"]))
      c410 = int(round(mod_qtd * fatores["C410"]))
      total_bandejas = c790 + c530 + c410

      lista_detalhada.append({
          "CLIENTE": cliente,
          "ORFA / OP": orfa,
          "MODELO": modelo,
          "MÓDULOS (1,25m)": mod_qtd,
          "BANDEJA C790": c790,
          "BANDEJA C530": c530,
          "BANDEJA C410": c410,
          "TOTAL ABS": total_bandejas,
      })

  df_detalhado = pd.DataFrame(lista_detalhada)

  df_por_cliente = df_detalhado.groupby(
      ["CLIENTE", "ORFA / OP", "MODELO"], as_index=False
  ).agg({
      "MÓDULOS (1,25m)": "sum",
      "BANDEJA C790": "sum",
      "BANDEJA C530": "sum",
      "BANDEJA C410": "sum",
      "TOTAL ABS": "sum",
  })

  df_geral_modelo = df_detalhado.groupby("MODELO", as_index=False).agg({
      "MÓDULOS (1,25m)": "sum",
      "BANDEJA C790": "sum",
      "BANDEJA C530": "sum",
      "BANDEJA C410": "sum",
      "TOTAL ABS": "sum",
  })
  df_geral_modelo.rename(
      columns={"MODELO": "MODELO DO EXPOSITOR"}, inplace=True
  )

  tot_c790 = df_geral_modelo["BANDEJA C790"].sum()
  tot_c530 = df_geral_modelo["BANDEJA C530"].sum()
  tot_c410 = df_geral_modelo["BANDEJA C410"].sum()

  placas_c790 = math.ceil(tot_c790 / 1.0)
  placas_c530 = math.ceil(tot_c530 / 1.0)
  placas_c410 = math.ceil(tot_c410 / 2.0)

  pead_data = [
      {
          "MEDIDA DA PLACA": "1000 × 510 mm",
          "ESPESSURA": "5,5 mm",
          "TIPO DE BANDEJA": "L290 × C790 (REF. MAIOR)",
          "QTD PLACAS": int(placas_c790),
          "PESO UNIT. (kg)": 2.50,
          "PESO TOTAL (kg)": round(placas_c790 * 2.50, 2),
      },
      {
          "MEDIDA DA PLACA": "750 × 500 mm",
          "ESPESSURA": "5,5 mm",
          "TIPO DE BANDEJA": "L290 × C530 (REF. MÉDIA)",
          "QTD PLACAS": int(placas_c530),
          "PESO UNIT. (kg)": 1.84,
          "PESO TOTAL (kg)": round(placas_c530 * 1.84, 2),
      },
      {
          "MEDIDA DA PLACA": "780 × 600 mm",
          "ESPESSURA": "5,5 mm",
          "TIPO DE BANDEJA": "L290 × C410 (REF. MENOR - 2 un/placa)",
          "QTD PLACAS": int(placas_c410),
          "PESO UNIT. (kg)": 2.29,
          "PESO TOTAL (kg)": round(placas_c410 * 2.29, 2),
      },
  ]
  df_pead = pd.DataFrame(pead_data)

  total_peso_pead = df_pead["PESO TOTAL (kg)"].sum()
  total_clientes = df_por_cliente["CLIENTE"].nunique()
  total_modulos = df_detalhado["MÓDULOS (1,25m)"].sum()
  total_bandejas_abs = int(df_detalhado["TOTAL ABS"].sum())

  # --- PAINEL VISUAL NA TELA ---
  st.write("---")
  col1, col2, col3, col4 = st.columns(4)
  col1.metric("TOTAL DE CLIENTES", total_clientes)
  col2.metric("MÓDULOS PROGRAMADOS", f"{total_modulos:.2f}")
  col3.metric("TOTAL BANDEJAS ABS", f"{total_bandejas_abs:,}".replace(",", "."))
  col4.metric(
      "PESO PEAD (kg)",
      f"{total_peso_pead:,.2f} kg"
      .replace(",", "X")
      .replace(".", ",")
      .replace("X", "."),
  )

  st.write("")
  st.write("### 1. SEPARAÇÃO DE MATÉRIA-PRIMA PEAD (PLACAS)")
  st.dataframe(df_pead, use_container_width=True, hide_index=True)

  st.write("### 2. RESUMO DE BANDEJAS POR MODELO DE EXPOSITOR")
  st.dataframe(df_geral_modelo, use_container_width=True, hide_index=True)

  st.write("### 3. DEMANDA DETALHADA POR CLIENTE E ORFA")
  st.dataframe(df_por_cliente, use_container_width=True, hide_index=True)

  # --- FUNÇÃO DE GERAÇÃO DO EXCEL EXECUTIVO IDÊNTICO AO EXEMPLO ---
  def gerar_excel_exemplo_diretoria(
      total_cli,
      total_mods,
      total_bnds,
      total_peso,
      df_pead_in,
      df_modelo_in,
      df_cliente_in,
  ):
    output_bytes = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Relatório Executivo"
    ws.views.sheetView[0].showGridLines = True

    cor_header = "1F497D"
    cor_kpi_bg = "F2F5F9"
    cor_zebrado = "F8F9FA"

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
    border_kpi = Border(
        left=Side(style="medium", color="1F497D"),
        right=Side(style="medium", color="1F497D"),
        top=Side(style="medium", color="1F497D"),
        bottom=Side(style="medium", color="1F497D"),
    )

    # --- BLOCO DE KPIS SUPERIORES (Linhas 1 e 2) ---
    kpis = [
        ("TOTAL DE CLIENTES", total_cli, "#,##0", "A", "B"),
        ("MÓDULOS PROGRAMADOS", total_mods, "#,##0.00", "C", "D"),
        ("TOTAL BANDEJAS ABS", total_bnds, "#,##0", "E", "F"),
        ("PESO PEAD (kg)", total_peso, '#,##0.00" kg"', "G", "H"),
    ]

    for titulo, valor, formato, col_ini, col_fim in kpis:
      ws.merge_cells(f"{col_ini}1:{col_fim}1")
      ws.merge_cells(f"{col_ini}2:{col_fim}2")

      c_t = ws[f"{col_ini}1"]
      c_t.value = titulo
      c_t.font = font_kpi_title
      c_t.alignment = Alignment(horizontal="center", vertical="center")

      c_v = ws[f"{col_ini}2"]
      c_v.value = valor
      c_v.font = font_kpi_val
      c_v.number_format = formato
      c_v.alignment = Alignment(horizontal="center", vertical="center")

      ws.row_dimensions[1].height = 16
      ws.row_dimensions[2].height = 22

    current_row = 4

    blocos = [
        ("1. SEPARAÇÃO DE MATÉRIA-PRIMA PEAD (PLACAS)", df_pead_in),
        ("2. RESUMO DE BANDEJAS POR MODELO DE EXPOSITOR", df_modelo_in),
        ("3. DEMANDA DETALHADA POR CLIENTE E ORFA", df_cliente_in),
    ]

    for titulo_secao, df_data in blocos:
      # Título da Seção
      ws.cell(row=current_row, column=1, value=titulo_secao).font = font_section
      ws.row_dimensions[current_row].height = 22
      current_row += 1

      # Cabeçalhos
      for col_idx, col_name in enumerate(df_data.columns, 1):
        c = ws.cell(row=current_row, column=col_idx, value=col_name)
        c.font = font_header
        c.fill = PatternFill(
            start_color=cor_header, end_color=cor_header, fill_type="solid"
        )
        c.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        c.border = border_thin
      ws.row_dimensions[current_row].height = 22
      header_row = current_row
      current_row += 1

      # Dados
      start_data_row = current_row
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

      # Linha de Totais
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

    # Ajuste automático de largura de colunas
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

  processed_data = gerar_excel_exemplo_diretoria(
      total_clientes,
      total_modulos,
      total_bandejas_abs,
      total_peso_pead,
      df_pead,
      df_geral_modelo,
      df_por_cliente,
  )

  st.write("")
  st.download_button(
      label="📥 Descarregar Relatório Executivo Padrão Diretoria (.xlsx)",
      data=processed_data,
      file_name="Relatorio_Executivo_Bandejas_Com_PEAD.xlsx",
      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      use_container_width=True,
  )