import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

# --- ESTILIZAÇÃO DO CABEÇALHO DA PÁGINA ---
st.markdown(
    """
    <div style="background-color: #1F497D; color: white; padding: 20px; border-radius: 8px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
        <h2 style="margin: 0; font-weight: 700;">🚪 CÁLCULO DE PORTAS E FRAMES (VERTICAL ALTO)</h2>
        <p style="margin: 5px 0 0 0; opacity: 0.9; font-size: 14px;">Painel de Engenharia & PCP - Relatório Gerencial Executivo</p>
    </div>
""",
    unsafe_allow_html=True,
)

st.write("")

uploaded_file = st.file_uploader(
    "Arraste ou selecione o seu ficheiro Excel de montagem de expositores (.xlsx, .xls)",
    type=["xlsx", "xls", "csv"],
)

# TABELA OFICIAL DE CÓDIGOS DE PORTAS
CODIGOS_PORTAS = {
    ("PRETA", "SERIGRAFADA", "DIR"): "153828",
    ("PRETA", "SERIGRAFADA", "ESQ"): "153829",
    ("PRETA", "TRANSPARENTE", "DIR"): "153830",
    ("PRETA", "TRANSPARENTE", "ESQ"): "153831",
    ("CINZA", "SERIGRAFADA", "DIR"): "153318",
    ("CINZA", "SERIGRAFADA", "ESQ"): "153319",
    ("CINZA", "TRANSPARENTE", "DIR"): "153832",
    ("CINZA", "TRANSPARENTE", "ESQ"): "153833",
}

# TABELA OFICIAL DE CÓDIGOS DE FRAMES (por cor e dimensão em mm)
CODIGOS_FRAMES = {
    ("CINZA", 1250): "154183",
    ("CINZA", 1875): "154678",
    ("CINZA", 2500): "154157",
    ("CINZA", 3750): "154158",
    ("PRETA", 1250): "154184",
    ("PRETA", 1875): "162726",
    ("PRETA", 2500): "153826",
    ("PRETA", 3750): "153827",
}


def converter_dimensao_para_mm(valor_dim):
  try:
    val = float(str(valor_dim).replace(",", "."))
  except:
    return 1250

  if 1.85 <= val <= 1.89:
    return 1875

  if val <= 10.0:
    mm = round(val * 1000)
  else:
    mm = round(val)

  chaves_validas = [1250, 1875, 2500, 3750]
  return min(chaves_validas, key=lambda x: abs(x - mm))


if uploaded_file is not None:
  try:
    if uploaded_file.name.endswith(".xls"):
      df = pd.read_excel(uploaded_file, engine="xlrd")
    else:
      df = pd.read_excel(uploaded_file, engine="openpyxl")
  except Exception:
    df = pd.read_excel(uploaded_file)

  df.columns = [str(col).strip().upper() for col in df.columns]

  # Filtrar apenas a linha VERTICAL ALTO
  if "LINHA" in df.columns:
    df_va = df[
        df["LINHA"].astype(str).str.upper().str.contains("VERTICAL ALTO|VA")
    ].copy()
  else:
    df_va = df.copy()

  if df_va.empty:
    st.warning(
        "Nenhum expositor da linha 'VERTICAL ALTO' foi encontrado na planilha."
    )
  else:
    lista_detalhada = []
    lista_resumo_portas = []
    lista_resumo_frames = []

    for _, row in df_va.iterrows():
      cliente = str(row.get("NAME", row.get("CLIENTE", "N/A"))).strip()
      orfa = str(row.get("ORFA", row.get("ORFA / OP", "N/A"))).strip()
      alinhamento = str(row.get("ALINHAMENTO", "N/A")).strip()
      serigrafia = str(row.get("SERIGRAFIA", "PRETA")).strip().upper()
      qtd_exp = int(row.get("QTD", row.get("QUANTIDADE", 1)))

      try:
        dim_val = float(
            str(row.get("DIMENSÃO", row.get("DIMENSAO", 1.25))).replace(
                ",", "."
            )
        )
      except:
        dim_val = 1.25

      dim_mm = converter_dimensao_para_mm(dim_val)
      is_pvt = "PVT" in alinhamento.upper()
      tipo_porta_desc = "TRANSPARENTE" if is_pvt else "SERIGRAFADA"

      # Regra de portas ESQ / DIR
      if dim_mm == 1875:
        qtd_esq = 1 * qtd_exp
        qtd_dir = 2 * qtd_exp
      else:
        pares = round(dim_mm / 1250)
        qtd_esq = pares * qtd_exp
        qtd_dir = pares * qtd_exp

      cod_esq = CODIGOS_PORTAS.get(
          (serigrafia, tipo_porta_desc, "ESQ"), "S/COD"
      )
      cod_dir = CODIGOS_PORTAS.get(
          (serigrafia, tipo_porta_desc, "DIR"), "S/COD"
      )

      # Código e Descrição do Frame
      cod_frame = CODIGOS_FRAMES.get((serigrafia, dim_mm), "S/COD")
      desc_frame_txt = (
          f"CJ FRAME RESFR. {dim_mm} {serigrafia}"
          if serigrafia == "CINZA"
          else f"CJ FRAME RESFR {dim_mm} PRETO"
      )
      qtd_frames = 1 * qtd_exp

      # Adicionar ao resumo consolidado de portas
      lista_resumo_portas.append({
          "CÓDIGO": cod_esq,
          "MODELO DE PORTA": (
              f"PORTA {serigrafia} {tipo_porta_desc} ESQUERDA"
          ),
          "QTD PORTAS (UN)": qtd_esq,
      })
      lista_resumo_portas.append({
          "CÓDIGO": cod_dir,
          "MODELO DE PORTA": f"PORTA {serigrafia} {tipo_porta_desc} DIREITA",
          "QTD PORTAS (UN)": qtd_dir,
      })

      # Adicionar ao resumo consolidado de frames
      lista_resumo_frames.append({
          "CÓDIGO": cod_frame,
          "MODELO DE FRAME": desc_frame_txt,
          "QTD FRAMES (UN)": qtd_frames,
      })

      # Detalhado por item
      lista_detalhada.append({
          "CLIENTE": cliente,
          "ORFA / OP": orfa,
          "ALINHAMENTO": alinhamento,
          "DIMENSÃO (m)": dim_val,
          "QTD EXP": qtd_exp,
          "CÓD. ESQ": cod_esq,
          "QTD ESQ": qtd_esq,
          "CÓD. DIR": cod_dir,
          "QTD DIR": qtd_dir,
          "TOTAL PORTAS": qtd_esq + qtd_dir,
          "CÓD. FRAME": cod_frame,
          "FRAME DESC": desc_frame_txt,
          "QTD FRAMES": qtd_frames,
      })

    df_itens = pd.DataFrame(lista_detalhada)
    df_portas_temp = pd.DataFrame(lista_resumo_portas)
    df_frames_temp = pd.DataFrame(lista_resumo_frames)

    # 1. Resumo Consolidado de Portas
    df_portas_resumo = (
        df_portas_temp.groupby(["CÓDIGO", "MODELO DE PORTA"], as_index=False)
        .agg({"QTD PORTAS (UN)": "sum"})
        .sort_values(by="CÓDIGO")
    )

    # 2. Resumo Consolidado de Frames por Código
    df_frames_resumo = (
        df_frames_temp.groupby(["CÓDIGO", "MODELO DE FRAME"], as_index=False)
        .agg({"QTD FRAMES (UN)": "sum"})
        .sort_values(by="CÓDIGO")
    )

    total_expositores = df_itens["QTD EXP"].sum()
    total_portas_geral = df_itens["TOTAL PORTAS"].sum()
    total_frames_geral = df_itens["QTD FRAMES"].sum()

    # --- PAINEL VISUAL NA TELA ---
    st.write("---")
    col1, col2, col3 = st.columns(3)
    col1.metric("TOTAL DE EXPOSITORES (VA)", total_expositores)
    col2.metric("TOTAL DE PORTAS (UN)", total_portas_geral)
    col3.metric("TOTAL DE FRAMES (UN)", total_frames_geral)

    st.write("")
    st.write(
        "### 1. Resumo Consolidado de Portas por Código e Lado (ESQ / DIR)"
    )
    st.dataframe(df_portas_resumo, use_container_width=True, hide_index=True)

    st.write("### 2. Resumo Consolidado de Frames por Código e Medida")
    st.dataframe(df_frames_resumo, use_container_width=True, hide_index=True)

    st.write("### 3. Demanda Detalhada por Cliente e Alinhamento")
    st.dataframe(df_itens, use_container_width=True, hide_index=True)


    # --- FUNÇÃO DE GERAÇÃO DO EXCEL EXECUTIVO ---
    def gerar_excel_portas(
        tot_exp, tot_portas, tot_frames, df_portas, df_frames, df_detalhado
    ):
      output_bytes = io.BytesIO()
      wb = openpyxl.Workbook()
      ws = wb.active
      ws.title = "Relatório Portas & Frames"
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

      ws.merge_cells("A1:M1")
      ws.merge_cells("A2:M2")

      c1 = ws["A1"]
      c1.value = (
          "RELATÓRIO EXECUTIVO - CÁLCULO DE PORTAS E FRAMES (VERTICAL ALTO)"
      )
      c1.font = font_main_title
      c1.fill = fill_header
      c1.alignment = Alignment(horizontal="center", vertical="center")

      c2 = ws["A2"]
      c2.value = (
          "Mapeamento Oficial por Código de Produto (Portas e Frames) | PCP &"
          " Diretoria"
      )
      c2.font = font_subtitle
      c2.fill = fill_header
      c2.alignment = Alignment(horizontal="center", vertical="center")

      ws.row_dimensions[1].height = 25
      ws.row_dimensions[2].height = 18

      kpis = [
          ("TOTAL EXPOSITORES", tot_exp, "#,##0", "A", "C"),
          ("TOTAL PORTAS (UN)", tot_portas, "#,##0", "D", "G"),
          ("TOTAL FRAMES (UN)", tot_frames, "#,##0", "H", "M"),
      ]

      for titulo, valor, formato, col_ini, col_fim in kpis:
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
          (
              "1. RESUMO CONSOLIDADO DE PORTAS POR CÓDIGO E MODELO",
              df_portas,
          ),
          ("2. RESUMO CONSOLIDADO DE FRAMES POR CÓDIGO", df_frames),
          ("3. DEMANDA DETALHADA POR CLIENTE E ALINHAMENTO", df_detalhado),
      ]

      for titulo_secao, df_data in blocos:
        ws.cell(row=current_row, column=1, value=titulo_secao).font = (
            font_section
        )
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
              c_tot.value = f"=SUM({col_letter}{start_data_row}:{col_letter}{end_data_row})"
              c_tot.alignment = Alignment(
                  horizontal="right", vertical="center"
              )
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

    excel_portas = gerar_excel_portas(
        total_expositores,
        total_portas_geral,
        total_frames_geral,
        df_portas_resumo,
        df_frames_resumo,
        df_itens,
    )

    st.write("")
    st.download_button(
        label="📥 Descarregar Relatório Executivo de Portas & Frames (.xlsx)",
        data=excel_portas,
        file_name="Relatorio_Executivo_Portas_Frames.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )