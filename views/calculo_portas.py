import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

# --- ESTILIZAÇÃO DO CABEÇALHO DA PÁGINA ---
st.markdown(
    """
    <div style="margin-bottom: 25px;">
        <h2 style="margin: 0; color: #FFFFFF; font-size: 24px; font-weight: 700; display: flex; align-items: center; gap: 10px;">
            🚪 CÁLCULO DE PORTAS, FRAMES & EXPOSITORES ABERTOS
        </h2>
        <p style="margin: 6px 0 0 0; color: #94A3B8; font-size: 14px;">
            Painel de Engenharia & PCP — Relatório Gerencial Executivo
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.write("")

uploaded_file = st.file_uploader(
    "Arraste ou selecione o seu ficheiro Excel de montagem de expositores (.xlsx, .xls)",
    type=["xlsx", "xls", "csv"],
)

# TABELA OFICIAL DE CÓDIGOS DE PORTAS VA (Standard com portas: VALP, VAHP, VACP, etc.)
CODIGOS_PORTAS_VA = {
    ("PRETA", "SERIGRAFADA", "DIR"): "153828",
    ("PRETA", "SERIGRAFADA", "ESQ"): "153829",
    ("PRETA", "TRANSPARENTE", "DIR"): "153830",
    ("PRETA", "TRANSPARENTE", "ESQ"): "153831",
    ("CINZA", "SERIGRAFADA", "DIR"): "153318",
    ("CINZA", "SERIGRAFADA", "ESQ"): "153319",
    ("CINZA", "TRANSPARENTE", "DIR"): "153832",
    ("CINZA", "TRANSPARENTE", "ESQ"): "153833",
}

# TABELA OFICIAL DE CÓDIGOS DE PORTAS VAP (Congelados - Apenas Direitas)
CODIGOS_PORTAS_VAP = {
    "CINZA": "154222",  # PORTA CONG DIR 730X1714 CINZA
    "PRETA": "154221",  # PORTA CONG DIR 730X1714 PRETA
}

# TABELA OFICIAL DE CÓDIGOS DE FRAMES / QUADROS VA
CODIGOS_FRAMES_VA = {
    ("CINZA", 1250): "154183",
    ("CINZA", 1875): "154678",
    ("CINZA", 2500): "154157",
    ("CINZA", 3750): "154158",
    ("PRETA", 1250): "154184",
    ("PRETA", 1875): "162726",
    ("PRETA", 2500): "153826",
    ("PRETA", 3750): "153827",
}

# TABELA OFICIAL DE CÓDIGOS DE QUADROS VAP (Congelados)
CODIGOS_QUADROS_VAP = {
    ("CINZA", 1500): "154220",  # QUADRO 2P CONG 1500X1735 CINZA
    ("CINZA", 2250): "154219",  # QUADRO 3P CONG 2250X1735 CINZA
    ("CINZA", 3000): "154218",  # QUADRO 4P CONG 3000X1735 CINZA
    ("CINZA", 3750): "154217",  # QUADRO 5P CONG 3750X1735 CINZA
    ("PRETA", 1500): "154216",  # QUADRO 2P CONG 1500X1735 PRETO
    ("PRETA", 2250): "154215",  # QUADRO 3P CONG 2250X1735 PRETO
    ("PRETA", 3000): "154214",  # QUADRO 4P CONG 3000X1735 PRETO
    ("PRETA", 3750): "154213",  # QUADRO 5P CONG 3750X1735 PRETO
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

  chaves_validas = [1250, 1500, 1875, 2250, 2500, 3000, 3750]
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

  # Filtrar apenas linhas que pertencem à série VA / VAP
  if "LINHA" in df.columns or "ALINHAMENTO" in df.columns:
    df_filtrado = df[
        df.apply(
            lambda r: any(
                k in str(r.get("LINHA", "")) or k in str(r.get("ALINHAMENTO", ""))
                for k in ["VA", "VAP"]
            ),
            axis=1,
        )
    ].copy()
  else:
    df_filtrado = df.copy()

  if df_filtrado.empty:
    st.warning("Nenhum expositor VA ou VAP foi encontrado na planilha.")
  else:
    lista_detalhada = []
    lista_resumo_portas = []
    lista_resumo_frames = []
    lista_abertos = []

    for _, row in df_filtrado.iterrows():
      cliente = str(row.get("NAME", row.get("CLIENTE", "N/A"))).strip()
      orfa = str(row.get("ORFA", row.get("ORFA / OP", "N/A"))).strip()
      alinhamento = str(row.get("ALINHAMENTO", "N/A")).strip()
      alinh_upper = alinhamento.upper()

      try:
        dim_val = float(
            str(row.get("DIMENSÃO", row.get("DIMENSAO", 1.25))).replace(
                ",", "."
            )
        )
      except:
        dim_val = 1.25

      qtd_exp = int(row.get("QTD", row.get("QUANTIDADE", 1)))

      # REGRA DE EXPOSITORES ABERTOS (Sem 'P' e sem 'PVT') -> VAH, VAL, etc.
      usa_porta = ("P" in alinh_upper) or ("PVT" in alinh_upper)
      if not usa_porta:
        lista_abertos.append({
            "CLIENTE": cliente,
            "ORFA / OP": orfa,
            "ALINHAMENTO": alinhamento,
            "TIPO": "EXPOSITOR ABERTO (S/ PORTA)",
            "DIMENSÃO (m)": dim_val,
            "QTD EXP": qtd_exp,
            "OBSERVAÇÃO": "Requer programação manual de vidros",
        })
        continue

      serigrafia = str(row.get("SERIGRAFIA", "PRETA")).strip().upper()
      if "CINZA" in alinh_upper:
        serigrafia = "CINZA"
      elif "PRETO" in alinh_upper or "PRETA" in alinh_upper:
        serigrafia = "PRETA"

      dim_mm = converter_dimensao_para_mm(dim_val)
      is_vap = "VAP" in alinh_upper

      if is_vap:
        # --- LÓGICA VAP (CONGELADOS) ---
        num_portas_map = {1500: 2, 2250: 3, 3000: 4, 3750: 5}
        qtd_portas_vap = num_portas_map.get(
            dim_mm, round(dim_mm / 750)
        ) * qtd_exp

        cod_porta = CODIGOS_PORTAS_VAP.get(serigrafia, "S/COD")
        desc_porta = f"PORTA CONG DIR 730X1714 {serigrafia}"

        cod_quadro = CODIGOS_QUADROS_VAP.get((serigrafia, dim_mm), "S/COD")
        desc_quadro = f"QUADRO {dim_mm}X1735 CONGELADOS {serigrafia}"
        qtd_quadros = 1 * qtd_exp

        lista_resumo_portas.append({
            "CÓDIGO": cod_porta,
            "MODELO DE PORTA": desc_porta,
            "QTD PORTAS (UN)": qtd_portas_vap,
        })
        lista_resumo_frames.append({
            "CÓDIGO": cod_quadro,
            "MODELO DE FRAME": desc_quadro,
            "QTD FRAMES (UN)": qtd_quadros,
        })

        lista_detalhada.append({
            "CLIENTE": cliente,
            "ORFA / OP": orfa,
            "ALINHAMENTO": alinhamento,
            "TIPO": "VAP (Congelados)",
            "DIMENSÃO (m)": dim_val,
            "QTD EXP": qtd_exp,
            "CÓD. PORTA DIR": cod_porta,
            "QTD PORTAS DIR": qtd_portas_vap,
            "TOTAL PORTAS": qtd_portas_vap,
            "CÓD. QUADRO": cod_quadro,
            "QUADRO DESC": desc_quadro,
            "QTD FRAMES": qtd_quadros,
        })

      else:
        # --- LÓGICA VA (STANDARD REFRIGERADO COM PORTA - EX: VALP, VAHP, VACP) ---
        is_pvt = "PVT" in alinh_upper
        tipo_porta_desc = "TRANSPARENTE" if is_pvt else "SERIGRAFADA"

        if dim_mm == 1875:
          qtd_esq = 1 * qtd_exp
          qtd_dir = 2 * qtd_exp
        else:
          pares = round(dim_mm / 1250)
          qtd_esq = pares * qtd_exp
          qtd_dir = pares * qtd_exp

        cod_esq = CODIGOS_PORTAS_VA.get(
            (serigrafia, tipo_porta_desc, "ESQ"), "S/COD"
        )
        cod_dir = CODIGOS_PORTAS_VA.get(
            (serigrafia, tipo_porta_desc, "DIR"), "S/COD"
        )

        cod_frame = CODIGOS_FRAMES_VA.get((serigrafia, dim_mm), "S/COD")
        desc_frame_txt = (
            f"CJ FRAME RESFR. {dim_mm} {serigrafia}"
            if serigrafia == "CINZA"
            else f"CJ FRAME RESFR {dim_mm} PRETO"
        )
        qtd_frames = 1 * qtd_exp

        lista_resumo_portas.append({
            "CÓDIGO": cod_esq,
            "MODELO DE PORTA": f"PORTA {serigrafia} {tipo_porta_desc} ESQUERDA",
            "QTD PORTAS (UN)": qtd_esq,
        })
        lista_resumo_portas.append({
            "CÓDIGO": cod_dir,
            "MODELO DE PORTA": f"PORTA {serigrafia} {tipo_porta_desc} DIREITA",
            "QTD PORTAS (UN)": qtd_dir,
        })
        lista_resumo_frames.append({
            "CÓDIGO": cod_frame,
            "MODELO DE FRAME": desc_frame_txt,
            "QTD FRAMES (UN)": qtd_frames,
        })

        lista_detalhada.append({
            "CLIENTE": cliente,
            "ORFA / OP": orfa,
            "ALINHAMENTO": alinhamento,
            "TIPO": "VA (Standard c/ Porta)",
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

    df_itens = pd.DataFrame(lista_detalhada) if lista_detalhada else pd.DataFrame()
    df_abertos = pd.DataFrame(lista_abertos) if lista_abertos else pd.DataFrame()
    df_portas_resumo = (
        pd.DataFrame(lista_resumo_portas)
        .groupby(["CÓDIGO", "MODELO DE PORTA"], as_index=False)
        .agg({"QTD PORTAS (UN)": "sum"})
        .sort_values(by="CÓDIGO")
        if lista_resumo_portas
        else pd.DataFrame()
    )
    df_frames_resumo = (
        pd.DataFrame(lista_resumo_frames)
        .groupby(["CÓDIGO", "MODELO DE FRAME"], as_index=False)
        .agg({"QTD FRAMES (UN)": "sum"})
        .sort_values(by="CÓDIGO")
        if lista_resumo_frames
        else pd.DataFrame()
    )

    total_expositores = df_itens["QTD EXP"].sum() if not df_itens.empty else 0
    total_portas_geral = df_itens["TOTAL PORTAS"].sum() if not df_itens.empty else 0
    total_frames_geral = df_itens["QTD FRAMES"].sum() if not df_itens.empty else 0
    total_abertos = df_abertos["QTD EXP"].sum() if not df_abertos.empty else 0

    # --- PAINEL VISUAL NA TELA ---
    st.write("---")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("EXPOSITORES C/ PORTA", total_expositores)
    col2.metric("TOTAL DE PORTAS (UN)", total_portas_geral)
    col3.metric("TOTAL DE FRAMES", total_frames_geral)
    col4.metric("EXPOSITORES ABERTOS", total_abertos, delta="Atenção Vidros")

    if not df_abertos.empty:
      st.warning(
          f"⚠️ Atenção: Foram encontrados **{total_abertos} expositor(es) aberto(s)**"
          " (sem fechamento de porta). Eles foram listados na seção exclusiva"
          " abaixo para conferência e programação manual de vidros."
      )

    st.write("")
    st.write("### 1. Resumo Consolidado de Portas (VA & VAP)")
    if not df_portas_resumo.empty:
      st.dataframe(df_portas_resumo, use_container_width=True, hide_index=True)
    else:
      st.info("Nenhuma porta computada.")

    st.write("### 2. Resumo Consolidado de Frames e Quadros Congelados")
    if not df_frames_resumo.empty:
      st.dataframe(df_frames_resumo, use_container_width=True, hide_index=True)
    else:
      st.info("Nenhum frame computado.")

    st.write("### 3. Demanda Detalhada por Cliente e Alinhamento")
    if not df_itens.empty:
      st.dataframe(df_itens, use_container_width=True, hide_index=True)
    else:
      st.info("Nenhum item com porta detalhado.")

    st.write(
        "### 4. 🪟 Área de Atenção: Expositores Abertos (Programação Manual de"
        " Vidros)"
    )
    if not df_abertos.empty:
      st.dataframe(df_abertos, use_container_width=True, hide_index=True)
    else:
      st.success("Nenhum expositor aberto encontrado na planilha atual.")

    # --- FUNÇÃO DE GERAÇÃO DO EXCEL EXECUTIVO ---
    def gerar_excel_portas(
        tot_exp,
        tot_portas,
        tot_frames,
        df_portas,
        df_frames,
        df_detalhado,
        df_abertos_rel,
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
      font_kpi_title = Font(
          name="Calibri", size=9, bold=True, color="555555"
      )
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
          "RELATÓRIO EXECUTIVO - CÁLCULO DE PORTAS, FRAMES & EXPOSITORES"
          " ABERTOS"
      )
      c1.font = font_main_title
      c1.fill = fill_header
      c1.alignment = Alignment(horizontal="center", vertical="center")

      c2 = ws["A2"]
      c2.value = (
          "Mapeamento Oficial | PCP & Diretoria"
      )
      c2.font = font_subtitle
      c2.fill = fill_header
      c2.alignment = Alignment(horizontal="center", vertical="center")

      ws.row_dimensions[1].height = 25
      ws.row_dimensions[2].height = 18

      kpis = [
          ("EXPOSITORES C/ PORTA", tot_exp, "#,##0", "A", "C"),
          ("TOTAL PORTAS (UN)", tot_portas, "#,##0", "D", "G"),
          ("TOTAL FRAMES / QUADROS", tot_frames, "#,##0", "H", "M"),
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
          ("2. RESUMO CONSOLIDADO DE FRAMES E QUADROS POR CÓDIGO", df_frames),
          ("3. DEMANDA DETALHADA POR CLIENTE E ALINHAMENTO", df_detalhado),
          (
              "4. EXPOSITORES ABERTOS (ATENÇÃO: PROGRAMAÇÃO MANUAL DE VIDROS)",
              df_abertos_rel,
          ),
      ]

      for titulo_secao, df_data in blocos:
        if df_data.empty:
          continue
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
        df_abertos,
    )

    st.write("")
    st.download_button(
        label="📥 Descarregar Relatório Executivo de Portas & Frames (.xlsx)",
        data=excel_portas,
        file_name="Relatorio_Executivo_Portas_Frames_VA_VAP.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )