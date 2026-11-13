import streamlit as st
import pandas as pd
import io
import re
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# --- CABEÇALHO DO MÓDULO ---
st.markdown(
    """
    <div style="padding: 10px 0 20px 0;">
        <h1 style="font-size: 26px; font-weight: 800; margin-bottom: 5px;">
            🧪 OP de Injeção de PU (Poliol + Iso)
        </h1>
        <p style="font-size: 14px; color: #94A3B8; margin: 0;">
            Geração automatizada de Ordens de Produção por Linha, Peça (Bojo, Costas, Teto/Tampo) e Identificação de Expositores Especiais.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

# --- FUNÇÃO AUXILIAR PARA LEITURA DE PDF OU EXCEL ---
def carregar_dados(uploaded_file):
    nome = uploaded_file.name.lower()
    if nome.endswith((".xlsx", ".xls")):
        xls = pd.ExcelFile(uploaded_file)
        return pd.read_excel(xls, xls.sheet_names[0])
    elif nome.endswith(".pdf"):
        try:
            import pdfplumber
            linhas_extraidas = []
            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        for row in table:
                            if row and any(row):
                                linhas_extraidas.append([cell.strip() if cell else "" for cell in row])
            if not linhas_extraidas:
                raise ValueError("Nenhuma tabela estruturada foi encontrada no PDF.")
            
            header_idx = 0
            for idx, r in enumerate(linhas_extraidas):
                row_upper = [str(c).upper() for c in r]
                if "ALINHAMENTO" in row_upper or "ORFA" in row_upper:
                    header_idx = idx
                    break
            
            headers = [str(c).strip().upper() for c in linhas_extraidas[header_idx]]
            data = linhas_extraidas[header_idx + 1:]
            df_pdf = pd.DataFrame(data, columns=headers)
            df_pdf = df_pdf[df_pdf[headers[0]].str.upper() != headers[0]]
            return df_pdf
        except ImportError:
            st.error("A biblioteca `pdfplumber` não está instalada no ambiente. Adicione `pdfplumber` ao seu `requirements.txt` para leitura direta de PDFs.")
            st.stop()

# --- MOTOR DE REGRAS DE ENGENHARIA (INJEÇÃO DE PU) ---
def processar_regras_injecao(df_raw):
    df = df_raw.copy()
    df.columns = [str(c).strip().upper() for c in df.columns]

    colunas_necessarias = ["ALINHAMENTO", "DIMENSÃO", "BOJO"]
    for col in colunas_necessarias:
        if col not in df.columns:
            raise ValueError(f"Coluna obrigatória '{col}' não encontrada no arquivo. Colunas detectadas: {list(df.columns)}")

    df = df.dropna(subset=["ALINHAMENTO"]).copy()
    df["DIMENSÃO"] = pd.to_numeric(df["DIMENSÃO"].astype(str).str.replace(",", "."), errors="coerce").fillna(0.0)
    df["BOJO_MAT"] = df["BOJO"].fillna("NÃO ESPECIFICADO").astype(str).str.strip().str.upper()
    df["ALINHAMENTO_STR"] = df["ALINHAMENTO"].astype(str).str.strip().str.upper()
    df["ORFA_STR"] = df["ORFA"].astype(str).str.strip() if "ORFA" in df.columns else "-"
    df["LINHA_STR"] = df["LINHA"].astype(str).str.strip().str.upper() if "LINHA" in df.columns else "-"
    df["CLIENTE_STR"] = df["NAME"].astype(str).str.strip() if "NAME" in df.columns else "-"

    pecas_padrao = []
    excecoes_supervisor = []

    termos_especiais = ["CANTO", "CT 180", "CT180", "CENTRAL", "PONTA", "90", "180"]

    for _, row in df.iterrows():
        alinhamento = row["ALINHAMENTO_STR"]
        dimensao = row["DIMENSÃO"]
        mat_bojo = row["BOJO_MAT"]
        orfa = row["ORFA_STR"]
        linha_orig = row["LINHA_STR"]
        cliente = row["CLIENTE_STR"]

        modelo_limpo = re.sub(r"^\d+(\.\d+)?\s*", "", alinhamento).strip()
        is_especial = any(termo in modelo_limpo for termo in termos_especiais)

        # 1. REGRA VRA (Açougue)
        if modelo_limpo.startswith("VRA") and not is_especial:
            pecas_padrao.append({"LINHA_PRODUTO": "AÇOUGUE (VRA)", "TIPO_PECA": "BOJO VRA", "MODELO_PECA": f"BOJO VRA - {mat_bojo}", "MATERIAL": mat_bojo, "DIMENSAO": dimensao, "QTD": 1, "ORFA": orfa})
            pecas_padrao.append({"LINHA_PRODUTO": "AÇOUGUE (VRA)", "TIPO_PECA": "COSTAS VRA", "MODELO_PECA": "COSTAS VRA", "MATERIAL": "PADRÃO", "DIMENSAO": dimensao, "QTD": 1, "ORFA": orfa})
            pecas_padrao.append({"LINHA_PRODUTO": "AÇOUGUE (VRA)", "TIPO_PECA": "TAMPO VRA", "MODELO_PECA": "TAMPOS VRA", "MATERIAL": "PADRÃO", "DIMENSAO": dimensao, "QTD": 1, "ORFA": orfa})

        # 2. REGRA VCA (Açougue)
        elif modelo_limpo.startswith("VCA") and not is_especial:
            pecas_padrao.append({"LINHA_PRODUTO": "AÇOUGUE (VCA)", "TIPO_PECA": "BOJO VCA", "MODELO_PECA": f"BOJO VCA - {mat_bojo}", "MATERIAL": mat_bojo, "DIMENSAO": dimensao, "QTD": 1, "ORFA": orfa})
            pecas_padrao.append({"LINHA_PRODUTO": "AÇOUGUE (VCA)", "TIPO_PECA": "TAMPO VCA", "MODELO_PECA": "TAMPOS VCA", "MATERIAL": "PADRÃO", "DIMENSAO": dimensao, "QTD": 1, "ORFA": orfa})

        # 3. REGRA VA (Vertical Alto)
        elif modelo_limpo.startswith("VA") and not is_especial:
            pecas_padrao.append({"LINHA_PRODUTO": "VERTICAL ALTO (VA)", "TIPO_PECA": "BOJO VA", "MODELO_PECA": f"BOJO VA - {mat_bojo}", "MATERIAL": mat_bojo, "DIMENSAO": dimensao, "QTD": 1, "ORFA": orfa})
            pecas_padrao.append({"LINHA_PRODUTO": "VERTICAL ALTO (VA)", "TIPO_PECA": "COSTAS VA", "MODELO_PECA": "COSTAS VA", "MATERIAL": "PADRÃO", "DIMENSAO": dimensao, "QTD": 1, "ORFA": orfa})
            pecas_padrao.append({"LINHA_PRODUTO": "VERTICAL ALTO (VA)", "TIPO_PECA": "TETO VA", "MODELO_PECA": "TETOS VA", "MATERIAL": "PADRÃO", "DIMENSAO": dimensao, "QTD": 1, "ORFA": orfa})

        # 4. REGRA VR (Padaria)
        elif modelo_limpo.startswith("VR") and not is_especial:
            pecas_padrao.append({"LINHA_PRODUTO": "PADARIA (VR)", "TIPO_PECA": "BOJO VILXVR", "MODELO_PECA": f"BOJO VILXVR - {mat_bojo}", "MATERIAL": mat_bojo, "DIMENSAO": dimensao, "QTD": 1, "ORFA": orfa})
            pecas_padrao.append({"LINHA_PRODUTO": "PADARIA (VR)", "TIPO_PECA": "COSTAS VRP", "MODELO_PECA": "COSTAS VRP", "MATERIAL": "PADRÃO", "DIMENSAO": dimensao, "QTD": 1, "ORFA": orfa})

        # 5. REGRA VIL LINEAR
        elif modelo_limpo.startswith("VIL") and not is_especial and ("2P" in modelo_limpo or "3P" in modelo_limpo):
            variante_costas = "COSTAS VIL 3P" if "3P" in modelo_limpo else "COSTAS VIL 2P"
            pecas_padrao.append({"LINHA_PRODUTO": "LINHA VIL (LINEAR)", "TIPO_PECA": "BOJO VILXVR", "MODELO_PECA": f"BOJO VILXVR - {mat_bojo}", "MATERIAL": mat_bojo, "DIMENSAO": dimensao, "QTD": 1, "ORFA": orfa})
            pecas_padrao.append({"LINHA_PRODUTO": "LINHA VIL (LINEAR)", "TIPO_PECA": "COSTAS VIL", "MODELO_PECA": variante_costas, "MATERIAL": "PADRÃO", "DIMENSAO": dimensao, "QTD": 1, "ORFA": orfa})

        # 6. EXCEÇÕES / ESPECIAIS
        else:
            motivo = "VIL Especial (Canto / CT 180 / Central)" if modelo_limpo.startswith("VIL") else "Sigla / Modelo Fora do Padrão Linear"
            excecoes_supervisor.append({
                "ORFA": orfa,
                "CLIENTE": cliente,
                "LINHA_PLANILHA": linha_orig,
                "ALINHAMENTO": alinhamento,
                "DIMENSAO": dimensao,
                "BOJO_INDICADO": mat_bojo,
                "MOTIVO_ATENCAO": motivo
            })

    df_pecas = pd.DataFrame(pecas_padrao)
    df_excecoes = pd.DataFrame(excecoes_supervisor)

    if not df_pecas.empty:
        df_op_consolidada = df_pecas.groupby(
            ["LINHA_PRODUTO", "TIPO_PECA", "MODELO_PECA", "MATERIAL", "DIMENSAO"], as_index=False
        ).agg(
            QTD_TOTAL=("QTD", "sum"),
            ORFAS=("ORFA", lambda x: ", ".join(sorted(set(x))))
        ).sort_values(by=["LINHA_PRODUTO", "TIPO_PECA", "MATERIAL", "DIMENSAO"])
    else:
        df_op_consolidada = pd.DataFrame()

    return df_op_consolidada, df_excecoes

# --- GERADOR DE EXCEL FORMATADO (MULTI-ABAS) ---
def gerar_excel_op(df_op, df_excecoes):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if not df_op.empty:
            df_op.to_excel(writer, index=False, sheet_name="OP_Injecao_PU")
        else:
            pd.DataFrame({"Aviso": ["Nenhuma peça padrão identificada."]}).to_excel(writer, index=False, sheet_name="OP_Injecao_PU")

        if not df_excecoes.empty:
            df_excecoes.to_excel(writer, index=False, sheet_name="Especiais_Supervisor")
        else:
            pd.DataFrame({"Status": ["Nenhum expositor especial/exceção nesta programação."]}).to_excel(writer, index=False, sheet_name="Especiais_Supervisor")

        wb = writer.book
        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        warn_fill = PatternFill(start_color="B45309", end_color="B45309", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=11)
        thin_border = Border(
            left=Side(style="thin", color="D1D5DB"),
            right=Side(style="thin", color="D1D5DB"),
            top=Side(style="thin", color="D1D5DB"),
            bottom=Side(style="thin", color="D1D5DB"),
        )

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            fill_atual = warn_fill if sheet_name == "Especiais_Supervisor" else header_fill
            for cell in ws[1]:
                cell.fill = fill_atual
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
            
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    cell.border = thin_border
                    cell.alignment = Alignment(horizontal="center", vertical="center")

            for col in ws.columns:
                max_len = max(len(str(cell.value or "")) for cell in col)
                ws.column_dimensions[col[0].column_letter].width = max(max_len + 4, 14)

    return output.getvalue()

# --- INTERFACE DE USUÁRIO ---
uploaded_file = st.file_uploader(
    "📂 Importar Programação de Montagem (PDF ou Excel)",
    type=["pdf", "xlsx", "xls"],
)

if uploaded_file is not None:
    try:
        df_entrada = carregar_dados(uploaded_file)
        df_op, df_excecoes = processar_regras_injecao(df_entrada)

        k1, k2, k3 = st.columns(3)
        total_pecas = int(df_op["QTD_TOTAL"].sum()) if not df_op.empty else 0
        total_bojos = int(df_op[df_op["TIPO_PECA"].str.contains("BOJO")]["QTD_TOTAL"].sum()) if not df_op.empty else 0
        total_especiais = len(df_excecoes)

        k1.metric("Total de Peças para Injeção", f"{total_pecas} un.")
        k2.metric("Total de Bojos (Inox / PP)", f"{total_bojos} un.")
        k3.metric("Expositores Especiais / Atenção", f"{total_especiais} itens")

        st.markdown("---")

        tab_op, tab_especiais, tab_bruto = st.tabs([
            "📋 OP de Injeção Consolidada",
            f"⚠️ Especiais / Supervisão ({total_especiais})",
            "📄 Programação Original"
        ])

        with tab_op:
            if not df_op.empty:
                st.dataframe(df_op, use_container_width=True, hide_index=True)
            else:
                st.warning("Nenhuma peça das linhas padrão (VA, VRA, VCA, VR, VIL Linear) foi encontrada.")

        with tab_especiais:
            if not df_excecoes.empty:
                st.warning("Os itens abaixo são modelos especiais (Cantos, CT 180, Centrais) ou siglas fora do padrão linear e requerem análise do supervisor:")
                st.dataframe(df_excecoes, use_container_width=True, hide_index=True)
            else:
                st.success("✅ Nenhum item especial ou fora do padrão encontrado nesta programação.")

        with tab_bruto:
            st.dataframe(df_entrada, use_container_width=True, hide_index=True)

        st.markdown("<br>", unsafe_allow_html=True)
        excel_bytes = gerar_excel_op(df_op, df_excecoes)
        st.download_button(
            label="📥 Baixar OP de Injeção Completa (Excel com Aba de Supervisão)",
            data=excel_bytes,
            file_name="OP_Injecao_PU_Consolidada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
else:
    st.info("💡 Faça o upload da planilha ou PDF de programação para gerar automaticamente as OPs de Bojos, Costas, Tetos/Tampos e a lista de itens especiais.")