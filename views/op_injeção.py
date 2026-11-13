import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# --- CABEÇALHO DO MÓDULO ---
st.markdown(
    """
    <div style="padding: 10px 0 20px 0;">
        <h1 style="font-size: 26px; font-weight: 800; margin-bottom: 5px;">
            🧪 OP de Injeção de PU (Poliol + Iso)
        </h1>
        <p style="font-size: 14px; color: #94A3B8; margin: 0;">
            Geração estruturada de Ordens de Produção (OP) com cabeçalho corporativo, separação por Bojos, Costas e Tetos/Tampos e resumo consolidado por linha.
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
    df["ALINHAMENTO_STR"] = df["ALINHAMENTO"].astype(str).str.strip().upper()
    df["ORFA_STR"] = df["ORFA"].astype(str).str.strip() if "ORFA" in df.columns else "-"
    df["LINHA_STR"] = df["LINHA"].astype(str).str.strip().upper() if "LINHA" in df.columns else "-"
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

        # 1. REGRA VRA (Açougue) -> Bojo, Costas VRA, Tampos VRA
        if modelo_limpo.startswith("VRA") and not is_especial:
            pecas_padrao.append({"LINHA_PRODUTO": "AÇOUGUE (VRA)", "GRUPO_PECA": "BOJOS", "TIPO_PECA": "BOJO VRA", "MODELO_PECA": f"BOJO VRA - {mat_bojo}", "MATERIAL": mat_bojo, "DIMENSAO": dimensao, "QTD_UNITARIA": 1, "ORFA": orfa})
            pecas_padrao.append({"LINHA_PRODUTO": "AÇOUGUE (VRA)", "GRUPO_PECA": "COSTAS", "TIPO_PECA": "COSTAS VRA", "MODELO_PECA": "COSTAS VRA", "MATERIAL": "PADRÃO", "DIMENSAO": dimensao, "QTD_UNITARIA": 1, "ORFA": orfa})
            pecas_padrao.append({"LINHA_PRODUTO": "AÇOUGUE (VRA)", "GRUPO_PECA": "TETOS - TAMPOS", "TIPO_PECA": "TAMPO VRA", "MODELO_PECA": "TAMPOS VRA", "MATERIAL": "PADRÃO", "DIMENSAO": dimensao, "QTD_UNITARIA": 1, "ORFA": orfa})

        # 2. REGRA VCA (Açougue) -> Bojo, Tampos VCA
        elif modelo_limpo.startswith("VCA") and not is_especial:
            pecas_padrao.append({"LINHA_PRODUTO": "AÇOUGUE (VCA)", "GRUPO_PECA": "BOJOS", "TIPO_PECA": "BOJO VCA", "MODELO_PECA": f"BOJO VCA - {mat_bojo}", "MATERIAL": mat_bojo, "DIMENSAO": dimensao, "QTD_UNITARIA": 1, "ORFA": orfa})
            pecas_padrao.append({"LINHA_PRODUTO": "AÇOUGUE (VCA)", "GRUPO_PECA": "TETOS - TAMPOS", "TIPO_PECA": "TAMPO VCA", "MODELO_PECA": "TAMPOS VCA", "MATERIAL": "PADRÃO", "DIMENSAO": dimensao, "QTD_UNITARIA": 1, "ORFA": orfa})

        # 3. REGRA VA (Vertical Alto) -> Bojo, Costas VA, Tetos VA
        elif modelo_limpo.startswith("VA") and not is_especial:
            pecas_padrao.append({"LINHA_PRODUTO": "VERTICAL ALTO (VA)", "GRUPO_PECA": "BOJOS", "TIPO_PECA": "BOJO VA", "MODELO_PECA": f"BOJO VA - {mat_bojo}", "MATERIAL": mat_bojo, "DIMENSAO": dimensao, "QTD_UNITARIA": 1, "ORFA": orfa})
            pecas_padrao.append({"LINHA_PRODUTO": "VERTICAL ALTO (VA)", "GRUPO_PECA": "COSTAS", "TIPO_PECA": "COSTAS VA", "MODELO_PECA": "COSTAS VA", "MATERIAL": "PADRÃO", "DIMENSAO": dimensao, "QTD_UNITARIA": 1, "ORFA": orfa})
            pecas_padrao.append({"LINHA_PRODUTO": "VERTICAL ALTO (VA)", "GRUPO_PECA": "TETOS - TAMPOS", "TIPO_PECA": "TETO VA", "MODELO_PECA": "TETOS VA", "MATERIAL": "PADRÃO", "DIMENSAO": dimensao, "QTD_UNITARIA": 1, "ORFA": orfa})

        # 4. REGRA VR (Padaria) -> Bojo VILXVR, Costas VRP, Teto VR
        elif modelo_limpo.startswith("VR") and not is_especial:
            pecas_padrao.append({"LINHA_PRODUTO": "PADARIA (VR)", "GRUPO_PECA": "BOJOS", "TIPO_PECA": "BOJO VILXVR", "MODELO_PECA": f"BOJO VILXVR - {mat_bojo}", "MATERIAL": mat_bojo, "DIMENSAO": dimensao, "QTD_UNITARIA": 1, "ORFA": orfa})
            pecas_padrao.append({"LINHA_PRODUTO": "PADARIA (VR)", "GRUPO_PECA": "COSTAS", "TIPO_PECA": "COSTAS VR", "MODELO_PECA": "COSTAS VRP", "MATERIAL": "PADRÃO", "DIMENSAO": dimensao, "QTD_UNITARIA": 1, "ORFA": orfa})
            pecas_padrao.append({"LINHA_PRODUTO": "PADARIA (VR)", "GRUPO_PECA": "TETOS - TAMPOS", "TIPO_PECA": "TETO VR", "MODELO_PECA": "TETOS VR", "MATERIAL": "PADRÃO", "DIMENSAO": dimensao, "QTD_UNITARIA": 1, "ORFA": orfa})

        # 5. REGRA VIL LINEAR -> Bojo VILXVR, Costas VIL (2P/3P), Teto VIL
        elif modelo_limpo.startswith("VIL") and not is_especial and ("2P" in modelo_limpo or "3P" in modelo_limpo):
            nome_costas = "COSTAS VIL 3P" if "3P" in modelo_limpo else "COSTAS VIL 2P"
            pecas_padrao.append({"LINHA_PRODUTO": "LINHA VIL (LINEAR)", "GRUPO_PECA": "BOJOS", "TIPO_PECA": "BOJO VILXVR", "MODELO_PECA": f"BOJO VILXVR - {mat_bojo}", "MATERIAL": mat_bojo, "DIMENSAO": dimensao, "QTD_UNITARIA": 1, "ORFA": orfa})
            pecas_padrao.append({"LINHA_PRODUTO": "LINHA VIL (LINEAR)", "GRUPO_PECA": "COSTAS", "TIPO_PECA": "COSTAS VIL", "MODELO_PECA": nome_costas, "MATERIAL": "PADRÃO", "DIMENSAO": dimensao, "QTD_UNITARIA": 1, "ORFA": orfa})
            pecas_padrao.append({"LINHA_PRODUTO": "LINHA VIL (LINEAR)", "GRUPO_PECA": "TETOS - TAMPOS", "TIPO_PECA": "TETO VIL", "MODELO_PECA": "TETOS VIL", "MATERIAL": "PADRÃO", "DIMENSAO": dimensao, "QTD_UNITARIA": 1, "ORFA": orfa})

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
        # Tabela 1: Consolidado detalhado por Grupo, Linha, Peça, Material e Dimensão
        df_op_consolidada = df_pecas.groupby(
            ["LINHA_PRODUTO", "GRUPO_PECA", "TIPO_PECA", "MODELO_PECA", "MATERIAL", "DIMENSAO"], as_index=False
        ).agg(
            QTD_UNITARIA=("QTD_UNITARIA", "first"),
            QTD_TOTAL=("QTD_UNITARIA", "sum"),
            ORFAS=("ORFA", lambda x: ", ".join(sorted(set(x))))
        ).sort_values(by=["LINHA_PRODUTO", "GRUPO_PECA", "MATERIAL", "DIMENSAO"])

        # Tabela 2: Resumo por Linha de Produção
        df_resumo_linha = df_pecas.groupby(
            ["LINHA_PRODUTO", "GRUPO_PECA"], as_index=False
        ).agg(
            TOTAL_PECAS=("QTD_UNITARIA", "sum")
        ).pivot(index="LINHA_PRODUTO", columns="GRUPO_PECA", values="TOTAL_PECAS").fillna(0).reset_index()
        
        # Garante que as colunas principais existam no resumo
        for col in ["BOJOS", "COSTAS", "TETOS - TAMPOS"]:
            if col not in df_resumo_linha.columns:
                df_resumo_linha[col] = 0
                
        df_resumo_linha["TOTAL_GERAL"] = df_resumo_linha[["BOJOS", "COSTAS", "TETOS - TAMPOS"]].sum(axis=1)
    else:
        df_op_consolidada = pd.DataFrame()
        df_resumo_linha = pd.DataFrame()

    return df_op_consolidada, df_resumo_linha, df_excecoes

# --- GERADOR DE EXCEL FORMATADO COM CABEÇALHO DE OP ---
def gerar_excel_op(df_op, df_resumo, df_excecoes):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        wb = writer.book
        ws = wb.create_sheet(title="OP_Injecao_PU")
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])

        # 1. Cabeçalho Corporativo da OP
        ws["A1"] = "EMPRESA:"
        ws["B1"] = "NSF TECNOLOGIA EM INSTALAÇÕES COMERCIAIS"
        ws["A2"] = "SETOR DE DESTINO:"
        ws["B2"] = "SETOR DE INJEÇÃO DE PU (POLIOL + ISO)"
        ws["A3"] = "EMITIDO POR:"
        ws["B3"] = "Portal PCP Enterprise"
        ws["A4"] = "DATA / HORA:"
        ws["B4"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        # Estilo do Cabeçalho
        title_font = Font(name="Calibri", size=11, bold=True, color="1E3A8A")
        val_font = Font(name="Calibri", size=11, color="000000")
        for r in range(1, 5):
            ws[f"A{r}"].font = title_font
            ws[f"B{r}"].font = val_font

        # 2. Seção 1: Relatório Detalhado de Peças
        start_row_det = 7
        ws[f"A{start_row_det - 1}"] = "1. RELATÓRIO DETALHADO DE ORDENS DE PRODUÇÃO (BOJOS, COSTAS, TETOS-TAMPOS)"
        ws[f"A{start_row_det - 1}"].font = Font(name="Calibri", size=12, bold=True, color="1E3A8A")

        if not df_op.empty:
            df_op.to_excel(writer, index=False, sheet_name="OP_Injecao_PU", startrow=start_row_det - 1, startcol=0)
        else:
            pd.DataFrame({"Aviso": ["Nenhuma peça padrão identificada."]}).to_excel(writer, index=False, sheet_name="OP_Injecao_PU", startrow=start_row_det - 1, startcol=0)

        # 3. Seção 2: Resumo por Linha de Produção
        row_after_det = start_row_det + len(df_op) + 3 if not df_op.empty else start_row_det + 3
        ws[f"A{row_after_det - 1}"] = "2. RESUMO CONSOLIDADO DE PEÇAS POR LINHA DE PRODUÇÃO"
        ws[f"A{row_after_det - 1}"].font = Font(name="Calibri", size=12, bold=True, color="1E3A8A")

        if not df_resumo.empty:
            df_resumo.to_excel(writer, index=False, sheet_name="OP_Injecao_PU", startrow=row_after_det - 1, startcol=0)

        # Aba 2: Especiais / Exceções para o Supervisor
        if not df_excecoes.empty:
            df_excecoes.to_excel(writer, index=False, sheet_name="Especiais_Supervisor")
        else:
            pd.DataFrame({"Status": ["Nenhum expositor especial/exceção nesta programação."]}).to_excel(writer, index=False, sheet_name="Especiais_Supervisor")

        # Estilização Geral do Excel (Bordas, Cores e Alinhamento)
        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        warn_fill = PatternFill(start_color="B45309", end_color="B45309", fill_type="solid")
        header_font = Font(name="Calibri", size=11, color="FFFFFF", bold=True)
        thin_border = Border(
            left=Side(style="thin", color="D1D5DB"),
            right=Side(style="thin", color="D1D5DB"),
            top=Side(style="thin", color="D1D5DB"),
            bottom=Side(style="thin", color="D1D5DB"),
        )

        ws_op = wb["OP_Injecao_PU"]
        # Formata cabeçalhos das tabelas na aba OP
        for r_idx in [start_row_det, row_after_det]:
            for col_idx in range(1, 20):
                cell = ws_op.cell(row=r_idx, column=col_idx)
                if cell.value is not None:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center")

        # Formata dados das tabelas da OP
        for row in ws_op.iter_rows(min_row=start_row_det + 1, max_row=ws_op.max_row):
            for cell in row:
                if cell.value is not None:
                    cell.border = thin_border
                    cell.alignment = Alignment(horizontal="center", vertical="center")

        # Formata aba de Supervisão
        if "Especiais_Supervisor" in wb.sheetnames:
            ws_sup = wb["Especiais_Supervisor"]
            for cell in ws_sup[1]:
                cell.fill = warn_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
            for row in ws_sup.iter_rows(min_row=2):
                for cell in row:
                    cell.border = thin_border
                    cell.alignment = Alignment(horizontal="center", vertical="center")

        # Ajuste de largura das colunas
        for sheet_name in wb.sheetnames:
            ws_sheet = wb[sheet_name]
            for col in ws_sheet.columns:
                max_len = max(len(str(cell.value or "")) for cell in col)
                col_letter = col[0].column_letter
                ws_sheet.column_dimensions[col_letter].width = max(max_len + 4, 15)

    return output.getvalue()

# --- INTERFACE DE USUÁRIO ---
uploaded_file = st.file_uploader(
    "📂 Importar Programação de Montagem (PDF ou Excel)",
    type=["pdf", "xlsx", "xls"],
)

if uploaded_file is not None:
    try:
        df_entrada = carregar_dados(uploaded_file)
        df_op, df_resumo, df_excecoes = processar_regras_injecao(df_entrada)

        k1, k2, k3 = st.columns(3)
        total_pecas = int(df_op["QTD_TOTAL"].sum()) if not df_op.empty else 0
        total_bojos = int(df_op[df_op["GRUPO_PECA"] == "BOJOS"]["QTD_TOTAL"].sum()) if not df_op.empty else 0
        total_especiais = len(df_excecoes)

        k1.metric("Total de Peças para Injeção", f"{total_pecas} un.")
        k2.metric("Total de Bojos (Inox / PP)", f"{total_bojos} un.")
        k3.metric("Expositores Especiais / Atenção", f"{total_especiais} itens")

        st.markdown("---")

        tab_op, tab_resumo, tab_especiais, tab_bruto = st.tabs([
            "📋 OP Consolidada Detalhada",
            "📊 Resumo por Linha",
            f"⚠️ Especiais / Supervisão ({total_especiais})",
            "📄 Programação Original"
        ])

        with tab_op:
            if not df_op.empty:
                st.dataframe(df_op, use_container_width=True, hide_index=True)
            else:
                st.warning("Nenhuma peça das linhas padrão foi encontrada.")

        with tab_resumo:
            if not df_resumo.empty:
                st.dataframe(df_resumo, use_container_width=True, hide_index=True)
            else:
                st.warning("Sem dados para o resumo.")

        with tab_especiais:
            if not df_excecoes.empty:
                st.warning("Os itens abaixo são modelos especiais (Cantos, CT 180, Centrais) e requerem análise do supervisor:")
                st.dataframe(df_excecoes, use_container_width=True, hide_index=True)
            else:
                st.success("✅ Nenhum item especial encontrado nesta programação.")

        with tab_bruto:
            st.dataframe(df_entrada, use_container_width=True, hide_index=True)

        st.markdown("<br>", unsafe_allow_html=True)
        excel_bytes = gerar_excel_op(df_op, df_resumo, df_excecoes)
        st.download_button(
            label="📥 Baixar OP de Injeção Profissional (Excel Estruturado)",
            data=excel_bytes,
            file_name="OP_Injecao_PU_Estruturada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
else:
    st.info("💡 Faça o upload da planilha ou PDF de programação para gerar automaticamente a OP estruturada com cabeçalho corporativo, separação por grupos e sumários.")