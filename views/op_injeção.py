import streamlit as st
import pandas as pd
import io
import re
import unicodedata
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
            Geração de Ordem de Produção unificada (Padrão Executivo), com correção definitiva de fundo branco para impressão e data exata.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

# --- ENTRADA DO RESPONSÁVEL DO PCP ---
responsavel_pcp = st.text_input("👤 Nome do Responsável pela Programação (PCP)", value="Gabriel Dalpra")

# --- FUNÇÃO AUXILIAR PARA NORMALIZAR TEXTO (REMOVER ACENTOS) ---
def normalizar_texto(texto):
    if not isinstance(texto, str):
        return str(texto)
    nfkd = unicodedata.normalize('NFKD', texto)
    return "".join([c for c in nfkd if not unicodedata.combining(c)]).strip().upper()

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
                row_upper = [normalizar_texto(str(c)) for c in r]
                if "ALINHAMENTO" in row_upper or "ORFA" in row_upper:
                    header_idx = idx
                    break
            
            headers = [normalizar_texto(str(c)) for c in linhas_extraidas[header_idx]]
            data = linhas_extraidas[header_idx + 1:]
            df_pdf = pd.DataFrame(data, columns=headers)
            df_pdf = df_pdf[df_pdf.iloc[:, 0].astype(str).str.upper() != headers[0]]
            return df_pdf
        except ImportError:
            st.error("A biblioteca `pdfplumber` não está instalada no ambiente. Adicione `pdfplumber` ao seu `requirements.txt` para leitura direta de PDFs.")
            st.stop()

# --- MOTOR DE REGRAS DE ENGENHARIA (INJEÇÃO DE PU) ---
def processar_regras_injecao(df_raw, linhas_selecionadas=None):
    df = df_raw.copy()
    
    novo_mapa_colunas = {}
    for c in df.columns:
        col_norm = normalizar_texto(str(c))
        if "DIMEN" in col_norm:
            col_norm = "DIMENSAO"
        novo_mapa_colunas[c] = col_norm
    df = df.rename(columns=novo_mapa_colunas)

    colunas_necessarias = ["ALINHAMENTO", "DIMENSAO", "BOJO"]
    for col in colunas_necessarias:
        if col not in df.columns:
            raise ValueError(f"Coluna obrigatória '{col}' não encontrada no arquivo. Colunas detectadas: {list(df.columns)}")

    df = df.dropna(subset=["ALINHAMENTO"]).copy()
    df["DIMENSAO"] = pd.to_numeric(df["DIMENSAO"].astype(str).str.replace(",", "."), errors="coerce").fillna(0.0)
    df["BOJO_MAT"] = df["BOJO"].fillna("NÃO ESPECIFICADO").astype(str).str.strip().str.upper()
    df["ALINHAMENTO_STR"] = df["ALINHAMENTO"].astype(str).str.strip().str.upper()
    df["ORFA_STR"] = df["ORFA"].astype(str).str.strip() if "ORFA" in df.columns else "-"
    df["LINHA_STR"] = df["LINHA"].astype(str).str.strip().str.upper() if "LINHA" in df.columns else "GERAL"
    df["CLIENTE_STR"] = df["NAME"].astype(str).str.strip() if "NAME" in df.columns else "-"

    pecas_padrao = []
    excecoes_supervisor = []
    expositores_unicos = []

    termos_especiais = ["CANTO", "CT 180", "CT180", "CENTRAL", "PONTA", "90", "180"]

    for idx, row in df.iterrows():
        alinhamento = row["ALINHAMENTO_STR"]
        dimensao = row["DIMENSAO"]
        mat_bojo = row["BOJO_MAT"]
        orfa = row["ORFA_STR"]
        linha_orig = row["LINHA_STR"]
        cliente = row["CLIENTE_STR"]

        if linhas_selecionadas and linha_orig not in linhas_selecionadas:
            continue

        modelo_limpo = re.sub(r"^\d+(\.\d+)?\s*", "", alinhamento).strip()
        is_especial = any(termo in modelo_limpo for termo in termos_especiais)

        if not is_especial:
            if modelo_limpo.startswith("VRA"):
                l_prod = "AÇOUGUE (VRA)"
            elif modelo_limpo.startswith("VCA"):
                l_prod = "AÇOUGUE (VCA)"
            elif modelo_limpo.startswith("VA"):
                l_prod = "VERTICAL ALTO (VA)"
            elif modelo_limpo.startswith("VR"):
                l_prod = "PADARIA (VR)"
            elif modelo_limpo.startswith("VIL"):
                l_prod = "LINHA VIL (LINEAR)"
            else:
                l_prod = linha_orig

            expositores_unicos.append({"LINHA_PRODUTO": l_prod, "DIMENSAO": dimensao, "ID_EXPOSITOR": idx})

        # 1. REGRA VRA (Açougue) -> Bojo, Costas VRA, Tampos VRA
        if modelo_limpo.startswith("VRA") and not is_especial:
            pecas_padrao.append({"LINHA_PRODUTO": "AÇOUGUE (VRA)", "GRUPO_PECA": "1. BOJOS", "TIPO_PECA": "BOJO VRA", "MODELO_PECA": f"BOJO VRA - {mat_bojo}", "MATERIAL": mat_bojo, "DIMENSAO": dimensao, "QTD_UNITARIA": 1, "ORFA": orfa, "CLIENTE": cliente, "ALINHAMENTO": alinhamento})
            pecas_padrao.append({"LINHA_PRODUTO": "AÇOUGUE (VRA)", "GRUPO_PECA": "2. COSTAS", "TIPO_PECA": "COSTAS VRA", "MODELO_PECA": "COSTAS VRA", "MATERIAL": "PADRÃO", "DIMENSAO": dimensao, "QTD_UNITARIA": 1, "ORFA": orfa, "CLIENTE": cliente, "ALINHAMENTO": alinhamento})
            pecas_padrao.append({"LINHA_PRODUTO": "AÇOUGUE (VRA)", "GRUPO_PECA": "3. TETOS - TAMPOS", "TIPO_PECA": "TAMPO VRA", "MODELO_PECA": "TAMPOS VRA", "MATERIAL": "PADRÃO", "DIMENSAO": dimensao, "QTD_UNITARIA": 1, "ORFA": orfa, "CLIENTE": cliente, "ALINHAMENTO": alinhamento})

        # 2. REGRA VCA (Açougue) -> Bojo, Tampos VCA
        elif modelo_limpo.startswith("VCA") and not is_especial:
            pecas_padrao.append({"LINHA_PRODUTO": "AÇOUGUE (VCA)", "GRUPO_PECA": "1. BOJOS", "TIPO_PECA": "BOJO VCA", "MODELO_PECA": f"BOJO VCA - {mat_bojo}", "MATERIAL": mat_bojo, "DIMENSAO": dimensao, "QTD_UNITARIA": 1, "ORFA": orfa, "CLIENTE": cliente, "ALINHAMENTO": alinhamento})
            pecas_padrao.append({"LINHA_PRODUTO": "AÇOUGUE (VCA)", "GRUPO_PECA": "3. TETOS - TAMPOS", "TIPO_PECA": "TAMPO VCA", "MODELO_PECA": "TAMPOS VCA", "MATERIAL": "PADRÃO", "DIMENSAO": dimensao, "QTD_UNITARIA": 1, "ORFA": orfa, "CLIENTE": cliente, "ALINHAMENTO": alinhamento})

        # 3. REGRA VA (Vertical Alto) -> Bojo, Costas VA, Tetos VA
        elif modelo_limpo.startswith("VA") and not is_especial:
            pecas_padrao.append({"LINHA_PRODUTO": "VERTICAL ALTO (VA)", "GRUPO_PECA": "1. BOJOS", "TIPO_PECA": "BOJO VA", "MODELO_PECA": f"BOJO VA - {mat_bojo}", "MATERIAL": mat_bojo, "DIMENSAO": dimensao, "QTD_UNITARIA": 1, "ORFA": orfa, "CLIENTE": cliente, "ALINHAMENTO": alinhamento})
            pecas_padrao.append({"LINHA_PRODUTO": "VERTICAL ALTO (VA)", "GRUPO_PECA": "2. COSTAS", "TIPO_PECA": "COSTAS VA", "MODELO_PECA": "COSTAS VA", "MATERIAL": "PADRÃO", "DIMENSAO": dimensao, "QTD_UNITARIA": 1, "ORFA": orfa, "CLIENTE": cliente, "ALINHAMENTO": alinhamento})
            pecas_padrao.append({"LINHA_PRODUTO": "VERTICAL ALTO (VA)", "GRUPO_PECA": "3. TETOS - TAMPOS", "TIPO_PECA": "TETO VA", "MODELO_PECA": "TETOS VA", "MATERIAL": "PADRÃO", "DIMENSAO": dimensao, "QTD_UNITARIA": 1, "ORFA": orfa, "CLIENTE": cliente, "ALINHAMENTO": alinhamento})

        # 4. REGRA VR (Padaria) -> Bojo VILXVR, Costas VRP, Teto VR
        elif modelo_limpo.startswith("VR") and not is_especial:
            pecas_padrao.append({"LINHA_PRODUTO": "PADARIA (VR)", "GRUPO_PECA": "1. BOJOS", "TIPO_PECA": "BOJO VILXVR", "MODELO_PECA": f"BOJO VILXVR - {mat_bojo}", "MATERIAL": mat_bojo, "DIMENSAO": dimensao, "QTD_UNITARIA": 1, "ORFA": orfa, "CLIENTE": cliente, "ALINHAMENTO": alinhamento})
            pecas_padrao.append({"LINHA_PRODUTO": "PADARIA (VR)", "GRUPO_PECA": "2. COSTAS", "TIPO_PECA": "COSTAS VR", "MODELO_PECA": "COSTAS VRP", "MATERIAL": "PADRÃO", "DIMENSAO": dimensao, "QTD_UNITARIA": 1, "ORFA": orfa, "CLIENTE": cliente, "ALINHAMENTO": alinhamento})
            pecas_padrao.append({"LINHA_PRODUTO": "PADARIA (VR)", "GRUPO_PECA": "3. TETOS - TAMPOS", "TIPO_PECA": "TETO VR", "MODELO_PECA": "TETOS VR", "MATERIAL": "PADRÃO", "DIMENSAO": dimensao, "QTD_UNITARIA": 1, "ORFA": orfa, "CLIENTE": cliente, "ALINHAMENTO": alinhamento})

        # 5. REGRA VIL LINEAR -> Bojo VILXVR, Costas VIL (2P/3P), Teto VIL
        elif modelo_limpo.startswith("VIL") and not is_especial and ("2P" in modelo_limpo or "3P" in modelo_limpo):
            nome_costas = "COSTAS VIL 3P" if "3P" in modelo_limpo else "COSTAS VIL 2P"
            pecas_padrao.append({"LINHA_PRODUTO": "LINHA VIL (LINEAR)", "GRUPO_PECA": "1. BOJOS", "TIPO_PECA": "BOJO VILXVR", "MODELO_PECA": f"BOJO VILXVR - {mat_bojo}", "MATERIAL": mat_bojo, "DIMENSAO": dimensao, "QTD_UNITARIA": 1, "ORFA": orfa, "CLIENTE": cliente, "ALINHAMENTO": alinhamento})
            pecas_padrao.append({"LINHA_PRODUTO": "LINHA VIL (LINEAR)", "GRUPO_PECA": "2. COSTAS", "TIPO_PECA": "COSTAS VIL", "MODELO_PECA": nome_costas, "MATERIAL": "PADRÃO", "DIMENSAO": dimensao, "QTD_UNITARIA": 1, "ORFA": orfa, "CLIENTE": cliente, "ALINHAMENTO": alinhamento})
            pecas_padrao.append({"LINHA_PRODUTO": "LINHA VIL (LINEAR)", "GRUPO_PECA": "3. TETOS - TAMPOS", "TIPO_PECA": "TETO VIL", "MODELO_PECA": "TETOS VIL", "MATERIAL": "PADRÃO", "DIMENSAO": dimensao, "QTD_UNITARIA": 1, "ORFA": orfa, "CLIENTE": cliente, "ALINHAMENTO": alinhamento})

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
    df_exp = pd.DataFrame(expositores_unicos)

    if not df_pecas.empty:
        df_op_consolidada = df_pecas.groupby(
            ["LINHA_PRODUTO", "GRUPO_PECA", "TIPO_PECA", "MODELO_PECA", "MATERIAL", "DIMENSAO"], as_index=False
        ).agg(
            QTD_UNITARIA=("QTD_UNITARIA", "first"),
            QTD_TOTAL=("QTD_UNITARIA", "sum"),
            ORFAS=("ORFA", lambda x: ", ".join(sorted(set(x))))
        ).sort_values(by=["GRUPO_PECA", "LINHA_PRODUTO", "MATERIAL", "DIMENSAO"])

        df_resumo_linha = df_pecas.groupby(
            ["LINHA_PRODUTO", "GRUPO_PECA"], as_index=False
        ).agg(
            TOTAL_PECAS=("QTD_UNITARIA", "sum")
        ).pivot(index="LINHA_PRODUTO", columns="GRUPO_PECA", values="TOTAL_PECAS").fillna(0).reset_index()
        
        for col in ["1. BOJOS", "2. COSTAS", "3. TETOS - TAMPOS"]:
            if col not in df_resumo_linha.columns:
                df_resumo_linha[col] = 0
                
        df_resumo_linha["TOTAL_GERAL"] = df_resumo_linha[["1. BOJOS", "2. COSTAS", "3. TETOS - TAMPOS"]].sum(axis=1)

        if not df_exp.empty:
            df_dimensoes_linha = df_exp.groupby(
                ["LINHA_PRODUTO", "DIMENSAO"], as_index=False
            ).agg(
                QTD_EXPOSITORES=("ID_EXPOSITOR", "count")
            ).sort_values(by=["LINHA_PRODUTO", "DIMENSAO"])
        else:
            df_dimensoes_linha = pd.DataFrame()

        df_rastreio = df_pecas[["ORFA", "CLIENTE", "LINHA_PRODUTO", "ALINHAMENTO", "DIMENSAO"]].drop_duplicates().sort_values(by=["ORFA", "LINHA_PRODUTO"])
    else:
        df_op_consolidada = pd.DataFrame()
        df_resumo_linha = pd.DataFrame()
        df_dimensoes_linha = pd.DataFrame()
        df_rastreio = pd.DataFrame()

    return df_op_consolidada, df_resumo_linha, df_dimensoes_linha, df_rastreio, df_excecoes

# --- GERADOR DE EXCEL UNIFICADO (COM CORREÇÃO DE FUNDO BRANCO PARA DADOS) ---
def gerar_excel_executivo(df_op, df_resumo, df_dimensoes, df_rastreio, df_excecoes, resp_pcp):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        wb = writer.book
        ws = wb.create_sheet(title="OP_Injecao_PU")
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])

        ws["A1"] = "EMPRESA:"
        ws["B1"] = "NSF TECNOLOGIA EM INSTALAÇÕES COMERCIAIS"
        ws["A2"] = "SETOR DE DESTINO:"
        ws["B2"] = "SETOR DE INJEÇÃO DE PU (POLIOL + ISO)"
        ws["A3"] = "EMITIDO POR (PCP):"
        ws["B3"] = resp_pcp
        ws["A4"] = "DATA / HORA:"
        ws["B4"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        title_font = Font(name="Calibri", size=11, bold=True, color="1E3A8A")
        val_font = Font(name="Calibri", size=11, color="000000")
        for r in range(1, 5):
            ws[f"A{r}"].font = title_font
            ws[f"B{r}"].font = val_font

        current_row = 7

        # Armazenaremos os números de linha exatos onde estão os cabeçalhos das tabelas
        linhas_cabecalho = set()

        def adicionar_tabela(df_tabela, titulo_secao):
            nonlocal current_row
            ws[f"A{current_row - 1}"] = titulo_secao
            ws[f"A{current_row - 1}"].font = Font(name="Calibri", size=12, bold=True, color="1E3A8A")
            if not df_tabela.empty:
                df_tabela.to_excel(writer, index=False, sheet_name="OP_Injecao_PU", startrow=current_row - 1, startcol=0)
                linhas_cabecalho.add(current_row)
                current_row += len(df_tabela) + 3
            else:
                current_row += 3

        adicionar_tabela(df_op, "1. RELATÓRIO DETALHADO DE ORDENS DE PRODUÇÃO (BOJOS, COSTAS, TETOS-TAMPOS)")
        adicionar_tabela(df_resumo, "2. RESUMO CONSOLIDADO POR LINHA DE PRODUÇÃO")
        adicionar_tabela(df_dimensoes, "3. SOMATÓRIO DE EXPOSITORES PROGRAMADOS POR LINHA E DIMENSÃO")
        adicionar_tabela(df_rastreio, "4. RASTREABILIDADE: ORIGEM DOS EXPOSITORES")
        
        # Seção de Exceções
        ws[f"A{current_row - 1}"] = "5. ITENS ESPECIAIS / ATENÇÃO DO SUPERVISOR (CANTOS, CT 180, CENTRAIS)"
        ws[f"A{current_row - 1}"].font = Font(name="Calibri", size=12, bold=True, color="B45309")
        if not df_excecoes.empty:
            df_excecoes.to_excel(writer, index=False, sheet_name="OP_Injecao_PU", startrow=current_row - 1, startcol=0)
            linhas_cabecalho.add(current_row)
        else:
            pd.DataFrame({"Status": ["Nenhum expositor especial/exceção nesta programação."]}).to_excel(writer, index=False, sheet_name="OP_Injecao_PU", startrow=current_row - 1, startcol=0)
            linhas_cabecalho.add(current_row)

        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        header_font = Font(name="Calibri", size=11, color="FFFFFF", bold=True)
        data_font = Font(name="Calibri", size=11, color="000000")
        thin_border = Border(
            left=Side(style="thin", color="D1D5DB"),
            right=Side(style="thin", color="D1D5DB"),
            top=Side(style="thin", color="D1D5DB"),
            bottom=Side(style="thin", color="D1D5DB"),
        )

        ws_op = wb["OP_Injecao_PU"]
        for row in ws_op.iter_rows(min_row=1, max_row=ws_op.max_row):
            for cell in row:
                if cell.row in linhas_cabecalho:
                    # Linha exata de cabeçalho da tabela: Azul corporativo e letra branca
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif cell.row >= 7 and cell.value is not None and not str(cell.value).startswith("1. ") and not str(cell.value).startswith("2. ") and not str(cell.value).startswith("3. ") and not str(cell.value).startswith("4. ") and not str(cell.value).startswith("5. "):
                    # Linhas de dados: Fundo branco limpo, borda fina e fonte preta
                    cell.border = thin_border
                    cell.fill = white_fill
                    cell.font = data_font
                    cell.alignment = Alignment(horizontal="center", vertical="center")

        for col in ws_op.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = col[0].column_letter
            ws_op.column_dimensions[col_letter].width = max(max_len + 4, 16)

    return output.getvalue()

# --- INTERFACE DE USUÁRIO ---
uploaded_file = st.file_uploader(
    "📂 Importar Programação de Montagem (PDF ou Excel)",
    type=["pdf", "xlsx", "xls"],
)

if uploaded_file is not None:
    try:
        df_entrada = carregar_dados(uploaded_file)
        
        df_temp = df_entrada.copy()
        for c in df_temp.columns:
            if "LINHA" in normalizar_texto(str(c)):
                df_temp = df_temp.rename(columns={c: "LINHA"})
                break
        
        linhas_disponiveis = sorted(df_temp["LINHA"].astype(str).str.strip().str.upper().unique().tolist()) if "LINHA" in df_temp.columns else []

        st.markdown("---")
        if linhas_disponiveis:
            st.markdown("### 🎛️ Filtro de Linhas de Produção")
            linhas_selecionadas = st.multiselect(
                "Selecione as Linhas de Fábrica que deseja processar (deixe em branco para incluir todas):",
                options=linhas_disponiveis,
                default=linhas_disponiveis
            )
        else:
            linhas_selecionadas = []

        df_op, df_resumo, df_dimensoes, df_rastreio, df_excecoes = processar_regras_injecao(df_entrada, linhas_selecionadas)

        st.markdown("---")
        k1, k2, k3 = st.columns(3)
        total_pecas = int(df_op["QTD_TOTAL"].sum()) if not df_op.empty else 0
        total_bojos = int(df_op[df_op["GRUPO_PECA"] == "1. BOJOS"]["QTD_TOTAL"].sum()) if not df_op.empty else 0
        total_especiais = len(df_excecoes)

        k1.metric("Total de Peças para Injeção", f"{total_pecas} un.")
        k2.metric("Total de Bojos (Inox / PP)", f"{total_bojos} un.")
        k3.metric("Expositores Especiais / Atenção", f"{total_especiais} itens")

        st.markdown("---")

        tab_op, tab_resumo, tab_dim, tab_rastreio, tab_especiais = st.tabs([
            "📋 OP Detalhada por Grupos",
            "📊 Resumo por Linha",
            "📐 Somatório por Dimensão",
            "🔍 Rastreabilidade",
            f"⚠️ Especiais ({total_especiais})"
        ])

        with tab_op:
            if not df_op.empty:
                st.dataframe(df_op, use_container_width=True, hide_index=True)
            else:
                st.warning("Nenhuma peça padrão encontrada.")

        with tab_resumo:
            if not df_resumo.empty:
                st.dataframe(df_resumo, use_container_width=True, hide_index=True)
            else:
                st.warning("Sem dados para o resumo.")

        with tab_dim:
            if not df_dimensoes.empty:
                st.dataframe(df_dimensoes, use_container_width=True, hide_index=True)
            else:
                st.warning("Sem dados de dimensões.")

        with tab_rastreio:
            if not df_rastreio.empty:
                st.dataframe(df_rastreio, use_container_width=True, hide_index=True)
            else:
                st.warning("Sem dados de rastreabilidade.")

        with tab_especiais:
            if not df_excecoes.empty:
                st.warning("Itens especiais para análise do supervisor:")
                st.dataframe(df_excecoes, use_container_width=True, hide_index=True)
            else:
                st.success("✅ Nenhum item especial encontrado.")

        st.markdown("<br>", unsafe_allow_html=True)
        excel_bytes = gerar_excel_executivo(df_op, df_resumo, df_dimensoes, df_rastreio, df_excecoes, responsavel_pcp)
        st.download_button(
            label="📥 Baixar OP de Injeção Executiva (Corrigida)",
            data=excel_bytes,
            file_name="OP_Injecao_PU_Executiva.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
else:
    st.info("💡 Faça o upload da programação para gerar a OP executiva perfeitamente formatada.")