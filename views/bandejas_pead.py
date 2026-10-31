import io
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# --- TÍTULO DA PÁGINA ---
st.title("📊 Portal Executivo | Cálculo de Bandejas & PEAD")
st.markdown("Painel gerencial para acompanhamento de produção, pesos e distribuição de materiais.")
st.markdown("---")

# --- DADOS DE EXEMPLO / SIMULAÇÃO DE DADOS DO PCP ---
# (Aqui entram os seus dados reais de cálculo de bandejas)
@st.cache_data
def carregar_dados():
    dados = {
        "MEDIDA": ["1000 × 510", "750 × 500", "780 × 600", "1200 × 600", "900 × 500"],
        "ESPESSURA": ["5,5 mm", "5,5 mm", "5,5 mm", "5,5 mm", "5,5 mm"],
        "TIPO DE BANDEJA": ["L290 × C790", "L290 × C530", "L290 × C410", "L290 × C950", "L290 × C680"],
        "QTD PEÇAS": [548, 328, 104, 420, 215],
        "PESO UNIT (kg)": [2.50, 1.84, 2.29, 3.10, 2.10],
    }
    df = pd.DataFrame(dados)
    df["PESO TOTAL (kg)"] = df["QTD PEÇAS"] * df["PESO UNIT (kg)"]
    return df

df_resultado = carregar_dados()

# --- CARTOES DE MÉTRICAS (KPIs EXECUTIVOS) ---
col1, col2, col3 = st.columns(3)

total_pecas = df_resultado["QTD PEÇAS"].sum()
peso_total_geral = df_resultado["PESO TOTAL (kg)"].sum()
total_itens = len(df_resultado)

with col1:
    st.metric(label="📦 Total de Peças Produzidas", value=f"{total_pecas:,}".replace(",", "."))
with col2:
    st.metric(label="⚖️ Peso Total Geral", value=f"{peso_total_geral:,.2f} kg".replace(",", "X").replace(".", ",").replace("X", "."))
with col3:
    st.metric(label="📋 Itens / Medidas Atendidas", value=total_itens)

st.markdown("---")

# --- SEÇÃO DA TABELA VISUAL NA TELA ---
st.subheader("📋 Detalhamento Executivo por Medida e Lote")
st.dataframe(df_resultado, use_container_width=True, hide_index=True)

# --- GRÁFICO EXECUTIVO ---
st.subheader("📈 Distribuição de Peso Total por Medida")
st.bar_chart(df_resultado.set_index("MEDIDA")["PESO TOTAL (kg)"])

st.markdown("---")

# --- FUNÇÃO DE GERAÇÃO DO EXCEL PROFISSIONAL (OPENPYXL) ---
def gerar_excel_executivo(df):
    output = io.BytesIO()
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Relatório Executivo"

    # Garantir visual limpo ativando linhas de grade
    sheet.views.sheetView[0].showGridLines = True

    # Cores e Estilos Corporativos
    cor_cabecalho = "1F497D"  # Azul Executivo Escuro
    cor_linha_par = "F2F5F9"  # Azul bem claro para listras
    fonte_branca = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fonte_padrao = Font(name="Calibri", size=11)
    fonte_negrito = Font(name="Calibri", size=11, bold=True)
    
    borda_fina = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    # Título do Relatório na Planilha
    sheet.merge_cells("A1:F1")
    sheet["A1"] = "RELATÓRIO EXECUTIVO - PRODUÇÃO DE BANDEJAS & PEAD"
    sheet["A1"].font = Font(name="Calibri", size=14, bold=True, color="1F497D")
    sheet["A1"].alignment = Alignment(horizontal="center", vertical="center")
    sheet.row_dimensions[1].height = 30

    # Cabeçalho da Tabela (Linha 3)
    headers = list(df.columns)
    for col_num, header in enumerate(headers, 1):
        cell = sheet.cell(row=3, column=col_num)
        cell.value = header
        cell.font = fonte_branca
        cell.fill = PatternFill(start_color=cor_cabecalho, end_color=cor_cabecalho, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = borda_fina
    sheet.row_dimensions[3].height = 25

    # Preenchimento dos Dados
    for row_idx, row_data in enumerate(df.values, 4):
        sheet.row_dimensions[row_idx].height = 20
        is_par = (row_idx % 2 == 0)
        
        for col_idx, value in enumerate(row_data, 1):
            cell = sheet.cell(row=row_idx, column=col_idx)
            cell.value = value
            cell.font = fonte_padrao
            cell.border = borda_fina
            
            # Fundo zebrado elegante
            if is_par:
                cell.fill = PatternFill(start_color=cor_linha_par, end_color=cor_linha_par, fill_type="solid")

            # Alinhamentos e Formatos Numéricos
            if isinstance(value, (int, float)):
                cell.alignment = Alignment(horizontal="right", vertical="center")
                if isinstance(value, float):
                    cell.number_format = '#,##0.00'
                else:
                    cell.number_format = '#,##0'
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    # Linha de Totais
    total_row_idx = len(df) + 4
    sheet.row_dimensions[total_row_idx].height = 22
    
    cell_label = sheet.cell(row=total_row_idx, column=1)
    cell_label.value = "TOTAL GERAL"
    cell_label.font = fonte_negrito
    cell_label.alignment = Alignment(horizontal="left", vertical="center")
    cell_label.border = borda_fina

    for col_idx in range(2, len(headers) + 1):
        cell = sheet.cell(row=total_row_idx, column=col_idx)
        cell.border = borda_fina
        if col_idx == 4: # Coluna QTD PEÇAS
            col_letter = get_column_letter(col_idx)
            cell.value = f"=SUM({col_letter}4:{col_letter}{total_row_idx-1})"
            cell.font = fonte_negrito
            cell.number_format = '#,##0'
            cell.alignment = Alignment(horizontal="right", vertical="center")
        elif col_idx == 6: # Coluna PESO TOTAL
            col_letter = get_column_letter(col_idx)
            cell.value = f"=SUM({col_letter}4:{col_letter}{total_row_idx-1})"
            cell.font = fonte_negrito
            cell.number_format = '#,##0.00'
            cell.alignment = Alignment(horizontal="right", vertical="center")

    # Ajuste automático inteligente da largura das colunas
    for col in sheet.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row >= 3: # Ignorar título principal para o cálculo de largura
                val_str = str(cell.value or '')
                if len(val_str) > max_len:
                    max_len = len(val_str)
        sheet.column_dimensions[col_letter].width = max(max_len + 5, 15)

    workbook.save(output)
    output.seek(0)
    return output

# --- BOTÃO DE DOWNLOAD EXECUTIVO ---
st.markdown("### 📥 Exportar Relatório para a Diretoria")
excel_data = gerar_excel_executivo(df_resultado)

st.download_button(
    label="📥 Baixar Relatório Executivo em Excel (Impecável)",
    data=excel_data,
    file_name="Relatorio_Executivo_Bandejas_PEAD.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)