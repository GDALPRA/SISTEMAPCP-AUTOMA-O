import io
import math
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import streamlit as st

# --- CABEÇALHO EXECUTIVO SLIM (PADRÃO NSF DARK) ---
st.markdown(
    """
    <div style="margin-bottom: 20px;">
        <h2 style="margin: 0; color: #FFFFFF; font-size: 22px; font-weight: 700; display: flex; align-items: center; gap: 10px;">
            🛡️ Módulo de Chapas de Aço Inox — Expositores VA, VCA & VRA
        </h2>
        <p style="margin: 4px 0 0 0; color: #94A3B8; font-size: 13px;">
            Cálculo automatizado de blanks (Inox 304 e Inox 430) para Vertical Alto, Açougue VCA e VRA (Bojos, Complementos VCA, Painéis, Costas, Tampos, Colunas, Bandejas e Prateleiras).
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

# --- CONSTANTES TÉCNICAS E PESOS ---
DENSIDADE_INOX = 7.93

# Pesos Unitários dos Blanks de Inox (com códigos oficiais)
PESO_CH_157219 = 0.80 * 1.24 * 3.90 * DENSIDADE_INOX  # Cód 157219 (~30.68 kg)
PESO_CH_053583 = 0.60 * 1.24 * 2.56 * DENSIDADE_INOX  # Cód 053583 (~15.10 kg)
PESO_CH_104818 = 0.60 * 1.24 * 2.00 * DENSIDADE_INOX  # Cód 104818 (~11.80 kg)
PESO_CH_157228 = 0.60 * 1.24 * 3.84 * DENSIDADE_INOX  # Cód 157228 (~22.65 kg)
PESO_CH_156102 = 1.20 * 1.24 * 2.96 * DENSIDADE_INOX  # Cód 156102 (~27.60 kg)
PESO_CH_153549 = 0.40 * 1.24 * 2.54 * DENSIDADE_INOX  # Cód 153549 (~9.99 kg)
PESO_CH_156082 = 0.40 * 1.24 * 3.78 * DENSIDADE_INOX  # Cód 156082 (~14.87 kg)
PESO_CH_151947 = 0.40 * 1.24 * 2.55 * DENSIDADE_INOX  # Cód 151947 (~10.03 kg)
PESO_CH_157101 = 0.40 * 1.24 * 2.06 * DENSIDADE_INOX  # Cód 157101 (~8.10 kg)


def extrair_modulos_acougue(texto_alinhamento):
    import re
    texto = str(texto_alinhamento).upper()
    res = {"1P": 0, "2P": 0, "modulos_total": 0}
    matches = re.findall(r'(1P|2P|G)?\s*\(?([0-9]+[,\.][0-9]+)\)?', texto)
    if not matches:
        return res
    for tipo, medida_str in matches:
        try:
            medida = float(medida_str.replace(",", "."))
            modulos_125 = medida / 1.25
            tipo_limpo = tipo.strip()
            if "2P" in tipo_limpo:
                res["2P"] += modulos_125
            elif "1P" in tipo_limpo:
                res["1P"] += modulos_125
        except:
            pass
    return res


def calcular_inox_geral(df_input):
    detalhamento = []

    total_ch_157219 = 0.0
    total_ch_053583 = 0.0
    total_ch_104818 = 0.0
    total_ch_157228 = 0.0
    total_ch_156102 = 0.0
    total_ch_153549 = 0.0
    total_ch_156082 = 0.0
    total_ch_151947 = 0.0
    total_ch_157101 = 0.0

    for idx, row in df_input.iterrows():
        cliente = str(row.get("Name", row.get("Cliente / Projeto", "Cliente")))
        orfa = str(row.get("ORFA", row.get("OP / Pedido", f"OP-{idx+1}")))
        linha = str(row.get("LINHA", "VERTICAL ALTO")).upper()
        alinhamento = str(row.get("ALINHAMENTO", "-"))

        dim_raw = row.get("DIMENSÃO", row.get("Modelo Expositor (VA)", "0"))
        try:
            dim_float = float(str(dim_raw).replace(",", "."))
        except Exception:
            dim_float = 0.0

        qtd = int(row.get("Quantidade (Unid)", 1))
        is_vap = "VAP" in alinhamento.upper()
        is_vca = "VCA" in linha or "AÇOUGUE" in linha and "VRA" not in linha
        is_vra = "VRA" in linha
        is_acougue = is_vca or is_vra

        ch_157219_item = 0.0
        ch_053583_item = 0.0
        ch_104818_item = 0.0
        ch_157228_item = 0.0
        ch_156102_item = 0.0
        ch_153549_item = 0.0
        ch_156082_item = 0.0
        ch_151947_item = 0.0
        ch_157101_item = 0.0

        # ==========================================
        # 1. PRATELEIRAS, BANDEJAS E COMPLEMENTOS VCA (Inox 0.80mm - Chapa 3900 / Cód 157219)
        # ==========================================
        if is_vra:
            modulos_equiv = dim_float / 1.25 if dim_float > 0 else 1.0
            if dim_float in [1.87, 1.875]:
                bandejas_tot = (2 * qtd)
            else:
                bandejas_tot = (modulos_equiv * 3 * qtd)
            ch_157219_item += bandejas_tot / 8.0

            info_prat = extrair_modulos_acougue(alinhamento)
            if info_prat["1P"] > 0 or info_prat["2P"] > 0:
                total_prat_pecas = ((info_prat["1P"] * 1) + (info_prat["2P"] * 2)) * qtd
                ch_157219_item += total_prat_pecas / 6.0
            elif " 1P" in alinhamento.upper() or " 2P" in alinhamento.upper():
                n_niveis = 2 if " 2P" in alinhamento.upper() else 1
                modulos_total = round(dim_float / 1.25, 2)
                total_prat_pecas = (modulos_total * n_niveis) * qtd
                ch_157219_item += total_prat_pecas / 6.0

        elif is_vca:
            modulos_equiv = dim_float / 1.25 if dim_float > 0 else 1.0
            if dim_float in [1.87, 1.875]:
                bandejas_tot = (2 * qtd)
            else:
                bandejas_tot = (modulos_equiv * 3 * qtd)
            ch_157219_item += bandejas_tot / 8.0

            if dim_float == 2.0:
                complementos_tot = (1 * qtd)
                ch_157219_item += complementos_tot / 4.0
            elif dim_float == 2.5:
                complementos_tot = (2 * qtd)
                ch_157219_item += complementos_tot / 12.0
            elif dim_float == 3.75:
                complementos_tot = (3 * qtd)
                ch_157219_item += complementos_tot / 12.0
            elif dim_float in [1.87, 1.875]:
                complementos_tot = (1 * qtd)
                ch_157219_item += complementos_tot / 8.0
            else:
                complementos_tot = (1 * qtd)
                ch_157219_item += complementos_tot / 12.0

        elif not is_acougue:
            if dim_float == 1.87:
                prat_1245_tot = 10 * qtd
                ch_157219_item += prat_1245_tot / 12.0
                prat_610_tot = 10 * qtd
                ch_157219_item += prat_610_tot / 22.0
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
        # 2. BOJOS (Inox 0,40mm - VA, VCA & VRA)
        # ==========================================
        if dim_float in [1.25, 1.5, 2.0, 2.25, 2.5]:
            if dim_float == 1.25:
                ch_153549_item = math.ceil(qtd / 2.0)
            else:
                ch_153549_item = float(qtd)
        elif dim_float in [1.87, 1.875, 3.0, 3.75]:
            if dim_float in [1.87, 1.875]:
                ch_156082_item = math.ceil(qtd / 2.0)
            else:
                ch_156082_item = float(qtd)

        # ==========================================
        # 3. COSTAS, COLUNAS E TAMPOS (VRA & VCA)
        # ==========================================
        if is_vra:
            if dim_float in [2.0, 2.5]:
                ch_053583_item += qtd * 2.0
            elif dim_float == 1.25:
                ch_053583_item += float(qtd)
            elif dim_float == 3.75:
                ch_157228_item += qtd * 2.0
            elif dim_float in [1.87, 1.875]:
                ch_157228_item += float(qtd)

            if dim_float == 1.25:
                ch_156102_item += (2 * qtd) / 4.0
            elif dim_float in [1.87, 1.875, 2.0, 2.5]:
                ch_156102_item += (4 * qtd) / 4.0
            elif dim_float == 3.75:
                ch_156102_item += (6 * qtd) / 4.0

        if is_vca:
            if dim_float == 1.25:
                ch_156102_item += (2 * qtd) / 5.0
            elif dim_float in [1.87, 1.875, 2.0, 2.5]:
                ch_156102_item += (4 * qtd) / 5.0
            elif dim_float == 3.75:
                ch_156102_item += (6 * qtd) / 5.0

        if is_acougue:
            if dim_float == 2.5:
                ch_151947_item += qtd / 2.0
            elif dim_float == 1.25:
                ch_151947_item += qtd / 4.0
            elif dim_float == 3.75:
                ch_157101_item += float(qtd)
            elif dim_float in [2.0, 1.87, 1.875]:
                ch_157101_item += qtd / 2.0

            if dim_float == 2.5:
                ch_053583_item += qtd / 2.0
            elif dim_float == 1.25:
                ch_053583_item += qtd / 4.0
            elif dim_float == 3.75:
                ch_104818_item += float(qtd)
            elif dim_float in [2.0, 1.87, 1.875]:
                ch_104818_item += qtd / 2.0

        elif not is_acougue:
            if dim_float == 1.25:
                paineis_fundo = (1 * 3) * qtd
                ch_053583_item += paineis_fundo / 4.0
                ch_053583_item += qtd / 4.0
            elif dim_float == 1.5:
                paineis_075 = (2 * 3) * qtd
                ch_053583_item += paineis_075 / 6.0
                ch_053583_item += qtd / 2.0
            elif dim_float in [1.87, 1.875]:
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
        total_ch_157228 += ch_157228_item
        total_ch_156102 += ch_156102_item
        total_ch_153549 += ch_153549_item
        total_ch_156082 += ch_156082_item
        total_ch_151947 += ch_151947_item
        total_ch_157101 += ch_157101_item

        peso_item = (
            (ch_157219_item * PESO_CH_157219)
            + (ch_053583_item * PESO_CH_053583)
            + (ch_104818_item * PESO_CH_104818)
            + (ch_157228_item * PESO_CH_157228)
            + (ch_156102_item * PESO_CH_156102)
            + (ch_153549_item * PESO_CH_153549)
            + (ch_156082_item * PESO_CH_156082)
            + (ch_151947_item * PESO_CH_151947)
            + (ch_157101_item * PESO_CH_157101)
        )

        detalhamento.append(
            {
                "CLIENTE": cliente,
                "ORFA / OP": orfa,
                "LINHA": linha,
                "DIMENSÃO (m)": dim_raw,
                "ALINHAMENTO": alinhamento,
                "CÓD 157219": round(ch_157219_item, 2),
                "CÓD 053583": round(ch_053583_item, 2),
                "CÓD 104818": round(ch_104818_item, 2),
                "CÓD 157228": round(ch_157228_item, 2),
                "CÓD 156102": round(ch_156102_item, 2),
                "CÓD 153549": round(ch_153549_item, 2),
                "CÓD 156082": round(ch_156082_item, 2),
                "CÓD 151947": round(ch_151947_item, 2),
                "CÓD 157101": round(ch_157101_item, 2),
                "PESO TOTAL (kg)": round(peso_item, 2),
            }
        )

    peso_total_geral = (
        (total_ch_157219 * PESO_CH_157219)
        + (total_ch_053583 * PESO_CH_053583)
        + (total_ch_104818 * PESO_CH_104818)
        + (total_ch_157228 * PESO_CH_157228)
        + (total_ch_156102 * PESO_CH_156102)
        + (total_ch_153549 * PESO_CH_153549)
        + (total_ch_156082 * PESO_CH_156082)
        + (total_ch_151947 * PESO_CH_151947)
        + (total_ch_157101 * PESO_CH_157101)
    )

    resumo = {
        "ch_157219_un": round(total_ch_157219, 2),
        "peso_157219_kg": total_ch_157219 * PESO_CH_157219,
        "ch_053583_un": round(total_ch_053583, 2),
        "peso_053583_kg": total_ch_053583 * PESO_CH_053583,
        "ch_104818_un": round(total_ch_104818, 2),
        "peso_104818_kg": total_ch_104818 * PESO_CH_104818,
        "ch_157228_un": round(total_ch_157228, 2),
        "peso_3840_06_kg": total_ch_157228 * PESO_CH_157228,
        "ch_156102_un": round(total_ch_156102, 2),
        "peso_colunas_v_kg": total_ch_156102 * PESO_CH_156102,
        "ch_153549_un": round(total_ch_153549, 2),
        "peso_153549_kg": total_ch_153549 * PESO_CH_153549,
        "ch_156082_un": round(total_ch_156082, 2),
        "peso_3780_kg": total_ch_156082 * PESO_CH_156082,
        "ch_151947_un": round(total_ch_151947, 2),
        "peso_151947_kg": total_ch_151947 * PESO_CH_151947,
        "ch_157101_un": round(total_ch_157101, 2),
        "peso_157101_kg": total_ch_157101 * PESO_CH_157101,
        "total_chapas_geral": round(
            total_ch_157219 + total_ch_053583 + total_ch_104818 + total_ch_157228 + total_ch_156102 + total_ch_153549 + total_ch_156082 + total_ch_151947 + total_ch_157101,
            2,
        ),
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

    ws.cell(row=1, column=1, value="RELATÓRIO EXECUTIVO - PRODUÇÃO DE CHAPAS DE AÇO INOX (VA, VCA & VRA)").font = font_title
    ws.cell(row=2, column=1, value="PCP & Engenharia - Visão Estratégica para Diretoria e Compras").font = font_subtitle

    ws.cell(row=4, column=1, value="INDICADORES GERAIS DE INOX").font = font_section

    kpi_headers = [
        "TOTAL DE CHAPAS (UN)",
        "CÓD 157219",
        "CÓD 053583",
        "CÓD 104818",
        "CÓD 157228",
        "CÓD 156102",
        "CÓD 153549",
        "CÓD 156082",
        "CÓD 151947",
        "CÓD 157101",
        "PESO TOTAL (kg)",
    ]
    kpi_values = [
        resumo["total_chapas_geral"],
        resumo["ch_157219_un"],
        resumo["ch_053583_un"],
        resumo["ch_104818_un"],
        resumo["ch_157228_un"],
        resumo["ch_156102_un"],
        resumo["ch_153549_un"],
        resumo["ch_156082_un"],
        resumo["ch_151947_un"],
        resumo["ch_157101_un"],
        round(resumo["peso_total_kg"], 2),
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

    # Lista completa de todos os blanks candidatos
    candidatos_resumo = [
        ["157219", "CH 0.80 x 1240 x 3900 INOX 304", resumo["ch_157219_un"], round(resumo["peso_157219_kg"], 2)],
        ["053583", "CH 0.60 x 1240 x 2560 INOX 304", resumo["ch_053583_un"], round(resumo["peso_053583_kg"], 2)],
        ["104818", "CH 0.60 x 1240 x 2000 INOX 304", resumo["ch_104818_un"], round(resumo["peso_104818_kg"], 2)],
        ["157228", "CH 0.60 x 1240 x 3840 INOX 304", resumo["ch_157228_un"], round(resumo["peso_3840_06_kg"], 2)],
        ["156102", "CH 1.20 x 1240 x 2960 INOX 304", resumo["ch_156102_un"], round(resumo["peso_colunas_v_kg"], 2)],
        ["153549", "CH 0.40 x 1240 x 2540 INOX 304", resumo["ch_153549_un"], round(resumo["peso_153549_kg"], 2)],
        ["156082", "CH 0.40 x 1240 x 3780 INOX 304", resumo["ch_156082_un"], round(resumo["peso_3780_kg"], 2)],
        ["151947", "CH 0.40 x 1240 x 2550 INOX 430", resumo["ch_151947_un"], round(resumo["peso_151947_kg"], 2)],
        ["157101", "CH 0.40 x 1240 x 2060 INOX 430", resumo["ch_157101_un"], round(resumo["peso_157101_kg"], 2)],
    ]

    # Filtra apenas os blanks com quantidade maior que zero
    resumo_data = [item for item in candidatos_resumo if item[2] > 0]

    for r_idx, row_vals in enumerate(resumo_data, 11):
        for c_idx, val in enumerate(row_vals, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_body
            cell.border = thin_border
            if c_idx in [1, 3, 4]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    start_row = 11 + len(resumo_data) + 3
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
            if c_idx >= 6:
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
        elif c_idx >= 6:
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


# --- INTERFACE DE UPLOAD E PROCESSAMENTO ---
st.markdown("---")

uploaded_file = st.file_uploader("Envie a planilha de Montagem em Excel (.xlsx)", type=["xlsx"], key="inox_uploader")

if uploaded_file:
    df_pedidos = pd.read_excel(uploaded_file)
    st.subheader("📋 Dados de Montagem Carregados")
    st.dataframe(df_pedidos.head(10), use_container_width=True)

    if st.button("🚀 Processar Cálculo de Aço Inox Completo", type="primary"):
        df_resultado, resumo = calcular_inox_geral(df_pedidos)

        st.markdown("---")
        st.subheader("📊 Indicadores Totais de Inox Necessários (VA, VCA & VRA)")

        col1, col2, col3, col4, col5, col6, col7, col8, col9, col10 = st.columns(10)
        col1.metric("CÓD 157219", f"{resumo['ch_157219_un']} un")
        col2.metric("CÓD 053583", f"{resumo['ch_053583_un']} un")
        col3.metric("CÓD 104818", f"{resumo['ch_104818_un']} un")
        col4.metric("CÓD 157228", f"{resumo['ch_157228_un']} un")
        col5.metric("CÓD 156102", f"{resumo['ch_156102_un']} un")
        col6.metric("CÓD 153549", f"{resumo['ch_153549_un']} un")
        col7.metric("CÓD 156082", f"{resumo['ch_156082_un']} un")
        col8.metric("CÓD 151947", f"{resumo['ch_151947_un']} un")
        col9.metric("CÓD 157101", f"{resumo['ch_157101_un']} un")
        col10.metric("Peso Total", f"{resumo['peso_total_kg']:.2f} kg")

        st.markdown("---")
        st.subheader("📄 Detalhamento Completo por Item")
        st.dataframe(df_resultado, use_container_width=True)

        excel_data = gerar_excel_estilo_tubos(df_resultado, resumo)

        st.download_button(
            label="📥 Baixar Relatório Executivo de Aço Inox em Excel",
            data=excel_data,
            file_name="Relatorio_Executivo_Inox_Geral.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )