import streamlit as st
import pandas as pd
import openpyxl
import re
import math
import io

st.markdown("""
    <div style="background-color: #1F497D; color: white; padding: 15px; border-radius: 6px; text-align: center; font-size: 20px; font-weight: bold;">
        📦 CÁLCULO DE BANDEJAS ABS E MATÉRIA-PRIMA PEAD
    </div>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Arraste ou selecione o seu ficheiro Excel de produção (.xlsx, .xls)", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.xls'):
            df = pd.read_excel(uploaded_file, engine='xlrd')
        else:
            df = pd.read_excel(uploaded_file, engine='openpyxl')
    except Exception as e:
        df = pd.read_excel(uploaded_file)

    df.columns = [str(col).strip().upper() for col in df.columns]

    FATORES_BANDEJAS = {
        'VCA2P': {'C790': 4, 'C530': 8, 'C410': 0},
        'VRA2P': {'C790': 4, 'C530': 0, 'C410': 8},
        'VRA1P': {'C790': 4, 'C530': 0, 'C410': 4},
        'VRAG':  {'C790': 4, 'C530': 0, 'C410': 0},
        'VCAG':  {'C790': 4, 'C530': 0, 'C410': 0},
    }
    MODELOS_VALIDOS = list(FATORES_BANDEJAS.keys())

    def extrair_modulos_por_modelo(texto_alinhamento, dimensao_total):
        texto = str(texto_alinhamento).upper().replace(',', '.')
        modulos_por_modelo = {}
        encontrados = re.findall(r'([A-Z0-9]+)\s*\(([\d\.]+)\)', texto)

        if encontrados:
            for modelo, tamanho_str in encontrados:
                if modelo in MODELOS_VALIDOS:
                    try:
                        metragem = float(tamanho_str)
                        mods = metragem / 1.25
                        modulos_por_modelo[modelo] = modulos_por_modelo.get(modelo, 0.0) + mods
                    except:
                        pass

        if not modulos_por_modelo:
            for modelo in MODELOS_VALIDOS:
                if modelo in texto:
                    modulos_por_modelo[modelo] = dimensao_total / 1.25 if dimensao_total > 0 else 0
                    break
        return modulos_por_modelo

    lista_detalhada = []
    for _, row in df.iterrows():
        cliente = str(row.get('NAME', row.get('CLIENTE', 'N/A'))).strip()
        orfa = str(row.get('ORFA', 'N/A')).strip()
        linha = str(row.get('LINHA', 'N/A')).strip()
        dimensao_val = row.get('DIMENSÃO', row.get('DIMENSAO', row.get('TAMANHO', 0)))
        alinhamento = str(row.get('ALINHAMENTO', row.get('MODELO', ''))).strip().upper()

        try:
            dimensao = float(dimensao_val)
        except:
            dimensao = 0.0

        modulos_dict = extrair_modulos_por_modelo(alinhamento, dimensao)

        for modelo, mod_qtd in modulos_dict.items():
            fatores = FATORES_BANDEJAS.get(modelo, {'C790': 0, 'C530': 0, 'C410': 0})
            c790 = int(round(mod_qtd * fatores['C790']))
            c530 = int(round(mod_qtd * fatores['C530']))
            c410 = int(round(mod_qtd * fatores['C410']))
            total_bandejas = c790 + c530 + c410

            lista_detalhada.append({
                'CLIENTE': cliente,
                'ORFA': orfa,
                'LINHA': linha,
                'ALINHAMENTO': alinhamento,
                'MODELO': modelo,
                'MÓDULOS (1,25m)': mod_qtd,
                'BANDEJA L290 C790': c790,
                'BANDEJA L290 C530': c530,
                'BANDEJA L290 C410': c410,
                'TOTAL BANDEJAS ABS': total_bandejas
            })

    df_detalhado = pd.DataFrame(lista_detalhada)

    df_por_cliente = df_detalhado.groupby(['CLIENTE', 'ORFA', 'MODELO'], as_index=False).agg({
        'MÓDULOS (1,25m)': 'sum',
        'BANDEJA L290 C790': 'sum',
        'BANDEJA L290 C530': 'sum',
        'BANDEJA L290 C410': 'sum',
        'TOTAL BANDEJAS ABS': 'sum'
    })

    df_geral_modelo = df_detalhado.groupby('MODELO', as_index=False).agg({
        'MÓDULOS (1,25m)': 'sum',
        'BANDEJA L290 C790': 'sum',
        'BANDEJA L290 C530': 'sum',
        'BANDEJA L290 C410': 'sum',
        'TOTAL BANDEJAS ABS': 'sum'
    })

    tot_c790 = df_geral_modelo['BANDEJA L290 C790'].sum()
    tot_c530 = df_geral_modelo['BANDEJA L290 C530'].sum()
    tot_c410 = df_geral_modelo['BANDEJA L290 C410'].sum()

    placas_c790 = math.ceil(tot_c790 / 1.0)
    placas_c530 = math.ceil(tot_c530 / 1.0)
    placas_c410 = math.ceil(tot_c410 / 2.0)

    pead_data = [
        {'MEDIDA DA PLACA': '1000 × 510 mm', 'ESPESSURA': '5,5 mm', 'TIPO DE BANDEJA': 'L290 × C790 (REF. MAIOR)', 'QTD PLACAS': int(placas_c790), 'PESO UNITÁRIO (kg)': 2.50, 'PESO TOTAL (kg)': round(placas_c790 * 2.50, 2)},
        {'MEDIDA DA PLACA': '750 × 500 mm', 'ESPESSURA': '5,5 mm', 'TIPO DE BANDEJA': 'L290 × C530 (REF. MÉDIA)', 'QTD PLACAS': int(placas_c530), 'PESO UNITÁRIO (kg)': 1.84, 'PESO TOTAL (kg)': round(placas_c530 * 1.84, 2)},
        {'MEDIDA DA PLACA': '780 × 600 mm', 'ESPESSURA': '5,5 mm', 'TIPO DE BANDEJA': 'L290 × C410 (REF. MENOR - 2 un/placa)', 'QTD PLACAS': int(placas_c410), 'PESO UNITÁRIO (kg)': 2.29, 'PESO TOTAL (kg)': round(placas_c410 * 2.29, 2)}
    ]
    df_pead = pd.DataFrame(pead_data)

    total_peso_pead = df_pead['PESO TOTAL (kg)'].sum()

    st.write("---")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total de Clientes", df_por_cliente['CLIENTE'].nunique())
    col2.metric("Módulos Programados", f"{df_detalhado['MÓDULOS (1,25m)'].sum():.2f}")
    col3.metric("Total Bandejas ABS", int(df_detalhado['TOTAL BANDEJAS ABS'].sum()))
    col4.metric("Peso Total PEAD", f"{total_peso_pead:,.2f} kg")

    st.write("### 1. Separação de Matéria-Prima PEAD (Placas)")
    st.dataframe(df_pead, use_container_width=True)

    st.write("### 2. Resumo de Bandejas por Modelo de Expositor")
    st.dataframe(df_geral_modelo, use_container_width=True)

    st.write("### 3. Demanda Detalhada por Cliente e ORFA")
    st.dataframe(df_por_cliente, use_container_width=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_pead.to_excel(writer, sheet_name='Separacao_PEAD', index=False)
        df_geral_modelo.to_excel(writer, sheet_name='Resumo_Modelo', index=False)
        df_por_cliente.to_excel(writer, sheet_name='Detalhamento_Itens', index=False)
    processed_data = output.getvalue()

    st.download_button(
        label="📥 Descarregar Relatório Executivo Completo (.xlsx)",
        data=processed_data,
        file_name="Relatorio_Executivo_Bandejas_Com_PEAD.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )