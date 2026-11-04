import io
import pandas as pd
import streamlit as st

# --- CONSTANTES TÉCNICAS E CÁLCULO DE PESO ---
# Densidade do Aço Cristal 0,43mm: ~3.3759 kg/m²
DENSIDADE_M2 = 3.3759

PESO_CHAPA_2550 = (1.20 * 2.55) * DENSIDADE_M2  # ~10.33 kg/chapa
PESO_CHAPA_3780 = (1.20 * 3.78) * DENSIDADE_M2  # ~15.31 kg/chapa
PESO_CHAPA_4000 = (1.20 * 4.00) * DENSIDADE_M2  # ~16.20 kg/chapa


def calcular_chapas_va(df_input):
  """Aplica as regras de consumo de chapas cristal (0,43mm) com base no layout real da planilha de Montagem."""
  total_ch_2550 = 0
  total_ch_3780 = 0
  total_ch_4000 = 0

  detalhamento = []

  for idx, row in df_input.iterrows():
    # Mapeamento dinâmico para aceitar colunas da nova planilha ou colunas genéricas
    cliente = str(row.get("Name", row.get("Cliente / Projeto", "Cliente")))
    orfa = str(row.get("ORFA", row.get("OP / Pedido", f"OP-{idx+1}")))
    linha = str(row.get("LINHA", "VERTICAL ALTO"))
    alinhamento = str(row.get("ALINHAMENTO", "-"))

    # Tratamento da dimensão (Ex: 3.75, 2.50, 1.87, etc)
    dim_raw = row.get("DIMENSÃO", row.get("Modelo Expositor (VA)", "0"))
    dim_str = (
        str(dim_raw)
        .replace(",", ".")
        .strip()
        .rstrip("0")
        .rstrip(".")
        if "." in str(dim_raw)
        else str(dim_raw).strip()
    )

    # Cada linha da planilha representa 1 expositor montado (Qtd = 1 por linha)
    qtd = int(row.get("Quantidade (Unid)", 1))

    ch_2550 = 0
    ch_3780 = 0
    ch_4000 = 0
    observacao = ""

    # REGRAS DE NEGÓCIO PCP (CHAPAS CRISTAL 0,43mm):
    if dim_str in ["1.25", "1.5", "2.25", "2.5"]:
      # 1 bojo + 3 costas + 2 teto + 1 cx forçadora = 7 chapas 2550
      ch_2550 = 7 * qtd
      observacao = f"Uso de CH 2550 para {dim_str}m. Sobram retalhos no corte."

    elif dim_str in ["1.87", "3"]:
      # 1 bojo + 3 costas + 2 teto + 1 cx forçadora = 7 chapas 3780
      ch_3780 = 7 * qtd
      observacao = f"Uso integral de CH 3780 para {dim_str}m."

    elif dim_str in ["3.75"]:
      # 1 CH 4000 (bojo) + 6 CH 3780 (3 costas + 2 teto + 1 cx forçadora)
      ch_4000 = 1 * qtd
      ch_3780 = 6 * qtd
      observacao = "Mistura Padrão 3.75m: 1x CH 4000 (Bojo) + 6x CH 3780."

    else:
      observacao = f"⚠️ Dimensão '{dim_raw}' não reconhecida na regra VA."

    # Acumula Totais Gerais
    total_ch_2550 += ch_2550
    total_ch_3780 += ch_3780
    total_ch_4000 += ch_4000

    detalhamento.append({
        "ORFA / Pedido": orfa,
        "Cliente": cliente,
        "Linha": linha,
        "Dimensão (m)": dim_raw,
        "Alinhamento": alinhamento,
        "CH 2550 (0,43x1200x2550)": ch_2550,
        "CH 3780 (0,43x1200x3780)": ch_3780,
        "CH 4000 (0,43x1200x4000)": ch_4000,
        "Total Chapas": ch_2550 + ch_3780 + ch_4000,
        "Observação Técnica": observacao,
    })

  # Cálculo dos Pesos Totais
  peso_2550 = total_ch_2550 * PESO_CHAPA_2550
  peso_3780 = total_ch_3780 * PESO_CHAPA_3780
  peso_4000 = total_ch_4000 * PESO_CHAPA_4000
  peso_total_geral = peso_2550 + peso_3780 + peso_4000

  resumo_chapas = {
      "total_ch_2550": total_ch_2550,
      "total_ch_3780": total_ch_3780,
      "total_ch_4000": total_ch_4000,
      "total_chapas_geral": total_ch_2550 + total_ch_3780 + total_ch_4000,
      "peso_2550_kg": peso_2550,
      "peso_3780_kg": peso_3780,
      "peso_4000_kg": peso_4000,
      "peso_total_kg": peso_total_geral,
  }

  return pd.DataFrame(detalhamento), resumo_chapas


# --- INTERFACE EXECUTIVA STREAMLIT ---
st.title("📦 Módulo de Chapas Cristal - Expositores VA")
st.markdown(
    "Cálculo automatizado de blanks cristal 0,43mm (2550, 3780 e 4000) baseado"
    " na Planilha de Montagem Oficial."
)

st.markdown("---")

uploaded_file = st.file_uploader(
    "Envie a planilha de Montagem em Excel (.xlsx)", type=["xlsx"]
)

if uploaded_file:
  df_pedidos = pd.read_excel(uploaded_file)
  st.subheader("📋 Dados de Montagem Carregados")
  st.dataframe(df_pedidos.head(10), use_container_width=True)

  if st.button("🚀 Processar Cálculo de Chapas", type="primary"):
    df_resultado, resumo = calcular_chapas_va(df_pedidos)

    # --- METRICAS EXECUTIVAS ---
    st.markdown("---")
    st.subheader("📊 Indicadores Totais de Materiais Necessários")

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("CH 2550 mm", f"{resumo['total_ch_2550']} un")
    col2.metric("CH 3780 mm", f"{resumo['total_ch_3780']} un")
    col3.metric("CH 4000 mm", f"{resumo['total_ch_4000']} un")
    col4.metric(
        "Total de Chapas", f"{resumo['total_chapas_geral']} un", delta="Blanks"
    )
    col5.metric(
        "Peso Total Material",
        f"{resumo['peso_total_kg']:.2f} kg",
        delta="Aço 0,43mm",
    )

    st.markdown("---")
    st.subheader("📄 Detalhamento da Insumos por Item da Montagem")
    st.dataframe(df_resultado, use_container_width=True)

    # GERADOR DE RELATÓRIO EXCEL COMPLETO
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      df_resultado.to_excel(
          writer, index=False, sheet_name="Detalhamento_Chapas"
      )

      # Resumo Executivo para Compras e Almoxarifado
      df_resumo_export = pd.DataFrame([{
          "Chapa 0,43x1200x2550 mm (Qtd)": resumo["total_ch_2550"],
          "Peso CH 2550 (Kg)": round(resumo["peso_2550_kg"], 2),
          "Chapa 0,43x1200x3780 mm (Qtd)": resumo["total_ch_3780"],
          "Peso CH 3780 (Kg)": round(resumo["peso_3780_kg"], 2),
          "Chapa 0,43x1200x4000 mm (Qtd)": resumo["total_ch_4000"],
          "Peso CH 4000 (Kg)": round(resumo["peso_4000_kg"], 2),
          "TOTAL GERAL DE CHAPAS (UN)": resumo["total_chapas_geral"],
          "PESO TOTAL DE MATERIAL (KG)": round(resumo["peso_total_kg"], 2),
      }])
      df_resumo_export.to_excel(
          writer, index=False, sheet_name="Resumo_Material_PCP"
      )

    excel_data = output.getvalue()

    st.download_button(
        label="📥 Baixar Relatório Executivo de Chapas em Excel",
        data=excel_data,
        file_name="Relatorio_PCP_Chapas_Cristal_VA.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )