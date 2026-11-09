mport os
import glob
import re
from pypdf import PdfReader
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from google.colab import files

def extrair_dados_orfa_pdf(caminho_pdf):
    reader = PdfReader(caminho_pdf)
    texto = ""
    for page in reader.pages:
        texto += page.extract_text() + "\n"

    # 1. Extração Ultra-Flexível de Metadados (Pula quebras de linha)
    orfa_match = re.search(r'N[º°]?\s*ORFA:\s*\n?([^\n]+)', texto, re.IGNORECASE)
    orfa_str = orfa_match.group(1).strip() if orfa_match else os.path.basename(caminho_pdf)

    cliente_match = re.search(r'CLIENTE:\s*\n?([^\n]+)', texto, re.IGNORECASE)
    cliente_str = cliente_match.group(1).strip() if cliente_match else "NÃO INFORMADO"

    pedido_match = re.search(r'N[º°]?\s*PEDIDO:\s*\n?([^\n]+)', texto, re.IGNORECASE)
    pedido_str = pedido_match.group(1).strip() if pedido_match else "S/N"

    entrega_match = re.search(r'DATA\s*ENTREGA:\s*\n?([^\n]+)', texto, re.IGNORECASE)
    entrega_str = entrega_match.group(1).strip() if entrega_match else "-"

    linhas = [l.strip() for l in texto.split("\n") if l.strip()]
    equipamentos = []

    # 2. Busca de Equipamentos Sem Rigidez Extrema
    # Procura por padrões do tipo: QTD + NOME MODELO + ( ENDEREÇO )
    padrao_item = re.compile(r'(\d{1,2})\s+([A-Z0-9\+\s\/]{2,30}?)\s*\(\s*([\d\.\s\-]+)\s*\)', re.IGNORECASE)

    palavras_descartar = ["INTERMEDIARIA", "PRATELEIRAS", "CARENAGEM", "ANTEPARO", "RODAPE", "DIVISORIA", "RESFRIADOS", "HORTIFRUTI", "LATICINIOS", "CARNES"]

    for i, linha in enumerate(linhas):
        match = padrao_item.search(linha)

        if match:
            qtd = match.group(1).zfill(2)
            modelo = match.group(2).strip()
            end = match.group(3).strip()

            # Descarta linhas de cabeçalho de categoria ou acessórios
            if any(p in modelo.upper() for p in palavras_descartar) or len(modelo) < 3:
                continue

            # Bloco de texto ao redor do item para pegar os acabamentos
            bloco_contexto = " ".join(linhas[max(0, i-2):min(len(linhas), i+35)])

            # Acabamento Testeira/Painel
            tt_match = re.search(r'Acb\.\s*TT/PNL:\s*([^\n\t]+?)(?=Acb\.|Ach\.|Bojo|Grelha|Caren\.|Porta|Informações|Fluido|$)', bloco_contexto, re.IGNORECASE)
            tt_pnl = tt_match.group(1).strip() if tt_match else "INOX"

            # Carenagem Frontal
            caren_match = re.search(r'Caren\.\s*Frontal:\s*([^\n\t]+?)(?=Porta|Serigrafia|Testeira|Informações|Fluido|$)', bloco_contexto, re.IGNORECASE)
            caren = caren_match.group(1).strip() if caren_match else "PRETO (P1902)"

            # Acabamento Externo
            ext_match = re.search(r'Acb\.\s*Externo:\s*([^\n\t]+?)(?=Bojo|Grelha|Perf\.|Caren\.|Porta|Informações|Fluido|$)', bloco_contexto, re.IGNORECASE)
            ext = ext_match.group(1).strip() if ext_match else "PRETO (P1902)"

            # Fluido / Gás (Ex: R22, GLICOL, R404)
            gas_match = re.search(r'Fluido/Gás:\s*([^\n\t]+?)(?=Válvula|Micromotor|Cx\.|Evaporador|Acessórios|$)', bloco_contexto, re.IGNORECASE)
            gas = gas_match.group(1).strip() if (gas_match and gas_match.group(1).strip()) else "GLICOL"

            # Iluminação
            ilum_match = re.search(r'Iluminação:\s*([^\n\t]+?)(?=Acessórios|PRATELEIRAS|ANTEPARO|Observações|$)', bloco_contexto, re.IGNORECASE)
            ilum = ilum_match.group(1).strip() if (ilum_match and ilum_match.group(1).strip()) else "3000K"

            # Evita duplicatas consecutivas idênticas
            if not any(e['modelo'] == modelo and e['end'] == end for e in equipamentos):
                equipamentos.append({
                    'qtd': qtd,
                    'modelo': modelo,
                    'end': end,
                    'tt_pnl': tt_pnl,
                    'caren': caren,
                    'ext': ext,
                    'gas': gas,
                    'ilum': ilum
                })

    return {
        'orfa': orfa_str,
        'cliente': cliente_str,
        'pedido': pedido_str,
        'entrega': entrega_str,
        'equipamentos': equipamentos
    }


def gerar_relatorio_consolidado(lista_orfas, nome_saida="RELATORIO_CONSOLIDADO_ORFAS.pdf"):
    doc = SimpleDocTemplate(
        nome_saida,
        pagesize=A4,
        rightMargin=15,
        leftMargin=15,
        topMargin=15,
        bottomMargin=15
    )
    story = []
    styles = getSampleStyleSheet()

    titulo_principal = ParagraphStyle('TP', parent=styles['Heading1'], fontSize=14, leading=16, textColor=colors.HexColor('#1A2B4C'), alignment=1)
    sub_titulo = ParagraphStyle('ST', parent=styles['Normal'], fontSize=9, leading=11, textColor=colors.HexColor('#4A5568'), alignment=1)
    orfa_header_style = ParagraphStyle('OH', parent=styles['Heading2'], fontSize=10, leading=13, textColor=colors.HexColor('#2B6CB0'), spaceBefore=6, spaceAfter=2)
    info_style = ParagraphStyle('INF', parent=styles['Normal'], fontSize=8, leading=10, textColor=colors.HexColor('#2D3748'))

    cell_bold = ParagraphStyle('CB', fontSize=7.5, leading=9, fontName='Helvetica-Bold')
    cell_norm = ParagraphStyle('CN', fontSize=7.5, leading=9, fontName='Helvetica')
    cell_highlight = ParagraphStyle('CH', fontSize=7.5, leading=9, fontName='Helvetica-Bold', textColor=colors.HexColor('#C53030'))
    cell_inox = ParagraphStyle('CI', fontSize=7.5, leading=9, fontName='Helvetica-Bold', textColor=colors.HexColor('#2B6CB0'))

    story.append(Paragraph("<b>RELATÓRIO CONSOLIDADO DE SEPARAÇÃO DE PEÇAS</b>", titulo_principal))
    story.append(Paragraph(f"Total de ORFAs no lote: <b>{len(lista_orfas)}</b>", sub_titulo))
    story.append(Spacer(1, 8))

    for idx, dados in enumerate(lista_orfas):
        story.append(Paragraph(f"📌 <b>ORFA: {dados['orfa']}</b> | Pedido: {dados['pedido']} | Entrega: {dados['entrega']}", orfa_header_style))
        story.append(Paragraph(f"<b>Cliente:</b> {dados['cliente']}", info_style))
        story.append(Spacer(1, 4))

        headers = ['Qtd', 'Item / Modelo', 'End.', '🚨 ACB. TT/PNL', '🚨 CAREN. FRONTAL', '🚨 ACB. EXTERNO', '❄️ GÁS', '💡 ILUM.']
        table_data = [[Paragraph(f"<b>{h}</b>", cell_bold) for h in headers]]

        if dados['equipamentos']:
            for eq in dados['equipamentos']:
                row = [
                    Paragraph(eq['qtd'], cell_norm),
                    Paragraph(eq['modelo'], cell_bold),
                    Paragraph(eq['end'], cell_norm),
                    Paragraph(f"<b>{eq['tt_pnl']}</b>", cell_highlight if "BRANCO" in eq['tt_pnl'].upper() else cell_inox),
                    Paragraph(f"<b>{eq['caren']}</b>", cell_norm),
                    Paragraph(f"<b>{eq['ext']}</b>", cell_norm),
                    Paragraph(f"<b>{eq['gas']}</b>", cell_bold),
                    Paragraph(f"<b>{eq['ilum']}</b>", cell_bold)
                ]
                table_data.append(row)
        else:
            table_data.append([
                Paragraph("-", cell_norm),
                Paragraph("Nenhum expositor identificado", cell_norm),
                Paragraph("-", cell_norm), Paragraph("-", cell_norm), Paragraph("-", cell_norm),
                Paragraph("-", cell_norm), Paragraph("-", cell_norm), Paragraph("-", cell_norm)
            ])

        col_widths = [25, 125, 55, 80, 80, 80, 50, 45]
        t = Table(table_data, colWidths=col_widths)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#EDF2F7')),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E0')),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]))
        story.append(t)
        story.append(Spacer(1, 8))

    doc.build(story)

# --- EXECUÇÃO AUTOMÁTICA ---
arquivos_pdf = glob.glob("*.pdf")
arquivos_pdf = [f for f in arquivos_pdf if "RELATORIO_" not in f]

if not arquivos_pdf:
    print("❌ NENHUM ARQUIVO PDF ENCONTRADO NA PASTA LATERAL!")
    print("👉 Arraste os arquivos PDF para a pasta lateral esquerda (ícone 📁) e aperte o Play ▶️ novamente.")
else:
    print(f"🔍 Processando {len(arquivos_pdf)} arquivo(s)...")
    lista_orfas = [extrair_dados_orfa_pdf(f) for f in arquivos_pdf]

    nome_pdf = "RELATORIO_CONSOLIDADO_ORFAS.pdf"
    gerar_relatorio_consolidado(lista_orfas, nome_pdf)

    print("✅ PDF Consolidado gerado com sucesso!")
    print("⬇️ Baixando arquivo...")
    files.download(nome_pdf)