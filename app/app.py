import streamlit as st
import pandas as pd
import os
import xmltodict
from datetime import datetime
import shutil
from pathlib import Path
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER
import calendar
import zipfile
import io
import unicodedata
import xml.parsers.expat

# Configuração da página
st.set_page_config(
    page_title="NFCe Relatório",
    page_icon="📄",
    layout="wide"
)

# Logo no canto direito
col1, col2 = st.columns([8, 1])
with col1:
    st.write("")  # Espaço reservado
with col2:
    st.image("app/logo_empresa.png", width=100)

# Título da aplicação
st.title("📄 NFCe Relatório")
st.markdown("Faça upload de arquivos XML de NFCe para gerar relatórios.")

# Função para processar arquivo XML
def process_xml_file(xml_content, filename):
    try:
        try:
            xml_dict = xmltodict.parse(xml_content)
        except xml.parsers.expat.ExpatError as e:
            st.error(f"Arquivo {filename} está corrompido ou não é um XML válido: {str(e)}")
            return None
        
        # Função auxiliar para buscar uma chave ignorando namespace
        def find_key_ignore_ns(d, key):
            for k in d.keys():
                if k.split(':')[-1].lower() == key.lower() or k.split('}')[-1].lower() == key.lower():
                    return k
            return None
        
        # Verifica se é um arquivo de inutilização
        proc_inut_key = find_key_ignore_ns(xml_dict, 'ProcInutNFe')
        if proc_inut_key:
            inutil = xml_dict[proc_inut_key].get('inutNFe', {})
            inf_inut = inutil.get('infInut', {})
            ret_inut = xml_dict[proc_inut_key].get('retInutNFe', {}).get('infInut', {})
            
            # Número inicial e final
            nNFIni = inf_inut.get('nNFIni', '')
            nNFFin = inf_inut.get('nNFFin', '')
            faixa = f"{nNFIni} - {nNFFin}" if nNFIni != nNFFin else nNFIni
            
            # Data do evento
            data_emissao = ret_inut.get('dhRecbto', '')
            if data_emissao:
                try:
                    data_emissao = datetime.strptime(data_emissao[:19], '%Y-%m-%dT%H:%M:%S')
                except Exception:
                    data_emissao = data_emissao
            
            # Justificativa da inutilização
            justificativa = inf_inut.get('xJust', '')
            
            # Status e protocolo
            status = ret_inut.get('xMotivo', 'INUTILIZADO')
            protocolo = ret_inut.get('nProt', '')
            
            return {
                'Data Emissão': data_emissao,
                'Chave da Nota': faixa,
                'Número NFCe': nNFIni,
                'Destinatário': 'Não se aplica',
                'CPF/CNPJ Destinatário': 'Não se aplica',
                'Valor Total': 0.0,
                'Status': status,
                'Protocolo': protocolo,
                'Justificativa': justificativa
            }
            
        # Verifica se é um envelope de eventos (envEvento)
        env_evento_key = find_key_ignore_ns(xml_dict, 'envEvento')
        if env_evento_key:
            eventos = xml_dict[env_evento_key].get('evento', [])
            # Garante que eventos seja uma lista
            if not isinstance(eventos, list):
                eventos = [eventos]
            resultados = []
            for evento in eventos:
                inf_evento = evento.get('infEvento', {})
                tp_evento = str(inf_evento.get('tpEvento', '')).strip()
                if tp_evento in ['110111', '110112']:
                    chNFe = inf_evento.get('chNFe', '')
                    data_emissao = inf_evento.get('dhEvento', '')
                    if data_emissao:
                        try:
                            data_emissao = datetime.strptime(data_emissao[:19], '%Y-%m-%dT%H:%M:%S')
                        except Exception:
                            data_emissao = data_emissao
                    justificativa = inf_evento.get('detEvento', {}).get('xJust', '')
                    protocolo = inf_evento.get('detEvento', {}).get('nProt', '')
                    status = 'CANCELADO'
                    nNF = 'Não identificado'
                    if len(chNFe) == 44:
                        try:
                            nNF = chNFe[25:34]
                        except Exception:
                            nNF = 'Não identificado'
                    resultados.append({
                        'Data Emissão': data_emissao,
                        'Chave da Nota': chNFe,
                        'Número NFCe': nNF,
                        'Destinatário': 'Não se aplica',
                        'CPF/CNPJ Destinatário': 'Não se aplica',
                        'Valor Total': 0.0,
                        'Status': status,
                        'Protocolo': protocolo,
                        'Justificativa': justificativa
                    })
            if resultados:
                return resultados
            
        # Verifica se é um inutilizado simples (inutNFe)
        inutnfe_key = find_key_ignore_ns(xml_dict, 'inutNFe')
        if inutnfe_key:
            inutil = xml_dict[inutnfe_key]
            inf_inut = inutil.get('infInut', {})
            nNFIni = inf_inut.get('nNFIni', '')
            nNFFin = inf_inut.get('nNFFin', '')
            faixa = f"{nNFIni} - {nNFFin}" if nNFIni != nNFFin else nNFIni
            data_emissao = inf_inut.get('dhRecbto', '')
            if data_emissao:
                try:
                    data_emissao = datetime.strptime(data_emissao[:19], '%Y-%m-%dT%H:%M:%S')
                except Exception:
                    data_emissao = data_emissao
            justificativa = inf_inut.get('xJust', '')
            status = 'INUTILIZADO'
            protocolo = ''
            return {
                'Data Emissão': data_emissao,
                'Chave da Nota': faixa,
                'Número NFCe': nNFIni,
                'Destinatário': 'Não se aplica',
                'CPF/CNPJ Destinatário': 'Não se aplica',
                'Valor Total': 0.0,
                'Status': status,
                'Protocolo': protocolo,
                'Justificativa': justificativa
            }
            
        # Verifica se é um envelope de envio (enviNFe)
        envinfe_key = find_key_ignore_ns(xml_dict, 'enviNFe')
        if envinfe_key:
            nfes = xml_dict[envinfe_key].get('NFe', [])
            if not isinstance(nfes, list):
                nfes = [nfes]
            resultados = []
            for nfe in nfes:
                infNFe = nfe.get('infNFe', {})
                ide = infNFe.get('ide', {}) if isinstance(infNFe.get('ide', {}), dict) else {}
                data_emissao = ide.get('dhEmi', '')
                if data_emissao:
                    try:
                        data_emissao = datetime.strptime(data_emissao[:19], '%Y-%m-%dT%H:%M:%S')
                    except Exception:
                        data_emissao = data_emissao
                chNFe = infNFe.get('@Id', '').replace('NFe','') if infNFe.get('@Id', '') else ''
                destinatario = infNFe.get('dest', {}) if isinstance(infNFe.get('dest', {}), dict) else {}
                total = infNFe.get('total', {}) if isinstance(infNFe.get('total', {}), dict) else {}
                icms_tot = total.get('ICMSTot', {}) if isinstance(total.get('ICMSTot', {}), dict) else {}
                vnf = icms_tot.get('vNF', 0)
                try:
                    valor_total = float(vnf)
                except Exception:
                    valor_total = 0.0
                return_dict = {
                    'Data Emissão': data_emissao,
                    'Chave da Nota': chNFe,
                    'Número NFCe': ide.get('nNF', ''),
                    'Destinatário': destinatario.get('xNome', '') if destinatario else 'Consumidor não identificado',
                    'CPF/CNPJ Destinatário': destinatario.get('CPF', '') or destinatario.get('CNPJ', '') or 'Não informado',
                    'Valor Total': valor_total,
                    'Status': 'ENVIADO (SEM PROTOCOLO)',
                    'Protocolo': '',
                    'Justificativa': ''
                }
                resultados.append(return_dict)
            if resultados:
                return resultados
            
        # Verifica se é um evento processado (procEventoNFe)
        proc_evento_key = find_key_ignore_ns(xml_dict, 'procEventoNFe')
        if proc_evento_key:
            proc_evento = xml_dict[proc_evento_key]
            evento = proc_evento.get('evento', {})
            inf_evento = evento.get('infEvento', {})
            tp_evento = str(inf_evento.get('tpEvento', '')).strip()
            if tp_evento == '110111':  # Cancelamento
                chNFe = inf_evento.get('chNFe', '')
                data_emissao = inf_evento.get('dhEvento', '')
                if data_emissao:
                    try:
                        data_emissao = datetime.strptime(data_emissao[:19], '%Y-%m-%dT%H:%M:%S')
                    except Exception:
                        data_emissao = data_emissao
                justificativa = inf_evento.get('detEvento', {}).get('xJust', '')
                protocolo = inf_evento.get('detEvento', {}).get('nProt', '') or proc_evento.get('retEvento', {}).get('infEvento', {}).get('nProt', '')
                status = 'CANCELADO'
                nNF = 'Não identificado'
                if len(chNFe) == 44:
                    try:
                        nNF = chNFe[25:34]
                    except Exception:
                        nNF = 'Não identificado'
                return {
                    'Data Emissão': data_emissao,
                    'Chave da Nota': chNFe,
                    'Número NFCe': nNF,
                    'Destinatário': 'Não se aplica',
                    'CPF/CNPJ Destinatário': 'Não se aplica',
                    'Valor Total': 0.0,
                    'Status': status,
                    'Protocolo': protocolo,
                    'Justificativa': justificativa
                }
            
        # Processamento de NFCe normal
        nfeproc_key = find_key_ignore_ns(xml_dict, 'nfeProc')
        if not nfeproc_key:
            st.warning(f"Arquivo {filename} não contém informações de NFCe válidas")
            return None
        nfeproc = xml_dict[nfeproc_key]
        nfe_key = find_key_ignore_ns(nfeproc, 'NFe')
        if not nfe_key:
            st.warning(f"Arquivo {filename} não contém informações de NFCe válidas")
            return None
        nfce = nfeproc[nfe_key]
        infNFe_key = find_key_ignore_ns(nfce, 'infNFe')
        if not infNFe_key:
            st.warning(f"Arquivo {filename} não contém informações de NFCe válidas")
            return None
        infNFe = nfce[infNFe_key]
        data_emissao = infNFe.get('ide', {}).get('dhEmi', '')
        if data_emissao:
            try:
                data_emissao = datetime.strptime(data_emissao[:19], '%Y-%m-%dT%H:%M:%S')
            except Exception:
                data_emissao = data_emissao
        protNFe_key = find_key_ignore_ns(nfeproc, 'protNFe')
        prot = nfeproc.get(protNFe_key, {}).get('infProt', {}) if protNFe_key else {}
        destinatario = infNFe.get('dest', {})
        total = infNFe.get('total', {}).get('ICMSTot', {})
        emitente = infNFe.get('emit', {})
        return {
            'Data Emissão': data_emissao,
            'Chave da Nota': prot.get('chNFe', ''),
            'Número NFCe': infNFe.get('ide', {}).get('nNF', ''),
            'Destinatário': destinatario.get('xNome', '') if destinatario else 'Consumidor não identificado',
            'CPF/CNPJ Destinatário': destinatario.get('CPF', '') or destinatario.get('CNPJ', '') or 'Não informado',
            'Valor Total': float(total.get('vNF', 0)),
            'Status': prot.get('xMotivo', ''),
            'Protocolo': prot.get('nProt', ''),
            'Justificativa': ''
        }
    except Exception as e:
        st.error(f"Erro ao processar arquivo {filename}: {str(e)}")
        import traceback
        st.text(traceback.format_exc())
        return None

# Função para gerar PDF
def generate_pdf(df, totais, output_path):
    doc = SimpleDocTemplate(
        output_path,
        pagesize=landscape(A4),
        leftMargin=20,
        rightMargin=20,
        topMargin=20,
        bottomMargin=20
    )
    styles = getSampleStyleSheet()
    elements = []
    
    # Estilo personalizado para títulos centralizados
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=16,
        spaceAfter=30,
        alignment=TA_CENTER
    )
    subtitle_style = ParagraphStyle(
        'CustomSubtitle',
        parent=styles['Heading2'],
        fontSize=13,
        spaceAfter=10,
        alignment=TA_CENTER
    )
    
    # Adicionar logo e título na mesma linha usando uma tabela
    try:
        logo_path = 'app/logo_empresa.png'
        logo = Image(logo_path, width=60, height=60, hAlign='RIGHT')
        logo_col_width = 70 # Largura da coluna da logo, com um pequeno padding
        
        # Ajuste para centralizar o título visualmente, empurrando-o um pouco mais para a direita.
        # Este valor pode ser ajustado conforme a necessidade visual.
        offset_to_right = 0 # Valor em pontos para deslocar o título para a direita
        left_spacer_width = logo_col_width + offset_to_right 
        title_col_width = doc.width - left_spacer_width - logo_col_width 
        
        header_table = Table(
            [[Spacer(left_spacer_width, 1), Paragraph("Relatório de NFCe", title_style), logo]],
            colWidths=[left_spacer_width, title_col_width, logo_col_width]
        )

    except Exception:
        # Caso a logo não carregue, apenas o título centralizado
        header_table = Table(
            [[Paragraph("Relatório de NFCe", title_style)]],
            colWidths=[doc.width]
        )
    
    header_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        # Remover alinhamento específico para a coluna da logo, pois já está no Image e no colWidths
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))
    elements.append(header_table)
    elements.append(Spacer(1, 20))
    
    # Resumo centralizado
    elements.append(Paragraph("Resumo", subtitle_style))
    elements.append(Spacer(1, 10))
    
    # Criar tabela de resumo
    resumo_data = [['Métrica', 'Valor']]
    for _, row in totais.iterrows():
        resumo_data.append([row['Métrica'], row['Valor']])
    
    resumo_table = Table(resumo_data, colWidths=[3*inch, 2*inch])
    resumo_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    
    elements.append(resumo_table)
    elements.append(Spacer(1, 30))
    
    # Separar dados em normais, inutilizados e cancelados
    df_canceladas = df[df['Status'].str.upper().str.contains('CANCELADO', na=False)]
    df_inutilizados = df[df['Status'].str.upper().str.contains('INUTILIZADO', na=False)]
    df_normais = df[~df['Status'].str.upper().str.contains('CANCELADO|INUTILIZADO', na=False)]
    
    # Remover colunas indesejadas para o PDF
    # Agora incluindo a coluna 'Justificativa' para exibição
    colunas_pdf = [
        c for c in df.columns
        if c not in [
            'Emitente', 'CNPJ Emitente', 'Base ICMS', 'ICMS', 'PIS', 'COFINS', 'Total de Produtos'
        ]
    ]
    
    # Ajustar larguras das colunas
    # Adapte as larguras conforme as novas colunas e a necessidade de espaço
    col_widths = [1*inch, 2.5*inch, 0.6*inch, 1.1*inch, 0.8*inch, 0.8*inch, 1.2*inch, 1.2*inch, 1.5*inch] # Adicionado espaço para Justificativa
    if len(colunas_pdf) != len(col_widths):
        col_widths = [1.1*inch] * len(colunas_pdf) # Fallback se as larguras não baterem
    
    # Função auxiliar para criar tabela
    def create_table(df_subset, title):
        if len(df_subset) == 0:
            return None
        elements.append(Paragraph(title, subtitle_style))
        elements.append(Spacer(1, 10))
        df_pdf = df_subset[colunas_pdf]
        table_data = [[Paragraph(str(col), ParagraphStyle(name='HeaderCell', fontSize=9, alignment=TA_CENTER, leading=10, wordWrap='CJK')) for col in df_pdf.columns]]
        for _, row in df_pdf.iterrows():
            table_data.append([
                Paragraph(str(cell), ParagraphStyle(
                    name='TableCell', fontSize=7, alignment=TA_CENTER, leading=8, wordWrap='CJK'))
                for cell in row.tolist()
            ])
        table = Table(table_data, colWidths=col_widths, repeatRows=1)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
            ('TOPPADDING', (0, 0), (-1, 0), 6),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 7),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 2),
            ('TOPPADDING', (0, 1), (-1, 1), 2),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        elements.append(table)
        elements.append(Spacer(1, 20))
    
    # Criar tabelas separadas
    create_table(df_normais, "NFCe Emitidas")
    create_table(df_inutilizados, "NFCe Inutilizadas")
    create_table(df_canceladas, "NFCe Canceladas")
    
    # Gerar PDF
    doc.build(elements)

def normalize_str(s):
    if not isinstance(s, str):
        return ''
    return unicodedata.normalize('NFKD', s).encode('ASCII', 'ignore').decode('ASCII').upper()

def generate_pdf_resumido(df, output_path):
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
    from reportlab.lib.pagesizes import landscape, A4
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER
    from reportlab.lib.units import inch
    # Lista de meses em português
    meses_pt = [
        '', 'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
        'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
    ]

    doc = SimpleDocTemplate(
        output_path,
        pagesize=landscape(A4),
        leftMargin=20,
        rightMargin=20,
        topMargin=20,
        bottomMargin=20
    )
    styles = getSampleStyleSheet()
    elements = []

    # Título centralizado e logo
    try:
        logo_path = 'app/logo_empresa.png'
        logo = Image(logo_path, width=60, height=60, hAlign='RIGHT')
        logo_col_width = 70
        left_spacer_width = logo_col_width
        title_col_width = doc.width - left_spacer_width - logo_col_width
        header_table = Table(
            [[Spacer(left_spacer_width, 1), Paragraph("Relatório Resumido de NFCe por Mês", ParagraphStyle('ResumoTitle', parent=styles['Heading1'], fontSize=16, spaceAfter=30, alignment=TA_CENTER)), logo]],
            colWidths=[left_spacer_width, title_col_width, logo_col_width]
        )
    except Exception:
        header_table = Table(
            [[Paragraph("Relatório Resumido de NFCe por Mês", ParagraphStyle('ResumoTitle', parent=styles['Heading1'], fontSize=16, spaceAfter=30, alignment=TA_CENTER))]],
            colWidths=[doc.width]
        )
    header_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))
    elements.append(header_table)
    elements.append(Spacer(1, 20))

    # Normalizar status
    df = df.copy()
    df['StatusNorm'] = df['Status'].apply(normalize_str)

    # Separar inutilizadas
    inutilizadas_total = df[df['StatusNorm'].str.contains('INUTILIZADO|INUTILIZACAO', na=False)]
    df_util = df[~df['StatusNorm'].str.contains('INUTILIZADO|INUTILIZACAO', na=False)]

    # Garantir que a coluna Data Emissão está em datetime
    df_util['Data Emissão'] = pd.to_datetime(df_util['Data Emissão'], errors='coerce')
    df_util = df_util.dropna(subset=['Data Emissão'])
    df_util['AnoMes'] = df_util['Data Emissão'].dt.to_period('M')

    # Agrupar por mês (apenas autorizadas e canceladas)
    meses = sorted(df_util['AnoMes'].unique())
    total_autorizadas_ano = 0
    valor_autorizadas_ano = 0.0
    total_canceladas_ano = 0
    for mes in meses:
        mes_df = df_util[df_util['AnoMes'] == mes]
        ano = mes.year
        mes_num = mes.month
        mes_nome = f"{meses_pt[mes_num]}/{ano}"
        # Autorizadas
        autorizadas = mes_df[~mes_df['StatusNorm'].str.contains('CANCELADO', na=False)]
        total_autorizadas = len(autorizadas)
        valor_autorizadas = autorizadas['Valor Total'].sum()
        # Canceladas
        canceladas = mes_df[mes_df['StatusNorm'].str.contains('CANCELADO', na=False)]
        total_canceladas = len(canceladas)
        # Acumular totais
        total_autorizadas_ano += total_autorizadas
        valor_autorizadas_ano += valor_autorizadas
        total_canceladas_ano += total_canceladas
        # Tabela resumo do mês
        resumo_data = [
            ["Mês/Ano", mes_nome],
            ["Notas Autorizadas", total_autorizadas],
            ["Valor Total Autorizadas", f"R$ {valor_autorizadas:,.2f}"],
            ["Notas Canceladas", total_canceladas],
        ]
        table = Table(resumo_data, colWidths=[2.5*inch, 3*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        elements.append(table)
        elements.append(Spacer(1, 20))
    # Totalizador final (inutilizadas só no total geral)
    totalizador_data = [
        ["TOTAL DE NOTAS AUTORIZADAS", total_autorizadas_ano],
        ["VALOR TOTAL AUTORIZADAS", f"R$ {valor_autorizadas_ano:,.2f}"],
        ["TOTAL DE NOTAS INUTILIZADAS", len(inutilizadas_total)],
        ["TOTAL DE NOTAS CANCELADAS", total_canceladas_ano],
    ]
    totalizador_table = Table(totalizador_data, colWidths=[3*inch, 2.5*inch])
    totalizador_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#333333')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 11),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    elements.append(Spacer(1, 30))
    elements.append(Paragraph("Totalizador Geral", ParagraphStyle('TotalTitle', parent=styles['Heading2'], alignment=TA_CENTER, fontSize=14, spaceAfter=10)))
    elements.append(totalizador_table)
    doc.build(elements)

# Função para detectar numerações descontínuas
def detectar_numeracoes_puladas(df):
    """
    Detecta numerações descontínuas nas NFCe e retorna os números pulados
    """
    # Filtrar apenas NFCe normais (não canceladas, não inutilizadas)
    df_normais = df[~df['Status'].str.upper().str.contains('CANCELADO|INUTILIZADO', na=False)]
    
    if len(df_normais) == 0:
        return []
    
    # Converter números para inteiros e ordenar
    numeros_nfce = []
    for _, row in df_normais.iterrows():
        try:
            numero = int(row['Número NFCe'])
            numeros_nfce.append(numero)
        except (ValueError, TypeError):
            continue
    
    if len(numeros_nfce) < 2:
        return []
    
    numeros_nfce.sort()
    numeros_pulados = []
    
    # Verificar se há gaps na numeração
    for i in range(len(numeros_nfce) - 1):
        atual = numeros_nfce[i]
        proximo = numeros_nfce[i + 1]
        
        if proximo - atual > 1:
            # Há números pulados entre 'atual' e 'proximo'
            for numero_pulado in range(atual + 1, proximo):
                numeros_pulados.append(numero_pulado)
    
    return numeros_pulados

def agrupar_em_intervalos(numeros):
    """
    Recebe uma lista de inteiros e retorna uma lista de strings com intervalos agrupados.
    Exemplo: [2,3,4,5,6,7,8,9] -> ['2 até 9']
    """
    if not numeros:
        return []
    numeros = sorted(numeros)
    intervalos = []
    inicio = numeros[0]
    fim = numeros[0]
    for n in numeros[1:]:
        if n == fim + 1:
            fim = n
        else:
            if inicio == fim:
                intervalos.append(f"{inicio}")
            else:
                intervalos.append(f"{inicio} até {fim}")
            inicio = fim = n
    if inicio == fim:
        intervalos.append(f"{inicio}")
    else:
        intervalos.append(f"{inicio} até {fim}")
    return intervalos

# Função para gerar relatório de números pulados
def generate_pdf_numeros_pulados(numeros_pulados, output_path):
    """
    Gera um relatório PDF específico para números pulados
    """
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=20,
        rightMargin=20,
        topMargin=20,
        bottomMargin=20
    )
    styles = getSampleStyleSheet()
    elements = []
    
    # Estilo personalizado para títulos
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=16,
        spaceAfter=30,
        alignment=TA_CENTER
    )
    
    # Adicionar logo e título
    try:
        logo_path = 'app/logo_empresa.png'
        logo = Image(logo_path, width=60, height=60, hAlign='RIGHT')
        logo_col_width = 70
        left_spacer_width = logo_col_width
        title_col_width = doc.width - left_spacer_width - logo_col_width
        header_table = Table(
            [[Spacer(left_spacer_width, 1), Paragraph("Relatório de Numerações Puladas", title_style), logo]],
            colWidths=[left_spacer_width, title_col_width, logo_col_width]
        )
    except Exception:
        header_table = Table(
            [[Paragraph("Relatório de Numerações Puladas", title_style)]],
            colWidths=[doc.width]
        )
    
    header_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
    ]))
    elements.append(header_table)
    elements.append(Spacer(1, 20))
    
    if numeros_pulados:
        intervalos = agrupar_em_intervalos(numeros_pulados)
        data = [['Intervalos de Numeração Pulada']]
        for intervalo in intervalos:
            data.append([intervalo])
        table = Table(data, colWidths=[3*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.red),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        elements.append(table)
        elements.append(Spacer(1, 20))
        resumo_text = f"Total de numerações puladas encontradas: {len(numeros_pulados)}"
        elements.append(Paragraph(resumo_text, styles['Normal']))
    else:
        elements.append(Paragraph("Nenhuma numeração pulada foi encontrada.", styles['Normal']))
    
    # Gerar PDF
    doc.build(elements)

# Inicializar contador de reset no session_state
if 'reset_xml_upload' not in st.session_state:
    st.session_state['reset_xml_upload'] = 0

# Interface de upload com chave dinâmica
uploaded_files = st.file_uploader(
    "Escolha os arquivos XML ou ZIP",
    type=['xml', 'zip'],
    accept_multiple_files=True,
    key=f"xml_upload_{st.session_state['reset_xml_upload']}"
)

# Botão para remover arquivos enviados
if uploaded_files:
    if st.button("Remover arquivos enviados", type="primary"):
        st.session_state['reset_xml_upload'] += 1
        st.rerun()

if uploaded_files:
    # Criar diretório para relatórios se não existir
    os.makedirs('reports', exist_ok=True)
    
    # Lista para armazenar os dados processados
    data = []
    
    # Mostrar apenas um resumo do upload
    st.info(f"{len(uploaded_files)} arquivo(s) XML enviado(s).")
    
    # Processar cada arquivo XML
    with st.spinner('Processando arquivos XML...'):
        progress_bar = st.progress(0)
        total_files = len(uploaded_files)
        arquivos_nao_processados = []
        total_itens_processados = 0
        arquivos_para_processar = []
        for uploaded_file in uploaded_files:
            if uploaded_file.name.lower().endswith('.zip'):
                try:
                    with zipfile.ZipFile(io.BytesIO(uploaded_file.read())) as z:
                        for zipinfo in z.infolist():
                            # Ignorar arquivos ocultos do MacOS
                            if (
                                zipinfo.filename.startswith('__MACOSX') or
                                os.path.basename(zipinfo.filename).startswith('._')
                            ):
                                continue
                            if zipinfo.filename.lower().endswith('.xml'):
                                with z.open(zipinfo) as xmlfile:
                                    conteudo_xml = xmlfile.read().decode('utf-8', errors='ignore')
                                    arquivos_para_processar.append((conteudo_xml, zipinfo.filename))
                except Exception as e:
                    st.error(f"Erro ao descompactar {uploaded_file.name}: {str(e)}")
            elif uploaded_file.name.lower().endswith('.xml'):
                conteudo_xml = uploaded_file.read().decode('utf-8', errors='ignore')
                arquivos_para_processar.append((conteudo_xml, uploaded_file.name))
        total_files = len(arquivos_para_processar)
        progress_bar = st.progress(0)
        for i, (xml_content, nome_arquivo) in enumerate(arquivos_para_processar):
            try:
                progress = (i + 1) / total_files
                progress_bar.progress(progress)
                result = process_xml_file(xml_content, nome_arquivo)
                if result:
                    if isinstance(result, list):
                        data.extend(result)
                        total_itens_processados += len(result)
                    else:
                        data.append(result)
                        total_itens_processados += 1
                else:
                    arquivos_nao_processados.append(nome_arquivo)
            except Exception as e:
                st.error(f"Erro ao processar arquivo {nome_arquivo}: {str(e)}")
                arquivos_nao_processados.append(nome_arquivo)
        progress_bar.empty()
        st.success(f"Total de itens processados: {total_itens_processados}")
        if arquivos_nao_processados:
            st.warning(f"Os seguintes arquivos não foram processados: {', '.join(arquivos_nao_processados)}")
        
        if data:
            # Criar DataFrame
            df = pd.DataFrame(data)
            
            # Ordenar pelo valor numérico da chave da nota
            def chave_int(chave):
                try:
                    return int(chave)
                except Exception:
                    return 0
            df = df.sort_values(by='Chave da Nota', key=lambda x: x.apply(chave_int))
            
            # Exibir estatísticas
            st.subheader("Estatísticas")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total de NFCe", len(df))
            with col2:
                st.metric("Valor Total", f"R$ {df['Valor Total'].sum():,.2f}")
            with col3:
                st.metric("Média por NFCe", f"R$ {df['Valor Total'].mean():,.2f}")
            
            # Verificar numerações puladas
            numeros_pulados = detectar_numeracoes_puladas(df)
            
            # Alerta visual para numerações puladas
            if numeros_pulados:
                intervalos = agrupar_em_intervalos(numeros_pulados)
                st.error(f"⚠️ **ATENÇÃO:** Foram detectadas {len(numeros_pulados)} numeração(ões) pulada(s) nas NFCe!")
                st.warning(f"**Números pulados:** {', '.join(intervalos)}")
                st.info("Um relatório específico será gerado com os números pulados.")
            else:
                st.success("✅ **Verificação de numeração:** Todas as NFCe estão com numeração contínua.")
            
            # Exibir dados
            st.subheader("Dados das NFCe")
            st.dataframe(df)
            
            # Criar DataFrame com totais
            totais = pd.DataFrame({
                'Métrica': [
                    'Total de NFCe',
                    'Valor Total das Notas',
                    'Média por NFCe'
                ],
                'Valor': [
                    len(df),
                    f"R$ {df['Valor Total'].sum():,.2f}",
                    f"R$ {df['Valor Total'].mean():,.2f}"
                ]
            })
            
            # Exibir totais em formato de tabela
            st.table(totais)
            
            # Gerar relatório Excel
            excel_path = os.path.join('reports', f'relatorio_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
            
            # Criar um ExcelWriter
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                # Separar dados em normais, inutilizados e cancelados
                df_canceladas_excel = df[df['Status'].str.upper().str.contains('CANCELADO', na=False)]
                df_inutilizados_excel = df[df['Status'].str.upper().str.contains('INUTILIZADO', na=False)]
                df_normais_excel = df[~df['Status'].str.upper().str.contains('CANCELADO|INUTILIZADO', na=False)]
                
                # Escrever cada DataFrame em uma planilha separada
                df_normais_excel.to_excel(writer, sheet_name='NFCe Emitidas', index=False)
                df_inutilizados_excel.to_excel(writer, sheet_name='NFCe Inutilizadas', index=False)
                df_canceladas_excel.to_excel(writer, sheet_name='NFCe Canceladas', index=False)
                
                # Ajustar largura das colunas
                for sheet_name in writer.sheets:
                    worksheet = writer.sheets[sheet_name]
                    # Obter o DataFrame correspondente à planilha atual
                    if sheet_name == 'NFCe Emitidas':
                        current_df = df_normais_excel
                    elif sheet_name == 'NFCe Inutilizadas':
                        current_df = df_inutilizados_excel
                    elif sheet_name == 'NFCe Canceladas':
                        current_df = df_canceladas_excel
                    else:
                        current_df = pd.DataFrame() # Fallback

                    for idx, col in enumerate(current_df.columns):
                        max_length = max(
                            current_df[col].astype(str).apply(len).max(),
                            len(col)
                        )
                        worksheet.column_dimensions[chr(65 + idx)].width = min(max_length + 2, 50)
            
            # Padronizar e remover colunas indesejadas
            if 'CPF/CNPJ Destinatario' in df.columns:
                df = df.drop(columns=['CPF/CNPJ Destinatario'])
            
            # Gerar relatório PDF
            pdf_path = os.path.join('reports', f'relatorio_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pdf')
            generate_pdf(df, totais, pdf_path)
            
            # Gerar relatório resumido PDF
            resumo_pdf_path = os.path.join('reports', f'relatorio_resumido_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pdf')
            generate_pdf_resumido(df, resumo_pdf_path)
            
            # Download dos relatórios
            if numeros_pulados:
                # Se há números pulados, mostrar 4 colunas
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    with open(excel_path, 'rb') as f:
                        st.download_button(
                            label="📥 Baixar Relatório Excel",
                            data=f,
                            file_name=os.path.basename(excel_path),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                with col2:
                    with open(pdf_path, 'rb') as f:
                        st.download_button(
                            label="📄 Baixar Relatório PDF Detalhado",
                            data=f,
                            file_name=os.path.basename(pdf_path),
                            mime="application/pdf"
                        )
                with col3:
                    with open(resumo_pdf_path, 'rb') as f:
                        st.download_button(
                            label="📄 Baixar Relatório PDF Resumido",
                            data=f,
                            file_name=os.path.basename(resumo_pdf_path),
                            mime="application/pdf"
                        )
                with col4:
                    numeros_pulados_pdf_path = os.path.join('reports', f'relatorio_numeros_pulados_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pdf')
                    generate_pdf_numeros_pulados(numeros_pulados, numeros_pulados_pdf_path)
                    with open(numeros_pulados_pdf_path, 'rb') as f:
                        st.download_button(
                            label="⚠️ Baixar Relatório de Numerações Puladas",
                            data=f,
                            file_name=os.path.basename(numeros_pulados_pdf_path),
                            mime="application/pdf"
                        )
            else:
                # Se não há números pulados, mostrar 3 colunas
                col1, col2, col3 = st.columns(3)
                with col1:
                    with open(excel_path, 'rb') as f:
                        st.download_button(
                            label="📥 Baixar Relatório Excel",
                            data=f,
                            file_name=os.path.basename(excel_path),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                with col2:
                    with open(pdf_path, 'rb') as f:
                        st.download_button(
                            label="📄 Baixar Relatório PDF Detalhado",
                            data=f,
                            file_name=os.path.basename(pdf_path),
                            mime="application/pdf"
                        )
                with col3:
                    with open(resumo_pdf_path, 'rb') as f:
                        st.download_button(
                            label="📄 Baixar Relatório PDF Resumido",
                            data=f,
                            file_name=os.path.basename(resumo_pdf_path),
                            mime="application/pdf"
                        )
        else:
            st.warning("Nenhum arquivo XML válido foi processado.")
