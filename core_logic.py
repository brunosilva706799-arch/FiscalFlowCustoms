import os
import re
import logging
from collections import defaultdict
from lxml import etree
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
import requests

def extrair_dados_nf(arquivo_xml):
    try:
        tree = etree.parse(arquivo_xml)
        root_xml = tree.getroot()
        namespace_match = re.search(r'\{(.+?)\}', root_xml.tag)
        namespace = namespace_match.group(1) if namespace_match else ''
        
        def find(element, tag):
            if namespace: return element.find(f'{{{namespace}}}{tag}')
            return element.find(tag)

        def find_all(element, tag):
            if namespace: return element.findall(f'{{{namespace}}}{tag}')
            return element.findall(tag)

        nfe_element = find(root_xml, 'NFe')
        if nfe_element is None:
            nfe_element = root_xml
        inf_nfe_tag = find(nfe_element, 'infNFe')

        if inf_nfe_tag is None: return None
        ide_tag = find(inf_nfe_tag, 'ide')
        if ide_tag is None: return None
        
        numero_nf = find(ide_tag, 'nNF').text if find(ide_tag, 'nNF') is not None else 'N/A'
        
        try:
            cfop_tag_det = find_all(inf_nfe_tag, 'det')
            if cfop_tag_det:
                prod_tag = find(cfop_tag_det[0], 'prod')
                if prod_tag is not None:
                    cfop_nf = find(prod_tag, 'CFOP').text
            else:
                cfop_nf = 'N/A'
        except (AttributeError, IndexError):
            cfop_nf = 'N/A'
            
        data_emissao_completa = find(ide_tag, 'dhEmi').text if find(ide_tag, 'dhEmi') is not None else 'N/A'
        data_emissao = data_emissao_completa.split('T')[0] if data_emissao_completa != 'N/A' else 'N/A'
        tipo_nota_val = find(ide_tag, 'tpNF').text if find(ide_tag, 'tpNF') is not None else '1'
        tipo_nota = "Saída" if tipo_nota_val == '1' else "Entrada"
        dest_tag = find(inf_nfe_tag, 'dest')
        nome_cliente = find(dest_tag, 'xNome').text if dest_tag is not None and find(dest_tag, 'xNome') is not None else 'N/A'
        total_tag = find(inf_nfe_tag, 'total')
        valor_total_nf, valor_total_produtos = 'N/A', 'N/A'
        impostos_totais = defaultdict(lambda: '0.00')
        v_ii, v_outras, v_afrmm = 0.0, 0.0, 0.0
        
        if total_tag is not None:
            icms_total_tag = find(total_tag, 'ICMSTot')
            if icms_total_tag is not None:
                valor_total_nf = find(icms_total_tag, 'vNF').text if find(icms_total_tag, 'vNF') is not None else 'N/A'
                valor_total_produtos = find(icms_total_tag, 'vProd').text if find(icms_total_tag, 'vProd') is not None else 'N/A'
                for imposto_tag in icms_total_tag:
                    tag_nome = imposto_tag.tag.split('}')[-1]
                    if tag_nome.startswith('v') and imposto_tag.text:
                        impostos_totais[tag_nome] = imposto_tag.text
                ii_tag = find(icms_total_tag, 'vII')
                if ii_tag is not None and ii_tag.text: v_ii = float(ii_tag.text)
                v_outras_tag = find(icms_total_tag, 'vOutro')
                if v_outras_tag is not None and v_outras_tag.text: v_outras = float(v_outras_tag.text)
        
        det_tags = find_all(inf_nfe_tag, 'det')
        for det_tag in det_tags:
            prod_tag = find(det_tag, 'prod')
            if prod_tag is not None:
                di_tag = find(prod_tag, 'DI')
                if di_tag is not None:
                    vafrmm_tag = find(di_tag, 'vAFRMM')
                    if vafrmm_tag is not None and vafrmm_tag.text:
                        try:
                            v_afrmm = float(vafrmm_tag.text)
                            if v_afrmm > 0: break 
                        except (ValueError, TypeError):
                            continue
                            
        inf_adicionais_tag = find(inf_nfe_tag, 'infAdic')
        if v_afrmm == 0.0 and inf_adicionais_tag is not None:
            inf_cpl_tag = find(inf_adicionais_tag, 'infCpl')
            if inf_cpl_tag is not None and inf_cpl_tag.text:
                text_info_adicional = inf_cpl_tag.text
                afrmm_match = re.search(r'AFRMM.*?R\$?\s*([\d.,]+)', text_info_adicional, re.IGNORECASE | re.DOTALL)
                if afrmm_match:
                    valor_str = afrmm_match.group(1).replace('.', '').replace(',', '.')
                    v_afrmm = float(valor_str)
                    
        nome_processo = 'N/A'
        if inf_adicionais_tag is not None:
            inf_cpl_tag = find(inf_adicionais_tag, 'infCpl')
            if inf_cpl_tag is not None and inf_cpl_tag.text:
                text_info_adicional = inf_cpl_tag.text
                padrao_cpl = re.search(r'PROCESSO\s*([A-Z0-9]+)\.?', text_info_adicional, re.IGNORECASE)
                if padrao_cpl: nome_processo = padrao_cpl.group(1).strip()
                if nome_processo == 'N/A' and tipo_nota == 'Entrada':
                    padrao_entrada = re.search(r'([A-Z]+(?:IMP|EXP)\d{3}\d{4})', text_info_adicional, re.IGNORECASE)
                    if padrao_entrada: nome_processo = padrao_entrada.group(1)
                elif nome_processo == 'N/A' and cfop_nf == '6108':
                    padrao_cfop_6108 = re.search(r'([A-Z0-9\s]+ (?:IMP|EXP)\.\d{3}\.\d{4})', text_info_adicional, re.IGNORECASE)
                    if padrao_cfop_6108: nome_processo = padrao_cfop_6108.group(1)

        if nome_processo == 'N/A' and tipo_nota == "Saída":
            processos_por_item = set()
            for item_tag in find_all(inf_nfe_tag, 'det'):
                inf_ad_prod_tag = find(item_tag, 'infAdProd')
                if inf_ad_prod_tag is not None and inf_ad_prod_tag.text:
                    padrao_ad_prod = re.search(r'(?:PROCESSO|REGISTRO):\s*([^.\s]+)', inf_ad_prod_tag.text, re.IGNORECASE)
                    if padrao_ad_prod: processos_por_item.add(padrao_ad_prod.group(1))
            if processos_por_item: nome_processo = ', '.join(processos_por_item)
            
        if nome_processo != 'N/A': nome_processo = nome_processo.replace(' ', '').replace('.', '')
        
        valor_servico_trading = 'N/A'
        if inf_adicionais_tag is not None and find(inf_adicionais_tag, 'infCpl') is not None and find(inf_adicionais_tag, 'infCpl').text:
            match_trading = re.search(r'trading[^0-9]*([\d.,]+)', find(inf_adicionais_tag, 'infCpl').text.lower())
            if match_trading: valor_servico_trading = match_trading.group(1).replace('.', '').replace(',', '.')
            
        status_nf = "N/A"
        prot_nfe_tag = find(root_xml, 'protNFe')
        if prot_nfe_tag is not None and find(prot_nfe_tag, 'infProt') is not None:
            cstat_tag = find(find(prot_nfe_tag, 'infProt'), 'cStat')
            if cstat_tag is not None:
                if cstat_tag.text == '100': status_nf = "Autorizada"
                elif cstat_tag.text == '101': status_nf = "Cancelada"
                
        return {'nome_arquivo': os.path.basename(arquivo_xml), 'numero_nf': numero_nf, 'cfop_nf': cfop_nf, 'data_emissao': data_emissao,'nome_cliente': nome_cliente, 'valor_total_nf': float(valor_total_nf) if valor_total_nf != 'N/A' else 0.0, 'valor_total_produtos': float(valor_total_produtos) if valor_total_produtos != 'N/A' else 0.0, 'impostos': impostos_totais, 'valor_servico_trading': float(valor_servico_trading) if valor_servico_trading != 'N/A' else 0.0, 'tipo_nota': tipo_nota,'nome_processo': nome_processo, 'vII': v_ii, 'vAFRMM': v_afrmm, 'vOutras': v_outras, 'status_nf': status_nf}
    except Exception as e:
        logging.error(f"Erro ao processar o arquivo {os.path.basename(arquivo_xml)}: {e}", exc_info=True)
        return None

def setup_headers(ws, headers):
    fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    alignment = Alignment(horizontal="center", vertical="center")
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    for col, title in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=title)
        cell.fill, cell.alignment, cell.border = fill, alignment, border

def write_data_to_excel(ws, row_index, data, headers):
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    currency_format = 'R$ #,##0.00'
    for col, header in enumerate(headers, 1):
        value_map = {'Arquivo': data.get('nome_arquivo'), 'Número da NF': data.get('numero_nf'), 'CFOP': data.get('cfop_nf'),'Cliente': data.get('nome_cliente'), 'Data de Emissão': data.get('data_emissao'), 'Nome do Processo': data.get('nome_processo'),'Valor Total dos Produtos': data.get('valor_total_produtos', 0.0), 'Valor Total': data.get('valor_total_nf', 0.0),'Valor ICMS': float(data.get('impostos', {}).get('vICMS', '0.00')), 'Valor IPI': float(data.get('impostos', {}).get('vIPI', '0.00')),'Valor PIS': float(data.get('impostos', {}).get('vPIS', '0.00')), 'Valor COFINS': float(data.get('impostos', {}).get('vCOFINS', '0.00')),'Valor Serviço Trading': data.get('valor_servico_trading', 0.0), 'Valor II': data.get('vII', 0.0),'Valor AFRMM': data.get('vAFRMM', 0.0), 'Outras Despesas': data.get('vOutras', 0.0), 'Status da NF': data.get('status_nf', 'N/A')}
        value = value_map.get(header, '')
        cell = ws.cell(row=row_index, column=col, value=value)
        cell.border = border
        if isinstance(value, (int, float)) and (header.startswith('Valor') or header == 'Outras Despesas' or header == 'Valor AFRMM' or header == 'Valor Serviço Trading'):
            cell.number_format = currency_format

def add_totals_row(ws, headers):
    last_row = ws.max_row
    if last_row < 2: 
        return
        
    totals_row = last_row + 1
    
    total_font = Font(bold=True)
    total_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
    
    first_cell = ws.cell(row=totals_row, column=1, value="TOTAIS")
    first_cell.font = total_font
    first_cell.fill = total_fill

    for col_idx, header in enumerate(headers, 1):
        if header.startswith('Valor') or header == 'Outras Despesas' or header == 'Valor AFRMM' or header == 'Valor Serviço Trading':
            col_letter = get_column_letter(col_idx)
            formula = f"=SUM({col_letter}2:{col_letter}{last_row})"
            
            cell = ws.cell(row=totals_row, column=col_idx, value=formula)
            cell.font = total_font
            cell.fill = total_fill
            cell.number_format = 'R$ #,##0.00'

# --- NOVA FUNÇÃO DE ATUALIZAÇÃO ---
def check_for_updates(current_version, repo_owner, repo_name):
    """
    Verifica no repositório do GitHub pela última versão lançada.
    Retorna um dicionário com os detalhes se houver uma atualização.
    """
    api_url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/releases/latest"
    try:
        response = requests.get(api_url, timeout=5)
        response.raise_for_status()
        
        latest_release = response.json()
        latest_version_str = latest_release['tag_name'].lstrip('v')
        
        # Comparação de versões robusta (ex: 2.3.1 > 2.3)
        current_v = tuple(map(int, (current_version.split('.'))))
        latest_v = tuple(map(int, (latest_version_str.split('.'))))

        if latest_v > current_v:
            assets = latest_release.get('assets', [])
            if assets:
                download_url = assets[0].get('browser_download_url')
                return {
                    "update_available": True,
                    "latest_version": latest_version_str,
                    "download_url": download_url,
                    "release_notes": latest_release.get("body", "Sem notas de versão.")
                }
    except requests.exceptions.RequestException as e:
        logging.error(f"Erro ao verificar atualizações (rede): {e}")
    except Exception as e:
        logging.error(f"Erro ao processar a resposta da API do GitHub: {e}")
        
    return {"update_available": False}