# -*- coding: utf-8 -*-
import os
import time
import json
import logging
import datetime
import xml.etree.ElementTree as ET

import pandas as pd
from tqdm import tqdm

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# ============================= CONFIGURAÇÕES =============================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, 'Portal_Nacional_config.json')
DEFAULT_CONFIG = {
    "pasta_downloads": r"C:\\NFS-e\\PortalNacional",
    "competencia_desejada": "11/2025",
    "timeout": 30,
    "headless": False
}
# =========================================================================

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')
logging.getLogger('selenium.webdriver').setLevel(logging.WARNING)
logger = logging.getLogger(__name__)

URL_PORTAL = "https://www.nfse.gov.br/EmissorNacional"

# Conjunto global com notas já existentes (para evitar duplicidade de XML)
# chave: (emitente_cnpj, numero_nota)
NOTAS_EXISTENTES = set()


def carregar_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
        except Exception:
            config = DEFAULT_CONFIG.copy()
    else:
        config = DEFAULT_CONFIG.copy()
        salvar_config(config)
    return config


def salvar_config(config):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=4, ensure_ascii=False)


CONFIG = carregar_config()
PASTA_DOWNLOADS = CONFIG.get('pasta_downloads', DEFAULT_CONFIG['pasta_downloads'])
COMPETENCIA_DESEJADA = CONFIG.get('competencia_desejada', DEFAULT_CONFIG['competencia_desejada'])
TIMEOUT = CONFIG.get('timeout', 30)


def criar_pasta_downloads(pasta):
    os.makedirs(pasta, exist_ok=True)


def criar_driver(headless=False):
    chrome_options = Options()
    prefs = {
        "download.default_directory": os.path.abspath(PASTA_DOWNLOADS),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "safebrowsing.disable_download_protection": True,
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    if headless:
        chrome_options.add_argument("--headless=new")

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=chrome_options
    )
    driver.maximize_window()
    return driver


def aguardar_downloads(pasta, timeout=600, log_fn=print):
    log_fn("Aguardando downloads terminarem...")
    start = time.time()
    while time.time() - start < timeout:
        if not any(f.endswith('.crdownload') for f in os.listdir(pasta)):
            time.sleep(2)
            if not any(f.endswith('.crdownload') for f in os.listdir(pasta)):
                log_fn("Downloads concluídos.")
                return
        time.sleep(1)
    log_fn("Timeout aguardando downloads; seguindo mesmo assim.")


def parse_competencia_str(comp_str):
    # "MM/AAAA" -> (ano, mes)
    try:
        mes, ano = comp_str.split('/')
        return int(ano), int(mes)
    except Exception:
        return None, None


def mesma_competencia(data_emissao_str, competencia_str):
    """Verifica se data_emissao (dd/mm/aaaa) está no mês/ano da competência MM/AAAA."""
    try:
        dt = datetime.datetime.strptime(data_emissao_str.strip(), "%d/%m/%Y").date()
        ano_c, mes_c = parse_competencia_str(competencia_str)
        if ano_c is None:
            return True
        return dt.year == ano_c and dt.month == mes_c
    except Exception:
        return False


def emissao_anterior_competencia(data_emissao_str, competencia_str):
    """Retorna True se data_emissao for de mês ANTERIOR à competência informada."""
    try:
        dt = datetime.datetime.strptime(data_emissao_str.strip(), "%d/%m/%Y").date()
        ano_c, mes_c = parse_competencia_str(competencia_str)
        if ano_c is None:
            return False
        if dt.year < ano_c:
            return True
        if dt.year == ano_c and dt.month < mes_c:
            return True
        return False
    except Exception:
        return False


def obter_situacao_e_numero_da_linha(linha):
    situacao = ""
    numero_nota = ""

    # Situação via ícone <img src="...tb-cancelada.svg"> ou "tb-gerada.svg"
    try:
        img = linha.find_element(
            By.XPATH,
            ".//img[contains(@src,'tb-cancelada.svg') or contains(@src,'tb-gerada.svg')]"
        )
        src = img.get_attribute("src") or ""
        if "tb-cancelada" in src:
            situacao = "Cancelada"
        elif "tb-gerada" in src:
            situacao = "Autorizada"
    except Exception:
        pass

    # Número da nota - tentar td com classe específica, depois fallback
    try:
        td_num = linha.find_element(By.XPATH, ".//td[contains(@class,'td-numero')]")
        numero_nota = td_num.text.strip()
    except Exception:
        try:
            # fallback: primeira coluna com dígitos
            tds = linha.find_elements(By.TAG_NAME, "td")
            for td in tds:
                txt = td.text.strip()
                if txt.isdigit():
                    numero_nota = txt
                    break
        except Exception:
            pass

    return situacao, numero_nota


def baixar_xml_da_linha(driver, linha, num, competencia_str,
                        situacoes_dict, log_fn=print):
    """
    Retornos possíveis:
    - True       => baixou XML (nota da competência, autorizada ou cancelada)
    - False      => ignorou linha (outra competência)
    - "ANTERIOR" => achou emissão anterior à competência (para parar varredura)
    """
    try:
        # Data de emissão (coluna td-data)
        data_emissao = linha.find_element(
            By.XPATH,
            ".//td[contains(@class, 'td-data')]"
        ).text.strip()

        # Situação + número da nota
        situacao, numero_nota = obter_situacao_e_numero_da_linha(linha)
        if numero_nota:
            situacoes_dict[numero_nota] = situacao

        # Se emissão não é da competência desejada
        if not mesma_competencia(data_emissao, competencia_str):
            log_fn(f"Linha {num}: Ignorada (emissão {data_emissao} | sit: {situacao or 'N/D'})")
            # Se for ANTERIOR à competência, sinalizar que deve parar tudo
            if emissao_anterior_competencia(data_emissao, competencia_str):
                return "ANTERIOR"
            return False

        # Se chegou aqui, é da competência desejada → BAIXA XML SEMPRE
        link_xml = linha.find_element(
            By.XPATH,
            ".//td[contains(@class,'td-opcoes')]"
            "//a[contains(@href,'/EmissorNacional/Notas/Download/NFSe/')]"
        )
        url_xml = link_xml.get_attribute("href")
        driver.get(url_xml)

        log_fn(
            f"Linha {num}: XML baixado → Emissão {data_emissao} | "
            f"Situação: {situacao or 'N/D'} | Nº {numero_nota or 'N/D'}"
        )
        return True

    except Exception as e:
        log_fn(f"Linha {num}: Falha no processamento da linha → {str(e)[:120]}")
        return False


def processar_pagina(driver, competencia_str, situacoes_dict,
                     log_fn=print):
    WebDriverWait(driver, TIMEOUT).until(
        EC.presence_of_all_elements_located((By.XPATH, "//table//tbody//tr[td]"))
    )
    linhas = driver.find_elements(By.XPATH, "//table//tbody//tr[td]")
    log_fn(f"Página atual: {len(linhas)} notas encontradas")

    baixadas_na_pagina = 0
    for i, linha in enumerate(linhas, 1):
        resultado = baixar_xml_da_linha(
            driver, linha, i, competencia_str, situacoes_dict, log_fn=log_fn
        )
        if resultado == "ANTERIOR":
            log_fn("Encontrada nota com emissão anterior à competência → encerrando varredura.")
            return -1
        if resultado is True:
            baixadas_na_pagina += 1
        time.sleep(0.8)

    log_fn(f"→ {baixadas_na_pagina} notas (XML baixado) da competência {competencia_str} nesta página")
    return baixadas_na_pagina



def tem_proxima_pagina(driver, log_fn=print):
    try:
        btn = driver.find_element(
            By.XPATH,
            "//a[contains(@href, 'Emitidas?pg=') and contains(@data-original-title, 'Próxima')]"
        )
        if "disabled" in btn.find_element(By.XPATH, "./ancestor::li").get_attribute("class"):
            log_fn("Botão 'Próxima' desabilitado.")
            return False
        driver.execute_script("arguments[0].click();", btn)
        time.sleep(2.5)
        return True
    except Exception:
        log_fn("Botão 'Próxima' não encontrado.")
        return False


def safe_float(val):
    try:
        s = str(val).strip()
        if not s:
            return 0.0
        # Apenas troca vírgula por ponto (não remove o ponto decimal)
        s = s.replace(',', '.')
        return float(s)
    except Exception:
        try:
            return float(val)
        except Exception:
            return 0.0


def parse_xml_por_nota(xml_path, situacoes_dict=None):
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
        nspace = 'http://www.sped.fazenda.gov.br/nfse'

        # Emitente
        emit = root.find(f'.//{{{nspace}}}emit')
        emit_nome = ''
        emit_cnpj = ''
        if emit is not None:
            xNome = emit.find(f'{{{nspace}}}xNome')
            cnpj_elt = emit.find(f'{{{nspace}}}CNPJ')
            emit_nome = xNome.text if xNome is not None else ''
            emit_cnpj = cnpj_elt.text if cnpj_elt is not None else ''

        # Tomador
        toma = root.find(f'.//{{{nspace}}}DPS/{{{nspace}}}infDPS/{{{nspace}}}toma')
        toma_nome = ''
        toma_cnpj = ''
        if toma is not None:
            xNome_t = toma.find(f'{{{nspace}}}xNome')
            cnpj_t = toma.find(f'{{{nspace}}}CNPJ')
            toma_nome = xNome_t.text if xNome_t is not None else ''
            toma_cnpj = cnpj_t.text if cnpj_t is not None else ''

        # Valores (top level em infNFSe)
        valores_inf = root.find(f'.//{{{nspace}}}infNFSe/{{{nspace}}}valores')
        v_bc = 0.0
        v_liq = 0.0
        v_total_ret = 0.0
        if valores_inf is not None:
            vBC = valores_inf.find(f'{{{nspace}}}vBC')
            vLiq = valores_inf.find(f'{{{nspace}}}vLiq')
            vTotRet = valores_inf.find(f'{{{nspace}}}vTotalRet')
            v_bc = safe_float(vBC.text) if vBC is not None else 0.0
            v_liq = safe_float(vLiq.text) if vLiq is not None else 0.0
            v_total_ret = safe_float(vTotRet.text) if vTotRet is not None else 0.0

        # Valores DPS (vServPrest e tributos federais)
        infDPS = root.find(f'.//{{{nspace}}}DPS/{{{nspace}}}infDPS')
        valores_dps = infDPS.find(f'.//{{{nspace}}}valores') if infDPS is not None else None

        v_serv = 0.0
        irrf = 0.0
        cp = 0.0
        csll = 0.0
        pis = 0.0
        cofins = 0.0

        if valores_dps is not None:
            # Valor do serviço
            vServPrest = valores_dps.find(f'.//{{{nspace}}}vServPrest/{{{nspace}}}vServ')
            v_serv = safe_float(vServPrest.text) if vServPrest is not None else 0.0

            # tribFed
            trib = valores_dps.find(f'.//{{{nspace}}}trib')
            tribFed = trib.find(f'.//{{{nspace}}}tribFed') if trib is not None else None

            if tribFed is not None:
                # IRRF
                elt_irrf = tribFed.find(f'{{{nspace}}}vRetIRRF')
                if elt_irrf is None:
                    elt_irrf = tribFed.find(f'{{{nspace}}}vIR')
                irrf = safe_float(elt_irrf.text) if elt_irrf is not None else 0.0

                # CP (INSS)
                elt_cp = tribFed.find(f'{{{nspace}}}vRetCP')
                cp = safe_float(elt_cp.text) if elt_cp is not None else 0.0

                # CSLL
                elt_csll = tribFed.find(f'{{{nspace}}}vRetCSLL')
                csll = safe_float(elt_csll.text) if elt_csll is not None else 0.0

                # PIS
                elt_pis = tribFed.find(f'{{{nspace}}}vRetPIS')
                if elt_pis is None:
                    elt_pis = tribFed.find(f'{{{nspace}}}vRetPis')
                pis = safe_float(elt_pis.text) if elt_pis is not None else 0.0

                # COFINS
                elt_cof = tribFed.find(f'{{{nspace}}}vRetCOFINS')
                if elt_cof is None:
                    elt_cof = tribFed.find(f'{{{nspace}}}vRetCofins')
                cofins = safe_float(elt_cof.text) if elt_cof is not None else 0.0

        # Outros
        infNFSe = root.find(f'.//{{{nspace}}}infNFSe')
        n_nfse = ''
        if infNFSe is not None:
            nN = infNFSe.find(f'{{{nspace}}}nNFSe')
            n_nfse = nN.text if nN is not None else ''

        # Data de emissão (dhEmi -> dd/mm/aaaa)
        d_emissao = ''
        if infDPS is not None:
            dhEmi = infDPS.find(f'{{{nspace}}}dhEmi')
            if dhEmi is not None and dhEmi.text:
                raw = dhEmi.text[:10]
                try:
                    dt = datetime.datetime.strptime(raw, "%Y-%m-%d").date()
                    d_emissao = dt.strftime("%d/%m/%Y")
                except Exception:
                    d_emissao = raw

        # Serviço
        serv = infDPS.find(f'{{{nspace}}}serv') if infDPS is not None else None
        x_desc = ''
        codigo_serv = ''
        if serv is not None:
            desc_elt = serv.find(f'{{{nspace}}}cServ/{{{nspace}}}xDescServ')
            cod_elt = serv.find(f'{{{nspace}}}cServ/{{{nspace}}}cTribNac')
            x_desc = desc_elt.text if desc_elt is not None else ''
            codigo_serv = cod_elt.text if cod_elt is not None else ''

        # Situação da nota (se foi capturada na tela)
        situacao = ''
        if situacoes_dict is not None and n_nfse:
            situacao = situacoes_dict.get(n_nfse, '')

        return {
            'arquivo': os.path.basename(xml_path),
            'numero_nota': n_nfse,
            'emitente_nome': emit_nome,
            'emitente_cnpj': emit_cnpj,
            'tomador_nome': toma_nome,
            'tomador_cnpj': toma_cnpj,
            'data_emissao': d_emissao,
            'valor_bc': v_bc,
            'valor_liq': v_liq,
            'valor_servico': v_serv,
            'descricao_serv': x_desc,
            'codigo_serv': codigo_serv,
            'situacao': situacao,
            'total_retencoes': v_total_ret,
            'irrf': irrf,
            'cp': cp,
            'csll': csll,
            'pis': pis,
            'cofins': cofins
        }
    except Exception as e:
        logger.error(f"Erro ao parsear {xml_path}: {e}")
        return None


def limpar_nome_empresa(nome):
    if not nome:
        return "SEM_NOME"
    invalidos = '<>:"/\\|?*'
    for ch in invalidos:
        nome = nome.replace(ch, ' ')
    return ' '.join(nome.split())


def carregar_notas_existentes(pasta_base, log_fn=print):
    global NOTAS_EXISTENTES
    NOTAS_EXISTENTES = set()
    if not os.path.exists(pasta_base):
        return
    log_fn("Carregando notas já existentes para evitar duplicidade de XML...")
    for root_dir, dirs, files in os.walk(pasta_base):
        for file in files:
            if not file.lower().endswith('.xml'):
                continue
            caminho = os.path.join(root_dir, file)
            data = parse_xml_por_nota(caminho)
            if not data:
                continue
            chave = (data['emitente_cnpj'], data['numero_nota'])
            if data['emitente_cnpj'] and data['numero_nota']:
                NOTAS_EXISTENTES.add(chave)
    log_fn(f"Total de notas já registradas: {len(NOTAS_EXISTENTES)}")


def gerar_relatorio_para_pasta(pasta_xmls, competencia_str,
                               situacoes_dict, log_fn=print):
    try:
        xml_files = [f for f in os.listdir(pasta_xmls) if f.lower().endswith('.xml')]
        dados = []
        cancelados_paths = []

        for xml_file in tqdm(xml_files, desc=f"Processando XMLs ({os.path.basename(pasta_xmls)})"):
            caminho = os.path.join(pasta_xmls, xml_file)
            data = parse_xml_por_nota(caminho, situacoes_dict=situacoes_dict)
            if not data:
                continue

            # Considera só notas da competência
            if not data['data_emissao'] or not mesma_competencia(data['data_emissao'], competencia_str):
                continue

            # Se não tiver situação preenchida, assume Autorizada
            if not data.get('situacao'):
                data['situacao'] = 'Autorizada'

            # Se for cancelada, marcar caminho para excluir depois
            if data['situacao'] == 'Cancelada':
                cancelados_paths.append(caminho)

            dados.append(data)

        if not dados:
            log_fn(f"Nenhum dado (autorizado ou cancelado) encontrado em {pasta_xmls} para a competência.")
            return

        df = pd.DataFrame(dados)

        # Renomear cabeçalhos
        df.rename(columns={
            'arquivo': 'Arquivo',
            'numero_nota': 'Número da Nota',
            'emitente_nome': 'Emitente',
            'emitente_cnpj': 'CNPJ Emitente',
            'tomador_nome': 'Tomador',
            'tomador_cnpj': 'CNPJ Tomador',
            'data_emissao': 'Data Emissão',
            'valor_bc': 'Valor BC',
            'valor_liq': 'Valor Líquido',
            'valor_servico': 'Valor Serviço',
            'descricao_serv': 'Descrição Serviço',
            'codigo_serv': 'Código Serviço',
            'situacao': 'Situação',
            'total_retencoes': 'Total Retenções',
            'irrf': 'IRRF',
            'cp': 'CP',
            'csll': 'CSLL',
            'pis': 'PIS',
            'cofins': 'COFINS'
        }, inplace=True)

        # Ordenar por Data Emissão e Número da Nota, se existirem
        cols_ordem = []
        if 'Data Emissão' in df.columns:
            cols_ordem.append('Data Emissão')
        if 'Número da Nota' in df.columns:
            cols_ordem.append('Número da Nota')
        if cols_ordem:
            df.sort_values(by=cols_ordem, inplace=True, ignore_index=True)

        # Caminho do relatório
        relatorio_path = os.path.join(
            pasta_xmls,
            f"Relatório Prestados {competencia_str.replace('/', '_')}.xlsx"
        )

        # Salvar Excel
        with pd.ExcelWriter(relatorio_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Detalhe_Notas', index=False)

        # Aplicar layout
        wb = load_workbook(relatorio_path)
        ws = wb['Detalhe_Notas']

        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        header_align = Alignment(horizontal="center", vertical="center")

        max_col = ws.max_column
        max_row = ws.max_row

        # Cabeçalho (linha 1) azul + fonte branca
        for col in range(1, max_col + 1):
            cell = ws.cell(row=1, column=col)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            col_letter = get_column_letter(col)
            ws.column_dimensions[col_letter].width = 19.57

        # Demais linhas: altura 35,25 + centralizado
        data_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for row in range(1, max_row + 1):
            ws.row_dimensions[row].height = 35.25
            for col in range(1, max_col + 1):
                cell = ws.cell(row=row, column=col)
                if row == 1:
                    cell.alignment = header_align
                else:
                    cell.alignment = data_align

        wb.save(relatorio_path)

        # Agora remove os XMLs das notas canceladas
        removidos = 0
        for caminho in cancelados_paths:
            try:
                if os.path.exists(caminho):
                    os.remove(caminho)
                    removidos += 1
            except Exception:
                pass

        total_notas = len(df)
        total_valor_liq = df['Valor Líquido'].sum() if 'Valor Líquido' in df.columns else 0
        total_valor_bc = df['Valor BC'].sum() if 'Valor BC' in df.columns else 0

        log_fn("")
        log_fn("=" * 80)
        log_fn(f"RELATÓRIO GERADO PARA: {os.path.basename(pasta_xmls)}")
        log_fn(f"Total de notas (autorizadas + canceladas): {total_notas}")
        log_fn(f"Total valor líquido: R$ {total_valor_liq:,.2f}")
        log_fn(f"Total base cálculo: R$ {total_valor_bc:,.2f}")
        log_fn(f"Arquivo: {relatorio_path}")
        log_fn(f"XMLs de notas CANCELADAS removidos da pasta: {removidos}")
        log_fn("=" * 80)
        log_fn("")

    except Exception as e:
        logger.error(f"Erro ao gerar relatório em {pasta_xmls}: {e}")
        log_fn(f"Falha na geração do relatório em {pasta_xmls}: {e}")



def organizar_xmls_e_gerar_relatorios_rodada(pasta_base, competencia_str,
                                             novos_xmls, situacoes_dict,
                                             log_fn=print):
    """
    Organiza XMLs recém-baixados (autorizados + cancelados) por empresa
    e gera relatórios para cada pasta de empresa afetada.
    Depois, remove os XMLs das notas canceladas (baixa mas não “guarda”).
    """
    global NOTAS_EXISTENTES
    pastas_afetadas = set()

    # 1. Mover XMLs para pastas de empresa
    for xml_file in novos_xmls:
        caminho = os.path.join(pasta_base, xml_file)
        data = parse_xml_por_nota(caminho, situacoes_dict=situacoes_dict)
        if not data:
            try:
                os.remove(caminho)
            except Exception:
                pass
            continue

        chave = (data['emitente_cnpj'], data['numero_nota'])
        if data['emitente_cnpj'] and data['numero_nota']:
            if chave in NOTAS_EXISTENTES:
                log_fn(f"XML duplicado ignorado (nota já existente): {data['numero_nota']} - {data['emitente_cnpj']}")
                try:
                    os.remove(caminho)
                except Exception:
                    pass
                continue
            NOTAS_EXISTENTES.add(chave)

        pasta_empresa = limpar_nome_empresa(data['emitente_nome'])
        pasta_empresa_path = os.path.join(pasta_base, pasta_empresa)
        os.makedirs(pasta_empresa_path, exist_ok=True)
        novo_caminho = os.path.join(pasta_empresa_path, os.path.basename(caminho))

        try:
            os.replace(caminho, novo_caminho)
        except Exception as e:
            logger.error(f"Erro ao mover XML {caminho} -> {novo_caminho}: {e}")
            continue

        pastas_afetadas.add(pasta_empresa_path)

    # 2. Gerar relatório + remover XML de canceladas
    for pasta in pastas_afetadas:
        gerar_relatorio_para_pasta(pasta, competencia_str, situacoes_dict, log_fn=log_fn)


# ============================= TKINTER APP =============================

class NFSeDownloaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Download NFS-e - Portal Nacional (Multiempresas)")
        self.root.geometry("840x540")
        self.root.minsize(820, 520)

        # Estilo geral
        self.root.configure(bg="#f3f4f6")
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure("TFrame", background="#f3f4f6")
        style.configure("Title.TLabel", font=("Segoe UI", 16, "bold"),
                        foreground="#f3f4f6", background="#f3f4f6")
        style.configure("TLabel", background="#f3f4f6", font=("Segoe UI", 10))
        style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=6)
        style.map("TButton",
                  foreground=[("disabled", "#f3f4f6")],
                  background=[("active", "#2563eb")])

        # Header (barra azul)
        header = tk.Frame(self.root, bg="#1d4ed8")
        header.pack(fill="x")

        lbl_title = tk.Label(
            header,
            text="Portal Nacional - Download de NFS-e (Multiempresas)",
            bg="#1d4ed8",
            fg="white",
            font=("Segoe UI", 14, "bold"),
            pady=10
        )
        lbl_title.pack(side="left", padx=20, pady=5)

        # Container principal
        container = ttk.Frame(self.root, padding=15)
        container.pack(fill="both", expand=True)

        # Frame de configurações
        frame_conf = ttk.LabelFrame(container, text="Configurações", padding=10)
        frame_conf.pack(fill="x", pady=(0, 10))

        # Pasta downloads
        ttk.Label(frame_conf, text="Pasta de downloads:").grid(row=0, column=0, sticky="w")
        self.var_pasta = tk.StringVar(value=PASTA_DOWNLOADS)
        self.entry_pasta = ttk.Entry(frame_conf, textvariable=self.var_pasta, width=60)
        self.entry_pasta.grid(row=0, column=1, padx=5, pady=2, sticky="we")

        btn_pasta = ttk.Button(frame_conf, text="Procurar...", command=self.escolher_pasta)
        btn_pasta.grid(row=0, column=2, padx=5, pady=2)

        frame_conf.columnconfigure(1, weight=1)

        # Competência
        ttk.Label(frame_conf, text="Competência (MM/AAAA):").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.var_comp = tk.StringVar(value=COMPETENCIA_DESEJADA)
        self.entry_comp = ttk.Entry(frame_conf, textvariable=self.var_comp, width=12)
        self.entry_comp.grid(row=1, column=1, sticky="w", padx=5, pady=(8, 0))

        # Frame de botões
        frame_btns = ttk.Frame(container)
        frame_btns.pack(fill="x", pady=(0, 10))

        self.btn_salvar = ttk.Button(frame_btns, text="Salvar Configurações",
                                     command=self.salvar_configuracoes)
        self.btn_salvar.pack(side="left")

        self.btn_limpar_log = ttk.Button(frame_btns, text="Limpar Log",
                                         command=self.limpar_log)
        self.btn_limpar_log.pack(side="left", padx=(10, 0))

        self.btn_baixar = ttk.Button(frame_btns, text="Baixar NFS-e (Multiempresas)",
                                     command=self.iniciar_download)
        self.btn_baixar.pack(side="right")

        # Frame de log
        frame_log = ttk.LabelFrame(container, text="Log da execução", padding=5)
        frame_log.pack(fill="both", expand=True)

        self.txt_log = tk.Text(frame_log, wrap="word", height=15, state="disabled",
                               font=("Consolas", 9))
        vsb = ttk.Scrollbar(frame_log, orient="vertical", command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=vsb.set)

        self.txt_log.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

    # ------------------ UTILITÁRIOS DE INTERFACE ------------------ #

    def log(self, msg: str):
        """Escreve no log da interface e no console."""
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", msg + "\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")
        self.root.update_idletasks()
        print(msg)

    def limpar_log(self):
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.configure(state="disabled")
        self.root.update_idletasks()

    def escolher_pasta(self):
        pasta = filedialog.askdirectory(initialdir=self.var_pasta.get() or os.getcwd())
        if pasta:
            self.var_pasta.set(pasta)

    def salvar_configuracoes(self):
        global PASTA_DOWNLOADS, COMPETENCIA_DESEJADA
        pasta = self.var_pasta.get().strip()
        comp = self.var_comp.get().strip()

        if not pasta:
            messagebox.showwarning("Atenção", "Informe a pasta de downloads.")
            return
        if not comp or len(comp) != 7 or "/" not in comp:
            messagebox.showwarning("Atenção", "Informe a competência no formato MM/AAAA.")
            return

        PASTA_DOWNLOADS = pasta
        COMPETENCIA_DESEJADA = comp

        CONFIG['pasta_downloads'] = pasta
        CONFIG['competencia_desejada'] = comp
        salvar_config(CONFIG)

        self.log("Configurações salvas com sucesso.")

    # ------------------ CONTROLE PRINCIPAL ------------------ #

    def iniciar_download(self):
        # Desabilitar botão para evitar múltiplos cliques
        self.btn_baixar.configure(state="disabled")
        try:
            self._rodar_multiempresas()
        finally:
            self.btn_baixar.configure(state="normal")

    def _rodar_multiempresas(self):
        """
        Fluxo principal multiempresas:
        - Lê a competência da interface
        - Para cada CNPJ (empresa), o usuário faz login e entra na tela de Notas Emitidas
        - O robô baixa os XMLs da competência (autorizadas + canceladas)
        - Organiza pastas por empresa e gera relatórios
        """
        global COMPETENCIA_DESEJADA

        # Lê competência da tela (ou usa padrão da config)
        competencia_str = self.var_comp.get().strip() or COMPETENCIA_DESEJADA
        COMPETENCIA_DESEJADA = competencia_str

        self.log("")
        self.log("=" * 90)
        self.log(f"Iniciando processamento multiempresas para competência {COMPETENCIA_DESEJADA}")
        self.log("=" * 90)

        # Zera o controle de notas já vistas (para não repetir entre empresas)
        global NOTAS_EXISTENTES
        NOTAS_EXISTENTES = set()

        continuar_empresas = True
        total_empresas = 0

        while continuar_empresas:
            total_empresas += 1
            self.log("")
            self.log(f"==================== EMPRESA #{total_empresas} ====================")

            driver = None
            try:
                # Garante pasta de downloads
                criar_pasta_downloads(PASTA_DOWNLOADS)

                # Snapshot dos XMLs antes desta empresa
                xml_antes = {
                    f for f in os.listdir(PASTA_DOWNLOADS)
                    if f.lower().endswith('.xml')
                }

                # Abre navegador
                driver = criar_driver()
                driver.get("https://www.nfse.gov.br/EmissorNacional")

                # Orientação para o usuário
                messagebox.showinfo(
                    "Atenção",
                    "1) Faça login com o CNPJ desejado (perfil PRESTADOR).\n"
                    "2) Acesse: Serviços Prestados → Notas Emitidas.\n"
                    "3) Quando a tabela de notas aparecer, clique em OK para iniciar o robô."
                )

                situacoes_dict = {}
                pagina = 1

                while True:
                    self.log(f"{'-' * 40} PÁGINA {pagina} {'-' * 40}")

                    resultado = processar_pagina(
                        driver,
                        COMPETENCIA_DESEJADA,
                        situacoes_dict,
                        log_fn=self.log
                    )

                    aguardar_downloads(PASTA_DOWNLOADS, log_fn=self.log)

                    # -1: encontramos emissão anterior à competência → pode parar
                    if resultado == -1:
                        self.log("Parando varredura: encontradas notas de emissão anterior à competência.")
                        break

                    # 0: nenhuma nota elegível na página
                    if resultado == 0:
                        self.log("Nenhuma nota elegível (competência) nesta página.")
                        if not tem_proxima_pagina(driver, log_fn=self.log):
                            break
                    else:
                        # Teve notas baixadas, tenta ir para próxima
                        if not tem_proxima_pagina(driver, log_fn=self.log):
                            break

                    pagina += 1

                # Snapshot após a empresa
                xml_depois = {
                    f for f in os.listdir(PASTA_DOWNLOADS)
                    if f.lower().endswith('.xml')
                }
                novos_xmls = sorted(list(xml_depois - xml_antes))

                self.log(f"XMLs novos nesta empresa (autorizadas + canceladas): {len(novos_xmls)}")

                if novos_xmls:
                    organizar_xmls_e_gerar_relatorios_rodada(
                        PASTA_DOWNLOADS,
                        COMPETENCIA_DESEJADA,
                        novos_xmls,
                        situacoes_dict,
                        log_fn=self.log
                    )
                else:
                    self.log("Nenhum dado para organizar/gerar relatório nesta empresa.")

            except Exception as e:
                logger.error(f"Erro na execução para empresa #{total_empresas}: {e}")
                self.log(f"ERRO na execução para esta empresa: {e}")
            finally:
                if driver is not None:
                    try:
                        driver.quit()
                    except Exception:
                        pass

            # Pergunta se vai rodar outra empresa (outro CNPJ)
            continuar_empresas = messagebox.askyesno(
                "Multiempresas",
                "Deseja processar outra empresa (outro CNPJ)?"
            )

        self.log("")
        self.log("=" * 90)
        self.log("PROCESSO FINALIZADO PARA TODAS AS EMPRESAS.")
        self.log("=" * 90)


def main():
    root = tk.Tk()
    app = NFSeDownloaderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

