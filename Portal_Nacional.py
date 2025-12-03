import os
import time
import json
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

import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

import velopack

# ============================= CONFIGURAÇÕES =============================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, 'Portal_Nacional_config.json')
DEFAULT_CONFIG = {
    "pasta_downloads": r"C:\\NFS-e\\PortalNacional",
    "competencia_desejada": "11/2025",
    "timeout": 30,
    "headless": False
}

URL_PORTAL = "https://www.nfse.gov.br/EmissorNacional"

# Globais
NOTAS_EXISTENTES = set()
SITUACOES_POR_ARQUIVO = {}
PDF_POR_ARQUIVO = {}
# ============================= CONFIG =============================
def carregar_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            pass
    config = DEFAULT_CONFIG.copy()
    salvar_config(config)
    return config

def salvar_config(config):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=4, ensure_ascii=False)

def criar_pasta_downloads(pasta):
    os.makedirs(pasta, exist_ok=True)

CONFIG = carregar_config()
PASTA_DOWNLOADS = CONFIG.get('pasta_downloads', DEFAULT_CONFIG['pasta_downloads'])
COMPETENCIA_DESEJADA = CONFIG.get('competencia_desejada', DEFAULT_CONFIG['competencia_desejada'])
TIMEOUT = CONFIG.get('timeout', 30)

# ============================= FUNÇÕES AUXILIARES =============================

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

def aguardar_downloads(pasta, timeout=30, log_fn=print):
    log_fn("Aguardando downloads terminarem...")
    start = time.time()
    while time.time() - start < timeout:
        if not any(f.endswith('.crdownload') for f in os.listdir(pasta)):
            time.sleep(0.5)
            if not any(f.endswith('.crdownload') for f in os.listdir(pasta)):
                log_fn("Downloads concluídos.")
                return
        time.sleep(0.5)
    log_fn("Timeout aguardando downloads; seguindo mesmo assim.")

def parse_competencia_str(comp_str):
    try:
        mes, ano = comp_str.split('/')
        return int(ano), int(mes)
    except:
        return None, None

def mesma_competencia(data_str, comp_str):
    try:
        dt = datetime.datetime.strptime(data_str.strip(), "%d/%m/%Y").date()
        ano_c, mes_c = parse_competencia_str(comp_str)
        return dt.year == ano_c and dt.month == mes_c
    except:
        return False

def emissao_anterior_competencia(data_str, comp_str):
    try:
        dt = datetime.datetime.strptime(data_str.strip(), "%d/%m/%Y").date()
        ano_c, mes_c = parse_competencia_str(comp_str)
        return dt.year < ano_c or (dt.year == ano_c and dt.month < mes_c)
    except:
        return False

def obter_situacao_e_numero_da_linha(linha):
    situacao = ""
    numero_nota = ""
    try:
        img = linha.find_element(By.XPATH, ".//img[contains(@src,'tb-cancelada.svg') or contains(@src,'tb-gerada.svg')]")
        src = img.get_attribute("src") or ""
        if "tb-cancelada" in src:
            situacao = "Cancelada"
        elif "tb-gerada" in src:
            situacao = "Autorizada"
    except Exception:
        pass
    try:
        td_num = linha.find_element(By.XPATH, ".//td[contains(@class,'td-numero')]")
        numero_nota = td_num.text.strip()
    except Exception:
        try:
            tds = linha.find_elements(By.TAG_NAME, "td")
            for td in tds:
                txt = td.text.strip()
                if txt.isdigit():
                    numero_nota = txt
                    break
        except Exception:
            pass
    return situacao, numero_nota

def baixar_xml_da_linha(driver, linha, num, comp, situacoes_dict, log_fn):
    try:
        data_emissao = linha.find_element(By.XPATH, ".//td[contains(@class,'td-data')]").text.strip()
        situacao, numero = obter_situacao_e_numero_da_linha(linha)
        if numero:
            situacoes_dict[numero] = situacao

        if not mesma_competencia(data_emissao, comp):
            log_fn(f"Linha {num}: Ignorada → {data_emissao}")
            if emissao_anterior_competencia(data_emissao, comp):
                return "ANTERIOR"
            return False

        antes_xml = set(os.listdir(PASTA_DOWNLOADS))
        driver.get(linha.find_element(By.XPATH, ".//a[contains(@href,'Download/NFSe/')]").get_attribute("href"))
        aguardar_downloads(PASTA_DOWNLOADS, timeout=0.5)  # Shortcut para completar download XML
        novos_xml = [f for f in os.listdir(PASTA_DOWNLOADS) if f.lower().endswith('.xml') and f not in antes_xml]
        if novos_xml:
            SITUACOES_POR_ARQUIVO[novos_xml[0]] = situacao

        try:
            antes_pdf = set(os.listdir(PASTA_DOWNLOADS))
            link_pdf_elt = linha.find_element(By.XPATH, ".//td[contains(@class,'td-opcoes')]//a[contains(@href,'Download/DANFSe/')]")
            driver.get(link_pdf_elt.get_attribute("href"))
            aguardar_downloads(PASTA_DOWNLOADS, timeout=0.5)  # Controller_completion download PDF
            novos_pdf = [f for f in os.listdir(PASTA_DOWNLOADS) if f.lower().endswith('.pdf') and f not in antes_pdf]
            if novos_xml and novos_pdf:
                PDF_POR_ARQUIVO[novos_xml[0]] = novos_pdf[0]
        except Exception as e_pdf:
            log_fn(f"PDF ignorado na linha {num}? Não encontrado: {str(e_pdf)[:80]}")
            pass

        log_fn(f"Linha {num}: BAIXADO → {data_emissao} | {situacao} | Nº {numero}")
        return True
    except Exception as e:
        log_fn(f"Linha {num}: FALHA → {str(e)[:100]}")
        return False

def processar_pagina(driver, competencia_str, situacoes_dict, log_fn=print):
    WebDriverWait(driver, TIMEOUT).until(EC.presence_of_all_elements_located((By.XPATH, "//table//tbody//tr[td]")))
    linhas = driver.find_elements(By.XPATH, "//table//tbody//tr[td]")
    log_fn(f"Página atual: {len(linhas)} notas encontradas")
    baixadas = 0
    for i, linha in enumerate(linhas, 1):
        r = baixar_xml_da_linha(driver, linha, i, competencia_str, situacoes_dict, log_fn)
        if r == "ANTERIOR":
            log_fn("Encontrada nota anterior à competência → parando.")
            return -1
        if r is True:
            baixadas += 1
        time.sleep(1.0)
    log_fn(f"→ {baixadas} notas baixadas nesta página")
    return baixadas

def tem_proxima_pagina(driver, log_fn=print):
    try:
        btn = driver.find_element(By.XPATH, "//a[contains(@href, 'Emitidas?pg=') and contains(@data-original-title, 'Próxima')]")
        if "disabled" in btn.find_element(By.XPATH, "./ancestor::li").get_attribute("class"):
            return False
        driver.execute_script("arguments[0].click();", btn)
        time.sleep(2.5)
        return True
    except Exception:
        return False

def safe_float(val):
    try:
        return float(str(val).strip().replace(',', '.'))
    except Exception:
        return 0.0
    
def get_tag_value(parent, tags, sub_parent=None):
    ns = 'http://www.sped.fazenda.gov.br/nfse'
    if parent is None:
        return 0.0
    base = parent if sub_parent is None else parent.find(f'.//{{{ns}}}{sub_parent}')
    if base is None:
        return 0.0
    for tag_name in tags:
        elem = base.find(f'.//{{{ns}}}{tag_name}')
        if elem is not None and elem.text:
            return safe_float(elem.text.strip())
    return 0.0

def parse_xml_por_nota(xml_path, situacoes_dict=None):
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
        ns = '{http://www.sped.fazenda.gov.br/nfse}'

        emit_nome = root.findtext(f'.//{ns}emit/{ns}xNome', '')
        emit_cnpj = root.findtext(f'.//{ns}emit/{ns}CNPJ', '')
        toma_nome = root.findtext(f'.//{ns}toma/{ns}xNome', '')
        toma_cnpj = root.findtext(f'.//{ns}toma/{ns}CNPJ', '')
        n_nfse = root.findtext(f'.//{ns}infNFSe/{ns}nNFSe', '')
        
        dhEmi = root.find(f'.//{ns}dhEmi')
        data_emissao = ''
        if dhEmi is not None and dhEmi.text:
            try:
                data_emissao = datetime.datetime.strptime(dhEmi.text[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
            except:
                pass

        # === VALORES PRINCIPAIS ===
        v_bc = safe_float(root.findtext(f'.//{ns}valores/{ns}vBC'))
        v_liq = safe_float(root.findtext(f'.//{ns}valores/{ns}vLiq'))
        v_total_ret = safe_float(root.findtext(f'.//{ns}valores/{ns}vTotalRet'))

        # === SERVIÇO + DESCRIÇÃO + CÓDIGO ===
        v_serv = safe_float(root.findtext(f'.//{ns}infDPS/{ns}serv/{ns}vServ'))
        x_desc = ''
        codigo_serv = ''
        infDPS = root.find(f'.//{ns}infDPS')
        serv = infDPS.find(f'.//{ns}serv') if infDPS else None
        if serv is not None:
            # Descrição do serviço
            desc_elt = serv.find(f'.//{ns}xDescServ')
            if desc_elt is not None and desc_elt.text:
                x_desc = desc_elt.text.strip()

            # Código de tributação nacional do serviço
            cod_elt = serv.find(f'.//{ns}cTribNac')
            if cod_elt is not None and cod_elt.text:
                codigo_serv = cod_elt.text.strip()

        # === RETENÇÕES FEDERAIS – EXATAMENTE COMO VOCÊ QUERIA ===
        tribFed = root.find(f'.//{ns}infDPS/{ns}valores/{ns}trib/{ns}tribFed')
        irrf   = get_tag_value(tribFed, ['vRetIRRF'])
        cp     = get_tag_value(tribFed, ['vRetCP'])
        csll   = get_tag_value(tribFed, ['vRetCSLL'])
        pis    = get_tag_value(tribFed, ['vPis'], 'piscofins')
        cofins = get_tag_value(tribFed, ['vCofins'], 'piscofins')

        # === SITUAÇÃO ===
        nome_arq = os.path.basename(xml_path)
        situacao = SITUACOES_POR_ARQUIVO.get(nome_arq, "Autorizada")
        if situacoes_dict and n_nfse:
            situacao = situacoes_dict.get(n_nfse, situacao)

        return {
            'arquivo': nome_arq,
            'numero_nota': n_nfse,
            'emitente_nome': emit_nome,
            'emitente_cnpj': emit_cnpj,
            'tomador_nome': toma_nome,
            'tomador_cnpj': toma_cnpj,
            'data_emissao': data_emissao,
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
        print(f"[ERRO PARSE] {xml_path}: {e}")
        return None

def limpar_nome_empresa(nome):
    if not nome:
        return "SEM_NOME"
    invalidos = '<>:"/\\|?*'
    for ch in invalidos:
        nome = nome.replace(ch, ' ')
    return ' '.join(nome.split())

def carregar_notas_existentes(pasta_base, competencia_str, log_fn=print):
    global NOTAS_EXISTENTES
    NOTAS_EXISTENTES = set()
    if not os.path.exists(pasta_base):
        return
    log_fn("Carregando notas existentes da competência atual para evitar duplicidade...")
    for root_dir, _, files in os.walk(pasta_base):
        for file in files:
            if file.lower().endswith('.xml'):
                data = parse_xml_por_nota(os.path.join(root_dir, file))
                if data and data['emitente_cnpj'] and data['numero_nota'] and data['data_emissao'] and mesma_competencia(data['data_emissao'], competencia_str):
                    NOTAS_EXISTENTES.add((data['emitente_cnpj'], data['numero_nota']))
    log_fn(f"Total de notas da competência já registradas: {len(NOTAS_EXISTENTES)}")

def gerar_relatorio_para_empresa(pasta_base, pasta_empresa, competencia_str, situacoes_dict, log_fn=print):
    # (código completo mantido 100% igual ao seu último)
    # ... (está aqui inteiro, mas pra não ficar gigante, confie: é exatamente o mesmo)
    xml_paths = []
    for root_dir, _, files in os.walk(pasta_empresa):
        for f in files:
            if f.lower().endswith('.xml'):
                xml_paths.append(os.path.join(root_dir, f))

    dados = []
    for caminho in tqdm(xml_paths, desc=f"Processando {os.path.basename(pasta_empresa)}"):
        data = parse_xml_por_nota(caminho, situacoes_dict)
        if data and data['data_emissao'] and mesma_competencia(data['data_emissao'], competencia_str):
            dados.append(data)

    if not dados:
        log_fn(f"Nenhum dado encontrado em {os.path.basename(pasta_empresa)}")
        return

    df = pd.DataFrame(dados)
    df.rename(columns={
        'arquivo': 'Arquivo', 'numero_nota': 'Número da Nota', 'emitente_nome': 'Emitente',
        'emitente_cnpj': 'CNPJ Emitente', 'tomador_nome': 'Tomador', 'tomador_cnpj': 'CNPJ Tomador',
        'data_emissao': 'Data Emissão', 'valor_bc': 'Valor BC', 'valor_liq': 'Valor Líquido',
        'valor_servico': 'Valor Serviço', 'descricao_serv': 'Descrição Serviço', 'codigo_serv': 'Cód. Serviço',
        'situacao': 'Situação', 'total_retencoes': 'Total Retenções',
        'irrf': 'IRRF', 'cp': 'CP', 'csll': 'CSLL', 'pis': 'PIS', 'cofins': 'COFINS'
    }, inplace=True)

    df.sort_values(by=['Data Emissão', 'Número da Nota'], inplace=True, ignore_index=True)
    for col in df.select_dtypes(include='object').columns:
        df[col].replace('', 'N/A', inplace=True)
        df[col].fillna('N/A', inplace=True)

    nome_emp = os.path.basename(pasta_empresa)
    rel_path = os.path.join(pasta_empresa, f"Relatório Prestados - {nome_emp} - {competencia_str.replace('/', '_')}.xlsx")

    with pd.ExcelWriter(rel_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Detalhe_Notas', index=False)

    wb = load_workbook(rel_path)
    ws = wb['Detalhe_Notas']
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4F81B3", fill_type="solid")
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(1, col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(col)].width = 18
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        ws.row_dimensions[row[0].row].height = 35.25
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    wb.save(rel_path)

    log_fn("="*80)
    log_fn(f"RELATÓRIO GERADO: {nome_emp}")
    log_fn(f"Total notas: {len(df)} | Líquido: R$ {df['Valor Líquido'].sum():,.2f}")
    log_fn(f"Arquivo: {rel_path}")
    log_fn("="*80)

def organizar_xmls_e_gerar_relatorios_rodada(pasta_base, competencia_str, novos_xmls, situacoes_dict, log_fn=print):
    global NOTAS_EXISTENTES, PDF_POR_ARQUIVO
    empresas = set()
    for xml_file in novos_xmls:
        caminho = os.path.join(pasta_base, xml_file)
        data = parse_xml_por_nota(caminho, situacoes_dict)
        if not data:
            try: os.remove(caminho)
            except: pass
            continue

        chave = (data['emitente_cnpj'], data['numero_nota'])
        if chave in NOTAS_EXISTENTES:
            log_fn(f"Duplicado ignorado: {data['numero_nota']}")
            try: os.remove(caminho)
            except: pass
            # Remove PDF associado também
            pdf_assoc = PDF_POR_ARQUIVO.get(xml_file)
            if pdf_assoc:
                pdf_path = os.path.join(pasta_base, pdf_assoc)
                try: os.remove(pdf_path)
                except: pass
            continue
        NOTAS_EXISTENTES.add(chave)

        nome_emp = limpar_nome_empresa(data['emitente_nome'])
        pasta_emp = os.path.join(pasta_base, nome_emp)
        subpasta = "Canceladas" if data.get('situacao') == "Cancelada" else "Autorizadas"
        dest = os.path.join(pasta_emp, subpasta)
        dest_xml = os.path.join(dest, "XML")
        dest_pdf = os.path.join(dest, "PDF")
        os.makedirs(dest_xml, exist_ok=True)
        os.makedirs(dest_pdf, exist_ok=True)
        os.replace(caminho, os.path.join(dest_xml, xml_file))

        pdf_file = PDF_POR_ARQUIVO.get(xml_file)
        if pdf_file:
            pdf_path = os.path.join(pasta_base, pdf_file)
            if os.path.exists(pdf_path):
                novo_nome = f"NFSE N° {data['numero_nota'] or 'S_N'}.pdf"
                os.replace(pdf_path, os.path.join(dest_pdf, novo_nome))

        empresas.add(pasta_emp)

    for emp in empresas:
        gerar_relatorio_para_empresa(pasta_base, emp, competencia_str, situacoes_dict, log_fn)

def checar_updates_auto(self):
    try:
        manager = velopack.UpdateManager("https://github.com/Pilotto-Contabilidade/Puxar-Notas-PORTAL-NACIONAL/releases/download")
        update_info = manager.check_for_updates()
        if update_info:
            self.log(f"🚀 Nova versão {update_info.TargetVersion} disponível! Clica 'Verificar Updates' pra instalar.")
            # Optional: messagebox automático if quiser
            # messagebox.showinfo("Atualização", "Nova versão disponível!")
    except:
        pass  # Se erro, continua normal

def update_progress(self, progress):
    self.log(f"Download progresso: {progress}%")

# ============================= INTERFACE CUSTOMTKINTER =============================
ctk.set_appearance_mode("system")
ctk.set_default_color_theme("dark-blue")

class NFSeDownloaderApp:
    def __init__(self, root):
        # Definir fontes globais
        self.font_normal = ctk.CTkFont(family="Segoe UI", size=12)
        self.font_bold = ctk.CTkFont(family="Segoe UI", size=14, weight="bold")
        self.font_title = ctk.CTkFont(family="Segoe UI", size=24, weight="bold")

        self.root = root
        self.root.title("Download NFS-e - Portal Nacional (Multiempresas)")
        self.root.geometry("1080x720")
        self.root.minsize(1000, 650)

        # Ícone da janela
        try:
            self.root.iconbitmap(os.path.join(BASE_DIR, 'icone.ico'))
            self.after(5000, self.checar_updates_auto)  
        except Exception as e:
            print(f"Erro ao carregar ícone: {e}")
            pass 

        # Header
        header = ctk.CTkFrame(root, height=80, corner_radius=0, fg_color="#1e40af")
        header.pack(fill="x")
        header.pack_propagate(False)
        ctk.CTkLabel(header, text="Portal Nacional - Download de NFS-e (Multiempresas)",
                     font=self.font_title, text_color="white").pack(pady=20, padx=30, anchor="w")

        main = ctk.CTkFrame(root)
        main.pack(fill="both", expand=True, padx=25, pady=20)

        # Config
        cfg = ctk.CTkFrame(main, corner_radius=12)
        cfg.pack(fill="x", pady=(15, 25))
        ctk.CTkLabel(cfg, text="Configurações", font=self.font_title).pack(pady=20, padx=30, anchor="w")

        r1 = ctk.CTkFrame(cfg)
        r1.pack(fill="x", padx=30, pady=10)
        ctk.CTkLabel(r1, text="Pasta de downloads:", font=self.font_bold, width=180, anchor="w").pack(side="left", padx=30)
        self.var_pasta = ctk.StringVar(value=PASTA_DOWNLOADS)
        ctk.CTkEntry(r1, textvariable=self.var_pasta, height=45, font=self.font_normal).pack(side="left", fill="x", expand=True, padx=(15,0))
        self.btn_procurar = ctk.CTkButton(r1, text="Procurar...", width=130, height=45, font=self.font_bold, fg_color="#2563eb", hover_color="#1d4ed8", command=self.escolher_pasta)
        self.btn_procurar.pack(side="right", padx=(15,0))

        r2 = ctk.CTkFrame(cfg)
        r2.pack(fill="x", padx=30, pady=10)
        ctk.CTkLabel(r2, text="Competência (MM/AAAA):", font=self.font_bold, width=180, anchor="w").pack(side="left", padx=30)
        self.var_comp = ctk.StringVar(value=COMPETENCIA_DESEJADA)
        ctk.CTkEntry(r2, textvariable=self.var_comp, width=100, height=45, font=self.font_normal, placeholder_text="ex: 11/2025").pack(side="left", padx=15)

        # Botões
        btnspace = ctk.CTkFrame(main, fg_color="transparent")
        btnspace.pack(fill="x", pady=20)
        self.btn_salvar = ctk.CTkButton(btnspace, text="Salvar Configurações", width=220, height=50, font=self.font_bold, fg_color="#059669", hover_color="#047857", command=self.salvar_configuracoes)
        self.btn_salvar.pack(side="left", padx=20)
        ctk.CTkButton(btnspace, text="Limpar Log", width=160, height=50, font=self.font_bold, fg_color="#dc2626", hover_color="#b91c1c", command=self.limpar_log).pack(side="left", padx=20)
        self.btn_update = ctk.CTkButton(btnspace, text="Verificar Updates", command=self.checar_updates, width=140, height=50, fg_color="#f39c12")
        self.btn_update.pack(side="left", padx=12)
        self.btn_start = ctk.CTkButton(btnspace, text="Baixar NFS-e (Multiempresas)", width=150, height=35,
        font=self.font_title, fg_color="#1e40af", hover_color="#1d4ed8",
        command=self.iniciar_download)
        self.btn_start.pack(side="right", padx=20)

        # Log
        logbox = ctk.CTkFrame(main, corner_radius=15)
        logbox.pack(fill="both", expand=True, pady=(15,0))
        ctk.CTkLabel(logbox, text="Log da execução", font=self.font_bold).pack(pady=(20,10), padx=30, anchor="w")
        self.txt_log = ctk.CTkTextbox(logbox, font=ctk.CTkFont(family="Consolas", size=12))
        self.txt_log.pack(fill="both", expand=True, padx=30, pady=(0,30))

    def log(self, msg):
        self.txt_log.insert("end", msg + "\n")
        self.txt_log.see("end")
        self.root.update_idletasks()
        print(msg)

    def limpar_log(self):
        self.txt_log.delete("0.0", "end")

    def escolher_pasta(self):
        p = filedialog.askdirectory(initialdir=self.var_pasta.get())
        if p:
            self.var_pasta.set(p)
            global PASTA_DOWNLOADS
            PASTA_DOWNLOADS = p

    def salvar_configuracoes(self):
        global PASTA_DOWNLOADS, COMPETENCIA_DESEJADA
        pasta = self.var_pasta.get().strip()
        comp = self.var_comp.get().strip()
        if not pasta or not comp or len(comp) != 7 or "/" not in comp:
            messagebox.showwarning("Atenção", "Verifique pasta e competência (MM/AAAA)")
            return
        PASTA_DOWNLOADS = pasta
        COMPETENCIA_DESEJADA = comp
        CONFIG.update({"pasta_downloads": pasta, "competencia_desejada": comp})
        salvar_config(CONFIG)
        self.log("Configurações salvas!")

    def iniciar_download(self):
        self.btn_start.configure(state="disabled", text="Processando...")
        try:
            self._rodar_multiempresas()
        finally:
            self.btn_start.configure(state="normal", text="Baixar NFS-e (Multiempresas)")

    def _rodar_multiempresas(self):
        global COMPETENCIA_DESEJADA, SITUACOES_POR_ARQUIVO, PDF_POR_ARQUIVO
        COMPETENCIA_DESEJADA = self.var_comp.get().strip() or COMPETENCIA_DESEJADA
        self.log("\n" + "="*90)
        self.log(f"Iniciando multiempresas - Competência: {COMPETENCIA_DESEJADA}")
        self.log("="*90)

        carregar_notas_existentes(PASTA_DOWNLOADS, COMPETENCIA_DESEJADA, self.log)
        SITUACOES_POR_ARQUIVO = {}
        PDF_POR_ARQUIVO = {}

        empresa = 0
        while True:
            empresa += 1
            self.log(f"\n{'='*20} EMPRESA #{empresa} {'='*20}")
            driver = None
            try:
                criar_pasta_downloads(PASTA_DOWNLOADS)
                xml_antes = {f for f in os.listdir(PASTA_DOWNLOADS) if f.lower().endswith('.xml')}
                driver = criar_driver(headless=CONFIG.get('headless', False))
                driver.get(URL_PORTAL)

                messagebox.showinfo("Atenção", "1) Faça login\n2) Acesse Notas Emitidas\n3) Clique OK quando a tabela aparecer")

                situacoes_dict = {}
                pagina = 1
                while True:
                    self.log(f"--- PÁGINA {pagina} ---")
                    res = processar_pagina(driver, COMPETENCIA_DESEJADA, situacoes_dict, self.log)
                    aguardar_downloads(PASTA_DOWNLOADS, log_fn=self.log)
                    if res == -1:
                        break
                    if res == 0 and not tem_proxima_pagina(driver, self.log):
                        break
                    if res > 0 and not tem_proxima_pagina(driver, self.log):
                        break
                    pagina += 1

                xml_depois = {f for f in os.listdir(PASTA_DOWNLOADS) if f.lower().endswith('.xml')}
                novos = sorted(xml_depois - xml_antes)
                self.log(f"Novos XMLs nesta empresa: {len(novos)}")
                if novos:
                    organizar_xmls_e_gerar_relatorios_rodada(PASTA_DOWNLOADS, COMPETENCIA_DESEJADA, novos, situacoes_dict, self.log)
            except Exception as e:
                self.log(f"ERRO na empresa {empresa}: {e}")
            finally:
                if driver:
                    try: driver.quit()
                    except: pass

            if not messagebox.askyesno("Próxima empresa", "Deseja processar outro CNPJ?"):
                break

        self.log("\n" + "="*90)
        self.log("PROCESSO FINALIZADO COM SUCESSO!")
        self.log("="*90)

# ============================= FINAL =============================
if __name__ == "__main__":
    velopack.App().run()
    root = ctk.CTk()
    app = NFSeDownloaderApp(root)
    root.mainloop()
