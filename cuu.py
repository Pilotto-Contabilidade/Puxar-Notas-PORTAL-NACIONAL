# -*- coding: utf-8 -*-
import os
import time
import json
import logging
import sys
import datetime
import xml.etree.ElementTree as ET
import xml.dom.minidom as minidom
import pandas as pd
from tqdm import tqdm
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# ============================= CONFIGURAÇÕES =============================
CONFIG_FILE = os.path.join(os.path.dirname(__file__), 'cuu_config.json')
DEFAULT_CONFIG = {
    "pasta_downloads": r"C:\NFS-e\PortalNacional",
    "competencia_desejada": "10/2025",
    "timeout": 30,
    "headless": False
}
# =========================================================================

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logging.getLogger('selenium.webdriver').setLevel(logging.WARNING)  # Suppress Selenium warnings
logger = logging.getLogger(__name__)

def carregar_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            config = json.load(f)
    else:
        config = DEFAULT_CONFIG.copy()
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4)
    return config

CONFIG = carregar_config()
PASTA_DOWNLOADS = CONFIG['pasta_downloads']
COMPETENCIA_DESEJADA = CONFIG['competencia_desejada']
TIMEOUT = CONFIG['timeout']

def criar_pasta_downloads(pasta):
    os.makedirs(pasta, exist_ok=True)
    print(f"Pasta de downloads: {pasta}")

def criar_driver():
    chrome_options = Options()
    prefs = {
        "download.default_directory": os.path.abspath(PASTA_DOWNLOADS),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "safebrowsing.disable_download_protection": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    if CONFIG.get('headless', False):
        chrome_options.add_argument("--headless")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    driver.maximize_window()
    return driver

def aguardar_downloads(pasta, timeout=600):
    print("Aguardando downloads terminarem...")
    start = time.time()
    while time.time() - start < timeout:
        if not any(f.endswith('.crdownload') for f in os.listdir(pasta)):
            time.sleep(2)
            if not any(f.endswith('.crdownload') for f in os.listdir(pasta)):
                print("Downloads concluídos.\n")
                return
        time.sleep(1)
    print("Timeout. Continuando...\n")

def baixar_xml_da_linha(driver, linha, num):
    try:
        competencia = linha.find_element(By.XPATH, ".//td[contains(@class, 'td-competencia')]").text.strip()
        if COMPETENCIA_DESEJADA not in competencia:
            print(f"Linha {num}: Ignorada → {competencia}")
            return False

        link_xml = linha.find_element(By.XPATH,
                                      ".//td[contains(@class,'td-opcoes')]//a[contains(@href,'/EmissorNacional/Notas/Download/NFSe/')]")
        url_xml = link_xml.get_attribute("href")
        driver.get(url_xml)
        print(f"Linha {num}: BAIXADO → {competencia}")
        return True

    except Exception as e:
        print(f"Linha {num}: Falha → {str(e)[:100]}")
        return False

def processar_pagina(driver, contador):
    WebDriverWait(driver, TIMEOUT).until(EC.presence_of_all_elements_located((By.XPATH, "//table//tbody//tr[td]")))
    linhas = driver.find_elements(By.XPATH, "//table//tbody//tr[td]")
    print(f"Página atual: {len(linhas)} notas encontradas")

    if linhas:
        competencia_primeira = linhas[0].find_element(By.XPATH, ".//td[contains(@class, 'td-competencia')]").text.strip()
        if competencia_primeira != COMPETENCIA_DESEJADA and competencia_primeira < COMPETENCIA_DESEJADA:
            print("Primeira nota da página é de competência anterior → parado!")
            return 0  # Encerra sem processar, economiza tempo

    baixadas_na_pagina = 0
    for i, linha in enumerate(linhas, 1):
        if baixar_xml_da_linha(driver, linha, i):
            contador[0] += 1
            baixadas_na_pagina += 1
        time.sleep(0.8)  # Velocidade aumentada

    print(f"→ {baixadas_na_pagina} notas de {COMPETENCIA_DESEJADA} baixadas nesta página\n")
    return baixadas_na_pagina

# ←←← XPATH CORRIGIDO 100% COM BASE NO SEU HTML
def tem_proxima_pagina(driver):
    try:
        # Funciona mesmo se estiver desabilitado (só checa se existe e tem pg=)
        btn = driver.find_element(By.XPATH, "//a[contains(@href, 'Emitidas?pg=') and contains(@data-original-title, 'Próxima')]")
        # Verifica se não está dentro de um <li class="disabled">
        if "disabled" in btn.find_element(By.XPATH, "./ancestor::li").get_attribute("class"):
            return False
        driver.execute_script("arguments[0].click();", btn)
        time.sleep(2.5)  # Velocidade aumentada
        return True
    except:
        return False

def safe_float(val):
    try:
        return float(val or 0)
    except (ValueError, TypeError):
        logger.warning(f"Falha ao converter {val} para float")
        return 0

def parse_xml_por_nota(xml_path):
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
        nspace = 'http://www.sped.fazenda.gov.br/nfse'

        # Emitente
        emit = root.find(f'.//{{{nspace}}}emit')
        emit_nome = (emit.find(f'{{{nspace}}}xNome').text if emit and emit.find(f'{{{nspace}}}xNome') is not None else '') if emit else ''
        emit_cnpj = (emit.find(f'{{{nspace}}}CNPJ').text if emit and emit.find(f'{{{nspace}}}CNPJ') is not None else '') if emit else ''

        # Tomador
        toma = root.find(f'.//{{{nspace}}}DPS/{{{nspace}}}infDPS/{{{nspace}}}toma')
        toma_nome = (toma.find(f'{{{nspace}}}xNome').text if toma and toma.find(f'{{{nspace}}}xNome') is not None else '') if toma else ''
        toma_cnpj = (toma.find(f'{{{nspace}}}CNPJ').text if toma and toma.find(f'{{{nspace}}}CNPJ') is not None else '') if toma else ''

        # Valores (top level)
        valores = root.find(f'.//{{{nspace}}}valores')
        v_bc = safe_float(valores.find(f'{{{nspace}}}vBC').text if valores and valores.find(f'{{{nspace}}}vBC') is not None else '0')
        v_liq = safe_float(valores.find(f'{{{nspace}}}vLiq').text if valores and valores.find(f'{{{nspace}}}vLiq') is not None else '0')

        # Valores DPS (mais confiável - vServPrest)
        infDPS = root.find(f'.//{{{nspace}}}DPS/{{{nspace}}}infDPS')
        valores_dps = infDPS.find(f'.//{{{nspace}}}valores') if infDPS else None
        v_serv = safe_float(valores_dps.find(f'.//{{{nspace}}}vServPrest/{{{nspace}}}vServ').text if valores_dps and valores_dps.find(f'.//{{{nspace}}}vServPrest/{{{nspace}}}vServ') is not None else '0') if valores_dps else 0

        # Outros
        infNFSe = root.find(f'.//{{{nspace}}}infNFSe')
        n_nfse = (infNFSe.find(f'{{{nspace}}}nNFSe').text if infNFSe and infNFSe.find(f'{{{nspace}}}nNFSe') is not None else '') if infNFSe else ''
        d_compet_raw = (infDPS.find(f'{{{nspace}}}dCompet').text if infDPS and infDPS.find(f'{{{nspace}}}dCompet') is not None else '') if infDPS else ''
        d_compet = d_compet_raw
        if d_compet_raw:
            try:
                # Formato brasileiro DD/MM/YYYY
                d_compet = datetime.datetime.strptime(d_compet_raw, "%Y-%m-%d").strftime("%d/%m/%Y")
            except ValueError:
                pass  # Mantém raw se erro
        serv = infDPS.find(f'{{{nspace}}}serv') if infDPS else None
        x_desc = (serv.find(f'{{{nspace}}}cServ/{{{nspace}}}xDescServ').text if serv and serv.find(f'{{{nspace}}}cServ/{{{nspace}}}xDescServ') is not None else '') if serv else ''
        codigo_serv = (serv.find(f'{{{nspace}}}cServ/{{{nspace}}}cTribNac').text if serv and serv.find(f'{{{nspace}}}cServ/{{{nspace}}}cTribNac') is not None else '') if serv else ''

        return {
            'arquivo': os.path.basename(xml_path),
            'numero_nota': n_nfse,
            'emitente_nome': emit_nome,
            'emitente_cnpj': emit_cnpj,
            'tomador_nome': toma_nome,
            'tomador_cnpj': toma_cnpj,
            'competencia': d_compet,
            'valor_bc': v_bc,
            'valor_liq': v_liq,
            'valor_servico': v_serv,
            'descricao_serv': x_desc,
            'codigo_serv': codigo_serv
        }
    except Exception as e:
        logger.error(f"Erro ao parsear {xml_path}: {e}")
        return None

def gerar_relatorio(pasta_xmls, competency):
    try:
        xml_files = [f for f in os.listdir(pasta_xmls) if f.endswith('.xml')]
        dados = []
        for xml_file in tqdm(xml_files, desc="Processando XMLs"):
            data = parse_xml_por_nota(os.path.join(pasta_xmls, xml_file))
            if data:
                dados.append(data)

        if not dados:
            print("Nenhum dado válido encontrado nos XMLs.")
            return

        df = pd.DataFrame(dados)

        # Renomear cabeçalhos para português brasileiro
        df.rename(columns={
            'arquivo': 'Arquivo',
            'numero_nota': 'Número da Nota',
            'emitente_nome': 'Emitente',
            'emitente_cnpj': 'CNPJ Emitente',
            'tomador_nome': 'Tomador',
            'tomador_cnpj': 'CNPJ Tomador',
            'competencia': 'Competência',
            'valor_bc': 'Valor BC',
            'valor_liq': 'Valor Líquido',
            'valor_servico': 'Valor Serviço',
            'descricao_serv': 'Descrição Serviço',
            'codigo_serv': 'Código Serviço'
        }, inplace=True)

        # Totais
        total_notas = len(df)
        total_valor_liq = df['Valor Líquido'].sum()
        total_valor_bc = df['Valor BC'].sum()

        # Por tomador (filtrar valores vazios)
        df_tomadores = df[df['CNPJ Tomador'].notna() & (df['CNPJ Tomador'] != '')]
        rel_tomadores = df_tomadores.groupby(['CNPJ Tomador', 'Tomador']).agg(
            total_notas=('Arquivo', 'count'),
            total_valor=('Valor Líquido', 'sum')
        ).sort_values('total_valor', ascending=False).head(10)

        # Salvar Excel
        relatorio_path = os.path.join(pasta_xmls, f'relatorio_nfse_{competency.replace("/", "_")}.xlsx')
        with pd.ExcelWriter(relatorio_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Detalhe_Notas', index=False)
            rel_tomadores.to_excel(writer, sheet_name='Top_Tomadores')

        print("\n" + "="*95)
        print("RELATÓRIO GERADO COM SUCESSO!")
        print(f"Total de notas processadas: {total_notas}")
        print(f"Total valor líquido: R$ {total_valor_liq:,.2f}")
        print(f"Total base cálculo: R$ {total_valor_bc:,.2f}")
        print(f"Arquivo: {relatorio_path}")
        print("="*95)

        # Exibir top tomadores no console
        print("\nTop 10 Tomadores por Valor:")
        for _, row in rel_tomadores.iterrows():
            print(f"{row.name[1]} ({row.name[0]}): {row['total_notas']} notas, R$ {row['total_valor']:,.2f}")

    except Exception as e:
        logger.error(f"Erro ao gerar relatório: {e}")
        print(f"Falha na geração do relatório: {e}")

def main(competency_override=None):
    global COMPETENCIA_DESEJADA
    if competency_override:
        COMPETENCIA_DESEJADA = competency_override

    criar_pasta_downloads(PASTA_DOWNLOADS)
    driver = criar_driver()

    try:
        driver.get("https://www.nfse.gov.br/EmissorNacional")
        print("\n" + "="*95)
        print(f"BAIXANDO TODAS AS NFS-e DE COMPETÊNCIA: {COMPETENCIA_DESEJADA}")
        print("1. Faz login → Perfil PRESTADOR")
        print("2. Vai em: Serviços Prestados → Notas Emitidas")
        print("3. Quando aparecer a lista → aperta ENTER aqui")
        print("="*95)
        input(">>> ENTER QUANDO ESTIVER NA TELA DAS NOTAS <<<")

        contador = [0]
        pagina = 1

        while True:
            print(f"{'='*45} PAGINA {pagina} {'='*45}")
            baixadas = processar_pagina(driver, contador)
            aguardar_downloads(PASTA_DOWNLOADS)

            if baixadas == 0:
                print("Nenhuma nota do mês encontrada → fim do download!")
                break

            if not tem_proxima_pagina(driver):
                print("Botão 'Próxima' não encontrado ou desabilitado → fim das páginas.")
                break

            pagina += 1

        print("\n" + "="*95)
        print(f"FINALIZADO COM SUCESSO!")
        print(f"Total de notas baixadas ({COMPETENCIA_DESEJADA}): {contador[0]}")
        print(f"Arquivos em: {PASTA_DOWNLOADS}")
        print("="*95)

        # Gerar relatório após downloads
        print("GERANDO RELATÓRIO DOS XMLS...")
        gerar_relatorio(PASTA_DOWNLOADS, COMPETENCIA_DESEJADA)

    except Exception as e:
        print(f"ERRO: {e}")
    finally:
        input("\nPressione ENTER para fechar...")
        driver.quit()

if __name__ == "__main__":
    # Parse args: allow --competencia MM/YYYY followed by optional --report-only
    COMPETENCIA_DESEJADA_USE = COMPETENCIA_DESEJADA
    is_report_only = False

    if len(sys.argv) >= 3 and sys.argv[1] == "--competencia":
        COMPETENCIA_DESEJADA_USE = sys.argv[2]
        if len(sys.argv) >= 4 and sys.argv[3] == "--report-only":
            is_report_only = True
    elif len(sys.argv) >= 2 and sys.argv[1] == "--report-only":
        is_report_only = True

    if is_report_only:
        gerar_relatorio(PASTA_DOWNLOADS, COMPETENCIA_DESEJADA_USE)
    else:
        main(COMPETENCIA_DESEJADA_USE)
