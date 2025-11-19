# -*- coding: utf-8 -*-
"""
Script de Download e Relatório de NFS-e do Portal Nacional Brasileiro.

Funcionalidades:
- Download automático de XMLs via Selenium.
- Parsing completo de dados NFS-e (emitente, tomador, valores).
- Geração de relatório Excel com totais e top tomadores.
- Configuração via JSON, suporte a CLI args para competência.
- Formatação brasileira de datas, cabeçalhos em português.

Refatorado para limpeza, modularidade e manutenção fácil.
"""

from __future__ import annotations

import json
import logging
import os
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from tqdm import tqdm
from webdriver_manager.chrome import ChromeDriverManager

# Configurações globais
CONFIG_FILE = Path(__file__).parent / "Portal_Nacional_config.json"
DEFAULT_CONFIG = {
    "pasta_downloads": "C:\\NFS-e\\PortalNacional",
    "competencia_desejada": "10/2025",
    "timeout": 30,
    "headless": False,
    "velocidade_linha": 0.8,
    "velocidade_pagina": 2.5
}

# Constants for XPATH
XPATH_TABELA = "//table//tbody//tr[td]"
XPATH_LINHA_OPCOES = ".//td[contains(@class,'td-opcoes')]//a[contains(@href,'/EmissorNacional/Notas/Download/NFSe/')]"
XPATH_COMPETENCIA = ".//td[contains(@class, 'td-competencia')]"
XPATH_PROXIMA_PAGINA = "//a[contains(@href, 'Emitidas?pg=') and contains(@data-original-title, 'Próxima')]"

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)


class Config:
    """Classe para carregar e validar configuração."""

    def __init__(self, config_path: Path = CONFIG_FILE) -> None:
        self.config_path = config_path
        self._data = self._load_config()

    def _load_config(self) -> Dict:
        if self.config_path.exists():
            with open(self.config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
        else:
            config = DEFAULT_CONFIG.copy()
            self._save_config(config)
            return config

        # Merge with defaults for missing keys
        config = {**DEFAULT_CONFIG, **config}
        # Save updated config if needed
        self._save_config(config)
        return config

    def _save_config(self, config: Dict) -> None:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4)

    def __getitem__(self, key: str):
        return self._data[key]

    def __contains__(self, key: str) -> bool:
        return key in self._data


class NFSeParser:
    """Parse XML de NFS-e."""

    namespace = 'http://www.sped.fazenda.gov.br/nfse'

    def __init__(self, xml_file: str) -> None:
        import xml.etree.ElementTree as ET
        self.xml_file = Path(xml_file)
        self.tree = ET.parse(self.xml_file)
        self.root = self.tree.getroot()

    def safe_text(self, element, tag: str) -> str:
        """Retorna texto seguro de tag XML."""
        if element is not None:
            child = element.find(f'{{{self.namespace}}}{tag}')
            if child is not None:
                return child.text or ''
        return ''

    def parse_note_data(self) -> Optional[Dict[str, any]]:
        """Parse dados principais da nota."""
        try:
            # Emitente
            emit = self.root.find(f'.//{{{self.namespace}}}emit')
            emit_nome = self.safe_text(emit, 'xNome')
            emit_cnpj = self.safe_text(emit, 'CNPJ')

            # Tomador
            toma = self.root.find(f'.//{{{self.namespace}}}DPS/{{{self.namespace}}}infDPS/{{{self.namespace}}}toma')
            toma_nome = self.safe_text(toma, 'xNome')
            toma_cnpj = self.safe_text(toma, 'CNPJ')

            # Valores
            valores = self.root.find(f'.//{{{self.namespace}}}valores')
            v_bc = self._safe_float(self.safe_text(valores, 'vBC'))
            v_liq = self._safe_float(self.safe_text(valores, 'vLiq'))

            # Valores DPS (mais confiável)
            infDPS = self.root.find(f'.//{{{self.namespace}}}DPS/{{{self.namespace}}}infDPS')
            valores_dps = infDPS.find(f'.//{{{self.namespace}}}valores') if infDPS else None
            v_serv = self._safe_float(self.safe_text(valores_dps, 'vServPrest/vServ')) if valores_dps else 0.0

            # Outros
            infNFSe = self.root.find(f'.//{{{self.namespace}}}infNFSe')
            n_nfse = self.safe_text(infNFSe, 'nNFSe') if infNFSe else ''

            d_compet = self.safe_text(infDPS, 'dCompet') if infDPS else ''
            d_compet_formatted = self._format_date(d_compet)

            serv = infDPS.find(f'{{{self.namespace}}}serv') if infDPS else None
            x_desc = self.safe_text(serv, 'cServ/xDescServ') if serv else ''
            codigo_serv = self.safe_text(serv, 'cServ/cTribNac') if serv else ''

            return {
                'arquivo': self.xml_file.name,
                'numero_nota': n_nfse,
                'emitente_nome': emit_nome,
                'emitente_cnpj': emit_cnpj,
                'tomador_nome': toma_nome,
                'tomador_cnpj': toma_cnpj,
                'competencia': d_compet_formatted,
                'valor_bc': v_bc,
                'valor_liq': v_liq,
                'valor_servico': v_serv,
                'descricao_serv': x_desc,
                'codigo_serv': codigo_serv
            }
        except Exception as e:
            logger.error(f"Erro ao parsear {self.xml_file}: {e}")
            return None

    def _safe_float(self, value: str) -> float:
        """Converte string para float de forma segura."""
        try:
            return float(value or 0)
        except (ValueError, TypeError):
            logger.warning(f"Falha ao converter '{value}' para float")
            return 0.0

    def _format_date(self, date_str: str) -> str:
        """Formata data para DD/MM/YYYY."""
        import datetime
        if date_str:
            try:
                return datetime.datetime.strptime(date_str, "%Y-%m-%d").strftime("%d/%m/%Y")
            except ValueError:
                pass
        return date_str


class NFSeDownloader:
    """Classe para download de XMLs via Selenium."""

    def __init__(self, config: Config) -> None:
        self.config = config
        self.driver = None

    def _create_driver(self) -> webdriver.Chrome:
        """Cria driver Chrome configurado."""
        chrome_options = Options()
        prefs = {
            "download.default_directory": os.path.abspath(self.config['pasta_downloads']),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        chrome_options.add_experimental_option("prefs", prefs)
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        if self.config['headless']:
            chrome_options.add_argument("--headless")

        return webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=chrome_options
        )

    def setup_download_folder(self) -> None:
        """Cria pasta de downloads."""
        os.makedirs(self.config['pasta_downloads'], exist_ok=True)

    def start_session(self) -> None:
        """Inicia sessão Selenium e navega ao portal."""
        self.driver = self._create_driver()
        self.driver.get("https://www.nfse.gov.br/EmissorNacional")
        self.driver.maximize_window()

    def download_notes_for_competencia(self, competencia: str) -> int:
        """Baixa notas de uma competência específica."""
        if not self.driver:
            raise RuntimeError("Sessão não iniciada. Chame start_session() primeiro.")

        print(f"\n{'='*95}\nBAIXANDO NFS-e DE COMPETÊNCIA: {competencia}\n{'='*95}")
        input(">>> PRESSIONE ENTER QUANDO ESTIVER NA TELA DAS NOTAS <<<")

        contador = [0]
        pagina = 1
        baixadas_total = 0

        while True:
            print(f"Página {pagina}: processando...")
            baixadas = self._process_page(contador, pagina, competencia)
            baixadas_total += baixadas

            if baixadas == 0:
                print("Nenhuma nota da competência encontrada → fim do download!")
                break

            if not self._next_page():
                print("Botão 'Próxima' não encontrado ou desabilitado → fim das páginas.")
                break

            pagina += 1

        print(f"\nFINALIZADO! Total baixadas: {baixadas_total}")
        return baixadas_total

    def _process_page(self, contador: List[int], pagina: int, competencia: str) -> int:
        """Processa uma página específica."""
        WebDriverWait(self.driver, self.config['timeout']).until(
            EC.presence_of_all_elements_located((By.XPATH, XPATH_TABELA))
        )

        linhas = self.driver.find_elements(By.XPATH, XPATH_TABELA)
        print(f"Página {pagina}: {len(linhas)} notas encontradas")

        # Verificar primeira linha
        if linhas:
            primeira_comp = linhas[0].find_element(By.XPATH, XPATH_COMPETENCIA).text.strip()
            if primeira_comp != competencia and primeira_comp < competencia:
                print("Primeira nota é anterior → parado!")
                return 0

        baixadas = 0
        for i, linha in enumerate(linhas, 1):
            if self._download_note_from_row(linha, contador, competencia):
                baixadas += 1
            from time import sleep
            sleep(self.config['velocidade_linha'])

        return baixadas

    def _download_note_from_row(self, linha, contador: List[int], competencia: str) -> bool:
        """Baixa XML de uma linha se compatível."""
        try:
            comp_linha = linha.find_element(By.XPATH, XPATH_COMPETENCIA).text.strip()
            if competencia not in comp_linha:
                return False

            link_xml = linha.find_element(By.XPATH, XPATH_LINHA_OPCOES)
            link_xml.click()

            contador[0] += 1
            return True
        except Exception as e:
            logger.debug(f"Falha em linha: {str(e)[:50]}")
            return False

    def _next_page(self) -> bool:
        """Navega para próxima página."""
        try:
            btn_proxima = self.driver.find_element(By.XPATH, XPATH_PROXIMA_PAGINA)
            parent = btn_proxima.find_element(By.XPATH, "./ancestor::li")
            if "disabled" in parent.get_attribute("class"):
                return False

            self.driver.execute_script("arguments[0].click();", btn_proxima)
            from time import sleep
            sleep(self.config['velocidade_pagina'])
            return True
        except:
            return False

    def calculate_totals(self, pasta_downloads: str):
        """Calcula totais básicos (placeholder para futuras expansões)."""
        # Implementação futura se necessário
        pass

    def close_session(self) -> None:
        """Fecha sessão Selenium."""
        if self.driver:
            input("\nPressione ENTER para fechar...\n")
            self.driver.quit()


class ReportGenerator:
    """Gera relatório Excel de notas processadas."""

    COLUMN_MAPPING = {
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
    }

    @staticmethod
    def generate_from_folder(pasta_xmls: str, competencia: str) -> None:
        """Gera relatório Excel completo."""
        xml_files = [f for f in os.listdir(pasta_xmls) if f.endswith('.xml')]

        if not xml_files:
            print("Nenhum XML encontrado.")
            return

        dados = []
        for xml_file in tqdm(xml_files, desc="Processando XMLs"):
            parser = NFSeParser(os.path.join(pasta_xmls, xml_file))
            data = parser.parse_note_data()
            if data:
                dados.append(data)

        if not dados:
            print("Nenhum dado válido.")
            return

        df = pd.DataFrame(dados)
        df.rename(columns=ReportGenerator.COLUMN_MAPPING, inplace=True)

        # Totais
        total_notas = len(df)
        total_liq = df['Valor Líquido'].sum()

        # Top tomadores
        df_tomadores = df[df['CNPJ Tomador'].notna() & (df['CNPJ Tomador'] != '')]
        top_tomadores = (
            df_tomadores.groupby(['CNPJ Tomador', 'Tomador'])
            .agg(total_notas=('Arquivo', 'count'), total_valor=('Valor Líquido', 'sum'))
            .sort_values('total_valor', ascending=False).head(10)
        )

        # Salvar Excel
        arquivo_xlsx = os.path.join(pasta_xmls, f'relatorio_nfse_{competencia.replace("/", "_")}.xlsx')
        with pd.ExcelWriter(arquivo_xlsx, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Detalhe Notas', index=False)
            top_tomadores.to_excel(writer, sheet_name='Top Tomadores')

        print(f"\n{'='*95}\nRELATÓRIO GERADO: {arquivo_xlsx}\nTotal notas: {total_notas}\nTotal líquido: R$ {total_liq:,.2f}\n{'='*95}")


def main(competency_override: Optional[str] = None) -> None:
    config = Config()
    competencia = competency_override or config['competencia_desejada']

    downloader = NFSeDownloader(config)
    downloader.setup_download_folder()
    downloader.start_session()

    try:
        baixadas = downloader.download_notes_for_competencia(competencia)
        if baixadas > 0:
            print("\nGerando relatório...")
            ReportGenerator.generate_from_folder(config['pasta_downloads'], competencia)
    except Exception as e:
        logger.error(f"Erro geral: {e}")
    finally:
        downloader.close_session()


if __name__ == "__main__":
    import sys

    competency = None
    report_only = '--report-only' in sys.argv

    if len(sys.argv) >= 3 and sys.argv[1] == '--competencia':
        competency = sys.argv[2]
    elif len(sys.argv) == 4 and '--report-only' in sys.argv:
        report_only = True

    if report_only:
        config = Config()
        pasta = config['pasta_downloads']
        comp = competency or config['competencia_desejada']
        ReportGenerator.generate_from_folder(pasta, comp)
    else:
        main(competency)
