import os
import logging
from files import leitura_excel
from carregar_configuracao import carrega_config

def configurar_logging(nome_arquivo_log="app.log", nivel=logging.INFO):
    os.makedirs("logs", exist_ok=True)  # Cria pasta se não existir

    caminho_log = os.path.join("logs", nome_arquivo_log)

    logging.basicConfig(
        filename=caminho_log,
        level=nivel,
        format="%(asctime)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        filemode='a'  # ou 'w' para sobrescrever cada vez
    )

    # Também mostra no terminal:
    console = logging.StreamHandler()
    console.setLevel(nivel)
    console.setFormatter(logging.Formatter("%(levelname)s - %(message)s"))
    logging.getLogger().addHandler(console)

configurar_logging()
logger = logging.getLogger(__name__)

try:
    logger.info("Carregando informações de configurações iniciais")
    url, caminho_arquivo_excel, nome_sheet, sheet_ingles, pasta_download = carrega_config.ler_xml()

    logger.info("Iniciando leitura do arquivo Excel para verificação de contratos a serem gerados")
    leitura_excel.ler_excel_e_inserir_dados(caminho_arquivo_excel, nome_sheet, sheet_ingles, url, pasta_download)
    
except Exception as e:
    logger.error(f"Erro durante o processamento - {e}")