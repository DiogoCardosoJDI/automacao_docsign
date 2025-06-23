import logging
import xml.etree.ElementTree as ET

logger = logging.getLogger(__name__)
def ler_xml():
    try:
        logger.info("Iniciando leitura do arquivo 'XML' para captura das informações")
        tree = ET.parse(r'xml\config_variaveis.xml')
        root = tree.getroot()

        for web in root.findall('web'):
           url = web.find('url').text
        
        for arquivo_entrada in root.findall('arquivo_entrada'):
           excel = arquivo_entrada.find('excel').text
           sheet = arquivo_entrada.find('sheet').text
           sheet_ingles = arquivo_entrada.find('sheet_ingles').text

        for diretorio in root.findall('diretorio'):
           pasta_download = diretorio.find('download').text
           
        logger.info("Finalizando leitura do arquivo 'XML' e retornando valores coletados")

        return url, excel, sheet, sheet_ingles, pasta_download

    except ET.ParseError as e:
        logger.info(f"Erro ao parsear o XML: {e}")
    except FileNotFoundError:
        logger.info(f"Arquivo não encontrado: {r'xml\config_variaveis.xml'}")
    except Exception as e:
        logger.info(f"Erro inesperado: {e}")
    
    return None
