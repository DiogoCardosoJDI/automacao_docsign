import os
import sys
import logging
import pandas as pd

# Adiciona o diretório web ao sys.path
sys.path.append(os.path.abspath('../web'))
from web import inserir_informacoes, acessar_docsign

logger = logging.getLogger(__name__)

def ler_excel_e_inserir_dados(caminho_arquivo_excel, nome_sheet, nome_sheet_ingles, url, pasta_download):
    try:
        logger.info("Abrindo o arquivo Excel para realizar a leitura das informações")
        try:
            logger.info("Abrindo o arquivo Excel com sheet nomeada em PT-BR")
            df = pd.read_excel(caminho_arquivo_excel, sheet_name=nome_sheet)
        except Exception as e:
            logger.warning(f"Erro - {e} - Abrindo o arquivo Excel com sheet nomeada em Ingles")
            df = pd.read_excel(caminho_arquivo_excel, sheet_name=nome_sheet_ingles)

        logger.info("Separando colunas com datas a serem formatadas '['T', 'AC', 'AG', 'AN']'")
        colunas_datas = ['T', 'AC', 'AG', 'AN']

        logger.info("Formatando as colunas de data no formato dd/mm/yyyy e tratando valores nulos")
        for coluna in colunas_datas:
            if coluna in df.columns:
                logger.info("Realizando a conversão das datas")
                try:
                    df[coluna] = pd.to_datetime(df[coluna], format='%d-%b-%y', errors='coerce', dayfirst=True)
                except Exception as e:
                    logger.error(f"Erro ao converter coluna {coluna}: {e}")

                logger.info("Substituindo NaT ou NaN por um valor padrão, como '01/01/1900' ou 'Data inválida'")
                df[coluna] = df[coluna].fillna(pd.Timestamp('1900-01-01'))  # Substitui por data padrão

                logger.info("Formatando no formato dd/mm/yyyy")
                df[coluna] = df[coluna].dt.strftime('%d/%m/%Y')  # Formata no formato pt-BR (dd/mm/yyyy)

        logger.info("Iterando sobre as linhas e processando as que contêm 'Elaborar' na coluna A")
        for idx, row in df.iterrows():
            if row.iloc[0] == 'Elaborar':  # Coluna A é o primeiro índice (0)
                logger.info(f"Processando linha {idx}: {row}")

                logger.info("Criando o dicionário com valores das colunas B a AM (colunas 1 a 38)")
                colunas_interesse = df.columns[1:39]
                dados = {coluna: row[coluna] for coluna in colunas_interesse}
                logger.info(f"Dados coletados: {dados}")

                logger.info("Iniciando Chrome e acessando URL")
                chrome = acessar_docsign.abrir_chrome(url, pasta_download)

                logger.info("Iniciando Inserção de informações na primeira parte do documento")
                chrome = inserir_informacoes.gerar_documento_1(chrome, dados)

                logger.info("Iniciando Inserção de informações na segundo parte do documento e geração do arquivo PDF")
                inserir_informacoes.gerar_documento_2(chrome, dados)

                logger.info("Preenchimento finalizado e PDF gerado, marcando como 'Gerado' na coluna A")
                df.at[idx, 'Unnamed: 0'] = 'Gerado'

                logger.info("Salvando o arquivo Excel com a alteração na informação da coluna 'A'")
                df.to_excel(caminho_arquivo_excel, index=False)

        logger.info(f"Arquivo processado e atualizado com sucesso em: {caminho_arquivo_excel}")

    except Exception as e:
        logger.error(f"Erro durante processamento {__name__} - Erro: {e}")