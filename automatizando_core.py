import os
import csv
import xmltodict
from datetime import datetime
import logging

logging.basicConfig()
logger = logging.getLogger('automatizando_core')
logger.setLevel(logging.INFO)

# --- DADOS FIXOS DO RESPONSÁVEL (Edite aqui uma vez só) ---
RESP_CPF = "00000000000"
RESP_NOME = "NOME DO SECRETARIO"
RESP_CARGO = "SECRETARIO"


def criar_csv_rascunho(pasta):
    """Gera o arquivo PLANILHA_CONTROLE_AUDESP.csv na pasta informada.
    Retorna (csv_path, count) ou raises Exception.
    """
    if not os.path.isdir(pasta):
        raise FileNotFoundError(f"Pasta não encontrada: {pasta}")

    arquivos = [f for f in os.listdir(pasta) if f.lower().endswith('.xml')]
    if not arquivos:
        raise FileNotFoundError("Nenhum XML encontrado nesta pasta.")

    csv_path = os.path.join(pasta, "PLANILHA_CONTROLE_AUDESP.csv")

    with open(csv_path, mode='w', newline='', encoding='utf-8-sig') as file:
        writer = csv.writer(file, delimiter=';')
        writer.writerow(['ARQUIVO_XML', 'FORNECEDOR', 'NUM_NOTA', 'DATA_EMISSAO', 'VALOR_TOTAL_NOTA', 'COD_AJUSTE', 'NUM_EMPENHO', 'DATA_EMPENHO', 'VALOR_TOTAL_EMPENHO', 'VALOR_PARCELA_PAGA'])

        count = 0
        for arq in arquivos:
            try:
                caminho_full = os.path.join(pasta, arq)
                with open(caminho_full, "rb") as f:
                    doc = xmltodict.parse(f)

                inf = doc['nfeProc']['NFe']['infNFe']
                num = inf['ide']['nNF']
                data = inf['ide']['dhEmi'].split('T')[0]
                valor = inf['total']['ICMSTot']['vNF']
                fornecedor = inf['emit']['xNome']

                writer.writerow([arq, fornecedor, num, data, valor, '', '', '', '', valor])
                count += 1
            except Exception as e:
                logger.warning(f"Erro ao ler {arq}: {e}")

    logger.info(f"Planilha criada: {csv_path} com {count} notas")
    return csv_path, count


def processar_final(caminho_csv, pasta_xmls, pasta_saida=None):
    """Processa o CSV preenchido e gera os arquivos XML na pasta de saída.
    Retorna número de documentos gerados.
    """
    if not os.path.isfile(caminho_csv):
        raise FileNotFoundError(f"CSV não encontrado: {caminho_csv}")
    if not os.path.isdir(pasta_xmls):
        raise FileNotFoundError(f"Pasta de XMLs não encontrada: {pasta_xmls}")

    if not pasta_saida:
        pasta_saida = os.path.join(pasta_xmls, "SAIDA_AUDESP_PRONTOS")
    if not os.path.exists(pasta_saida):
        os.makedirs(pasta_saida)

    sucesso = 0
    with open(caminho_csv, newline='', encoding='utf-8-sig') as csvfile:
        reader = csv.DictReader(csvfile, delimiter=';')

        for row in reader:
            if not row.get('COD_AJUSTE') or not row.get('NUM_EMPENHO'):
                logger.info(f"PULADO: Nota {row.get('NUM_NOTA')} sem dados de empenho.")
                continue

            arq_xml = row['ARQUIVO_XML']
            ajuste = row['COD_AJUSTE']
            num_empenho = row['NUM_EMPENHO']
            data_empenho = row['DATA_EMPENHO']

            try:
                val_total_nota = float(row['VALOR_TOTAL_NOTA'].replace(',', '.'))
                val_pago = float(row['VALOR_PARCELA_PAGA'].replace(',', '.'))
                val_total_empenho = float(row['VALOR_TOTAL_EMPENHO'].replace(',', '.')) if row.get('VALOR_TOTAL_EMPENHO') else val_total_nota

                percentual = (val_pago / val_total_empenho * 100) if val_total_empenho > 0 else 0

                gerar_xml_fisico(pasta_saida, row, val_pago, val_total_empenho, percentual)
                sucesso += 1

            except ValueError:
                logger.warning(f"Erro de valor na nota {row.get('NUM_NOTA')}")

    logger.info(f"Processo concluído! {sucesso} documentos gerados na pasta {pasta_saida}.")
    return sucesso


def gerar_xml_fisico(pasta, dados, valor_pago, valor_total_empenho, percentual):
    num_nota = dados['NUM_NOTA']
    num_empenho = dados['NUM_EMPENHO']
    cod_ajuste = dados['COD_AJUSTE']
    data_empenho = dados['DATA_EMPENHO']
    fornecedor = dados['FORNECEDOR']

    prefixo = f"{num_nota}_EMP{num_empenho}"

    try:
        xml_exec = f"""<?xml version="1.0" encoding="UTF-8"?>
<execucao>
    <numero_empenho>{num_empenho}</numero_empenho>
    <data_empenho>{data_empenho}</data_empenho>
    <codigo_ajuste>{cod_ajuste}</codigo_ajuste>
    <valor_total>{valor_total_empenho:.2f}</valor_total>
    <fornecedor>{fornecedor}</fornecedor>
    <data_criacao>{datetime.now().strftime('%Y-%m-%d')}</data_criacao>
</execucao>"""
        with open(os.path.join(pasta, f"{prefixo}_EXECUCAO.xml"), 'w', encoding='utf-8') as f:
            f.write(xml_exec)

        xml_fiscal = f"""<?xml version="1.0" encoding="UTF-8"?>
<fiscal>
    <numero_nota>{num_nota}</numero_nota>
    <numero_empenho>{num_empenho}</numero_empenho>
    <valor_nota>{valor_total_empenho:.2f}</valor_nota>
    <valor_fiscalizado>{valor_pago:.2f}</valor_fiscalizado>
    <percentual_pago>{percentual:.2f}</percentual_pago>
    <responsavel_cpf>{RESP_CPF}</responsavel_cpf>
    <responsavel_nome>{RESP_NOME}</responsavel_nome>
    <responsavel_cargo>{RESP_CARGO}</responsavel_cargo>
    <data_fiscalizacao>{datetime.now().strftime('%Y-%m-%d')}</data_fiscalizacao>
</fiscal>"""
        with open(os.path.join(pasta, f"{prefixo}_FISCAL.xml"), 'w', encoding='utf-8') as f:
            f.write(xml_fiscal)

        xml_pago = f"""<?xml version="1.0" encoding="UTF-8"?>
<pagamento>
    <numero_empenho>{num_empenho}</numero_empenho>
    <data_pagamento>{datetime.now().strftime('%Y-%m-%d')}</data_pagamento>
    <valor_pago>{valor_pago:.2f}</valor_pago>
    <percentual_pago>{percentual:.2f}</percentual_pago>
    <fornecedor>{fornecedor}</fornecedor>
    <status>PAGO</status>
</pagamento>"""
        with open(os.path.join(pasta, f"{prefixo}_PAGAMENTO.xml"), 'w', encoding='utf-8') as f:
            f.write(xml_pago)

        logger.info(f"✓ Gerados 3 XMLs para Nota {num_nota} (Empenho {num_empenho})")
    except Exception as e:
        logger.error(f"Erro ao gerar XMLs para {num_nota}: {e}")
