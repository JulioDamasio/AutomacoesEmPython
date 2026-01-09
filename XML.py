import tkinter as tk
from tkinter import ttk
import pandas as pd
from datetime import datetime
import os

# Leitura da planilha Excel
df = pd.read_excel(r'W:\B - TED\7 - AUTOMAÇÃO\Lote RC\Lote RC.xlsx')

def generate_xml(data_geracao, sequencial_geracao, ano_referencia, ug_responsavel, cpf_responsavel, output_directory):
    
    print('Acessando arquivo excel do lote para a geração do XMl...')

    # Modelo XML
    modelo_xml = f"""
    <sb:arquivo xmlns:sb="http://www.tesouro.gov.br/siafi/submissao">
        <sb:header>
            <sb:codigoLayout>DH001</sb:codigoLayout>
            <sb:dataGeracao>{data_geracao}</sb:dataGeracao>
            <sb:sequencialGeracao>{sequencial_geracao}</sb:sequencialGeracao>
            <sb:anoReferencia>{ano_referencia}</sb:anoReferencia>
            <sb:ugResponsavel>{ug_responsavel}</sb:ugResponsavel>
            <sb:cpfResponsavel>{cpf_responsavel}</sb:cpfResponsavel>
        </sb:header>
        <sb:detalhes>
        """
        
    # Inicialize uma variável para rastrear o número de documentos gerados
    quantidade_documentos = 0

    # Certifique-se de que sequencial_geracao seja um número inteiro
    sequencial_geracao = int(sequencial_geracao)

    # Inicialize numero_documento com 400001
    numero_documento = 520000
        
    # Defina num_seq_item fora do loop
    num_seq_item = sequencial_geracao
   
    # Itera sobre as linhas de dados da planilha
    for index, row in df.iloc[1:].iterrows():
        cod_tipo_dh = row[0]  # Pega o valor da coluna A para cada linha
        cod_ug_emit = row[1]
        dt_emissao = row[2]
        dt_venc = row[3]
        vl_doc = round(float(row[7]), 2) if pd.notna(row[7]) else 0.00
        obs = row[9]
        cod_credor = row[8]
        situacao = row[10]
        vl_ol = round(float(row[15]), 2) if pd.notna(row[15]) else 0.00
        siafi = row[14]
        ce = row[13]
        fonte = row[12]
        NormalEstorno = row[11]
        dt_emissao = datetime.strptime(str(row[2]), "%Y-%m-%d %H:%M:%S").date()
        dt_venc = datetime.strptime(str(row[3]), "%Y-%m-%d %H:%M:%S").date()

        vl_doc_str = f"{vl_doc:.2f}"  # Garante formato com duas casas decimais
        vl_ol_str = f"{vl_ol:.2f}"  # Garante formato com duas casas decimais

        # Processa os valores em listas
        if isinstance(fonte, list):
            fonte = fonte[0] if fonte else ""
        if isinstance(ce, list):
            ce = ce[0] if ce else ""
        if isinstance(NormalEstorno, list):
            NormalEstorno = NormalEstorno[0] if NormalEstorno else ""    
        quantidade_documentos += 1  # Incrementa o número de documentos

        # Use str.zfill para formatar o número do documento com 6 dígitos
        numero_documento = str(int(numero_documento) + 1).zfill(6)

        # Inicializar num_seq_item para cada iteração
        num_seq_item = str(int(num_seq_item) + 1).zfill(4)
        
        # Use str.zfill para formatar o número sequencial com 4 dígitos (baseado em num_seq_item)
        num_seq_item_str = str(num_seq_item).zfill(4)

        # Adicionar uma nova tag sb:detalhe para cada linha do Excel
        modelo_xml += f"""
        <sb:detalhe>    
            <cpr:CprDhCadastrar xmlns:cpr="http://services.docHabil.cpr.siafi.tesouro.fazenda.gov.br/">
                <codUgEmit>{ug_responsavel}</codUgEmit>
                <anoDH>{ano_referencia}</anoDH>
                <codTipoDH>{(cod_tipo_dh)}</codTipoDH>
                <numDH>{(numero_documento)}</numDH>
                <dadosBasicos>
                        <dtEmis>{(dt_emissao)}</dtEmis>
                        <dtVenc>{(dt_venc)}</dtVenc>
                        <codUgPgto>{(cod_ug_emit)}</codUgPgto>
                        <vlr>{(vl_doc_str)}</vlr>
                        <txtObser>{(obs)}</txtObser>
                        <vlrTaxaCambio>0.0000</vlrTaxaCambio>
                        <dtAteste>{(dt_emissao)}</dtAteste>
                        <codCredorDevedor>{(cod_credor)}</codCredorDevedor>
                    <docOrigem>
                        <codIdentEmit>{(cod_credor)}</codIdentEmit>
                        <dtEmis>{(dt_emissao)}</dtEmis>
                        <numDocOrigem>0</numDocOrigem>
                        <vlr>{(vl_doc_str)}</vlr>
                    </docOrigem>
                </dadosBasicos>
                <outrosLanc>
                    <numSeqItem>0001</numSeqItem>
                    <codSit>{(situacao)}</codSit>
                    <vlr>{(vl_ol_str)}</vlr>
                    <indrTemContrato>0</indrTemContrato>
                    <txtInscrA>{(fonte)}</txtInscrA>
                    <txtInscrB>{(ce)}</txtInscrB>
                    <txtInscrC>{(siafi)}</txtInscrC>
                    <tpNormalEstorno>{(NormalEstorno)}</tpNormalEstorno>
                </outrosLanc>
            </cpr:CprDhCadastrar>
        </sb:detalhe>
        """
    # Fecha o modelo XML e insere a quantidade de documentos
    modelo_xml += """
        </sb:detalhes>
        <sb:trailler>
            <sb:quantidadeDetalhe>{}</sb:quantidadeDetalhe>
        </sb:trailler>
    </sb:arquivo>
    """.format(quantidade_documentos)

    print('Lote Gerado com sucesso!')

    with open(os.path.join(output_directory, "lote.xml"), "w", encoding="UTF-8") as xml_file:
        xml_file.write(modelo_xml)