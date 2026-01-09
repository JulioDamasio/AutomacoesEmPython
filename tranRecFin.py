import pandas as pd
from datetime import datetime
import os

# Leitura da planilha Excel
df = pd.read_excel(r'W:\B - TED\7 - AUTOMAÇÃO\TranRecFin\TranRecFin.xlsx')

def generate_xml_tranRecFin(data_geracao, sequencial_geracao, ano_referencia, ug_responsavel, cpf_responsavel, output_directory, sequencial_doc):
    
    modelo_xml = f"""
        <sb:arquivo xmlns:sb="http://www.tesouro.gov.br/siafi/submissao">
            <sb:header>
                <sb:codigoLayout>PF003</sb:codigoLayout>
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

    # Inicialize numero_documento com o valor passado pelo usuário
    numero_documento = str(sequencial_doc)
        
    # Defina num_seq_item fora do loop
    num_seq_item = sequencial_geracao
    
    # Use str.zfill para formatar o número do documento com 6 dígitos
    numero_documento = str(int(numero_documento) + 1).zfill(6)

    # Inicializar num_seq_item para cada iteração
    num_seq_item = str(int(num_seq_item) + 1).zfill(4)
    
    # Use str.zfill para formatar o número sequencial com 4 dígitos (baseado em num_seq_item)
    num_seq_item_str = str(num_seq_item).zfill(4)
    
    # Inicialize as variáveis necessárias fora do loop
    siafi = ""
    codigo_pf = ""
   
    # Itera sobre as linhas de dados da planilha
    for index, row in df.iloc[0:].iterrows():
        ug_emitente = row[0] if pd.notna(row[0]) else " " # Pega o valor da coluna A para cada linha
        observacao = row[1] if pd.notna(row[1]) else " "
        dt_emissao = row[2] if pd.notna(row[2]) else " "
        ug_favorecida = str(row[3]).strip() if pd.notna(row[3]) else ""
        ano = row[4] if pd.notna(row[4]) else " "
        valor = row[5] if pd.notna(row[5]) else " "
        vinculacao_pagamento = row[6] if pd.notna(row[6]) else " "
        fonte = row[7] if pd.notna(row[7]) else " "
        categoria_gasto = row[8] if pd.notna(row[8]) else " "
        cod_situacao = row[9] if pd.notna(row[9]) else " "
        codigo_recurso = row[10] if pd.notna(row[10]) else " "
        codigo_pf = row[11]    
        saldo_pf = row[12] if pd.notna(row[12]) else " "
        siafi = row[13]
        dt_emissao = datetime.strptime(str(row[2]), "%Y-%m-%d %H:%M:%S").date()
       
        # Processa os valores em listas
        if isinstance(siafi, list):
            siafi = siafi[0] if siafi else ""
        if isinstance(codigo_pf, list):
            codigo_pf = codigo_pf[0] if codigo_pf else ""
        
        quantidade_documentos += 1  # Incrementa o número de documentos

        # Use str.zfill para formatar o número do documento com 6 dígitos
        numero_documento = str(int(numero_documento) + 1)

        # Inicializar num_seq_item para cada iteração
        num_seq_item = str(int(num_seq_item) + 1).zfill(4)
        
        # Use str.zfill para formatar o número sequencial com 4 dígitos (baseado em num_seq_item)
        num_seq_item_str = str(num_seq_item).zfill(4)
        
        # Adicionar uma nova tag sb:detalhe para cada linha do Excel
        detalhe_xml = f"""
            <sb:detalhe>
                <ns2:programacaoFinanceira xmlns:ns2="http://www.tesouro.gov.br/siafi/services/pf/manterProgramacaoFinanceira">
                    <codUgEmit>{(ug_emitente)}</codUgEmit>
                    <observacao>{(observacao)}</observacao>
                    <TRF>
                        <codUgFavorecida>{(ug_favorecida)}</codUgFavorecida>
                        <numeroDocumento>{(numero_documento)}</numeroDocumento>
                        <itemTRF>
                            <vlr>{(valor)}</vlr>
                            <codVinc>{(vinculacao_pagamento)}</codVinc>
                            <codFontRecur>{(fonte)}</codFontRecur>
                            <codCtgoGasto>{(categoria_gasto)}</codCtgoGasto>
                            <codSit>{(cod_situacao)}</codSit>"""
        
        # Condicional para adicionar a tag <txtInscrA> dependendo do valor de cod_situacao
        if cod_situacao not in ['TRF009', 'TRF007']:
            detalhe_xml += f"""
                            <txtInscrA>{(siafi)}</txtInscrA>"""
        
        detalhe_xml += f"""
                        </itemTRF> 
                    </TRF>    
                </ns2:programacaoFinanceira>
            </sb:detalhe>"""
        
        modelo_xml += detalhe_xml
        
    # Fecha o modelo XML e insere a quantidade de documentos
    modelo_xml += f"""
        </sb:detalhes>
        <sb:trailler>
            <sb:quantidadeDetalhe>{quantidade_documentos}</sb:quantidadeDetalhe>
        </sb:trailler>
    </sb:arquivo>"""

    print('Lote Gerado com sucesso!')

    with open(os.path.join(output_directory, "loteTranRecFin.xml"), "w", encoding="UTF-8") as xml_file:
        xml_file.write(modelo_xml)