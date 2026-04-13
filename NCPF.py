import pandas as pd
import csv
from datetime import datetime
import os
from openpyxl import load_workbook
import os  # Importar o módulo 'os' para manipular caminhos de arquivos
import re
import unicodedata

def extract_ted_number(observation):
    if observation is None:
        return None

    s = str(observation)
    s = unicodedata.normalize("NFKC", s).replace("\u00A0", " ")

    # casa:
    # "TED 15074"
    # "TED: 15074"
    # "TED Nº 15074,"
    # "TED N°15074"
    # "TED N º 15074;"
    pattern = r'\bTED\b\s*(?:[:\-]|N\s*[º°])?\s*(\d+)\s*(?=[\s,.;|)]|$)'
    m = re.search(pattern, s, flags=re.IGNORECASE)
    return m.group(1) if m else None

def process_nc_report(selected_dates, output_path):
    input_file_path = r"W:\B - TED\7 - AUTOMAÇÃO\NC e PF\NC funcionando - EXERCÍCIO 2026.xlsx"
    teds_file_path = r"W:\B - TED\7 - AUTOMAÇÃO\NC e PF\Teds da Administração Direta.xlsx"
    
    if not os.path.exists(input_file_path) or not os.path.exists(teds_file_path):
        print("Arquivo de NC Funcionando ou Teds da Administração Direta não encontrado na pasta.")
        return  # Retorna se o arquivo não for encontrado
    
    df = pd.read_excel(input_file_path, header=5)
    
    df['Emissão - Dia'] = pd.to_datetime(df['Emissão - Dia'], format='%d/%m/%Y')
    
    df_selecionado = df[df['Emissão - Dia'].dt.date.isin(selected_dates)]
    
        # tente achar o texto de observação em mais de uma coluna (fallback)
    OBS_COLS = ["Doc - Observação Texto", "Observação", "Histórico", "Descricao", "Descrição"]

    def pick_obs(row):
        for c in OBS_COLS:
            if c in row and pd.notna(row[c]) and str(row[c]).strip() != "":
                return row[c]
        return None

    df_selecionado["_OBS_SRC"] = df_selecionado.apply(pick_obs, axis=1)
    df_selecionado["TED"] = df_selecionado["_OBS_SRC"].apply(extract_ted_number)
    
    df_selecionado['TED'] = df_selecionado['Doc - Observação Texto'].apply(extract_ted_number)
    txt = "DEVOLUCAO ... REFERENTE AO TED Nº 15074, NOTA DE CREDITO ..."
    print(extract_ted_number(txt))
    
    df_sem_ted = df_selecionado[df_selecionado['TED'].isnull()]

    if not df_sem_ted.empty:
        teds_df = pd.read_excel(teds_file_path, header=None, engine='openpyxl')
        teds_df.columns = teds_df.iloc[0]  # Define a primeira linha como cabeçalho

        def fill_ted(row):
            siafi_value = str(row['NC - Transferência Código'])
            if 'NC - Transferência Código' in teds_df.columns:  # Verifica se a coluna 'SIAFI' existe no DataFrame
                matching_row = teds_df[teds_df['NC - Transferência Código'] == siafi_value]
                if not matching_row.empty:
                    ted_value = str(matching_row.iloc[0]['TED'])  # Convertendo para string
                    return ted_value
            return None

        df_sem_ted['TED'] = df_sem_ted.apply(fill_ted, axis=1)
        df_selecionado.update(df_sem_ted)
    
    df_dia_anterior_selecionado = df_selecionado.copy()

    df_dia_anterior_selecionado = df_dia_anterior_selecionado[
    ~df_dia_anterior_selecionado['RO - Evento Código']
        .astype(str)
        .str.contains(r'\b301206\b', regex=True, na=False)
]
    
    df_dia_anterior_selecionado['TED'] = df_dia_anterior_selecionado.apply(
        lambda row: row['Doc - Observação Texto'] if pd.isnull(row['TED']) else row['TED'],
        axis=1
    )
    
    colunas_selecionadas = [
        'Emissão - Dia Data Completa', 'Emitente - UG Nome', 'Favorecido Doc. Nome', 'RO - Evento Código',
        'NC - PTRES Código', 'NC Número', 'NC - Plano Interno Código PI', 'NC - Natureza Despesa Código',
        'NC - Transferência Código', 'NC - Valor Linha Valor', 'TED'
    ]
    
    df_dia_anterior_selecionado = df_dia_anterior_selecionado[colunas_selecionadas]
    
    df_dia_anterior_selecionado['Emissão - Dia Data Completa'] = df_dia_anterior_selecionado['Emissão - Dia Data Completa'].dt.strftime('%d/%m/%Y')
    
    df_dia_anterior_selecionado['Emissão - Dia Data Completa'] = df_dia_anterior_selecionado['Emissão - Dia Data Completa'].apply(
        lambda value: "{:,.2f}".format(
            float(
                str(value)
                .replace('.', '')   # remove milhar
                .replace(',', '.')  # troca decimal
            )
        ).replace(",", "_").replace(".", ",").replace("_", ".")
        if pd.notnull(value) and str(value).strip() != ''
        else None
    )
    
    # Verificar duplicatas de TED com diferentes NC Transferência e adicionar "DUPLICIDADE" se necessário
    ted_nc_mapping = df_dia_anterior_selecionado.groupby('TED')['NC - Transferência Código'].apply(set)
    df_dia_anterior_selecionado['NC - Transferência Código'] = df_dia_anterior_selecionado.apply(
        lambda row: 'DUPLICIDADE ' + row['NC - Transferência Código'] if len(ted_nc_mapping.get(row['TED'], [])) > 1 else row['NC - Transferência Código'],
        axis=1
    )

    # Verificar duplicatas de NC Transferência e adicionar "DUPLICIDADE SIAFI" se necessário
    siafi_nc_mapping = df_dia_anterior_selecionado.groupby('NC - Transferência Código')['TED'].apply(set)
    df_dia_anterior_selecionado['TED'] = df_dia_anterior_selecionado.apply(
        lambda row: 'DUPLICIDADE SIAFI ' + row['TED'] if row['NC - Transferência Código'] in siafi_nc_mapping and len(siafi_nc_mapping[row['NC - Transferência Código']]) > 1 and len(set(df_dia_anterior_selecionado[df_dia_anterior_selecionado['NC - Transferência Código'] == row['NC - Transferência Código']]['TED'])) > 1 else row['TED'],
        axis=1
    )

    # Adicionar lógica para gerar o nome do arquivo, se necessário
    output_file_name = selected_dates[0].strftime("%Y-%m-%d")

    if any(pd.notna(ted) and not (isinstance(ted, str) and ted.isdigit()) for ted in df_dia_anterior_selecionado['TED']):
        output_file_name += " VERIFICAR!"

    # Verificar duplicatas de TED com diferentes NC Transferência e adicionar "DUPLICIDADE" se necessário
    if any(pd.notna(ted) and not (isinstance(ted, float) or ted.isdigit()) for ted in df_dia_anterior_selecionado['TED']):
        output_file_name += " VERIFICAR!"

    # Verificar duplicatas de NC Transferência e adicionar "DUPLICIDADE SIAFI" se necessário
    if any(pd.notna(ted) and 'DUPLICIDADE SIAFI' in str(ted) for ted in df_dia_anterior_selecionado['TED']):
        output_file_name += " DUPLICIDADE SIAFI"

    output_file_path = os.path.join(output_path, f"NC {output_file_name}.csv")

    if df_dia_anterior_selecionado.empty:
        empty_df = pd.DataFrame(columns=colunas_selecionadas)  # Criar DataFrame vazio com as mesmas colunas
        output_file_path_nc_empty = os.path.join(output_path, f"NC {output_file_name} SEM DADOS.csv")
        empty_df.to_csv(output_file_path_nc_empty, index=False, sep=';', encoding='latin1', quoting=csv.QUOTE_NONNUMERIC)
        print(f"Arquivo NC sem dados gerado para a data {selected_dates[0].strftime('%Y-%m-%d')}")
    else:
        df_dia_anterior_selecionado.to_csv(output_file_path, index=False, sep=';', encoding='latin1', quoting=csv.QUOTE_NONNUMERIC)
        print("Arquivo NC salvo com sucesso em formato CSV!")

    return df_dia_anterior_selecionado  # Retornar o DataFrame para uso na função de PF

def process_pf_legado_report2(selected_dates, output_path):
    input_file_path_pf = r"W:\B - TED\7 - AUTOMAÇÃO\NC e PF\PF Legado - EXERCÍCIO 2026.xlsx"
    
    if not os.path.exists(input_file_path_pf):
        print("Arquivo de PF Legado não encontrado na pasta.")
        return  # Retorna se o arquivo não for encontrado
    
    df_pf_legado = pd.read_excel(input_file_path_pf, header=5)
    df_pf_legado['Emissão - Dia Data Completa'] = pd.to_datetime(df_pf_legado['Emissão - Dia Data Completa'], format='%d/%m/%Y')
    
    df_pf_legado_selecionado = df_pf_legado[df_pf_legado['Emissão - Dia Data Completa'].dt.date.isin(selected_dates)]
    df_pf_legado_selecionado = df_pf_legado_selecionado[df_pf_legado_selecionado['Emitente - UG Código'].astype(str) != "152734"]
    
    colunas_selecionadas_pf_legado = [
        'Emissão - Dia Data Completa', 'PF Número', 'Emitente - UG Nome', 'Emitente - Gestão Código',
        'Favorecido Doc. Número', 'PF - Evento Código', 'PF - Categoria Gasto Código',
        'PF - Fonte Recursos Código', 'PF - Vinculação Pagamento Código', 'PF - Inscrição Código',
        'PF - Valor Linha Valor'
    ]
    df_pf_legado_selecionado = df_pf_legado_selecionado[colunas_selecionadas_pf_legado]
    
    df_pf_legado_selecionado['Emitente - Gestão Nome'] = df_pf_legado_selecionado.iloc[:, 5]
    df_pf_legado_selecionado['Favorecido Doc. Nome'] = df_pf_legado_selecionado.iloc[:, 7]
    
    colunas_reordenadas = [
        'Emissão - Dia Data Completa', 'PF Número', 'Emitente - UG Nome', 'Emitente - Gestão Código',
        'Emitente - Gestão Nome', 'Favorecido Doc. Número', 'Favorecido Doc. Nome', 'PF - Evento Código',
        'PF - Categoria Gasto Código', 'PF - Fonte Recursos Código', 'PF - Vinculação Pagamento Código', 'PF - Inscrição Código', 'PF - Valor Linha Valor'
    ]
    df_pf_legado_selecionado = df_pf_legado_selecionado[colunas_reordenadas]
    
    df_pf_legado_selecionado['Emissão - Dia Data Completa'] = df_pf_legado_selecionado['Emissão - Dia Data Completa'].dt.strftime('%d/%m/%Y')
    
    # Caminho para o arquivo de saída da planilha PF Legado
    nome_arquivo_saida_pf_legado = selected_dates[0].strftime("%Y-%m-%d")
    output_file_path_pf_legado = os.path.join(output_path, f"PF {nome_arquivo_saida_pf_legado}.csv")  # Usa o caminho escolhido pelo usuário

    # Verifica se o DataFrame está vazio
    if df_pf_legado_selecionado.empty:
        # Criar DataFrame vazio com as colunas selecionadas
        empty_df = pd.DataFrame(columns=colunas_reordenadas)
        output_file_path_pf_legado_empty = os.path.join(output_path, f"PF {nome_arquivo_saida_pf_legado} SEM DADOS.csv")
        # Salva o DataFrame vazio em formato CSV separado por vírgula
        empty_df.to_csv(output_file_path_pf_legado_empty, index=False, sep=';', encoding='utf-8-sig', quoting=csv.QUOTE_NONNUMERIC)
        print(f"Arquivo PF Legado sem dados gerado para a data {selected_dates[0].strftime('%Y-%m-%d')}")
    else:
        # Salvar o DataFrame selecionado da planilha PF Legado em formato CSV separado por vírgula
        df_pf_legado_selecionado.to_csv(output_file_path_pf_legado, index=False, sep=';', encoding='utf-8-sig', quoting=csv.QUOTE_NONNUMERIC)
        print("Arquivo PF Legado salvo com sucesso em formato CSV!")