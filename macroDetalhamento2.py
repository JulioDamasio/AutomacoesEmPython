from openpyxl import load_workbook
from openpyxl.styles import Alignment  # Importe a classe Alignment
import shutil
import os
import pandas as pd
from datetime import datetime

def copiar_arquivo(origem, destino):
    try:
        shutil.copy(origem, destino)
        print(f"Arquivo copiado de {origem} para {destino} mantendo a formatação.")
    except Exception as e:
        print(f"Erro ao copiar o arquivo: {e}")
        
def adicionar_colunas_excel(caminho_arquivo_copia):
    df = pd.read_excel(caminho_arquivo_copia)
    df['Fonte 4 Digitos De'] = ''
    df['Fonte 6 Digitos De'] = ''
    df['Fonte 4 Digitos Para'] = ''
    df['Fonte 6 Digitos Para'] = ''
    df['Natureza Despesa 2 Digitos De'] = ''
    df['Natureza Despesa 4 Digitos De'] = ''
    df['Natureza Despesa 2 Digitos Para'] = ''
    df['Natureza Despesa 2 Digitos Para'] = ''
    df['Observação'] = 'RETIRADA DE DETALHAMENTO'
    
    # Adicionar a data do dia na nova coluna
    data_hoje = datetime.now().strftime('%d/%m/%Y')  # Formato de data brasileiro (DD/MM/AAAA)
    df['Data do dia'] = data_hoje  # Preencher com a data atual
    df.to_excel(caminho_arquivo_copia, index=False)

def preencher_colunas(caminho_arquivo_copia):
    try:
        # Carregar o arquivo Excel como DataFrame
        df = pd.read_excel(caminho_arquivo_copia)
        
        # Criar as colunas, se não existirem
        if 'Fonte 4 Digitos De' not in df.columns:
            df['Fonte 4 Digitos De'] = ''
        if 'Fonte 6 Digitos De' not in df.columns:
            df['Fonte 6 Digitos De'] = ''
        if 'Fonte 4 Digitos Para' not in df.columns:
            df['Fonte 4 Digitos Para'] = ''
        if 'Fonte 6 Digitos Para' not in df.columns:
            df['Fonte 6 Digitos Para'] = ''    
        if 'Natureza Despesa 2 Digitos De' not in df.columns:
            df['Natureza Despesa 2 Digitos De'] = ''
        if 'Natureza Despesa 4 Digitos De' not in df.columns:
            df['Natureza Despesa 4 Digitos De'] = ''
        if 'Natureza Despesa 2 Digitos Para' not in df.columns:
            df['Natureza Despesa 2 Digitos Para'] = ''
        if 'Natureza Despesa 4 Digitos Para' not in df.columns:
            df['Natureza Despesa 4 Digitos Para'] = ''
        if 'Natureza Despesa Centro' not in df.columns:
            df['Natureza Despesa Centro'] = ''        
        
        # Preencher as colunas com os valores extraídos
        df['Fonte 4 Digitos De'] = df['Fonte De'].astype(str).str[:4]  # 4 primeiros dígitos
        df['Fonte 6 Digitos De'] = df['Fonte De'].astype(str).str[-6:]  # 6 últimos dígitos
        df['Fonte 4 Digitos Para'] = df['Fonte Para'].astype(str).str[:4]  # 4 primeiros dígitos
        df['Fonte 6 Digitos Para'] = df['Fonte Para'].astype(str).str[-6:]  # 6 últimos dígitos
        df['Natureza Despesa 2 Digitos De'] = df['Natureza Despesa De'].astype(str).str[:2]  # 2 primeiros dígitos
        df['Natureza Despesa 4 Digitos De'] = df['Natureza Despesa De'].astype(str).str[-4:]
        df['Natureza Despesa 2 Digitos Para'] = df['Natureza Despesa Para'].astype(str).str[:2]  # 2 primeiros dígitos
        df['Natureza Despesa 4 Digitos Para'] = df['Natureza Despesa Para'].astype(str).str[-4:]
        df['Natureza Despesa Centro'] = df['Natureza Despesa De'].astype(str).str[2:4]

        # Salvar o arquivo de volta
        df.to_excel(caminho_arquivo_copia, index=False)
        print("Colunas preenchidas com sucesso!")
    
    except Exception as e:
        print(f"Erro ao preencher as colunas: {e}")
    
    
def formatar_data_do_dia(caminho_arquivo_copia):
    # Dicionário para traduzir os meses de inglês para português
    meses_pt = {
        'jan': 'Jan',
        'feb': 'Fev',
        'mar': 'Mar',
        'apr': 'Abr',
        'may': 'Mai',
        'jun': 'Jun',
        'jul': 'Jul',
        'aug': 'Ago',
        'sep': 'Set',
        'oct': 'Out',
        'nov': 'Nov',
        'dec': 'Dez'
    }
    
    # Obter a data atual no formato desejado
    data_do_dia = datetime.now().strftime('%d%b%Y').lower()
    
    # Traduzir a data atual para o formato em português
    dia = data_do_dia[:2]
    mes = meses_pt[data_do_dia[2:5]]
    ano = data_do_dia[5:][-2:]
    data_formatada = f"{dia}{mes}{ano}"
    
    def formatar_data(data):
        formatos = ['%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d/%m/%Y']
        for fmt in formatos:
            try:
                dt = datetime.strptime(data, fmt)
                return dt.strftime('%d') + meses_pt[dt.strftime('%b').lower()] + dt.strftime('%Y').lower()
            except ValueError:
                pass
        raise ValueError(f"Formato de data desconhecido: {data}")
    
    try:
        # Carrega o arquivo Excel com o Pandas
        df = pd.read_excel(caminho_arquivo_copia)
        
        # Adiciona a data formatada ao DataFrame
        df['Data do dia'] = data_formatada
        
        # Salva as modificações de volta no arquivo
        df.to_excel(caminho_arquivo_copia, index=False)
        
        print("Data de vigência formatada com sucesso.")
        
    except Exception as e:
        print(f"Erro ao formatar a data de vigência: {e}")  

def formatar_valores_arquivo(caminho_arquivo_copia):
    try:
        # Carregar o arquivo Excel
        df = pd.read_excel(caminho_arquivo_copia)
        
        # Verificar se a coluna 'Valor' existe
        if 'Valor' in df.columns:
            # Manter os centavos e remover apenas o ponto decimal
            df['Valor'] = df['Valor'].astype(float).apply(lambda x: f"{x:.2f}".replace('.', ''))
        
        # Salvar as alterações no arquivo
        df.to_excel(caminho_arquivo_copia, index=False)
        print("Valores formatados com sucesso, pontos removidos.")
    
    except Exception as e:
        print(f"Erro ao formatar os valores do arquivo: {e}")
        

def generate_macro_vigencia(output_directory, df):
    modelo_xml = f"""<HAScript name="detaorc" description="" timeout="60000" pausetime="300" promptall="true" blockinput="false" author="juliodamasio" creationdate="09/12/2024 17:35:29" supressclearevents="false" usevars="false" ignorepauseforenhancedtn="true" delayifnotenhancedtn="0" ignorepausetimeforenhancedtn="true">
    
    <screen name="Tela1" entryscreen="true" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="&gt;detaorc[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Tela2" />
        </nextscreens>
    </screen>
    """
    
    tela = 2  # Inicializa a variável com o número da primeira tela
    
    for index, row in df.iterrows():
        # Extrai os valores de cada linha
        esfera = str(row['Esfera'])
        ptres = str(row['PTRES'])
        fonte = str(row['Fonte De'])
        nd2 = str(row['Natureza Despesa 2 Digitos De'])
        nd4 = str(row['Natureza Despesa 4 Digitos De'])
        nd5 = str(row['Natureza Despesa 2 Digitos Para'])
        nd6 = str(row['Natureza Despesa 4 Digitos Para'])
        nd3 = str(row['Natureza Despesa Centro'])
        valor = str(row['Valor'])
        fontequatrode = str(row['Fonte 4 Digitos De'])
        fonteseisde = str(row['Fonte 6 Digitos De']).split('.')[0].strip()
        fonteseisde = fonteseisde.zfill(6)
        fontequatropara = str(row['Fonte 4 Digitos Para'])
        fonteseispara = str(row['Fonte 6 Digitos Para'])
        observacao = str(row['Observação'])
        data = str(row['Data do dia'])
        
        # Adiciona o conteúdo gerado ao modelo
        modelo_xml += f"""
    <screen name="Tela{tela}" entryscreen="true" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="{data}1[tab]1527341[tab]{esfera}{ptres}{fontequatrode}{nd2}1[tab]{data}9999[tab]{observacao}[tab][tab][tab]R{fonteseisde}{nd4}[tab][tab][tab]{valor}[tab]A{fonteseispara}{nd6}[tab][tab][tab]{valor}[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Tela{tela + 1}" />
        </nextscreens>
    </screen>

    <screen name="Tela{tela + 1}" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
            <numfields number="242" optional="false" invertmatch="false" />
            <numinputfields number="1" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="s[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Tela{tela + 2}" />
        </nextscreens>
    </screen>

    <screen name="Tela{tela + 2}" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
            <numfields number="64" optional="false" invertmatch="false" />
            <numinputfields number="0" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Tela{tela + 3}" />
        </nextscreens>
    </screen>
        """
        # Atualiza o número da tela para a próxima iteração
        tela += 3

    # Finaliza o modelo
    modelo_xml += """
</HAScript>
"""

    # Salva o arquivo
    output_filename = os.path.join(output_directory, "Retirada de detalhamento 2.MAC")
    with open(output_filename, "w") as file:
        file.write(modelo_xml)

    print(f"Nova macro de vigência gerada com sucesso. O arquivo está no caminho: {output_filename}")

def main():
    caminho_arquivo_original = r'W:\B - TED\7 - AUTOMAÇÃO\Detalhamento de credito\DETAORC dados 2.xlsx'
    caminho_arquivo_copia = r'W:\B - TED\7 - AUTOMAÇÃO\Detalhamento de credito\COPIA DETAORC dados 2.xlsx'
    
    copiar_arquivo(caminho_arquivo_original, caminho_arquivo_copia)
    adicionar_colunas_excel(caminho_arquivo_copia)
    
    # Carrega o arquivo Excel para um DataFrame
    df = pd.read_excel(caminho_arquivo_copia)
    
    preencher_colunas(caminho_arquivo_copia)
    # Salva o DataFrame modificado de volta no arquivo Excel
    formatar_data_do_dia(caminho_arquivo_copia)
    formatar_valores_arquivo(caminho_arquivo_copia)
    
    # Carrega o arquivo Excel novamente para obter o DataFrame atualizado
    df = pd.read_excel(caminho_arquivo_copia)

    # Chama a função generate_macro_vigencia com o argumento df
    generate_macro_vigencia(r'W:\B - TED\7 - AUTOMAÇÃO\Detalhamento de credito', df)
    
if __name__ == "__main__":
    main()
               