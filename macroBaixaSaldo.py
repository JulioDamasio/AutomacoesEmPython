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
    df['OBSERVAÇÃO'] = ''
    df['COD LANÇAMENTO 001'] = ''
    df['COD LANÇAMENTO 013'] = ''
    df.to_excel(caminho_arquivo_copia, index=False)

def preencher_colunas(caminho_arquivo_copia):
    try:
        # Carrega o arquivo Excel com o Pandas
        df = pd.read_excel(caminho_arquivo_copia)
        
        # Preenche a coluna 'observação' com o texto desejado
        df['OBSERVAÇÃO'] = 'Baixa no saldo contabil para finalizacao na prestacao de contas'
        df['COD LANÇAMENTO 001'] = '001'
        df['COD LANÇAMENTO 013'] = '013'
        
        # Salva as modificações de volta no arquivo
        df.to_excel(caminho_arquivo_copia, index=False)
        
        print("Coluna de observação preenchida com sucesso.")
        
    except Exception as e:
        print(f"Erro ao preencher a coluna de observação: {e}")

def formatar_valores_arquivo(caminho_arquivo, colunas_de_valores):
    try:
        # Carrega o arquivo Excel com o Pandas
        df = pd.read_excel(caminho_arquivo)

        for coluna in colunas_de_valores:
            if coluna in df.columns:
                # Converte para string
                df[coluna] = df[coluna].astype(str)

                # Remove "R$", espaços, pontos e vírgulas
                df[coluna] = (
                    df[coluna]
                    .str.replace(r'R\$', '', regex=True)
                    .str.replace('.', '', regex=False)
                    .str.replace(',', '', regex=False)
                    .str.strip()
                )

                # Se quiser garantir que só fiquem dígitos:
                df[coluna] = df[coluna].str.replace(r'[^0-9]', '', regex=True)

        # Salva as modificações de volta no arquivo
        df.to_excel(caminho_arquivo, index=False)

        print("Valores formatados com sucesso (somente dígitos, sem pontuação).")

    except Exception as e:
        print(f"Erro ao formatar os valores do arquivo: {e}")

def remover_pontos_virgulas(caminho_arquivo, colunas_de_valores):
    try:
        # Carrega o arquivo Excel com o Pandas
        df = pd.read_excel(caminho_arquivo)
        
        for coluna in colunas_de_valores:
            if coluna in df.columns:
                # Remove pontos e vírgulas da coluna
                df[coluna] = df[coluna].replace('[,.]', '', regex=True)
        
        # Salva as modificações de volta no arquivo
        df.to_excel(caminho_arquivo, index=False)
        
        print("Pontos e vírgulas removidos com sucesso.")
        
    except Exception as e:
        print(f"Erro ao remover pontos e vírgulas: {e}")
        
def filtrar_teds_aptos_para_comprovar_001(df):
    # Filtrar linhas com 'ok' na coluna 'TEDS APTOS PARA COMPROVAR'
    df_filtrado = df[df['TEDS APTOS PARA COMPROVAR'].str.lower() == 'ok']
    # Filtrar linhas onde 'A COMPROVAR' é diferente de 0 ou vazio
    df_resultado = df_filtrado[df_filtrado['A COMPROVAR'].notnull() & (df_filtrado['A COMPROVAR'] != 0)]
    return df_resultado

def filtrar_teds_aptos_para_comprovar_013(df):
    # Filtrar linhas com 'ok' na coluna 'TEDS APTOS PARA COMPROVAR'
    df_filtrado = df[df['TEDS APTOS PARA COMPROVAR'].str.lower() == 'ok']
    # Filtrar linhas onde 'A COMPROVAR' é diferente de 0 ou vazio
    df_resultado = df_filtrado[df_filtrado['A REPASSAR'].notnull() & (df_filtrado['A REPASSAR'] != 0)]
    return df_resultado

def generate_macro_001(output_directory, df):
    tela_inicial = 2  # Inicializa a variável com o número da primeira tela
    
    modelo_xml = f"""<HAScript name="Baixa de Saldos à Comprovar" description="" timeout="60000" pausetime="300" promptall="true" blockinput="false" author="AugustoCezar" creationdate="05/08/2024 16:58:56" supressclearevents="false" usevars="false" ignorepauseforenhancedtn="true" delayifnotenhancedtn="0" ignorepausetimeforenhancedtn="true">


    <screen name="Tela1" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
            <numfields number="107" optional="false" invertmatch="false" />
            <numinputfields number="1" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="&gt;exectransf[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Tela2" />
        </nextscreens>
    </screen>
    """

    tela = tela_inicial  # Inicializa a variável com o número da primeira tela
    
    for index, row in df.iterrows():
        siafi = str(row['SIAFI'])
        a_comprovar = str(row['A COMPROVAR'])
        observacao = str(row['OBSERVAÇÃO'])
    

        modelo_xml += f"""
    <screen name="Tela{tela}" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
            <numfields number="38" optional="false" invertmatch="false" />
            <numinputfields number="1" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="{siafi}[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Tela{tela + 1}" />
        </nextscreens>
    </screen>

    <screen name="Tela{tela + 1}" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
            <numfields number="81" optional="false" invertmatch="false" />
            <numinputfields number="6" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="001[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Tela{tela + 2}" />
        </nextscreens>
    </screen>

    <screen name="Tela{tela + 2}" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
            <numfields number="82" optional="false" invertmatch="false" />
            <numinputfields number="6" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="{a_comprovar}[tab]{observacao}[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Tela{tela + 3}" />
        </nextscreens>
    </screen>

    <screen name="Tela{tela + 3}" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
            <numfields number="93" optional="false" invertmatch="false" />
            <numinputfields number="1" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="s[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Tela{tela + 4}" />
        </nextscreens>
    </screen>

    <screen name="Tela{tela + 4}" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
            <numfields number="76" optional="false" invertmatch="false" />
            <numinputfields number="0" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Tela{tela + 5}" />
        </nextscreens>
    </screen>
    """

        tela += 5  # Atualiza a variável para o próximo número de tela conforme o seu modelo

    # fecha o modelo
    modelo_xml += """    
</HAScript> 
        """

    # Salva o arquivo
    output_filename = os.path.join(output_directory, "Macro_FinalizarTED_001.MAC")
    with open(output_filename, "w") as file:
        file.write(modelo_xml)

    print(f"Macro gerada com sucesso. Total de linhas processadas: {len(df)}")

def generate_macro_013(output_directory, df):
    tela_inicial = 2  # Inicializa a variável com o número da primeira tela
    
    modelo_xml = f"""<HAScript name="Baixa de Saldos à Comprovar" description="" timeout="60000" pausetime="300" promptall="true" blockinput="false" author="AugustoCezar" creationdate="05/08/2024 16:58:56" supressclearevents="false" usevars="false" ignorepauseforenhancedtn="true" delayifnotenhancedtn="0" ignorepausetimeforenhancedtn="true">

    <screen name="Tela1" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
            <numfields number="107" optional="false" invertmatch="false" />
            <numinputfields number="1" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="&gt;exectransf[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Tela2" />
        </nextscreens>
    </screen>
    """

    tela = 2  # Inicializa a variável com o número da primeira tela
    
    for index, row in df.iterrows():
        siafi = str(row['SIAFI'])
        a_repassar = str(row['A REPASSAR'])
        observacao = str(row['OBSERVAÇÃO'])

        modelo_xml += f"""
    <screen name="Tela{tela}" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
            <numfields number="38" optional="false" invertmatch="false" />
            <numinputfields number="1" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="{siafi}[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Tela{tela + 1}" />
        </nextscreens>
    </screen>

    <screen name="Tela{tela + 1}" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
            <numfields number="81" optional="false" invertmatch="false" />
            <numinputfields number="6" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="013[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Tela{tela + 2}" />
        </nextscreens>
    </screen>

    <screen name="Tela{tela + 2}" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
            <numfields number="82" optional="false" invertmatch="false" />
            <numinputfields number="6" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="{a_repassar}[tab]{observacao}[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Tela{tela + 3}" />
        </nextscreens>
    </screen>

    <screen name="Tela{tela + 3}" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
            <numfields number="93" optional="false" invertmatch="false" />
            <numinputfields number="1" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="s[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Tela{tela + 4}" />
        </nextscreens>
    </screen>

    <screen name="Tela{tela + 4}" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
            <numfields number="76" optional="false" invertmatch="false" />
            <numinputfields number="0" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Tela{tela + 5}" />
        </nextscreens>
    </screen>
    """

        tela += 5  # Atualiza a variável para o próximo número de tela conforme o seu modelo

    # fecha o modelo
    modelo_xml += """    
</HAScript> 
        """

    # Salva o arquivo
    output_filename = os.path.join(output_directory, "Macro_FinalizarTED_013.MAC")
    with open(output_filename, "w") as file:
        file.write(modelo_xml)

    print(f"Macro gerada com sucesso. Total de linhas processadas: {len(df)}")
  
def main():    
    
    teds_para_finalizar = r'W:\B - TED\7 - AUTOMAÇÃO\Teds para finalizar\COPIA TEDS para Finalizar.xlsx'
    teds_para_finalizar_macro = r'W:\B - TED\7 - AUTOMAÇÃO\Teds para finalizar\MACRO TEDS para Finalizar.xlsx'
    output_directory = r'W:\B - TED\7 - AUTOMAÇÃO\Teds para finalizar'
    
    colunas_de_valores = [
        'VALOR ORÇAMENTÁRIO SIMEC', 'VALOR NC SIMEC', 'VALOR NC SIAFI', 'VALOR PF SIMEC', 
        'VALOR PF SIAFI', 'NC - PF SIMEC', 'NC - PF SIAFI', 'ORÇAMENTÁRIO - PF SIMEC', 
        'VALORES FIRMADOS', 'A REPASSAR', 'A COMPROVAR', 'COMPROVADO', 'NÃO REPASSADO/ DEVOLVIDO'
    ]
    
    copiar_arquivo(teds_para_finalizar, teds_para_finalizar_macro)
    adicionar_colunas_excel(teds_para_finalizar_macro)
    preencher_colunas(teds_para_finalizar_macro)
    formatar_valores_arquivo(teds_para_finalizar_macro, colunas_de_valores)
    remover_pontos_virgulas(teds_para_finalizar_macro, colunas_de_valores)
    
    # Leitura e filtragem dos TEDS
    df_teds = pd.read_excel(teds_para_finalizar_macro)
    df_filtrado_001 = filtrar_teds_aptos_para_comprovar_001(df_teds)
    df_filtrado_013 = filtrar_teds_aptos_para_comprovar_013(df_teds)
    
    generate_macro_001(output_directory, df_filtrado_001)
    generate_macro_013(output_directory, df_filtrado_013)

if __name__ == "__main__":
    main()   