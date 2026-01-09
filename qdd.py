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

def remover_ultima_linha_excel(copia_caminho_arquivo_TERMO):
    try:
        # Carregar o arquivo Excel
        df = pd.read_excel(copia_caminho_arquivo_TERMO)

        # Remover a última linha
        df = df.iloc[:-1]

        # Salvar de volta no mesmo arquivo
        df.to_excel(copia_caminho_arquivo_TERMO, index=False)
        print("Última linha removida com sucesso.")

    except Exception as e:
        print(f"Erro ao remover a última linha: {e}")        

def remover_12_primeiras_linhas_excel(caminho_arquivo):
    try:
        # Carrega o Excel pulando as 12 primeiras linhas
        df = pd.read_excel(caminho_arquivo, skiprows=12)

        # Salva o DataFrame no mesmo arquivo (sobrescreve)
        df.to_excel(caminho_arquivo, index=False)
        print("As 12 primeiras linhas foram removidas com sucesso.")

    except Exception as e:
        print(f"Erro ao remover as 12 primeiras linhas: {e}")

def remover_segunda_linha_excel(caminho_arquivo):
    try:
        # Lê todo o arquivo normalmente
        df = pd.read_excel(caminho_arquivo)

        # Remove a segunda linha (index 1)
        df = df.drop(index=1).reset_index(drop=True)

        # Salva de volta
        df.to_excel(caminho_arquivo, index=False)
        print("✅ Segunda linha removida com sucesso.")

    except Exception as e:
        print(f"❌ Erro ao remover a segunda linha: {e}")
        
def adicionar_colunas_excel(copia_caminho_arquivo_TERMO):
    df = pd.read_excel(copia_caminho_arquivo_TERMO)
    df['Esfera'] = ''
    df['Fonte 4 Digitos'] = ''
    df['Fonte 6 Digitos'] = ''
    df['Natureza Despesa 2 Digitos'] = ''
    df['Natureza Despesa 4 Digitos'] = ''
    df['Natureza Despesa Centro'] = ''
    df['Observação'] = 'Quadro de detalhamento de despesa'
    
    # Adicionar a data do dia na nova coluna
    data_hoje = datetime.now().strftime('%d/%m/%Y')  # Formato de data brasileiro (DD/MM/AAAA)
    df['Data do dia'] = data_hoje  # Preencher com a data atual
    df.to_excel(copia_caminho_arquivo_TERMO, index=False)
    

def preencher_esfera_por_ptres_linha_a_linha(
    caminho_termo=r'W:\B - TED\7 - AUTOMAÇÃO\QDD para descentralização\COPIA Termo Aprovado aguardando descentralização.xlsx',
    caminho_credito=r'W:\B - TED\7 - AUTOMAÇÃO\QDD para descentralização\COPIA Crédito Disponivel Geral com Esfera.xlsx'
):
    try:
        # Carregar arquivos
        df_termo   = pd.read_excel(caminho_termo)
        df_credito = pd.read_excel(caminho_credito)

        # Padronizar nomes de colunas (remover espaços extras)
        df_termo.columns   = df_termo.columns.str.strip()
        df_credito.columns = df_credito.columns.str.strip()

        # Conferir colunas necessárias
        col_termo_ptres = 'PTRES (Orçamentário)'
        col_credito_ptres = 'PTRES'
        col_credito_esfera = 'Esfera Orçamentária'
        col_destino_esfera = 'Esfera'

        for col, df_name in [(col_termo_ptres, 'TERMO'),
                             (col_credito_ptres, 'CRÉDITO'),
                             (col_credito_esfera, 'CRÉDITO')]:
            if col not in (df_termo.columns if df_name=='TERMO' else df_credito.columns):
                raise ValueError(f"Coluna '{col}' não encontrada no arquivo {df_name}.")

        # Remover linhas totalmente vazias no TERMO (inclui a 2ª linha em branco pós-desmesclagem)
        df_termo = df_termo[~df_termo.isna().all(axis=1)].reset_index(drop=True)

        # --- Normalizador de chave PTRES (robusto a float, '.0', vírgulas, etc.)
        def norm_ptres(v):
            if pd.isna(v):
                return None
            if isinstance(v, (int, float)):
                try:
                    return str(int(v))
                except Exception:
                    return None
            s = str(v).strip()
            if s == '' or s.lower() == 'nan':
                return None
            # tentar conversão numérica considerando formatações
            try:
                # se tiver vírgula e ponto: assume ponto como milhar e vírgula como decimal (BR)
                if ',' in s and '.' in s:
                    s2 = s.replace('.', '').replace(',', '.')
                    return str(int(float(s2)))
                # se só vírgula: vírgula como decimal
                if ',' in s:
                    s2 = s.replace(',', '.')
                    return str(int(float(s2)))
                # caso geral: ponto decimal ou inteiro puro
                return str(int(float(s)))
            except Exception:
                # fallback: pegar só dígitos
                digits = ''.join(ch for ch in s if ch.isdigit())
                return digits if digits else None

        # Criar coluna chave normalizada
        df_termo['_PTRES_KEY_']   = df_termo[col_termo_ptres].apply(norm_ptres)
        df_credito['_PTRES_KEY_'] = df_credito[col_credito_ptres].apply(norm_ptres)

        # Mostrar amostras para debug
        print("Amostra PTRES TERMO (normalizado):", df_termo['_PTRES_KEY_'].dropna().unique()[:10])
        print("Amostra PTRES CRÉDITO (normalizado):", df_credito['_PTRES_KEY_'].dropna().unique()[:10])

        # Criar índice rápido (dicionário) de PTRES -> Esfera Orçamentária
        mapa = {}
        for _, r in df_credito.dropna(subset=['_PTRES_KEY_']).iterrows():
            k = r['_PTRES_KEY_']
            if k not in mapa:  # mantém a primeira ocorrência
                mapa[k] = r[col_credito_esfera]

        # Garantir coluna destino
        if col_destino_esfera not in df_termo.columns:
            df_termo[col_destino_esfera] = pd.NA

        # Preencher linha a linha com prints de depuração limitados
        limite_prints = 12
        faltas = 0
        acertos = 0

        for i, row in df_termo.iterrows():
            chave = row['_PTRES_KEY_']
            if not chave:
                if faltas < limite_prints:
                    print(f"[{i}] PTRES vazio/NaN no TERMO → Esfera=0")
                df_termo.at[i, col_destino_esfera] = 0
                faltas += 1
                continue

            valor_esfera = mapa.get(chave, None)
            if pd.isna(valor_esfera) or valor_esfera is None:
                if faltas < limite_prints:
                    print(f"[{i}] PTRES {chave} não encontrado no CRÉDITO → Esfera=0")
                df_termo.at[i, col_destino_esfera] = 0
                faltas += 1
            else:
                df_termo.at[i, col_destino_esfera] = valor_esfera
                if acertos < limite_prints:
                    print(f"[{i}] PTRES {chave} encontrado → Esfera={valor_esfera}")
                acertos += 1

        print(f"\nResumo: {acertos} correspondências | {faltas} não encontradas (mostradas até {limite_prints}).")

        # Limpar colunas auxiliares e salvar
        df_termo = df_termo.drop(columns=['_PTRES_KEY_'])
        df_termo.to_excel(caminho_termo, index=False)
        print("✅ Coluna 'Esfera' preenchida e arquivo salvo com sucesso.")

    except Exception as e:
        print(f"❌ Erro ao preencher a coluna 'Esfera': {e}")

def preencher_colunas(copia_caminho_arquivo_TERMO):
    try:
        # Carregar o arquivo Excel como DataFrame
        df = pd.read_excel(copia_caminho_arquivo_TERMO)
        
        # Criar as colunas, se não existirem
        if 'Fonte 4 Digitos' not in df.columns:
            df['Fonte 4 Digitos'] = ''
        if 'Fonte 6 Digitos' not in df.columns:
            df['Fonte 6 Digitos'] = ''
        if 'Natureza Despesa 2 Digitos' not in df.columns:
            df['Natureza Despesa 2 Digitos'] = ''
        if 'Natureza Despesa 4 Digitos' not in df.columns:
            df['Natureza Despesa 4 Digitos'] = ''
        if 'Natureza Despesa Centro' not in df.columns:
            df['Natureza Despesa Centro'] = ''        
        
        # Preencher as colunas com os valores extraídos
        df['Fonte 4 Digitos'] = df['Fonte (Orçamentário)'].astype(str).str[:4]  # 4 primeiros dígitos
        df['Fonte 6 Digitos'] = df['Fonte (Orçamentário)'].astype(str).str[-6:]  # 6 últimos dígitos
        df['Natureza Despesa 2 Digitos'] = df['Natureza da Despesa'].astype(str).str[:2]  # 2 primeiros dígitos
        df['Natureza Despesa 4 Digitos'] = df['Natureza da Despesa'].astype(str).str[2:6]
        df['Natureza Despesa Centro'] = df['Natureza da Despesa'].astype(str).str[2:4]

        # Salvar o arquivo de volta
        df.to_excel(copia_caminho_arquivo_TERMO, index=False)
        print("Colunas preenchidas com sucesso!")
    
    except Exception as e:
        print(f"Erro ao preencher as colunas: {e}")
        
def remover_valores_zerados(caminho_arquivo):
    try:
        # Lê o arquivo
        df = pd.read_excel(caminho_arquivo)

        # Identifica a última coluna
        ultima_coluna = df.columns[-1]

        # Converte valores da última coluna para número (tratando "0,00", "0.00", etc.)
        df[ultima_coluna] = (
            df[ultima_coluna]
            .astype(str)
            .str.replace(',', '.', regex=False)
            .astype(float)
        )

        # Remove linhas onde a última coluna é zero
        df = df[df[ultima_coluna] != 0].reset_index(drop=True)

        # Salva de volta
        df.to_excel(caminho_arquivo, index=False)
        print("✅ Linhas com valor zero removidas com sucesso.")

    except Exception as e:
        print(f"❌ Erro ao remover linhas com valor zero: {e}")        
        
def formatar_data_do_dia(copia_caminho_arquivo_TERMO):
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
        df = pd.read_excel(copia_caminho_arquivo_TERMO)
        
        # Adiciona a data formatada ao DataFrame
        df['Data do dia'] = data_formatada
        
        # Salva as modificações de volta no arquivo
        df.to_excel(copia_caminho_arquivo_TERMO, index=False)
        
        print("Data de vigência formatada com sucesso.")
        
    except Exception as e:
        print(f"Erro ao formatar a data de vigência: {e}")  

def formatar_valores_arquivo(copia_caminho_arquivo_TERMO):
    try:
        # Carregar o arquivo Excel
        df = pd.read_excel(copia_caminho_arquivo_TERMO)
        
        # Verificar se a coluna 'Valor' existe
        if 'Valor' in df.columns:
            # Manter os centavos e remover apenas o ponto decimal
            df['Valor'] = df['Valor'].astype(float).apply(lambda x: f"{x:.2f}".replace('.', ''))
        
        # Salvar as alterações no arquivo
        df.to_excel(copia_caminho_arquivo_TERMO, index=False)
        print("Valores formatados com sucesso, pontos removidos.")
    
    except Exception as e:
        print(f"Erro ao formatar os valores do arquivo: {e}")
        

def generate_macro_vigencia(output_directory, df):
    # Garantir que a coluna de valores esteja como float arredondado
    df['Valor Autorizado (R$)'] = df['Valor Autorizado (R$)'].astype(float).round(2)

    modelo_xml = f"""<HAScript name="detaorc" description="" timeout="60000" pausetime="300" promptall="true" blockinput="false" author="juliodamasio" creationdate="09/12/2024 17:35:29" supressclearevents="false" usevars="false" ignorepauseforenhancedtn="true" delayifnotenhancedtn="0" ignorepausetimeforenhancedtn="true">
    
    <screen name="Tela1" entryscreen="true" exitscreen="false" transient="false">
        <description>
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="&gt;detaorc[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0">
            <nextscreen name="Tela2" />
        </nextscreens>
    </screen>
    """

    tela = 2  # Número da primeira tela

    for index, row in df.iterrows():
        # Extrair e preparar os valores
        esfera = str(row['Esfera']).split('.')[0].zfill(1)
        ptres = str(row['PTRES (Orçamentário)']).split('.')[0].zfill(6)
        fonte = str(row['Fonte (Orçamentário)'])
        pi = str(row['PI (Orçamentário)'])
        nd2 = str(row['Natureza Despesa 2 Digitos']).split('.')[0].zfill(2)
        nd4 = str(row['Natureza Despesa 4 Digitos']).split('.')[0].zfill(4)
        nd3 = str(row['Natureza Despesa Centro']).split('.')[0]
        valor = float(row['Valor Autorizado (R$)'])
        valor_formatado = f"{valor:.2f}".replace('.', '')  # ARREDONDA e remove ponto
        fontequatro = str(row['Fonte 4 Digitos']).split('.')[0].zfill(4)
        fonteseis = str(row['Fonte 6 Digitos'])
        observacao = str(row['Observação'])
        data = str(row['Data do dia'])

        # Adiciona o conteúdo gerado ao modelo
        modelo_xml += f"""
    <screen name="Tela{tela}" entryscreen="true" exitscreen="false" transient="false">
        <description>
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="{data}1[tab]1527341[tab]{esfera}{ptres}{fontequatro}{nd2}1[tab]{data}9999[tab]{observacao}[tab][tab][tab]A{fonteseis}{nd4}[tab][tab]{pi}{valor_formatado}[tab]R{fonteseis}{nd3}00[tab][tab][tab]{valor_formatado}[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0">
            <nextscreen name="Tela{tela + 1}" />
        </nextscreens>
    </screen>

    <screen name="Tela{tela + 1}" entryscreen="false" exitscreen="false" transient="false">
        <description>
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
            <numfields number="242" optional="false" invertmatch="false" />
            <numinputfields number="1" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="s[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0">
            <nextscreen name="Tela{tela + 2}" />
        </nextscreens>
    </screen>

    <screen name="Tela{tela + 2}" entryscreen="false" exitscreen="false" transient="false">
        <description>
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
            <numfields number="64" optional="false" invertmatch="false" />
            <numinputfields number="0" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0">
            <nextscreen name="Tela{tela + 3}" />
        </nextscreens>
    </screen>
        """

        tela += 3

    # Finaliza o modelo
    modelo_xml += """
</HAScript>
    """

    # Salva o arquivo
    output_filename = os.path.join(output_directory, "Quadro de Detalhamento de Despesa.MAC")
    with open(output_filename, "w") as file:
        file.write(modelo_xml)

    print(f"Nova macro de vigência gerada com sucesso. O arquivo está no caminho: {output_filename}")

def main():

    print('iNICIANDO PROCESSO AGUARDE...')
    caminho_arquivo_PTRES = r'W:\B - TED\7 - AUTOMAÇÃO\QDD para descentralização\Crédito Disponivel Geral com Esfera.xlsx'
    copia_caminho_arquivo_PTRES = r'W:\B - TED\7 - AUTOMAÇÃO\QDD para descentralização\COPIA Crédito Disponivel Geral com Esfera.xlsx'
    caminho_arquivo_TERMO = r'W:\B - TED\7 - AUTOMAÇÃO\QDD para descentralização\Termo Aprovado aguardando descentralização.xlsx'
    copia_caminho_arquivo_TERMO = r'W:\B - TED\7 - AUTOMAÇÃO\QDD para descentralização\COPIA Termo Aprovado aguardando descentralização.xlsx'
    
    copiar_arquivo(caminho_arquivo_TERMO, copia_caminho_arquivo_TERMO)
    # remover_segunda_linha_excel(copia_caminho_arquivo_TERMO)
    copiar_arquivo(caminho_arquivo_PTRES, copia_caminho_arquivo_PTRES)
    remover_ultima_linha_excel(copia_caminho_arquivo_TERMO)
    remover_12_primeiras_linhas_excel(copia_caminho_arquivo_PTRES)
    remover_valores_zerados(copia_caminho_arquivo_TERMO)
    adicionar_colunas_excel(copia_caminho_arquivo_TERMO)
    
    # Carrega o arquivo Excel para um DataFrame
    df = pd.read_excel(copia_caminho_arquivo_TERMO)
    
    preencher_esfera_por_ptres_linha_a_linha(
    r'W:\B - TED\7 - AUTOMAÇÃO\QDD para descentralização\COPIA Termo Aprovado aguardando descentralização.xlsx',
    r'W:\B - TED\7 - AUTOMAÇÃO\QDD para descentralização\COPIA Crédito Disponivel Geral com Esfera.xlsx'
    )
    
    preencher_colunas(copia_caminho_arquivo_TERMO)
    
    # Salva o DataFrame modificado de volta no arquivo Excel
    formatar_data_do_dia(copia_caminho_arquivo_TERMO)
    formatar_valores_arquivo(copia_caminho_arquivo_TERMO)
    
    # Carrega o arquivo Excel novamente para obter o DataFrame atualizado
    df = pd.read_excel(copia_caminho_arquivo_TERMO)
    
    # Pula as 2 primeiras linhas do arquivo (mantém o cabeçalho original) 
    df_macro = df.iloc[2:].reset_index(drop=True) 
    
    #começa da linha 3 do arquivo Agora a macro usa o DF CORRETO já sem a segunda linha            
    generate_macro_vigencia(r'W:\B - TED\7 - AUTOMAÇÃO\QDD para descentralização', df_macro) 
    print("Processo finalizado com sucesso")                     

if __name__ == "__main__":
    main()
