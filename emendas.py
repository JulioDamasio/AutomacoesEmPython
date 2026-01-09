import pandas as pd
import re

def extrair_empenho(arquivo_excel, aba, coluna_origem, coluna_destino):
    # Ler o arquivo Excel
    df = pd.read_excel(arquivo_excel, sheet_name=aba)
    
    # Função para extrair o texto completo após "EMPENHO:" até o próximo espaço ou final da string
    def extrair_numero_empenho(texto):
        if isinstance(texto, str) and "EMPENHO:" in texto:
            # Captura tudo após "EMPENHO:" até o final da string ou próximo espaço
            match = re.search(r'EMPENHO: ([^\s]+)', texto)
            if match:
                return match.group(1)
        return None
    
    # Aplicar a função na coluna de origem e salvar na coluna de destino
    df[coluna_destino] = df[coluna_origem].apply(extrair_numero_empenho)
    
    # Salvar o resultado no mesmo arquivo Excel
    with pd.ExcelWriter(arquivo_excel, mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name=aba, index=False)
    
    print(f"Extração completa. Dados salvos na coluna '{coluna_destino}' da aba '{aba}'.")
    
extrair_empenho(r'C:\Users\juliodamasio\Desktop\Pasta1.xlsx', 'Planilha1', 'OBSERVAÇÃO DA PF', 'Empenho')
extrair_empenho(r'C:\Users\juliodamasio\Downloads\PF Legado - Exercício 2024.xlsx', 'PF Legado - Exercício 2024', 'Doc - Observação', 'Empenho')