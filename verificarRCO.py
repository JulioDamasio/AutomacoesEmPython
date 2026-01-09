import pandas as pd
import os

def fazer_copia(origem, destino):
    df = pd.read_excel(origem)
    df.to_excel(destino, index=False)

def apagar_linhas(arquivo):
    # Carregar o arquivo Excel
    df = pd.read_excel(arquivo)

    # Verificar se o DataFrame não está vazio
    if df.empty:
        print("O arquivo está vazio. Nenhuma linha foi apagada.")
        return

    # Verificar os índices existentes e apagar a segunda e a última linha
    indices_para_apagar = []
    if 0 in df.index:  # Verifica se o índice 0 existe
        indices_para_apagar.append(0)
    if (len(df) - 1) in df.index:  # Verifica se o índice da última linha existe
        indices_para_apagar.append(len(df) - 1)

    if indices_para_apagar:
        df = df.drop(indices_para_apagar)
        print(f"Linhas {indices_para_apagar} removidas.")
    else:
        print("Nenhuma linha correspondente encontrada para apagar.")

    # Salvar o arquivo atualizado
    df.to_excel(arquivo, index=False)

def filtrar_e_apagar_linhas(arquivo):
    # Carregar o arquivo Excel
    df = pd.read_excel(arquivo)

    # Contar o número de linhas antes da filtragem
    total_linhas_antes = df.shape[0]

    # Filtrar as linhas com base na coluna "Situação Documento"
    df = df[(df['Situação Documento'] != 'Arquivado') & (df['Situação Documento'] != 'Comprovado no SIAFI.')]

    # Salvar o arquivo atualizado
    df.to_excel(arquivo, index=False)

    # Contar o número de linhas após a filtragem
    total_linhas_depois = df.shape[0]

    # Calcular o número de linhas apagadas
    linhas_apagadas = total_linhas_antes - total_linhas_depois

    print(f"{linhas_apagadas} linhas foram apagadas do arquivo.")

def apagar_linhas_siafi_vazio(arquivo):
    # Carregar o arquivo Excel
    df = pd.read_excel(arquivo)

    # Filtrar as linhas onde a coluna "SIAFI" não é "-" e não está vazia
    df_filtrado = df[~df['SIAFI'].isin(['-', ''])]

    # Salvar o arquivo atualizado
    df_filtrado.to_excel(arquivo, index=False)          

def comparar_unidades_gestoras(arquivo1, arquivo2):
    df1 = pd.read_excel(arquivo1)
    df2 = pd.read_excel(arquivo2)

    unidades_gestoras_arquivo1 = set(df1["Unidade Gestora Descentralizada"].astype(str).str[:6].unique())
    unidades_gestoras_arquivo2 = set(df2["Unidade Gestora Descentralizada"].astype(str).str[:6].unique())

    unidades_em_comum = unidades_gestoras_arquivo1.intersection(unidades_gestoras_arquivo2)

    if unidades_em_comum:
        print("Aviso: Algumas unidades gestoras estão presentes em ambos os arquivos.")
        print("Unidades Gestoras em Comum:", unidades_em_comum)
    else:
        print("Não há unidades gestoras em comum entre os arquivos.")
        
def comparar_termo_ted(arquivo_analise_final, arquivo_rcos):
    # Leitura dos arquivos
    df_analise_final = pd.read_excel(arquivo_analise_final)
    df_rcos_vencidos = pd.read_excel("W:\\B - TED\\7 - AUTOMAÇÃO\\Relatório pendencia RCO\\Copia vencidos + 120 dias.xlsx")
    df_rcos_entrega = pd.read_excel("W:\\B - TED\\7 - AUTOMAÇÃO\\Relatório pendencia RCO\\Copia Relatório de Entrega do RCO.xlsx")

    # Convertendo as colunas para o mesmo tipo de dados
    df_rcos_vencidos["Termo"] = df_rcos_vencidos["Termo"].astype(str)
    df_rcos_entrega["TED"] = df_rcos_entrega["TED"].astype(str)

    unidades_gestoras = df_analise_final["Unidade Gestora Descentralizada"].astype(str)

    for unidade_gestora in unidades_gestoras:
        print(f"Processando unidade gestora: {unidade_gestora}")

        linhas_filtradas = df_rcos_vencidos[
            df_rcos_vencidos["Unidade Gestora Descentralizada"].astype(str).str.startswith(unidade_gestora[:6])
        ]

        print(f"Quantidade de linhas encontradas para a unidade gestora {unidade_gestora}: {len(linhas_filtradas)}")

        termos_encontrados = set()
        termos_nao_encontrados = set()
        for termo in linhas_filtradas["Termo"]:
            if termo in set(df_rcos_entrega["TED"].astype(str)):
                termos_encontrados.add(termo)
            else:
                termos_nao_encontrados.add(termo)

        print(f"Quantidade de termos encontrados para a unidade gestora {unidade_gestora}: {len(termos_encontrados)} das {len(linhas_filtradas)} linhas")

        if len(termos_encontrados) != len(linhas_filtradas):
            mensagem_pendencia = f"Há pendência na entrega dos teds para a unidade gestora {unidade_gestora}: "
            if termos_nao_encontrados:
                mensagem_pendencia += f"Termos não encontrados: {', '.join(map(str, termos_nao_encontrados))}"
            print(mensagem_pendencia)

        indices_unidade = df_analise_final["Unidade Gestora Descentralizada"].astype(str) == unidade_gestora
        termos_nao_encontrados_str = ", ".join(map(str, termos_nao_encontrados))
        if len(termos_nao_encontrados) > 0:
            df_analise_final.loc[indices_unidade, "Tem Pendência de RCO?"] = f"Há pendência na entrega dos teds: {termos_nao_encontrados_str}"
        else:
            df_analise_final.loc[indices_unidade, "Tem Pendência de RCO?"] = "Não"

    # Salvar o arquivo de forma segura
    with pd.ExcelWriter(arquivo_analise_final, mode="w", engine="openpyxl") as writer:
        df_analise_final.to_excel(writer, index=False)

def adicionar_coluna_ted_pactuado(arquivo_analise_final):
    # Carregar o arquivo Excel
    df_analise_final = pd.read_excel(arquivo_analise_final)

    # Adicionar a coluna "TED Pactuado?"
    df_analise_final.insert(len(df_analise_final.columns) - 1, "TED Pactuado?", "")

    # Preencher a coluna de acordo com a condição (considerando "-" e valores vazios como "Não")
    df_analise_final["TED Pactuado?"] = df_analise_final.apply(lambda row: "Sim" if str(row["SIAFI"]).strip() not in ["-", ""] else "Não", axis=1)

    # Salvar as alterações de volta no arquivo Excel
    df_analise_final.to_excel(arquivo_analise_final, index=False)

    # Fechar o arquivo após salvar as alterações
    del df_analise_final  # Remover o DataFrame da memória para garantir que o arquivo seja fechado
    
def main():
    # Definindo os caminhos dos arquivos
    arquivo1_origem = "W:\\B - TED\\7 - AUTOMAÇÃO\\Relatório pendencia RCO\\Teds em análise pela ug intermediária.xlsx"
    arquivo1_copia = "W:\\B - TED\\7 - AUTOMAÇÃO\\Relatório pendencia RCO\\Copia Teds em análise pela ug intermediária.xlsx"
    arquivo2_origem = "W:\\B - TED\\7 - AUTOMAÇÃO\\Relatório pendencia RCO\\vencidos + 120 dias.xlsx"
    arquivo2_copia = "W:\\B - TED\\7 - AUTOMAÇÃO\\Relatório pendencia RCO\\Copia vencidos + 120 dias.xlsx"
    arquivo3_origem = "W:\\B - TED\\7 - AUTOMAÇÃO\\Relatório pendencia RCO\\Relatório de Entrega do RCO.xlsx"
    arquivo3_copia = "W:\\B - TED\\7 - AUTOMAÇÃO\\Relatório pendencia RCO\\Copia Relatório de Entrega do RCO.xlsx"
    arquivo_rcos = "W:\\B - TED\\7 - AUTOMAÇÃO\\Relatório pendencia RCO\\Copia Relatório de Entrega do RCO.xlsx"
    arquivo_analise_final = "W:\\B - TED\\7 - AUTOMAÇÃO\\Relatório pendencia RCO\\Análise Final.xlsx"

    # Fazer a cópia dos arquivos
    print("Iniciando o processamento aguarde...")
    fazer_copia(arquivo1_origem, arquivo1_copia)
    fazer_copia(arquivo2_origem, arquivo2_copia)
    fazer_copia(arquivo3_origem, arquivo3_copia)
    apagar_linhas(arquivo_analise_final)
    apagar_linhas(arquivo1_copia)
    apagar_linhas(arquivo2_copia)
    apagar_linhas_siafi_vazio(arquivo2_copia)
    filtrar_e_apagar_linhas(arquivo2_copia)
    fazer_copia(arquivo1_copia, arquivo_analise_final)
    comparar_unidades_gestoras(arquivo1_copia, arquivo2_copia)
    comparar_termo_ted(arquivo_analise_final, arquivo_rcos)
    adicionar_coluna_ted_pactuado(arquivo_analise_final)
    print("Processo totalmente finalizado. O arquivo se encontra em: "'W:\B - TED\7 - AUTOMAÇÃO\Relatório pendencia RCO\Análise Final.xlsx')

if __name__ == "__main__":
    main()