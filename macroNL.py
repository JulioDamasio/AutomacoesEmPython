import tkinter as tk
from tkinter import ttk
import pandas as pd
from datetime import datetime
import os

df = pd.read_excel(r'W:\B - TED\7 - AUTOMAÇÃO\macro lote NL\NL exercícios anteriores.xlsx')

print("Iniciando processo da macro, aguarde...")

def generate_screen(tela, col_ug, col_gestao, col_observacao, col_evento, col_fonte, col_categoria, col_siafi, col_valor):
    return f"""
    <screen name="Tela{tela}" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
            <numfields number="80" optional="false" invertmatch="false" />
            <numinputfields number="16" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="[tab][tab][tab][tab][tab]{col_ug}[tab]{col_gestao}[tab][tab][tab]{col_observacao}[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Tela{tela + 1}" />
        </nextscreens>
    </screen>

    <screen name="Tela{tela + 1}" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
            <numfields number="146" optional="false" invertmatch="false" />
            <numinputfields number="48" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="{col_evento + col_fonte + col_categoria}[tab][tab][tab]{col_siafi}[tab][tab][tab]{col_valor}[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Tela{tela + 2}" />
        </nextscreens>
    </screen>

    <screen name="Tela{tela + 2}" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
            <numfields number="144" optional="false" invertmatch="false" />
            <numinputfields number="1" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="s[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Tela{tela + 3}" />
        </nextscreens>
    </screen>

    <screen name="Tela{tela + 3}" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
            <numfields number="64" optional="false" invertmatch="false" />
            <numinputfields number="0" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Tela{tela + 4}" />
        </nextscreens>
    </screen>
        """

def generate_macro(output_directory):
    modelo_xml = """<HAScript name="gravacao de macro exercicio 2024 nl" description="" timeout="60000" pausetime="300" promptall="true" blockinput="false" author="AugustoCezar" creationdate="09/01/2024 11:10:32" supressclearevents="false" usevars="false" ignorepauseforenhancedtn="true" delayifnotenhancedtn="0" ignorepausetimeforenhancedtn="true">
    
    <screen name="Tela1" entryscreen="true" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="&gt;nl[enter]" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Tela2" />
        </nextscreens>
    </screen> 
    """
    tela = 2  # Inicializa a variável com o número da primeira tela

    for index, row in df.iterrows():
        col_ug = str(row['UG'])
        col_gestao = str(row['Gestão']).zfill(5)
        col_observacao = str(row['Observação'])
        col_evento = str(row['Evento'])
        col_fonte = str(row['Fonte'])
        col_categoria = str(row['Categoria'])
        col_siafi = "" if pd.isna(row['Siafi']) else str(row['Siafi'])
        # Tratamento do valor
        valor_bruto = str(row[7])  # Valor bruto como string
        try:
            # Converte para float, arredonda para 2 casas e converte de volta para string
            valor_formatado = f"{float(valor_bruto):.2f}"  
            # Remove separadores (ponto e vírgula)
            col_valor = valor_formatado.replace('.', '').replace(',', '')
        except ValueError:
            col_valor = '0'  # Caso o valor não seja numérico, usa 0 como padrão
        
        modelo_xml += generate_screen(tela, col_ug, col_gestao, col_observacao, col_evento, col_fonte, col_categoria, col_siafi, col_valor)

        tela += 4  # Atualiza a variável para o próximo número de tela

    # fecha o modelo
    modelo_xml += """    
</HAScript> 
        """

    # Salva o arquivo
    output_filename = os.path.join(output_directory, "MacroNL.MAC")
    with open(output_filename, "w") as file:
        file.write(modelo_xml)

    print(f"Macro gerada com sucesso. O arquivo está no caminho: {output_filename}")

# Substitua df pelo seu DataFrame real
df = pd.read_excel(r'W:\B - TED\7 - AUTOMAÇÃO\macro lote NL\NL exercícios anteriores.xlsx')

# Chame a função com o diretório desejado
generate_macro(r'W:\B - TED\7 - AUTOMAÇÃO\macro lote NL')