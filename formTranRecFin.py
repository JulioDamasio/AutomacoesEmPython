import tkinter as tk
from tkinter import ttk, filedialog
from tkinter import messagebox
from tranRecFin import generate_xml_tranRecFin  # Importe a função generate_xml do seu módulo XML

# Defina as variáveis globais no início do arquivo
data_geracao_entry = None
sequencial_geracao_entry = None
sequencial_doc_entry = None
ano_referencia_entry = None
ug_responsavel_entry = None
cpf_responsavel_entry = None
output_directory = None

def choose_output_path():
    global output_directory  # Use a variável global
    output_directory = filedialog.askdirectory()
    if not output_directory:
        return

    # Atualize o rótulo para exibir o diretório selecionado
    output_path_entry.delete(0, tk.END)
    output_path_entry.insert(0, output_directory)

def generate_lote_TranRecFin():
    global data_geracao_entry, sequencial_geracao_entry, sequencial_doc_entry, ano_referencia_entry, ug_responsavel_entry, cpf_responsavel_entry, output_directory

    # Verifique se o usuário selecionou um diretório e todos os campos estão preenchidos
    if not output_directory:
        messagebox.showerror("Erro", "Por favor, escolha um caminho de saída.")
        return

    data_geracao = data_geracao_entry.get()
    sequencial_geracao = sequencial_geracao_entry.get()
    sequencial_doc = sequencial_doc_entry.get()
    ano_referencia = ano_referencia_entry.get()
    ug_responsavel = ug_responsavel_entry.get()
    cpf_responsavel = cpf_responsavel_entry.get()
    
    if not data_geracao or not sequencial_geracao or not ano_referencia or not ug_responsavel or not cpf_responsavel:
        messagebox.showerror("Erro", "Preencha todos os campos antes de gerar o lote.")
        return

    # Chame a função em XML.py para gerar o arquivo XML com os dados
    generate_xml_tranRecFin(data_geracao, sequencial_geracao, ano_referencia, ug_responsavel, cpf_responsavel, output_directory, sequencial_doc)

    # Exiba um popup com a mensagem de sucesso
    messagebox.showinfo("Sucesso", "Lote Gerado com Sucesso")
    
root = tk.Tk()
root.title("Gerar Lote de TranRecFIn")

# Formulário para preencher os campos do cabeçalho
form_frame = ttk.Frame(root)
form_frame.grid(row=0, column=0, padx=20, pady=10, sticky="w")

data_geracao_label = ttk.Label(form_frame, text="Data Geração (DD/MM/AAAA):")
data_geracao_label.grid(row=0, column=0, sticky="w")
data_geracao_entry = ttk.Entry(form_frame)
data_geracao_entry.grid(row=0, column=1, padx=(0,280), pady=5)

sequencial_geracao_label = ttk.Label(form_frame, text="Número de Sequência (4 digitos aleatórios):")
sequencial_geracao_label.grid(row=1, column=0, sticky="w")
sequencial_geracao_entry = ttk.Entry(form_frame)
sequencial_geracao_entry.grid(row=1, column=1, padx=(0,280), pady=5)

sequencial_doc_label = ttk.Label(form_frame, text="Número do documento (de 400001 à 800000):")
sequencial_doc_label.grid(row=2, column=0, sticky="w")
sequencial_doc_entry = ttk.Entry(form_frame)
sequencial_doc_entry.grid(row=2, column=1, padx=(0,280), pady=5)

ano_referencia_label = ttk.Label(form_frame, text="Ano Referência:")
ano_referencia_label.grid(row=3, column=0, sticky="w")
ano_referencia_entry = ttk.Entry(form_frame)
ano_referencia_entry.grid(row=3, column=1, padx=(0,280), pady=5)

ug_responsavel_label = ttk.Label(form_frame, text="UG Responsável:")
ug_responsavel_label.grid(row=4, column=0, sticky="w")
ug_responsavel_entry = ttk.Entry(form_frame)
ug_responsavel_entry.grid(row=4, column=1, padx=(0,280), pady=5)

cpf_responsavel_label = ttk.Label(form_frame, text="CPF Responsável:")
cpf_responsavel_label.grid(row=5, column=0, sticky="w")
cpf_responsavel_entry = ttk.Entry(form_frame)
cpf_responsavel_entry.grid(row=5, column=1, padx=(0,280), pady=5)

# Rótulo e campo de entrada para o caminho de saída no mesmo frame
output_path_label = ttk.Label(form_frame, text="Caminho de Saída:")
output_path_label.grid(row=6, column=0, sticky="w")
output_path_entry = ttk.Entry(form_frame, width=50)
output_path_entry.grid(row=6, column=1, columnspan=5, padx=(0, 100), pady=5)

choose_output_button = ttk.Button(form_frame, text="Escolher Caminho", command=choose_output_path)
choose_output_button.grid(row=7, column=1, padx=(0,270), pady=5)# Ainda pode ajustar o pady como preferir

generate_button = ttk.Button(root, text="Gerar Lote de TranRecFin", command=generate_lote_TranRecFin)  # Remova os parênteses
generate_button.grid(row=1, column=0, padx=(90, 120), pady=5)

window_width = 800
window_height = 320
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x_coordinate = (screen_width - window_width) // 2
y_coordinate = (screen_height - window_height) // 2
root.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")

root.mainloop()