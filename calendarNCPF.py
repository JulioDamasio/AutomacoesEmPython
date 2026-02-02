import pandas as pd
import csv
from datetime import datetime, timedelta
import tkinter as tk
from tkcalendar import Calendar
from tkinter import filedialog
from tkinter import messagebox
from database.reports.nc_report import generate_nc_report
from database.reports.pf_report import generate_pf_legado_report

# Lista para armazenar as datas selecionadas
selected_dates = []

# Função para adicionar uma data selecionada
def add_date():
    selected_date = cal.selection_get()
    if selected_date not in selected_dates:
        selected_dates.append(selected_date)
        update_selected_dates()

# Função para remover uma data selecionada
def remove_date():
    selected_date = cal.selection_get()
    if selected_date in selected_dates:
        selected_dates.remove(selected_date)
        update_selected_dates()

# Atualiza a exibição das datas selecionadas
def update_selected_dates():
    selected_dates_label.config(text=", ".join([date.strftime('%d/%m/%Y') for date in selected_dates]))

# Função para escolher o caminho de saída dos relatórios
def choose_output_path():
    output_path = filedialog.askdirectory()
    if output_path:
        output_path_entry.delete(0, tk.END)
        output_path_entry.insert(0, output_path)

# Função para gerar os relatórios
def generate_reports():
    output_path = output_path_entry.get()
    if not output_path:
        messagebox.showerror("Erro", "Por favor, escolha um caminho de saída.")
        return

    for selected_date in selected_dates:
        generate_nc_report([selected_date], output_path)
        generate_pf_legado_report([selected_date], output_path)
    
    messagebox.showinfo("Relatórios Gerados", "Relatórios gerados com sucesso!")

# Configuração da interface
root = tk.Tk()
root.title("Geração de NC e PF")

# Configurar o estilo do calendário
cal = Calendar(root, selectmode="day", year=datetime.now().year, month=datetime.now().month, day=datetime.now().day, 
               background='lightblue', foreground='black', font=('Arial', 12), selectbackground='green', 
               selectforeground='white', borderwidth=2, relief='solid')
cal.pack(pady=10)

# Exibe as datas selecionadas
selected_dates_label = tk.Label(root, text="", pady=5, font=('Arial', 10))
selected_dates_label.pack()

# Botões para adicionar e remover datas selecionadas
button_frame = tk.Frame(root)
button_frame.pack()

add_button = tk.Button(button_frame, text="Adicionar", command=add_date, font=('Arial', 10), bg='lightblue', relief='flat')
add_button.pack(side=tk.LEFT, padx=5)

remove_button = tk.Button(button_frame, text="Remover", command=remove_date, font=('Arial', 10), bg='lightcoral', relief='flat')
remove_button.pack(side=tk.LEFT, padx=5)

# Espaço entre os botões e o próximo elemento
tk.Label(root, text="").pack()

# Campo de entrada para o caminho de saída
output_path_label = tk.Label(root, text="Caminho de Saída:", font=('Arial', 10))
output_path_label.pack()

output_path_entry = tk.Entry(root, width=50, font=('Arial', 10))
output_path_entry.pack(padx=50, pady=(0, 10))

choose_output_button = tk.Button(root, text="Escolher Caminho", command=choose_output_path, font=('Arial', 10), bg='lightgray', relief='flat')
choose_output_button.pack(pady=(0, 10))

# Botão para gerar relatórios (modificado para azul)
generate_button = tk.Button(root, text="Gerar Relatórios", command=generate_reports, font=('Arial', 12), bg='lightblue', fg='black', relief='flat')
generate_button.pack(pady=10)

# Centralizar a janela na tela
window_width = 500
window_height = 455
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x_coordinate = (screen_width - window_width) // 2
y_coordinate = (screen_height - window_height) // 2
root.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")

# Iniciar interface
root.mainloop()