import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import subprocess

# Funções para os botões
def install_libraries():
    try:
        subprocess.run(["pip", "install", "-r", "W:/B - TED/AUTOMAÇÃO/Scripts/requirements.txt"])
        messagebox.showinfo("Instalação Concluída", "As bibliotecas foram instaladas com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro na Instalação", f"Erro: {str(e)}")

def generate_nc_pf_reports(): subprocess.Popen(["python", "calendarNCPF.py"])
def generate_liquidação_report(): subprocess.Popen(["python", "tedLiquidacao.py"])
def generate_painel_reports(): subprocess.Popen(["python", "painel.py"])
def generate_XML(): subprocess.Popen(["python", "formXML.py"])
def generate_macro(): subprocess.Popen(["python", "macroNL.py"])
def generate_painel_interno(): subprocess.Popen(["python", "orcFin.py"])
def generate_aditivo(): subprocess.Popen(["python", "macroAditivo.py"])
def generate_pendencia_rco(): subprocess.Popen(["python", "verificarRCO.py"])
def generate_conformidade(): subprocess.Popen(["python", "conformidade.py"])
def generate_imprenssa(): subprocess.Popen(["python", "ImprenssaNacional.py" ])
def generate_cadastrarEmpenho(): subprocess.Popen(["python", "cadastrarEmpenho.py" ])
def generate_finalizaTermo(): subprocess.Popen(["python", "finalizarTermo.py"])
def generate_MacroFinalizarTed(): subprocess.Popen(["python", "macroBaixaSaldo.py"])
def generate_auditoria(): subprocess.Popen(["python", "auditoria.py"])
def generate_residencia(): subprocess.Popen(["python", "residencia.py"])
def generate_QDD(): subprocess.Popen(["python", "qdd.py"])
def generate_RP(): subprocess.Popen(["python", "integrar_RP_TED.py"])

# Configuração da interface
root = tk.Tk()
root.title("TED Gerencial")
root.geometry("850x530") 
root.configure(bg="#e3f2fd")  # Cor de fundo suave

# Garantir que a janela fique no primeiro plano
root.focus_force()

# Cabeçalho
header = tk.Label(root, text="TED Gerencial", bg="#1565c0", fg="white", font=("Arial", 16, "bold"))
header.pack(fill=tk.X, pady=(10, 5))  # Diminuir o espaço entre o cabeçalho e a tabela

# Criar um Frame para os botões com cor de fundo azul claro
frame = ttk.Frame(root, padding=10)
frame.pack(pady=(10, 20))  # Diminuir o espaço entre a tabela e o rodapé

# Definindo a cor de fundo da tabela
frame.configure(style="TFrame")

# Estilo dos botões
style = ttk.Style()
style.configure("TButton",
                padding=10,
                font=("Arial", 11),
                width=25,
                background="#42a5f5",  # Azul médio para os botões
                focuscolor="#1976d2",
                borderwidth=2,
                relief="solid")
style.map("TButton",
          background=[('active', '#1976d2')],
          foreground=[('active', 'black')])  # Cor do texto no estado ativo para preto

# Estilo para o fundo da tabela
style.configure("TFrame", background="#bbdefb")  # Azul claro para a tabela

# Lista de botões (12 botões organizados em uma tabela 4x5)
botoes = [
    ("Instalar Bibliotecas", install_libraries),
    ("NC e PF", generate_nc_pf_reports),
    ("Liquidação", generate_liquidação_report),
    ("Painel", generate_painel_reports),
    ("Lote de RC", generate_XML),
    ("Macro de NL", generate_macro),
    ("Painel Interno", generate_painel_interno),
    ("Aditivo de Vigência", generate_aditivo),
    ("Verificar Pendência de RCO", generate_pendencia_rco),
    ("Conformidade", generate_conformidade),
    ("Imprensa Nacional", generate_imprenssa),
    ("Cadastrar Empenho", generate_cadastrarEmpenho),
    ("Finalizar TEDS planilha", generate_finalizaTermo),
    ("Baixa Saldo SIAFI", generate_MacroFinalizarTed),
    ("Auditoria", generate_auditoria),
    ("Residência", generate_residencia),
    ("QDD", generate_QDD),
    ("Integrar RP", generate_RP)
]

# Criar os botões na grade e centralizar a tabela
for i, (text, command) in enumerate(botoes):
    row, col = divmod(i, 3)  # Distribui os botões em 3 colunas
    btn = ttk.Button(frame, text=text, command=command)
    btn.grid(row=row, column=col, padx=10, pady=10)

# Rodapé
footer = tk.Label(root, text="Desenvolvido por CGSO/CTED - 2023", bg="#1565c0", fg="white", font=("Arial", 12, "bold"))
footer.pack(fill=tk.X, pady=(10, 5))  # Diminuir o espaço entre a tabela e o rodapé

# Não minimizar automaticamente a janela
# root.state('iconic')  # Removido, caso queira que a janela abra normalmente

# Iniciar interface
tk.mainloop()
