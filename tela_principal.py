from tkinter import ttk
from tkinter import *
import tela_registros
import tela_amostras 
from tela_registros import Dados_Cotacoes
from tela_amostras import Dados_Amostras
from tela_registros import Funcs_Cotacoes
from tela_amostras import Funcs_Amostras 

registros_dados = Dados_Cotacoes # sqlToExcel, backupSqliteDados
amostras_dados = Dados_Amostras # sqlToExcel, backupSqliteDados

registros_funcs = Funcs_Cotacoes # mostrarTodos
amostras_funcs = Funcs_Amostras # mostrarTodos

root = Tk()
root.title("Registros de Fornecedores & Controle de Amostras")
root.geometry("1250x650")
root.maxsize(width=1350, height=750)  
root.minsize(width=1200, height=600)
root.configure(background='#1e3743', border=5)

# Define Icone
# root.iconbitmap(bitmap='icon.ico')

# Criando um Widget Notebook
notebook = ttk.Notebook(root)
notebook.pack(expand=True, fill='both')

# Tela 1 : Registros
registros_screen = tela_registros.RegistrosScreen(notebook)
notebook.add(registros_screen, text='Registros de Fornecedores')

# Tela 2 : Amostras
amostras_screen = tela_amostras.AmostrasScreen(notebook)
notebook.add(amostras_screen, text='Controle de Amostras')

# Crie um menu 
menu_bar = Menu(root)
root.config(menu=menu_bar)

# Crie um menu "Arquivo"
menu_arquivo = Menu(menu_bar, tearoff=0)
menu_arquivo.add_command(label="Fechar programa", command=root.quit)
menu_bar.add_cascade(label="Arquivo", menu=menu_arquivo)

# Exporta
menu_exporta = Menu(menu_bar, tearoff=0)
menu_exporta.add_command(label="Fornecedores", command=registros_dados.sqlToExcel)
menu_exporta.add_separator()
menu_exporta.add_command(label="Amostras", command=amostras_dados.sqlToExcel)
menu_bar.add_cascade(label="Exporta Excel", menu=menu_exporta)

#Backup
menu_backup = Menu(menu_bar, tearoff=0)
menu_backup.add_command(label="Fornecedores", command=registros_dados.backupSqliteDados)
menu_backup.add_separator()
menu_backup.add_command(label="Amostras", command=amostras_dados.backupSqliteDados)
menu_bar.add_cascade(label="Criar Backup", menu=menu_backup)

root.mainloop()

# pyinstaller --onefile --windowed --noconsole --icon=icon.ico tela_principal.py
# Após concluido a arquivo executável, adicione o "Icon.icon" a pasta onde está o executável.