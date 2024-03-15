from tkinter import *
from tkinter import ttk
import tkinter as tk
import sqlite3
import shutil
import pandas as pd
import pyperclip as pc
from tkinter import messagebox
from datetime import datetime
import os

nomeArquivoRegistros = 'regs_registros.db'

class Dados_Cotacoes():
    
    def sqlToExcel():

        resposta = messagebox.askquestion("Confirmação", "Deseja Exportação os Registros para o Excel?")

        if resposta == "yes":

            # Crie uma pasta para salvar os arquivos XLSX (se não existir)
            output_folder = 'exporta_xlsx'
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)

            # Obter a data atual
            data_atual = datetime.now().strftime("%d%m%y")
            nome_arquivo_xlsx = f"cotacoesExcel_{data_atual}.xlsx"
            pasta_destino = "exporta_xlsx"
            caminho_completo = f"{pasta_destino}/{nome_arquivo_xlsx}"

            conexao = sqlite3.connect(nomeArquivoRegistros)
            c =  conexao.cursor()
            c.execute("SELECT * FROM registros")
            tabelaSql = c.fetchall()
            tabelaSql = pd.DataFrame(tabelaSql, columns=['Código','Produto','Marcas','Frete %','Suframa %','ICMS %','Contato','E-mail'])
            tabelaSql.to_excel(caminho_completo, index=False)
            
            messagebox.showinfo("Operação Concluída!" ,"Acesso o diretório do programa para obter o arquivo 'cotacoesExcel.xlsx'.")
        else:
            messagebox.showinfo("Cancela!" ,"Exportação Cancelada!")   

    def backupSqliteDados():
        resposta = messagebox.askquestion("Confirmação", "Deseja Fazer o Backup dos Registros de Cotações")
        
        # Crie uma pasta para salvar os arquivos XLSX (se não existir)
        output_folder = 'backups_sqlite'
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        # Obter a data atual
            data_atual = datetime.now().strftime("%d%m%Y")
            nome_novo = f"regs_registros_{data_atual}.db"
        
        if resposta == "yes":
            caminhoEntradaSqlDados = "regs_registros.db"
            try:
                shutil.copy(caminhoEntradaSqlDados, os.path.join(output_folder, nome_novo))
                messagebox.showinfo("Concluído!" ,"Backup Criado!")
            except Exception as e:
                messagebox.showinfo(message=f"Erro ao criar o backup: {str(e)}")
        else:
             messagebox.showinfo("Cancelado!" ,"Backup Cancelado!")

class Funcs_Cotacoes():
       
    def limpar(self): # LIMPAR CAMPOS
        self.entry_cod.delete(0, END)
        self.entry_produto.delete(0, END)
        self.entry_marca.delete(0, END)
        self.entry_frete.delete(0, END)
        self.entry_suframa.delete(0, END)
        self.entry_icms.delete(0, END)
        self.entry_contato.delete(0, END)
        self.entry_email.delete(0, END)

    def conecta_bd(self):
        self.conn = sqlite3.connect(nomeArquivoRegistros)
        self.cursor = self.conn.cursor()

    def desconecta_bd(self):
        self.conn.close(); print('Desconectando ao Banco de dados...')
     
    def criar_tabela(self): ### CRIAR TABELA
        self.conecta_bd()
        self.cursor.execute(""" CREATE TABLE IF NOT EXISTS registros (
                            cod INTEGER PRIMARY KEY,
                            produto CHAR(40) NOT NULL,
                            marca CHAR(40),
                            frete INTEGER(10),
                            suframa INTEGER(10),
                            icms INTEGER(10),
                            contato INTEGER(20),
                            email CHAR(40));
                        """)
        self.conn.commit(); 
        print("Banco de dados criado...")
        self.desconecta_bd()

    def variaveis(self): ### Variaveis
        self.cod = self.entry_cod.get()
        self.produto = self.entry_produto.get()
        self.marca = self.entry_marca.get()
        self.frete = self.entry_frete.get()
        self.suframa = self.entry_suframa.get()
        self.icms = self.entry_icms.get()
        self.contato = self.entry_contato.get()
        self.email = self.entry_email.get()

    def cadastrar(self): ### NOVOS CADASTROS
        self.variaveis()

        if self.entry_produto.get()  == "":
            messagebox.showinfo(title="Erro de operação!" ,
                                message="Digite o Nome do Produto para Cadastrar!")
            self.entry_produto.focus()
        else:
            self.conecta_bd()
            self.cursor.execute(""" INSERT INTO registros 
                                (produto, 
                                marca, 
                                frete, 
                                suframa, 
                                icms, 
                                contato, 
                                email) 
                                VALUES(?, ?, ?, ?, ?, ?, ?) """,
                                (self.produto, 
                                 self.marca, 
                                 self.frete, 
                                 self.suframa, 
                                 self.icms, 
                                 self.contato, 
                                 self.email))
            self.conn.commit()   
            self.desconecta_bd()
            self.selecionar()
            self.limpar()
        
    def selecionar(self): ### Selecionar NA Lista
        self.layout_consultar()
        self.registrados.delete(*self.registrados.get_children())
        
        self.conecta_bd()
        lista = self.cursor.execute(
            """ SELECT cod, 
            produto, marca, frete, 
            suframa, icms, contato, email FROM registros ORDER BY produto ASC; """)
        
        for i in lista:
            self.registrados.insert("", END, values=i)
        self.desconecta_bd()

    def filtrar(self):
        texto_pesquisa = self.pesquisar_entry.get()

        for row in self.registrados.get_children():
            self.registrados.delete(row)

        self.conn = sqlite3.connect(nomeArquivoRegistros)
        self.cursor = self.conn.cursor()
        
        # Consulta SQL para buscar dados com base no critério de pesquisa
        # query = f"SELECT * FROM registros WHERE produto LIKE '%{texto_pesquisa}%'"
        query = f"SELECT * FROM registros WHERE produto LIKE '%{texto_pesquisa}%' OR marca LIKE '%{texto_pesquisa}%' OR contato LIKE '%{texto_pesquisa}%'"

        self.cursor.execute(query)
        rows = self.cursor.fetchall()

        # Insira os dados filtrados no treeview
        for row in rows:
            self.registrados.insert('', 'end', values=row)

        # Feche a conexão com o banco de dados SQLite
        self.conn.close()
    
    def mostrarTodos(self):  # Mostrar todos
        self.pesquisar_entry.delete(0, END)
        self.pesquisar_entry.focus()
    
        self.conecta_bd()
        self.registrados.delete(*self.registrados.get_children())
        self.pesquisar_entry.insert(END, "%")
        
        nome = self.pesquisar_entry.get()
        self.cursor.execute(
            """ SELECT * FROM registros WHERE produto LIKE '%s' ORDER BY produto ASC """ % nome)
        buscanome = self.cursor.fetchall()
        
        for i in buscanome:
            self.registrados.insert("", END, values=i)
            
        self.pesquisar_entry.delete(0, END)
        self.desconecta_bd() 
  
    def OnDoubleClick(self, event):
        self.limpar()  
        self.registrados.selection()
        for n in self.registrados.selection():
            col1, col2, col3, col4, col5, col6, col7, col8 = self.registrados.item(n, 'values')
            self.entry_cod.insert(END, col1)
            self.entry_produto.insert(END, col2)
            self.entry_marca.insert(END, col3)
            self.entry_frete.insert(END, col4)
            self.entry_suframa.insert(END, col5)
            self.entry_icms.insert(END, col6)
            self.entry_contato.insert(END, col7)
            self.entry_email.insert(END, col8)

    def deletar(self): ### Deletar Cliente
        self.variaveis()
        
        if self.entry_cod.get()  == "":
            messagebox.showinfo(title="Erro de operação!" ,
                                message="Digite o número 'Cod' do Registro no campo 'Cod' ou dê um click duplo no Registro que deseja excluir e aperte no botão Excluir!")
            self.entry_cod.focus()
        else: 
            verify = messagebox.askyesno(title="Excluir Registro?", message="Deseja Excluir esse Registro?")
            if  verify == True:  
                self.conecta_bd()
                self.cursor.execute(""" DELETE FROM registros  WHERE cod = ? """, ([self.cod]))
                self.conn.commit()
                self.desconecta_bd()
                self.limpar()
                self.selecionar()
            else:
                messagebox.showinfo(message="Operação Cancelada!")
   
    def alterar(self): ### Alterar Cliente
        self.variaveis()
        
        if self.entry_produto.get()  == "":
            messagebox.showinfo(title="Erro de operação!" ,message="Escolha um registro já cadastrado para Alterar as Informações!")
        else:
            self.conecta_bd()
            self.cursor.execute(""" UPDATE registros SET produto = ?, marca = ?, frete = ?, suframa = ?, icms = ?, contato = ?, email = ? WHERE cod = ? """, (self.produto , self.marca , self.frete , self.suframa , self.icms , self.contato , self.email, self.cod))
            
            self.conn.commit()
            self.desconecta_bd()
            self.selecionar()
            self.limpar()
                   
    def copy_email(self):
        self.variaveis()   
        pc.copy(self.email)        

class RegistrosScreen(tk.Frame, Funcs_Cotacoes, Dados_Cotacoes):

    def __init__(self, master=None):
        super().__init__(master)
        self.layout()
        self.layout_consultar()
        self.layout_cadastrados()
        self.criar_tabela()
        self.selecionar()

    def layout(self):

        self.frameNovo = ttk.Frame(self)
        self.frameNovo.place(
            relx=0.01, 
            rely=0.02, 
            relwidth=0.98, 
            relheight=0.22)
         
        self.labelFrameNovoRegistro = ttk.LabelFrame(self.frameNovo, text="Cadastrar Fornecedor")
        self.labelFrameNovoRegistro.place(
            relx=0.01, 
            rely=0.01, 
            relwidth=0.98, 
            relheight=0.98)
        
        ## cod
        self.lb_cod = ttk.Label(self.labelFrameNovoRegistro,text="Cod").place(
            relx=0.01, 
            rely=0.02)
        self.entry_cod = ttk.Entry(self.labelFrameNovoRegistro)
        self.entry_cod.place(
            relx=0.01, 
            rely=0.20, 
            relwidth=0.03, 
            relheight=0.20)
        
        ## produto
        self.lb_produto = ttk.Label( self.labelFrameNovoRegistro, text="Descriçãp do Produto").place(
            relx=0.05, 
            rely=0.02)
        self.entry_produto = ttk.Entry(self.labelFrameNovoRegistro)
        self.entry_produto.place(
            relx=0.05, 
            rely=0.20, 
            relwidth=0.20, 
            relheight=0.20)
        
        ### marca
        self.lb_marca = ttk.Label(self.labelFrameNovoRegistro, text="Marca | Fabricante").place(
            relx=0.26, 
            rely=0.02)
        self.entry_marca = ttk.Entry(self.labelFrameNovoRegistro)
        self.entry_marca.place(
            relx=0.26, 
            rely=0.20, 
            relwidth=0.15, 
            relheight=0.20)  


         ### frete
        self.lb_frete = ttk.Label(self.labelFrameNovoRegistro, text="Frete").place(
            relx=0.42, 
            rely=0.02)
        self.entry_frete = ttk.Spinbox(self.labelFrameNovoRegistro, from_=10, to=50)
        self.entry_frete.place(
            relx=0.42, 
            rely=0.20, 
            relwidth=0.055, 
            relheight=0.20)
        
        ### suframa
        self.lb_suframa = ttk.Label(self.labelFrameNovoRegistro, text="Suframa").place(
            relx=0.48, 
            rely=0.02)
        self.entry_suframa = ttk.Spinbox(self.labelFrameNovoRegistro, from_=0, to=50)
        self.entry_suframa.place(
            relx=0.48, 
            rely=0.20, 
            relwidth=0.05, 
            relheight=0.20)
        
        ### icms
        self.lb_icms = ttk.Label(self.labelFrameNovoRegistro, text="ICMS").place(
            relx=0.54, 
            rely=0.02)
        self.entry_icms = ttk.Spinbox(self.labelFrameNovoRegistro, from_=11, to=50)
        self.entry_icms.place(
            relx=0.54, 
            rely=0.20, 
            relwidth=0.05, 
            relheight=0.20) 
        
        ### contato 
        self.lb_contato = ttk.Label(self.labelFrameNovoRegistro, text="Nome | Contato").place(
            relx=0.01, 
            rely=0.42)
        self.entry_contato = ttk.Entry(self.labelFrameNovoRegistro, )
        self.entry_contato.place(
            relx=0.01, 
            rely=0.62, 
            relwidth=0.13, 
            relheight=0.20)

        ### email
        self.lb_email = ttk.Label(self.labelFrameNovoRegistro, text="E-mail | Telefone").place(
            relx=0.15, 
            rely=0.42)
        self.entry_email = ttk.Entry(self.labelFrameNovoRegistro)
        self.entry_email.place(
            relx=0.15, 
            rely=0.62, 
            relwidth=0.20, 
            relheight=0.20)

        self.bt_email = ttk.Button(self.labelFrameNovoRegistro, text="Copiar", command=self.copy_email)
        self.bt_email.place(
            relx=0.36, 
            rely=0.62)
        
        # Botões Cadastrar
        self.bt_novo = ttk.Button(self.labelFrameNovoRegistro, text="Cadastrar", command=self.cadastrar)
        self.bt_novo.place(
            relx=0.44, 
            rely=0.60, 
            relwidth=0.15)
        
        # Botões Alterar
        self.bt_alterar = ttk.Button(self.labelFrameNovoRegistro, text="Alterar Registro", command=self.alterar)
        self.bt_alterar.place(
            relx=0.68, 
            rely=0.65, 
            relwidth=0.10)
        
        # Botões Excluir
        self.bt_apagar = ttk.Button(self.labelFrameNovoRegistro, text="Excluir Registro", command=self.deletar).place(
            relx=0.78, 
            rely=0.65, 
            relwidth=0.10)  
        
        # Botões Limpar
        self.bt_limpar = ttk.Button(self.labelFrameNovoRegistro, text="Limpar Campos", command=self.limpar)
        self.bt_limpar.place(
            relx=0.88, 
            rely=0.65, 
            relwidth=0.10)   
             
    def layout_consultar(self):

        self.labelFrameConsultarRegistro = ttk.LabelFrame(self.labelFrameNovoRegistro, text=" Consultar ")
        self.labelFrameConsultarRegistro.place(
            relx=0.68, 
            rely=0.01, 
            relwidth=0.30,
            relheight=0.50)
        
        # ENTRY PARA PESQUISAR
        self.pesquisar_entry = ttk.Entry(self.labelFrameConsultarRegistro)
        self.pesquisar_entry.place(
            relx=0.02, 
            rely=0.02, 
            relwidth=0.55,
            relheight=0.85) 
        self.pesquisar_entry.bind("<KeyRelease>", lambda event: self.filtrar())
        
        ## PESQUISAR TODOS
        btn_todos = ttk.Button(self.labelFrameConsultarRegistro, text="Mostrar Todos", command=self.mostrarTodos)
        btn_todos.place(
                relx=0.60, 
                rely=0.02,
                relwidth=0.35,
                relheight=0.85)
            
    def layout_cadastrados(self):

        self.frameCadastrados = ttk.Frame(self)
        self.frameCadastrados.place(
            relx=0.01, 
            rely=0.25, 
            relwidth=0.98, 
            relheight=0.78)
        
        self.labelFrame_Cadastrados = ttk.LabelFrame(self.frameCadastrados, text=" Registros de Fornecedores ")
        self.labelFrame_Cadastrados.place(
            relx=0.01, 
            rely=0.01, 
            relwidth=0.98, 
            relheight=0.90)
        
        self.registrados = ttk.Treeview(self.labelFrame_Cadastrados, height=3,columns=( "col1", "col2", "col3", "col4", "col5", "col6", "col7", "col8" ))
        
        self.registrados.heading("#0", text="")
        self.registrados.heading("#1", text="Cod")
        self.registrados.heading("#2", text="Produto")
        self.registrados.heading("#3", text="Marca")
        self.registrados.heading("#4", text="% Frete")
        self.registrados.heading("#5", text="% Suframa")
        self.registrados.heading("#6", text="% ICMS")
        self.registrados.heading("#7", text="Contato")
        self.registrados.heading("#8", text="Email")
        
        self.registrados.column("#0", width=0, stretch=NO)
        self.registrados.column("#1", anchor="center", width=1)
        self.registrados.column("#2", width=200)
        self.registrados.column("#3", width=100)
        self.registrados.column("#4", anchor="center", width=1)
        self.registrados.column("#5", anchor="center", width=20)
        self.registrados.column("#6", anchor="center", width=1)
        self.registrados.column("#7", width=70)
        self.registrados.column("#8", width=200)
        self.registrados.place(
            relx=0.01, 
            rely=0.02, 
            relwidth=0.97, 
            relheight=0.90)
        
        ### Barra de Rolagem Vertical
        self.scroolLista = ttk.Scrollbar(self.labelFrame_Cadastrados, orient='vertical',
                                        command=self.registrados.yview)
        
        # Configure as barras de rolagem para o TreeView
        self.registrados.configure(yscroll=self.scroolLista.set)
        self.scroolLista.place(
            relx=0.98, 
            rely=0.02, 
            relwidth=0.02, 
            relheight=0.96)
        
        # ### Barra de Rolagem Horizontal
        # self.scroolLista = ttk.Scrollbar(self.labelFrame_Cadastrados, orient='vertical',
        #                                 command=self.registrados.yview)
        
        # # Configure as barras de rolagem para o TreeView
        # self.registrados.configure(yscroll=self.scroolLista.set)
        # self.scroolLista.place(
        #     relx=0.98, 
        #     rely=0.02, 
        #     relwidth=0.02, 
        #     relheight=0.96)
        
        # Vincule eventos de rolagem às barras de rolagem
        self.registrados.bind("<Double-1>", self.OnDoubleClick)
        self.pesquisar_entry.bind("<KeyRelease>", lambda event: self.filtrar())

