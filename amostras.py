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

nomeArquivoAmostras = 'regs_amostras.db'

class Dados_Amostras():
    
    def sqlToExcel():

        resposta = messagebox.askquestion("Confirmação", "Deseja Exportação os Registros para o Excel?")
        if resposta == "yes":
            
            # Crie uma pasta para salvar os arquivos XLSX (se não existir)
            output_folder = 'exporta_xlsx'
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)

            # Obter a data atual
            data_atual = datetime.now().strftime("%d%m%y")
            nome_arquivo_xlsx = f"amostrasExcel_{data_atual}.xlsx"
            pasta_destino = "exporta_xlsx"
            caminho_completo = f"{pasta_destino}/{nome_arquivo_xlsx}"

            conexao = sqlite3.connect(nomeArquivoAmostras)
            c =  conexao.cursor()
        
            c.execute("SELECT * FROM registros")
            tabelaSql = c.fetchall()
            tabelaSql = pd.DataFrame(tabelaSql, columns=['Código','Tipo','Nº','Orgão %','Descrição','Empresa','Código Rastreamento','Valor R$','Entregue Data','Solicitada Data','Entregar Data','Itens Participando','Observações'])
            tabelaSql.to_excel(caminho_completo, index=False)       

            messagebox.showinfo("Operação Concluída!" ,"Acesso o diretório do programa para obter o arquivo 'amostrasExcel.xlsx'.")
        else:
            messagebox.showinfo("Cancela!" ,"Exportação Cancelada!")  

    def backupSqliteDados():
        resposta = messagebox.askquestion("Confirmação", "Deseja Fazer o Backup dos Registros de Amostras")
        
        # Crie uma pasta para salvar os arquivos XLSX (se não existir)
        output_folder = 'backups_sqlite'
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        # Obter a data atual
        data_atual = datetime.now().strftime("%d%m%Y")
        nome_novo_arquivo = f"regs_amostras_{data_atual}.db"
               
        if resposta == "yes":
            caminhoEntradaSqlDados = "regs_amostras.db"
            try:
                shutil.copy(caminhoEntradaSqlDados, os.path.join(output_folder, nome_novo_arquivo))
                messagebox.showinfo("Concluído!" ,"Backup Criado!")
            except Exception as e:
                messagebox.showinfo(message=f"Erro ao criar o backup: {str(e)}")
        else:
            messagebox.showinfo("Cancelado!" ,"Backup Cancelado!")

class Funcs_Amostras():
        
    def limpar(self): ## Limpar os Campos
        self.entry_cod.delete(0, END)
        self.OptionMenuTipo.delete(0, END)
        self.OptionMenuOrgao.delete(0, END)
        self.OptionMenuEmpresa.delete(0, END)
        self.entry_observacao.delete(0, END)
        self.entry_numero.delete(0, END)
        self.dateEntry_dataParaEntrega.delete(0, END)
        self.entry_descricao.delete(0, END)
        self.dateEntry_dataDaEntrega.delete(0, END)
        self.dateEntry_dataSolicitada.delete(0, END)
        self.entry_codRastreamento.delete(0, END)
        self.entry_valor.delete(0, END)
        self.entry_item.delete(0, END)

    def conecta_bd(self):
        self.conn = sqlite3.connect(nomeArquivoAmostras)
        self.cursor = self.conn.cursor()

    def desconecta_bd(self):
        self.conn.close(); 
        print('Desconectando ao Banco de dados...')
        
    def criar_tabela(self):
        self.conecta_bd()
        self.cursor.execute(""" CREATE TABLE IF NOT EXISTS registros (
                            cod INTEGER PRIMARY KEY,
                            tipo TEXT,
                            num TEXT,
                            org TEXT,
                            desc TEXT,
                            emp TEXT,
                            codRast TEXT,
                            valor REAL,
                            dataEntreg TEXT,
                            dataSolic TEXT,
                            entregData TEXT,
                            item TEXT,
                            obs TEXT);
                        """)
        self.conn.commit(); 
        print("Banco de dados criado...")
        self.desconecta_bd()

    def variaveis(self):
        self.cod_var = self.entry_cod.get()
        self.tipo_var = self.OptionMenuTipo.get()
        self.org_var = self.OptionMenuOrgao.get()
        self.empresa_var = self.OptionMenuEmpresa.get()
        self.dataParaEntregar_var = self.dateEntry_dataParaEntrega.get()
        self.dataSolicitada_var = self.dateEntry_dataSolicitada.get()
        self.dataDaEntrega_var = self.dateEntry_dataDaEntrega.get()
        self.obs_var = self.entry_observacao.get()
        self.num_var = self.entry_numero.get()
        self.desc_var = self.entry_descricao.get()
        self.codRast_var = self.entry_codRastreamento.get()
        self.valor_var = self.entry_valor.get()
        self.item_var = self.entry_item.get()
    
    def cadastrar(self): ## NOVOS CADASTROS
        self.variaveis()

        if self.entry_descricao.get()  == "":
            messagebox.showinfo(title="Erro de operação!" ,
                                message="Digite a Descrição do Produto para Cadastrar!")
            self.entry_descricao.focus()
        else:
            self.conecta_bd()
            self.cursor.execute(
                """ INSERT INTO registros 
                    (tipo, num, org, desc, emp, codRast, 
                    valor, dataEntreg, dataSolic, entregData, item, obs) 
                    VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) """,
                        (self.tipo_var, self.num_var, self.org_var, self.desc_var, self.empresa_var, self.codRast_var, self.valor_var, 
                         self.dataParaEntregar_var, self.dataSolicitada_var, self.dataDaEntrega_var, self.item_var, self.obs_var))
            self.conn.commit()   
            self.desconecta_bd()
            self.selecionar()
            self.limpar()
        
    def selecionar(self): ## Selecionar NA Lista
        self.layout_consultar()
        self.registrados.delete(*self.registrados.get_children())
        self.conecta_bd()
        
        lista = self.cursor.execute(""" SELECT cod, 
            tipo, num, org, desc, emp, codRast, valor, dataEntreg, dataSolic, 
            entregData, item, obs FROM registros ORDER BY desc ASC; """)

        for i in lista:
            self.registrados.insert("", END, values=i)
            
        self.desconecta_bd()

    def filtrar(self):
                
        texto_pesquisa = self.pesquisar_entry.get()

        for row in self.registrados.get_children():
            self.registrados.delete(row)

        self.conn = sqlite3.connect(nomeArquivoAmostras)
        self.cursor = self.conn.cursor()
        
        # Consulta SQL para buscar dados com base no critério de pesquisa
        # query = f"SELECT * FROM registros WHERE produto LIKE '%{texto_pesquisa}%'"
        query = f"SELECT * FROM registros WHERE num LIKE '%{texto_pesquisa}%' OR desc LIKE '%{texto_pesquisa}%' OR org LIKE '%{texto_pesquisa}%'"

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
            """ SELECT * FROM registros WHERE desc LIKE '%s' ORDER BY desc ASC """ % nome)
        buscanome = self.cursor.fetchall()
        
        for i in buscanome:
            self.registrados.insert("", END, values=i)
            
        self.pesquisar_entry.delete(0, END)
        self.desconecta_bd() 
  
    def OnDoubleClick(self, event):
        self.limpar()  
        self.registrados.selection()
        for n in self.registrados.selection():
            col1, col2, col3, col4, col5, col6, col7, col8, col9, col10, col11, col12, col13 = self.registrados.item(n, 'values')
            self.entry_cod.insert(END, col1)
            self.OptionMenuTipo.insert(END, col2)
            self.entry_numero.insert(END, col3)
            self.OptionMenuOrgao.insert(END, col4)
            self.entry_descricao.insert(END, col5)
            self.OptionMenuEmpresa.insert(END, col6)
            self.entry_codRastreamento.insert(END, col7)
            self.entry_valor.insert(END, col8)
            self.dateEntry_dataParaEntrega.insert(END, col9)
            self.dateEntry_dataSolicitada.insert(END, col10)
            self.dateEntry_dataDaEntrega.insert(END, col11)
            self.entry_item.insert(END, col12)          
            self.entry_observacao.insert(END, col13)

    def deletar(self): ## Deletar Cliente
        self.variaveis()
        
        if self.entry_cod.get()  == "":
            messagebox.showinfo(
                title="Erro de operação!" ,
                message="Digite o número 'Cod' do Registro no campo 'Cod' ou dê um click duplo no Registro que deseja excluir e aperte no botão Excluir!")
            self.entry_cod.focus()
        else: 
            verify = messagebox.askyesno(
                title="Excluir Registro?", 
                message="Deseja Excluir esse Registro?\n=" + self.desc_var)
            if  verify == True:  
                self.conecta_bd()
                self.cursor.execute(""" DELETE FROM registros  WHERE cod = ? """, ([self.cod_var]))
                self.conn.commit()
                self.desconecta_bd()
                self.limpar()
                self.selecionar()
            else:
                messagebox.showinfo(message="Operação Cancelada!")
   
    def alterar(self): ## Alterar Cliente
        self.variaveis()
        
        if self.entry_descricao.get()  == "":
            messagebox.showinfo(
                title="Erro de operação!" ,
                message="Escolha um registro já cadastrado para Alterar as Informações!")
        else:
            self.conecta_bd()
            self.cursor.execute(""" UPDATE registros SET tipo = ?, num = ?, org = ?, desc = ?,
                                emp = ?, codRast = ?, valor = ?, dataEntreg = ?, dataSolic = ?, 
                                entregData = ?, item = ?, obs = ? WHERE cod = ? """, 
                                (self.tipo_var, self.num_var, self.org_var, self.desc_var, self.empresa_var, self.codRast_var, self.valor_var, 
                                 self.dataParaEntregar_var, self.dataSolicitada_var, self.dataDaEntrega_var, self.item_var, self.obs_var, self.cod_var))
            self.conn.commit() 
            self.desconecta_bd()
            self.selecionar()
            self.limpar()

class AmostrasScreen(tk.Frame, Funcs_Amostras):

    def __init__(self, master=None):
        super().__init__(master)

        self.layout()
        self.layout_pesquisar()
        self.layout_consultar()
        self.criar_tabela()
        self.selecionar()

    def layout(self):

        self.frameNovoRegs = ttk.Frame(self)
        self.frameNovoRegs.place(
            relx=0.01, 
            rely=0.02, 
            relwidth=0.98, 
            relheight=0.22)
    
        #LabelFrame
        self.labelFrameNovoRegistro = ttk.LabelFrame(self.frameNovoRegs, text=" Cadastar Amostra ")
        self.labelFrameNovoRegistro.place(
            relx=0.01, 
            rely=0.01, 
            relwidth=0.98, 
            relheight=0.98)
        
        ## código
        ttk.Label(self.labelFrameNovoRegistro, text="cod" ).place(
            relx=0.01,
            rely=0.01)
        self.entry_cod = ttk.Entry(self.labelFrameNovoRegistro)
        self.entry_cod.place(
            relx=0.01, 
            rely=0.20, 
            relwidth=0.03, 
            relheight=0.20)

        ### Tipo
        ttk.Label(self.labelFrameNovoRegistro, text="Tipo").place(
            relx=0.05, 
            rely=0.01)
        self.optionsTipos = ["PE","PP","TP",]
        self.OptionMenuTipo = ttk.Combobox(self.labelFrameNovoRegistro, values=self.optionsTipos)
        self.OptionMenuTipo.set("Escolher")
        self.OptionMenuTipo.place(
            relx=0.05, 
            rely=0.20, 
            relwidth=0.06, 
            relheight=0.20)

        ### Orgão
        ttk.Label(self.labelFrameNovoRegistro, text="Orgão").place(
            relx=0.12, 
            rely=0.01)
        self.optionsOrgao = ["GOV AM","PREF MANAUS","COMPRASNET"] 
        self.OptionMenuOrgao = ttk.Combobox(self.labelFrameNovoRegistro, value=self.optionsOrgao)
        self.OptionMenuOrgao.set('Escolher')
        self.OptionMenuOrgao.place(
            relx=0.12, 
            rely=0.20, 
            relwidth=0.08, 
            relheight=0.20)
        
        ### Empresa
        ttk.Label(self.labelFrameNovoRegistro, text="Empresa").place(
            relx=0.21, 
            rely=0.01)
        self.optionsEmpresa = ["FARMA","COMÉRCIO",] 
        self.OptionMenuEmpresa = ttk.Combobox(self.labelFrameNovoRegistro, value=self.optionsEmpresa)
        self.OptionMenuEmpresa.set('Escolher')
        self.OptionMenuEmpresa.place(
            relx=0.21, 
            rely=0.20, 
            relwidth=0.07, 
            relheight=0.20)

        ## Observações
        ttk.Label(self.labelFrameNovoRegistro, text="Observações").place(
            relx=0.29, 
            rely=0.01)
        self.entry_observacao = ttk.Entry(self.labelFrameNovoRegistro)
        self.entry_observacao.place(
            relx=0.29, 
            rely=0.20, 
            relwidth=0.13, 
            relheight=0.20)
        
        ## Solicitada
        ttk.Label(self.labelFrameNovoRegistro, text="Solicitada").place(
            relx=0.43,
            rely=0.01)
        self.dateEntry_dataSolicitada = ttk.Entry(self.labelFrameNovoRegistro)
        self.dateEntry_dataSolicitada.place(
            relx=0.43, 
            rely=0.20, 
            relwidth=0.06, 
            relheight=0.20)

        ## Código de Rastreamento
        ttk.Label(self.labelFrameNovoRegistro, text="Cód Rastreamento").place(
            relx=0.50, 
            rely=0.01)
        self.entry_codRastreamento = ttk.Entry(self.labelFrameNovoRegistro)
        self.entry_codRastreamento.place(
            relx=0.50, 
            rely=0.20, 
            relwidth=0.10, 
            relheight=0.20)       
        
        ## Número
        ttk.Label(self.labelFrameNovoRegistro, text="Nº Pregão").place(
            relx=0.01,
            rely=0.42)
        self.entry_numero = ttk.Entry(self.labelFrameNovoRegistro)
        self.entry_numero.place(
            relx=0.01, 
            rely=0.62, 
            relwidth=0.05, 
            relheight=0.20)
        
        ## DATA/DIA PARA ENTREGAR 
        ttk.Label(self.labelFrameNovoRegistro, text="Entregar").place(
            relx=0.07,
            rely=0.42)
        self.dateEntry_dataParaEntrega = ttk.Entry(self.labelFrameNovoRegistro)
        self.dateEntry_dataParaEntrega.place(
            relx=0.07, 
            rely=0.62, 
            relwidth=0.06, 
            relheight=0.20)

        ## Descrição
        ttk.Label(self.labelFrameNovoRegistro, text="Descrição do Item").place(
            relx=0.14, 
            rely=0.42)
        self.entry_descricao = ttk.Entry(self.labelFrameNovoRegistro)
        self.entry_descricao.place(
            relx=0.14, 
            rely=0.62, 
            relwidth=0.15, 
            relheight=0.20)
        
        ## DATA/DIA PARA ENTREGAR
        ttk.Label(self.labelFrameNovoRegistro, text="Entregue").place(
            relx=0.30,
            rely=0.42)
        self.dateEntry_dataDaEntrega = ttk.Entry(self.labelFrameNovoRegistro)
        self.dateEntry_dataDaEntrega.place(
            relx=0.30, 
            rely=0.62, 
            relwidth=0.06, 
            relheight=0.20)

        ## Itens
        ttk.Label(self.labelFrameNovoRegistro, text="Participando do Itens").place(
            relx=0.37,
            rely=0.42)
        self.entry_item = ttk.Entry(self.labelFrameNovoRegistro)
        self.entry_item.place(
            relx=0.37, 
            rely=0.62, 
            relwidth=0.15, 
            relheight=0.20)    
    
        ## Valor
        ttk.Label(self.labelFrameNovoRegistro, text="Valor R$").place(
            relx=0.53,
            rely=0.42)
        self.entry_valor = ttk.Entry(self.labelFrameNovoRegistro)
        self.entry_valor.place(
            relx=0.53, 
            rely=0.62, 
            relwidth=0.05, 
            relheight=0.20)
        
        ## Botões Cadastrar
        self.bt_novo = ttk.Button(
            self.labelFrameNovoRegistro, text="Cadastrar", command=self.cadastrar).place(
            relx=0.59, 
            rely=0.60)
        
        ## Botões Alterar
        self.bt_alterar = ttk.Button(
            self.labelFrameNovoRegistro, text="Alterar", command=self.alterar).place(
                relx=0.68, 
                rely=0.65, 
                relwidth=0.100)
                    
        # Botões Excluir
        self.bt_apagar = ttk.Button(self.labelFrameNovoRegistro, text="Exclui Registro", command=self.deletar).place(
            relx=0.78, 
            rely=0.65,
            relwidth=0.10)

        ## Botões Limpar
        self.bt_limpar = ttk.Button(
            self.labelFrameNovoRegistro, text="Limpar", command=self.limpar).place(
            relx=0.88, 
            rely=0.65, 
            relwidth=0.10)             
    
    def layout_pesquisar(self):

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
        btn_todos = ttk.Button(self.labelFrameConsultarRegistro, text="Mostrar todos", command=self.mostrarTodos).place(
                                relx=0.60, 
                                rely=0.02,
                                relwidth=0.35,
                                relheight=0.85)
        
    def layout_consultar(self):
    
        self.frameCadastradosRegistros = ttk.Frame(self)
        self.frameCadastradosRegistros.place(
            relx=0.01, 
            rely=0.25, 
            relwidth=0.98, 
            relheight=0.78)
        
        self.labelFrameCadastradosRegistros = ttk.LabelFrame(self.frameCadastradosRegistros, text=" Registros de Amostras ")
        self.labelFrameCadastradosRegistros.place(
            relx=0.01, 
            rely=0.01, 
            relwidth=0.98, 
            relheight=0.90)
        
        self.registrados = ttk.Treeview(
            self.labelFrameCadastradosRegistros, 
            columns=("col1","col2", "col3", "col4", "col5","col6", "col7", "col8", "col9", "col10","col11", "col12", "col13" ))
    
        self.registrados.heading("#0",)
        self.registrados.heading("#1", text="Cod")
        self.registrados.heading("#2", text="Tipo")
        self.registrados.heading("#3", text="Nº")
        self.registrados.heading("#4", text="Orgão")
        self.registrados.heading("#5", text="Descrição")
        self.registrados.heading("#6", text="Empresa")
        self.registrados.heading("#7", text="Cód Rastreamento")
        self.registrados.heading("#8", text="Valor")
        self.registrados.heading("#9", text="Entregar")
        self.registrados.heading("#10", text="Solicitada")
        self.registrados.heading("#11", text="Entregue")
        self.registrados.heading("#12", text="Itens")
        self.registrados.heading("#13", text="Obs") 

        self.registrados.column("#0",  width=0, stretch=NO)
        self.registrados.column("#1", anchor="center", width=2)
        self.registrados.column("#2", anchor="center", width=2)
        self.registrados.column("#3", anchor="center", width=2)
        self.registrados.column("#4", anchor="center", width=5)
        self.registrados.column("#5", width=100)
        self.registrados.column("#6",  anchor="center", width=5)
        self.registrados.column("#7",  anchor="center", width=50)
        self.registrados.column("#8", anchor="center", width=7)
        self.registrados.column("#9", anchor="center", width=10)
        self.registrados.column("#10", anchor="center", width=10)
        self.registrados.column("#11", anchor="center", width=10)
        self.registrados.column("#12", width=5)
        self.registrados.column("#13", width=50)
        self.registrados.place(
            relx=0.01, 
            rely=0.02, 
            relwidth=0.97, 
            relheight=0.90)
        self.registrados.bind("<Double-1>", self.OnDoubleClick)

        ### Barra de Rolagem Horizontal
        self.scroolLista = ttk.Scrollbar(self.labelFrameCadastradosRegistros, orient='vertical',
                                        command=self.registrados.yview)
        
        # Configure as barras de rolagem para o TreeView
        self.registrados.configure(yscroll=self.scroolLista.set)
        self.scroolLista.place(
            relx=0.98, 
            rely=0.02, 
            relwidth=0.02, 
            relheight=0.96)
        
        # Vincule eventos de rolagem às barras de rolagem
        self.registrados.bind("<Double-1>", self.OnDoubleClick)
        self.pesquisar_entry.bind("<KeyRelease>", lambda event: self.filtrar())


