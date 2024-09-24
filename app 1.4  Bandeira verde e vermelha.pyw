import os
import pandas as pd
import ttkbootstrap as ttk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook
from ttkbootstrap.constants import *
from openpyxl.drawing.image import Image


class Fatpy(ttk.Frame):

    def __init__(self, master):

        super().__init__(master, padding=(0, 0))
        self.pack(fill=BOTH, expand=YES)
        self.cancelar = False
        self.todas_colunas = [
            'Cliente', 'CNPJ', 'Local', 'Medidor', 
            'Leitura Anterior', 'Leitura Atual', 'Diferença', 
            'Fator', 'CONS.EM kWh', 'VALOR DO CONSUMO', 
            'RATEIO DEMANDA', 'Fator Demanda',
            'Fator Bandeira', 'VALOR A COBRAR'
        ]
        
        self.nome_arquivo = ttk.StringVar()
        self.pular_linha = 0
        
        # CONTROLE DE ABAS
        self.controle_abas = ttk.Notebook(self)
        self.aba1 = ttk.Frame(self.controle_abas, padding=(10, 10))
        self.aba2 = ttk.Frame(self.controle_abas, padding=(10, 10))
        self.controle_abas.add(self.aba1, text="Faturas de Energia")
        self.controle_abas.pack(expand=YES, fill="both")

        self.titulo = ttk.Frame(self.aba1)
        self.titulo.pack(pady=18, padx=10, fill="x")
        ttk.Label(
            self.titulo, text=" TAB VII - ENERGIA ").pack(side=TOP, padx=5, pady=5)
       
        # VARIAVEIS ENERGIA
        self.fatura_energia = ttk.StringVar(value="")
        self.taxa_variavel_energia = ttk.StringVar(value="")
        self.taxa_fixa_energia = ttk.StringVar(value="")
        
        # FORM TAB VII - ENERGIA
        self.campo_form_entrada("FATURA", self.fatura_energia, self.aba1)
        self.campo_form_entrada(
            "VARIAVEL", self.taxa_variavel_energia, self.aba1)
        self.campo_form_entrada("FIXA", self.taxa_fixa_energia, self.aba1)
        self.campo_form_excel(self.aba1)
        self.campo_btnbox(self.aba1, self.gerar_fatura_energia)
    
    def campo_form_entrada(self, label, variable, tabs):

        container = ttk.Frame(tabs)
        container.pack(fill=X, expand=False, pady=10)

        lbl = ttk.Label(master=container, text=label.title(), width=10)
        lbl.pack(side=LEFT, padx=5, pady=10)

        ent = ttk.Entry(master=container, textvariable=variable)
        ent.pack(side=LEFT, padx=8, fill=X, pady=10, expand=YES)

    def campo_btnbox(self, tabs, command):

        container = ttk.Frame(tabs)
        container.pack(fill=X, expand=False, pady=(15, 10))

        sub_btn = ttk.Button(
            master=container,
            text="GERAR",
            command=command,
            bootstyle=SUCCESS,
            width=10,
        )
        sub_btn.pack(side=TOP, padx=5, pady=10)
        sub_btn.focus_set()

        cnl_btn = ttk.Button(
            master=container,
            text="CANCEL",
            bootstyle=DANGER,
            width=10,
            command=self.cancelar
        )
        cnl_btn.pack(side=TOP, padx=5, pady=10)

    def mostrar_msg(self,msg):
        messagebox.showinfo("", msg)

    def cancelar(self):
        if self.cancelar:
            self.cancelar = True
    
    def tratar_excel(self,file):
        
        print("iniciando tratamento")
        
        df = pd.read_excel(file, skiprows=self.pular_linha)
        
        # verificação se existe a primeira linha 
        if df.shape[0] == 0 or df.iloc[0].isnull().all():
            self.pular_linha = 1
            return
        
        for coluna in self.todas_colunas:
            
            if df.get(coluna) is None:
                
                self.mostrar_msg(f'coluna {coluna} não foi encontrada')
        
                return False 
                
        df['VALOR A COBRAR'] = df['VALOR A COBRAR'].round(2)
        df = df[df['VALOR A COBRAR'] != 0]
        df = df.dropna(subset=['VALOR A COBRAR'])
        df['RATEIO DEMANDA'] = df['RATEIO DEMANDA'].round(2)
        df['Fator Demanda'] = df['Fator Demanda'].round(2)
        df['Fator Bandeira'] = df['Fator Bandeira'].round(2)
        
        
        df['Leitura Anterior'] = pd.to_numeric(df['Leitura Anterior'],errors='coerce')
        df['Leitura Atual'] = pd.to_numeric(df['Leitura Atual'],errors='coerce')
        
        df['Leitura Anterior'] = df['Leitura Anterior'].fillna(0)
        df['Leitura Atual'] = df['Leitura Atual'].fillna(0)

        df.to_excel(file, index=False)

        print('tratamento ------- OK ')
        
    def selecionar_arquivo(self):

        self.arquivo_selecionado = filedialog.askopenfilename(
            filetypes=[("Arquivos Excel", "*.xlsx")])
        self.nome_arquivo.set(self.arquivo_selecionado.split("/")[-1])
        #print(self.arquivo_selecionado)

        return self.arquivo_selecionado

    def selecionar_diretorio(self):

        self.caminho_diretorio = filedialog.askdirectory()
        #print(self.caminho_diretorio)
        return self.caminho_diretorio

    def campo_form_excel(self, tabs):

        container = ttk.Frame(tabs)
        container.pack(fill=X, expand=False, pady=10)

        cnl_btn = ttk.Button(
            master=container,
            text="buscar",
            command=self.selecionar_arquivo,
            bootstyle=DANGER,
            width=6,
        )
        cnl_btn.pack(side=LEFT, padx=10)

        ent = ttk.Entry(master=container, textvariable=self.nome_arquivo)
        ent.pack(side=LEFT, padx=5, fill=X, expand=YES)

    def gerar_fatura_energia(self):
        
        self.cancelar = False
        self.camimho_faturas = self.selecionar_diretorio()

        tratado = self.tratar_excel(self.arquivo_selecionado)
        
        if  tratado is False:
            print('excel não tratado')
            return
        
        df = pd.read_excel(self.arquivo_selecionado)
        
        #nm_fatura = int(self.fatura_energia.get())
        #taxa_variavel = float(self.taxa_variavel_energia.get())
        #taxa_fixa = float(self.taxa_fixa_energia.get())
        
        nm_fatura = 1
        taxa_variavel = 1.18
        taxa_fixa = 0.37

        
        total_kw = ((df['Leitura Atual'] - df['Leitura Anterior']) * df['Fator']) + (df['Fator Demanda'] + df['Fator Bandeira'])

        total_taxa_variavel = total_kw * taxa_variavel
        total_taxa_fixo = total_kw * taxa_fixa

        total_fatura =total_taxa_fixo + total_taxa_variavel
        
        #adiconar coluna calculada 
        
        df['TOTAL CALCULADO'] = total_fatura.round(2)
        
        df.to_excel(self.arquivo_selecionado,index=False)
     
        Todos_ok = True
    
        img = Image('logo_docas.jpg')

        colunas_necessarias = (zip(df['Cliente'], df['CNPJ'], df['Local'], df['Leitura Atual'],
                   df['Leitura Anterior'], df['Fator'], df['Fator Demanda'], df['VALOR A COBRAR'],df['Fator Bandeira']))
        
        print('gerando faturas..')
        
        for i, (cliente, cnpj, local, lt_atual, lt_anterior, fator, ft_demanda, vl_a_cobrar,ft_bandeira) in enumerate(colunas_necessarias, start=0):
            
            if self.cancelar:
                print('cancelado')
                return
            
            workbook_energia = load_workbook(filename='modelo energia.xlsx')
            sheet_energia = workbook_energia.active
            
            sheet_energia.add_image(img, 'C1')
            
            sheet_energia['D15'] = taxa_variavel
            sheet_energia['D30'] = taxa_fixa

            nm_fatura = nm_fatura + 1
            sheet_energia['F6'] = nm_fatura

            if df['VALOR A COBRAR'].iloc[i] < 50:

                texto_menos_50 = "3  TABELA VII ITEM 2.3"
                tx_menos_50 = 50 - df['VALOR A COBRAR'].iloc[i]

                sheet_energia['A32'] = texto_menos_50
                sheet_energia['D32'] = tx_menos_50.round(2)
                sheet_energia['E32'] = total_kw.iloc[i]
                sheet_energia['F34'] = df['VALOR A COBRAR'].iloc[i] + \
                    tx_menos_50.round(2)
                sheet_energia['F32'] = tx_menos_50.round(2)

            else:
                sheet_energia['F34'] = vl_a_cobrar

            sheet_energia['E15'] = total_kw.iloc[i]
            sheet_energia['B26'] = total_kw.iloc[i]
            sheet_energia['E30'] = total_kw.iloc[i]
            sheet_energia['F15'] = total_taxa_variavel[i]
            sheet_energia['F30'] = total_taxa_fixo[i]

            sheet_energia['B6'] = cliente
            sheet_energia['B7'] = cnpj
            sheet_energia['A16'] = local
            sheet_energia['B19'] = lt_atual
            sheet_energia['B18'] = lt_anterior
            sheet_energia['B24'] = fator
            sheet_energia['B20'] = ft_demanda
            sheet_energia['B21'] = ft_bandeira

            nome_fatura = f"FATURA {nm_fatura}.xlsx"
            caminho_salvar = os.path.join(self.camimho_faturas, nome_fatura)
            workbook_energia.save(caminho_salvar)
            
            if (df['VALOR A COBRAR'].round(1) == df['TOTAL CALCULADO'].round(1)).all():
                print('VALORES ------- OK')
                print(f'{df["VALOR A COBRAR"].iloc[i]}')
                print(f'{df["TOTAL CALCULADO"].iloc[i]}')
                
            else:
                print('VALORES --------- EROOR')
                Todos_ok = False
            
        if Todos_ok:
            self.mostrar_msg('TODOS OS VALORES ESTÃO OK') 
        else:
            self.mostrar_msg('existem valores incorrestos')
        
        print('concluido')

if __name__ == "__main__":

    app = ttk.Window("FATPY", "superhero", resizable=(
        False, False),size=(400,550))
    Fatpy(app)
    app.mainloop()
