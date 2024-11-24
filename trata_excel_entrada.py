import pandas as pd
import tkinter as tk
from tkinter import simpledialog
from tkinter import messagebox

def converter_de_coluna_para_numero(coluna):
    numero = 0
    for char in coluna:
        numero = numero * 26 + (ord(char.upper()) - ord('A') + 1)
    return numero

def receber_informacoes_do_operador(mensagem):
    root = tk.Tk()
    root.withdraw()
    return simpledialog.askstring("Entrada", mensagem)

endereco_arquivo_entrada =  receber_informacoes_do_operador("Cole aqui o endereço onde está seu arquivo de entrada: ")+"/balancete_entrada.xls"
endereco_arquivo_saida = receber_informacoes_do_operador("Cole aqui o endereço onde você deseja descarregar o arquivo tratado?")+"/balancete_saida.xlsx"
str_coluna_1 = receber_informacoes_do_operador("Informe a primeira coluna que aparece o [Cód. Rel]: ")
str_coluna_2 = receber_informacoes_do_operador("Informe a primeira coluna que aparece a [Descrição da Conta]: ")
str_coluna_3 = receber_informacoes_do_operador("Informe a primeira coluna que aparece o [Saldo]: ")
str_linha_cabecalho = receber_informacoes_do_operador("Informe o número da linha que está o cabeçalho: ")

int_coluna_1 = converter_de_coluna_para_numero(str_coluna_1)-1
int_coluna_2 = converter_de_coluna_para_numero(str_coluna_2)-1
int_coluna_3 = converter_de_coluna_para_numero(str_coluna_3)-1
int_linha_cabecalho = int(str_linha_cabecalho)-1

def tratar_excel_e_converter_novo_formato(str_arquivo_entrada, str_arquivo_saida):
    arquivo_original = pd.read_excel(str_arquivo_entrada, skiprows=int_linha_cabecalho)
    arquivo_em_tratamento = arquivo_original.iloc[:,[int_coluna_1,int_coluna_2,int_coluna_3]]
    filtro = arquivo_em_tratamento.iloc[:, 0].astype(str).str.startswith("3")
    arquivo_de_saida = arquivo_em_tratamento[filtro]
    arquivo_de_saida = arquivo_de_saida.dropna(subset=[arquivo_de_saida.columns[0]])
    arquivo_de_saida.to_excel(str_arquivo_saida)
    print(f' Arquivo gerado com sucesso! \n Verique na pasta [file_base] o arquivo chamado [balancete_saida.xlsx] ')
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("Mensagem","Arquivo [balancete_saida.xlsx] gerado com sucesso na pasta. [" + endereco_arquivo_saida +"]. Clique em ok para sair.")
    root.destroy()


tratar_excel_e_converter_novo_formato(endereco_arquivo_entrada, endereco_arquivo_saida)