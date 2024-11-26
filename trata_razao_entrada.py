import pandas as pd
from pandas.plotting import table
import matplotlib.pyplot as plt

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


def excel_to_pdf(excel_file, pdf_file):
    # Ler o arquivo Excel
    df = pd.read_excel(excel_file)
    
    # Criar uma figura para o gráfico
    fig, ax = plt.subplots(figsize=(8.3, 11.7))  # Tamanho A4 em polegadas
    ax.axis('off')  # Remover os eixos

    # Criar uma tabela a partir do DataFrame
    tabela = table(ax, df, loc='center', colWidths=[0.2]*len(df.columns))
    tabela.auto_set_font_size(False)
    tabela.set_fontsize(5)
    tabela.scale(1.0, 1.2)  # Ajustar escala para caber na página
    
    # Salvar como PDF
    plt.savefig(pdf_file, orientation="landscape", bbox_inches='tight')
    print(f"PDF salvo como: {pdf_file}")

endereco_arquivo_entrada = "C:/Projetos/conciliacao-valores/conciliacao-valores/file_base/livro_diario_entrada.xls"
endereco_arquivo_saida = "C:/Projetos/conciliacao-valores/conciliacao-valores/file_base/razao_saida.xlsx"
endereco_arquivo_saida_resumido = "C:/Projetos/conciliacao-valores/conciliacao-valores/file_base/razao_saida_resumido.xlsx"
endereco_arquivo_saida_resumido_pdf = "C:/Projetos/conciliacao-valores/conciliacao-valores/file_base/razao_saida_resumido.pdf"

def tratar_excel_e_converter_novo_formato(str_arquivo_entrada, str_arquivo_saida, str_arquivo_saida_resumido):
    
    arquivo_original = pd.read_excel(str_arquivo_entrada, skiprows=13)
    arquivo_em_tratamento = arquivo_original.iloc[:,[0,3,6,11,16,18]]
    
    filtro = arquivo_em_tratamento.iloc[:, 5].astype(str) != "-"
    arquivo_em_tratamento = arquivo_em_tratamento[filtro]
    filtro = arquivo_em_tratamento.iloc[:,2].astype(str).str.startswith("3")
    arquivo_em_tratamento = arquivo_em_tratamento[filtro]
    filtro = arquivo_em_tratamento.iloc[:, 2].astype(str) != "3080 - MÃO DE OBRA VOLUNTÁRIA - MANUT PREVENTIVA"
    arquivo_em_tratamento = arquivo_em_tratamento[filtro]
    filtro = arquivo_em_tratamento.iloc[:, 2].astype(str) != "3031 - DESPESA TARIFA CARTÃO COLETA"
    arquivo_de_saida = arquivo_em_tratamento[filtro]
    arquivo_de_saida["Data Lçto"] = pd.to_datetime(arquivo_de_saida["Data Lçto"]).dt.strftime("%d/%m/%Y")

    #arquivo_de_saida = arquivo_de_saida.dropna(subset=[arquivo_de_saida.columns[0]])
    df = pd.DataFrame(arquivo_de_saida)
    df = df.dropna(axis=0)
    arquivo_de_saida_resumido = df.groupby(["Data Lçto", "Lçto", "Conta", "Histórico", "Centro Custo"], as_index=False)["Débito"].sum()
   

    #Gera o arquivo em Excel - Detalhado
    arquivo_de_saida.to_excel(str_arquivo_saida)

    #Gera o arquivo em Excel - Resumido
    arquivo_de_saida_resumido.to_excel(str_arquivo_saida_resumido)

    print(f' Arquivo gerado com sucesso! \n Verique na pasta [file_base] o arquivo chamado [razao_saida.xlsx] ')
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("Mensagem","Arquivo [razao_saida.xlsx] gerado com sucesso na pasta. [" + endereco_arquivo_saida +"]. Clique em ok para sair.")
    root.destroy()


tratar_excel_e_converter_novo_formato(endereco_arquivo_entrada, endereco_arquivo_saida, endereco_arquivo_saida_resumido)
excel_to_pdf(endereco_arquivo_saida_resumido, endereco_arquivo_saida_resumido_pdf)
