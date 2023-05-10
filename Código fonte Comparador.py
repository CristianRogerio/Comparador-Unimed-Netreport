import tkinter as tk
import pandas as pd
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import shutil
import os
import tempfile
import customtkinter
import subprocess
from PIL import Image, ImageTk


def juntar_linha(linha):
    return ','.join(map(str, linha))
def procurar_arquivo_Unimed():

    arquivo = filedialog.askopenfilename(initialdir='/',title=" Selecione um arquivo",filetypes=(("Arquivos Excel","*.xls"),('Arquivos Excel', '*.xlsx')))
    if arquivo:
        if arquivo.endswith(".xlsx"):
            os.rename(arquivo, os.path.splitext(arquivo)[0] + ".xlsx")
            novo_nome_arquivo = "Relatorio Unimed.xlsx"
            pasta_temp = tempfile.gettempdir()
            novo_caminho_arquivo = os.path.join(pasta_temp, novo_nome_arquivo)
            shutil.copy(arquivo, novo_caminho_arquivo)
            button1.configure(fg_color='#2e8b57')
            lista_arquivos.insert(tk.END, arquivo)
        elif arquivo.endswith('.xls'):
            df = pd.read_excel(arquivo)
            new_file_path = os.path.splitext(arquivo)[0] + ".xlsx"
            df.to_excel(new_file_path, index=False)
            novo_nome_arquivo = "Relatorio Unimed.xlsx"
            pasta_temp = tempfile.gettempdir()
            novo_caminho_arquivo = os.path.join(pasta_temp, novo_nome_arquivo)
            shutil.copy(new_file_path, novo_caminho_arquivo)
            button1.configure(fg_color='#2e8b57')
            lista_arquivos.insert(tk.END, arquivo)
        else:
            messagebox.showerror("Erro", "Formato de arquivo inválido!")
    else:
        # Exibe uma mensagem se o arquivo não for selecionado corretamente
        messagebox.showerror("Erro", "Nenhum arquivo selecionado.")
def procurar_arquivo_Netreport():

    arquivo = filedialog.askopenfilename(initialdir='/',title=" Selecione um arquivo",filetypes=(("Arquivos Excel","*.xls"),('Arquivos Excel', '*.xlsx')))
    if arquivo:
        if arquivo.endswith(".xlsx"):
            os.rename(arquivo, os.path.splitext(arquivo)[0] + ".xlsx")
            novo_nome_arquivo = "Relatorio Netreport.xlsx"
            pasta_temp = tempfile.gettempdir()
            novo_caminho_arquivo = os.path.join(pasta_temp, novo_nome_arquivo)
            shutil.copy(arquivo, novo_caminho_arquivo)
            button2.configure(fg_color='#2e8b57')
            lista_arquivos.insert(tk.END, arquivo)
        elif arquivo.endswith('.xls'):
            df = pd.read_excel(arquivo)
            new_file_path = os.path.splitext(arquivo)[0] + ".xlsx"
            df.to_excel(new_file_path, index=False)
            novo_nome_arquivo = "Relatorio Netreport.xlsx"
            pasta_temp = tempfile.gettempdir()
            novo_caminho_arquivo = os.path.join(pasta_temp, novo_nome_arquivo)
            shutil.copy(new_file_path, novo_caminho_arquivo)
            button2.configure(fg_color='#2e8b57')
            lista_arquivos.insert(tk.END, arquivo)
        else:
            messagebox.showerror("Erro", "Formato de arquivo inválido!")
    else:
        # Exibe uma mensagem se o arquivo não for selecionado corretamente
        messagebox.showerror("Erro", "Nenhum arquivo selecionado.")
def executar_comparacao(caixa_texto):
    try:
        # Manipulando Planilha do relário do Unimed--------------------------------------------------------------------------
        global dfresult, dfUnimedComparador
        temp_path = tempfile.gettempdir()
        if os.path.exists(temp_path + "\\Relatorio Unimed.xlsx"):
            df = pd.read_excel(temp_path+'\\Relatorio Unimed.xlsx')
            matricula = df['Nome Beneficiário'].copy()
            df["Matrícula"] = matricula
            df = df.rename(columns={'Nome Beneficiário': 'Nome'})
            df = df.rename(columns={'Valor Item Lib. Pagto': 'Valor'})
            df = df.rename(columns={'Cód. Item': 'Cód.Serviço'})
            df['Matrícula'] = df['Matrícula'].replace('[A-Z]', '', regex=True)
            df['Matrícula'] = df['Matrícula'].replace('[-]', '', regex=True)
            df['Matrícula'] = df['Matrícula'].replace('[Ç]', '', regex=True)
            df['Nome'] = df['Nome'].replace('[0-9]', '', regex=True)
            df['Nome'] = df['Nome'].replace('[-]', '', regex=True)
            df['Guia'] = df['Guia'].replace('[-]', '', regex=True)
            dfUnimedOrg = df[['Nome', 'Guia', 'Matrícula', 'Cód.Serviço', 'Valor']]
            dfUnimedOrg = dfUnimedOrg.sort_values('Nome', ignore_index=True)


            #Exportar Relatorio Unimed Organizado
            # dfUnimedOrg.to_excel("Relatorio Unimed Organizado.xlsx")

            # Agrupando Df pelo nome e somando os exames dos Pacientes

            dfUnimedOrganizado = dfUnimedOrg.groupby('Nome')['Valor'].sum().reset_index()
            dfUnimedOrganizado['Valor'] = dfUnimedOrganizado['Valor'].astype(float).round(2)

            dfUnimedOrganizado['Comparador'] = dfUnimedOrganizado.apply(juntar_linha, axis=1)
            colunaJunto = ['Comparador']
            dfUnimedComparador = dfUnimedOrganizado.drop(columns=[col for col in dfUnimedOrganizado.columns if col not in colunaJunto])

            # Gerando DataFrame para comparação
            # dfUnimedComparador.to_excel("Relatorio Unimed Comparador.xlsx")
        else:
            messagebox.showerror("Erro", "Planilha Unimed não Selecionada.")
        # Manipulando Planilha do relário do Net Report--------------------------------------------------------------------------
        if os.path.exists(temp_path + "\\Relatorio Netreport.xlsx"):
            df2 = pd.read_excel(temp_path+'\\Relatorio Netreport.xlsx', header=None)
            df2 = df2.iloc[:-1]
            df2.drop([0, 1, 2, 3, 4, 5, 6, 7], axis=0, inplace=True)
            df2.drop([0, 1, 2, 3, 5, 7, 8, 10, 11, 12, 13, 15, 16, 17], axis=1, inplace=True)
            df2 = df2.rename(columns={4: 'Matrícula', 6: 'Nome', 9: 'Cód.Serviço', 14: 'Guia', 18: 'Valor'})
            df2 = df2[['Nome', 'Guia', 'Matrícula', 'Cód.Serviço', 'Valor']]
            df2['Nome'] = df2['Nome'].str.strip()
            df2 = df2.sort_values('Nome', ignore_index=True)
            df2['Valor'] = df2['Valor'].astype(float).round(2)
            # df2.to_excel("Relatorio Netreport Organizado.xlsx")

            # Salvando Planilha Relatorio Netreport Organizado

            dfNetreportOrganizado = df2.groupby('Nome')['Valor'].sum().reset_index()
            # dfNetreportOrganizado.to_excel("Relatorio Netreport Nomes Somados.xlsx")
            dfNetreportOrganizado['Valor'] = dfNetreportOrganizado['Valor'].astype(float).round(2)

            dfNetreportOrganizado['Comparador'] = dfNetreportOrganizado.apply(juntar_linha, axis=1)
            colunaJunto = ['Comparador']
            dfNetreportComparador = dfNetreportOrganizado.drop(
                columns=[col for col in dfNetreportOrganizado.columns if col not in colunaJunto])

            # Gerando DataFrame para comparação
            # dfNetreportComparador.to_excel("Relatorio Netreport Comparador.xlsx")

            # Comparando planilhas ------------------------------------------------------------------------------------------------

            comparadorUnimedNetrepot = pd.merge(dfUnimedComparador[['Comparador']], dfNetreportComparador[['Comparador']],how='inner', on='Comparador')
            # comparadorUnimedNetrepot.to_excel("Corretos.xlsx")
    #Nomes que estao no netreport e há divergencia na Unimed***************************

            NomesNaoencontrados = pd.merge(dfNetreportComparador, comparadorUnimedNetrepot, how='outer', on='Comparador',indicator=True)
            NomesNaoencontrados = NomesNaoencontrados.loc[NomesNaoencontrados['_merge'] == 'left_only', 'Comparador']
            NomesNaoencontrados.to_excel("Inconsistencias.xlsx")
            df = pd.read_excel("Inconsistencias.xlsx")
            df['Comparador'] = df['Comparador'].replace('[0-9]', '', regex=True)
            df['Comparador'] = df['Comparador'].replace('[,]', '', regex=True)
            df['Comparador'] = df['Comparador'].replace('[.]', '', regex=True)
            df = df.rename(columns={'Comparador': "Nome"})
            dfresult = df[['Nome']]
            dfresult = dfresult.rename(columns={"Nome": 'Nomes com Divergência entre Netreport/Unimed'})
            n = len(dfresult)
            new_index = range(1, n + 1)
            dfresult = dfresult.set_index(pd.Index(new_index))
            df_string = dfresult.to_string()
    #Divergencia Nomes encontrados na Unimed que não Exitem na planilha Net Report
            '''
            comparadorNetrepotUnimed = pd.merge(dfNetreportComparador[['Comparador']], dfUnimedComparador[['Comparador']],how='inner', on='Comparador')
            NomesNaoencontrados = pd.merge(comparadorNetrepotUnimed, dfNetreportComparador, how='outer', on='Comparador',indicator=True)
            NomesNaoencontrados = NomesNaoencontrados.loc[NomesNaoencontrados['_merge'] == 'left_only', 'Comparador']
            NomesNaoencontrados.to_excel("NomesAmaisUnimed.xlsx")
            df = pd.read_excel("NomesAmaisUnimed.xlsx")
            df['Comparador'] = df['Comparador'].replace('[0-9]', '', regex=True)
            df['Comparador'] = df['Comparador'].replace('[,]', '', regex=True)
            df['Comparador'] = df['Comparador'].replace('[.]', '', regex=True)
            df = df.rename(columns={'Comparador': "Nome"})
            dfExistUnimed = df[['Nome']]
            dfExistUnimed = dfExistUnimed.rename(columns={"Nome": 'Nomes encontrados Somente na Unimed'})
            n = len(dfresult)
            new_index = range(1, n + 1)
            dfresult= pd.concat([dfresult, dfExistUnimed], ignore_index=True)
            dfresult = dfresult.set_index(pd.Index(new_index))
            df_string = dfresult.to_string()
            '''
            caixa_texto.delete("1.0", tk.END)
            dfresult.to_excel(temp_path + "\\Inconsistencias.xlsx")

            linhas = df_string.split("\n")

            caixa_texto.tag_configure("Cinza", background="#282725")
            caixa_texto.tag_configure("Cinzaescuro", background="#535251")

            for i, linha in enumerate(linhas):
                if i % 2 == 0:
                    caixa_texto.insert("end", linha + "\n", "Cinza")
                else:
                    caixa_texto.insert("end", linha + "\n", "Cinzaescuro")
            '''
            for index, row in df.iterrows():
                linha = f'{row["Divergencias Unimed"]}\t{row["Nomes encontrados Somente na Unime"]}\n'
                caixa_texto.insert('end', linha)'''

            button3.configure(fg_color='#2e8b57')
            button5.configure(fg_color='#2e8b57')
        else:
            messagebox.showerror("Erro", "Planilha Netreport não Selecionada.")
    except :
        messagebox.showerror("Erro entre Planilhas ","Por favor verificar as planilhas selecionadas")

def limpar_comparacao():

    lista_arquivos.selection_clear(0, tk.END)
    for widget in app.winfo_children():
        # Se o widget for um ScrolledText, limpa o seu conteúdo
        if isinstance(widget, tk.Text):
            widget.delete("1.0", tk.END)
    button1.configure(fg_color='#3a7ebf')
    button2.configure(fg_color='#3a7ebf')
    button3.configure(fg_color='#3a7ebf')
    button5.configure(fg_color='#3a7ebf')

    temp_path = tempfile.gettempdir()
    if os.path.exists(temp_path + "\\Inconsistencias.xlsx"):
        os.remove(temp_path + "\\Inconsistencias.xlsx")

    if os.path.exists(temp_path + "\\Relatorio Netreport.xlsx"):
        os.remove(temp_path + "\\Relatorio Netreport.xlsx")

    if os.path.exists(temp_path + "\\Relatorio Unimed.xlsx"):
        os.remove(temp_path + "\\Relatorio Unimed.xlsx")


    caixa_texto.insert('1.0', '\n\n\n\n\n\n\n\n\nComparação Reiniciada com sucesso!:)')
    caixa_texto.tag_configure('center', justify='center')
    caixa_texto.tag_add('center', '1.0', 'end')
def deletar_arquivo():
    temp_path = tempfile.gettempdir()
    if os.path.exists(temp_path + "\\Inconsistencias.xlsx"):
        os.remove(temp_path + "\\Inconsistencias.xlsx")
        app.destroy()  # destrói a janela principal
    else:
        app.destroy()  # destrói a janela principal

    if os.path.exists(temp_path + "\\Relatorio Netreport.xlsx"):
        os.remove(temp_path + "\\Relatorio Netreport.xlsx")

    if os.path.exists(temp_path + "\\Relatorio Unimed.xlsx"):
        os.remove(temp_path + "\\Relatorio Unimed.xlsx")
def block_edit(event):
    return "break"
def exportar_resultado():
    try:
        temp_path = tempfile.gettempdir()
        if os.path.exists(temp_path + "\\Inconsistencias.xlsx"):

            filename = filedialog.asksaveasfilename(initialfile='Resultado Comparação.xlsx', defaultextension='.xlsx',filetypes=[("Arquivos Excel", "*.xlsx")])
            temp_path = os.path.join(os.environ['TEMP'], 'Inconsistencias.xlsx')

            shutil.copy(temp_path, filename)


            subprocess.Popen(r'explorer /select,"{}"'.format(filename.replace('/', '\\')))

        else:
            messagebox.showerror("Erro", "Ainda Não foi feita a Comparação!.")
    except Exception:
        messagebox.showerror("Erro", "O arquivo não foi Salvo Corretamente")
def reset_selection(event):
    caixa_texto.tag_remove('sel', '1.0', 'end')
def selecionar_com_cor(event):
    # Verifica se há texto selecionado
    if caixa_texto.tag_ranges('sel'):
        # Remove a seleção atual
        caixa_texto.tag_remove('sel', '1.0', 'end')
    # Obtém a posição inicial e final da seleção
    posicao_inicial = caixa_texto.index('sel.first')
    posicao_final = caixa_texto.index('sel.last')
    # Adiciona a nova tag para personalizar a cor da seleção
    caixa_texto.tag_add('selecao', posicao_inicial, posicao_final)
def pular_linha(coluna,linha):
    empty_space = tk.Label(app, text="")
    empty_space.config(font=("Arial",1), foreground="#1a1a1a", background="#1a1a1a")
    empty_space.grid(column=coluna,row=linha)


customtkinter.set_default_color_theme("dark-blue")
customtkinter.set_appearance_mode("dark")


app = customtkinter.CTk()  # Criar Janela
app.geometry("550x654") #Dimenções da Janela
screen_width = app.winfo_screenwidth()
screen_height = app.winfo_screenheight()
# Definindo a posição da janela para o centro da tela
x = int((screen_width - app.winfo_reqwidth()) / 3)
y = int((screen_height - app.winfo_reqheight()) / 150)
app.geometry("+{}+{}".format(x, y))

app.title("Comparador Unimed/Netreport 2.0")
app.iconbitmap('Comparador.ico')
app.grid_columnconfigure(0, weight=1)
app.grid_columnconfigure(1, weight=1)
lista_arquivos = tk.Listbox(app)
logo = Image.open("logo.png")
photo = ImageTk.PhotoImage(logo)


label = tk.Label(app,image=photo)
label.grid(row=0, column=0,columnspan=2)
label.config(background="#1a1a1a")

#Linha e Botao Unimed---------------------------------------------------------------------------------------------------
texto_orientacao1 = Label(app,text="Clique em procurar para inserir a planilha da Unimed")
texto_orientacao1.grid(column=0, row=1)
texto_orientacao1.config(font=("Arial",10),foreground="White",background="#1a1a1a")
#Botao1
button1 = customtkinter.CTkButton(master=app, text="Procurar Unimed", command=procurar_arquivo_Unimed)
button1.grid(column=1, row=1)
pular_linha(0,2)

#Linha e Botao Netreport------------------------------------------------------------------------------------------------
texto_orientacao2 = Label(app,text= "Clique em procurar para inserir a planilha do Netreport")
texto_orientacao2.grid(column=0, row=3)
texto_orientacao2.config(font=("Arial",10),foreground="White",background="#1a1a1a")
#Botao2
button2 = customtkinter.CTkButton(master=app, text="Procurar NetReport", command=procurar_arquivo_Netreport)
button2.grid(column=1, row=3)
pular_linha(0,4)

#Botoes Ação------------------------------------------------------------------------------------------------
#Botao3
button3 = customtkinter.CTkButton(master=app, text="Comparar Planilhas",  command=lambda: executar_comparacao(caixa_texto))
button3.grid(column=0, row=5)
#Botao4
button4 = customtkinter.CTkButton(master=app, text="Reiniciar Comparação", command=limpar_comparacao)
button4.grid(column=1, row=5)
button4.configure(fg_color='#ff3a35')
pular_linha(0,6)

caixa_texto = tk.Text(app, height=25, width=120, exportselection=False)
caixa_texto.bind("<Key>", block_edit)

caixa_texto.config(font=("Arial",12),background="#292725",foreground="White")
caixa_texto.grid(column=0, row=7,columnspan=2)
scrollbar = tk.Scrollbar(caixa_texto, orient=tk.VERTICAL, command=caixa_texto.yview)
scrollbar.place(relx=1.0, rely=0, relheight=1.0, anchor=tk.NE)
caixa_texto.config(yscrollcommand=scrollbar.set)
caixa_texto.configure(state='normal')
caixa_texto.insert('1.0', '\n\n\n\n\n\n\n\n\nO resultado da Comparação irá aparecer nesta Janela :)')
caixa_texto.tag_configure('center', justify='center')
caixa_texto.tag_add('center', '1.0', 'end')
pular_linha(0,8)

#Botao5
button5 = customtkinter.CTkButton(master=app, text="Exportar Resultado",  command=exportar_resultado)
button5.grid(column=0, row=9,columnspan=2)
button5.configure(width=700)

caixa_texto.bind('<Button-1>', reset_selection)

app.protocol("WM_DELETE_WINDOW", deletar_arquivo)
app.mainloop()


