import pandas as pd
import csv as csv
import matplotlib.pyplot as plt
import numpy as np
import sys
import tkinter as tk
from tkinter import PhotoImage, scrolledtext, ttk, filedialog

#criando função para procurar arquivo
def selecionar_arquivo():
    global arquivo
    arquivo = filedialog.askopenfilename(initialdir="/", 
                                               title="Selecione um Arquivo",
                                               filetype=(("xlsx files", "*.xlsx"),("All Files", "*.*")))
    entry_arquivo.delete(0, tk.END)
    entry_arquivo.insert(0, arquivo)
    carregar_dataframe()

def carregar_dataframe():
    try:
        global df
        df = pd.read_excel(arquivo)
        print('')
    except pd.errors.EmptyDataError:
        print('')
    except pd.errors.ParserError:
        print('')

#função para criar um novo arquivo (nova turma)
def criar_turma():

    dados = {'Aluno': ['Inicial'],
            'Idade': [10],
            'Cpf': [10],
            'Telefone': [10],
            'Cidade': ['Inicial']}
    
    df = pd.DataFrame(dados)

    local_arquivo = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                   filetypes=[("Arquivos do Excel", "*.xlsx"), ("Todos os Arquivos", "*.*")])
    if not local_arquivo:
        return

    df.to_excel(local_arquivo, index=False)



#carregando df
def ver_turmas():
    #cria a janela de ver turmas
    global janela_ver_turmas
    janela_ver_turmas = tk.Toplevel(janela_login)
    janela_ver_turmas.title('')
    janela_ver_turmas.geometry('500x500+100+100')
    janela_login.withdraw()
    #frame para exibir a turma
    frame_ver_alunos = tk.LabelFrame(janela_ver_turmas, text="Alunos")
    frame_ver_alunos.place(height=250, width=500)
    #frame pra abrir o arquivo
    file_frame = tk.LabelFrame(janela_ver_turmas, text="Escolher Arquivo")
    file_frame.place(height=100, width=400, rely=0.65, relx=0)
    #botÕes
    btn_search = tk.Button(file_frame, text='Procurar Arquivo', command=lambda: File_dialog())
    btn_search.place(rely=0.65, relx=0.50)
    btn_loadfile = tk.Button(file_frame, text='Carregar Arquivo', command=lambda: Load_excel_data())
    btn_loadfile.place(rely=0.65, relx=0.20)
    #texto
    label_file = ttk.Label(file_frame, text='Nenhum arquivo Selecionado')
    label_file.place(rely=0, relx=0)
    #TREEVIEW
    tv1 = ttk.Treeview(frame_ver_alunos)
    tv1.place(relheight=1, relwidth=1)
    #cria os scroll 
    treescrolly = tk.Scrollbar(frame_ver_alunos, orient='vertical', command=tv1.yview)
    treescrollx = tk.Scrollbar(frame_ver_alunos, orient='horizontal', command=tv1.xview)
    tv1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set)
    treescrollx.pack(side='bottom', fill='x')
    treescrolly.pack(side='right', fill='y')
    #função pra escolher o arquivo
    def File_dialog():
        filename = filedialog.askopenfilename(initialdir="/", 
                                               title="Selecione um Arquivo",
                                               filetype=(("xlsx files", "*.xlsx"),("All Files", "*.*")))
        label_file["text"] = filename
    def Load_excel_data():
        file_path = label_file["text"]
        try:
            excel_filename = r"{}".format(file_path)
            df = pd.read_excel(excel_filename)
        except ValueError:
            tk.messagebox.showerror('Informação', 'O arquivo que você escolheu é inválido')
            return None
        except FileNotFoundError:
            tk.messagebox.showerror('Informação', f'{file_path}')
            return None
        clear_data()
        tv1['column'] = list(df.columns)
        tv1['show'] = 'headings'
        for column in tv1['columns']:
            tv1.heading(column, text=column)

        df_rows = df.to_numpy().tolist()
        for row in df_rows:
            tv1.insert('', 'end', values=row)
        return None
    def clear_data():
        tv1.delete(*tv1.get_children())
    #ao fechar a janela do menu mostra a janela de login novamente e destroi a janela de ver turmas
    def comando_ao_fechar0():
        janela_login.deiconify()
        janela_ver_turmas.destroy()
    #maneira de ativar um comando ao fechar uma janela (delete window)
    janela_ver_turmas.protocol("WM_DELETE_WINDOW", comando_ao_fechar0)

#função para abrir a janela da turma quando a senha estiver correta
def verificar_senha():
    usuario = inserir_usuario.get()
    senha = inserir_senha.get()
#abre o menu caso a acc/senha for admin/admin
    if 'arquivo' not in globals() or not arquivo:
        label_erro.config(text='Selecione um arquivo.')
        return
    if usuario == 'admin' and senha == 'admin':
        abrir_janela_turma1()
        janela_login.withdraw()
    else:
        label_erro.config(text='Senha Incorreta')
    #deleta o usuário e senha escritos após entrar.
    inserir_usuario.delete(0, tk.END)
    inserir_senha.delete(0, tk.END)
#função para abrir a janela do menu da turma1
def abrir_janela_turma1():
###começo do programa da turma1(funções)###
    #função adicionar novo aluno

        #função ver turma1
    def ver_turma1():
        #cria janela para ver os membros
        janela_ver_membros = tk.Toplevel(janela_menu_turma1)
        janela_ver_membros.title('')
        janela_ver_membros.geometry('800x500+120+120')
        
        df = pd.read_excel(arquivo)

        texto_df = scrolledtext.ScrolledText(janela_ver_membros, wrap=tk.WORD, width=100, height=100, font=('Arial', 12))
        texto_df.insert(tk.INSERT, df.to_string(index=False))
        texto_df.pack(pady=20, padx=20)
        texto_df.config(state=tk.DISABLED)
        
    def adicionar_novo_aluno():
        #lê o arquivo excel
        df = pd.read_excel(arquivo)
        #pega as informações do aluno
        aluno = input_aluno.get()         
        idade = int(input_idade.get())
        cpf = int(input_cpf.get())
        telefone = int(input_telefone.get())
        cidade = input_cidade.get()
        #puxa as informações do df e víncula as que pegou do botão
        novo_aluno = pd.DataFrame({'Aluno': [aluno], 'Idade': [idade], 'Cpf': [cpf], 'Telefone': [telefone], 'Cidade': [cidade]})
        #se o nome estiver no dataframe:
        if aluno in df['Aluno'].values:
            #puxa o comando do botao de fechar a janela
            def fechar_janela_notadd_aluno():
                janela_notadd_aluno.destroy()
            #cria a janela "esse aluno ja está na turma"
            janela_notadd_aluno = tk.Toplevel(janela_menu_turma1)
            janela_notadd_aluno.title('')
            janela_notadd_aluno.geometry('200x50+120+120')
            #texto dizendo que o aluno ja está na turma
            texto_aluno_ja_esta = tk.Label(janela_notadd_aluno, text='Esse aluno ja está na turma!')
            texto_aluno_ja_esta.pack()
            #botao para fechar a janela
            btn_fechar_notaddaluno = tk.Button(janela_notadd_aluno, text='Fechar', command=fechar_janela_notadd_aluno)
            btn_fechar_notaddaluno.pack()
        #se o nome ainda não estiver no dataframe
        else:
            #puxa o comando do botao de fechar a janela
            def fechar_janela_add_aluno():
                janela_add_aluno.destroy()
            #cria a janela dizendo que o aluno foi adicionado com sucesso
            janela_add_aluno = tk.Toplevel(janela_menu_turma1)
            janela_add_aluno.title('')
            janela_add_aluno.geometry('200x50+120+120')
            #texto
            texto_aluno_adicionado = tk.Label(janela_add_aluno, text='Aluno adicionado com sucesso!')
            texto_aluno_adicionado.pack()
            #botão fechar
            btn_add_aluno = tk.Button(janela_add_aluno, text='Fechar', command=fechar_janela_add_aluno)
            btn_add_aluno.pack()
            #adicionar o novo aluno no df
            df = pd.concat([df, novo_aluno], ignore_index = True)
            df.to_excel(arquivo, index = False)
        #apaga oque foi digitado nos Entry
        input_aluno.delete(0, tk.END)
        input_idade.delete(0, tk.END)
        input_cpf.delete(0, tk.END)
        input_telefone.delete(0, tk.END)
        input_cidade.delete(0, tk.END)
    #função para remover aluno da turma
    def remover_aluno():
        #lê o arquivo csv
        df = pd.read_excel(arquivo)
        #pega o nome do aluno que você deseja remover
        aluno_removido = input_aluno_removido.get()
        #se o aluno estiver no df:
        if aluno_removido in df['Aluno'].values:
            #puxa a função para fechar a janela do aluno removido com sucesso
            def fechar_janela_remover_aluno():
                janela_remover_aluno.destroy()
            #cria a janela remover_aluno
            janela_remover_aluno = tk.Toplevel(janela_menu_turma1)
            janela_remover_aluno.title('')
            janela_remover_aluno.geometry('200x50+120+120')
            #texto
            texto_aluno_removido = tk.Label(janela_remover_aluno, text='Aluno removido com sucesso!')
            texto_aluno_removido.pack()
            #botão
            btn_fechar_add_aluno = tk.Button(janela_remover_aluno, text='Fechar', command=fechar_janela_remover_aluno)
            btn_fechar_add_aluno.pack()
            #tira o nome do aluno do df
            df = df.drop(df[df['Aluno'] == aluno_removido].index)
            #salva o df
            df.to_excel(arquivo, index=False)
        #se o aluno não estiver no df, não tem como remove-lo
        else:
            #função para fechar a janela
            def fechar_janela_notexist_aluno():
                janela_notexist_aluno.destroy()
            #cria a janela falando que não existe aluno na turma
            janela_notexist_aluno = tk.Toplevel(janela_menu_turma1)
            janela_notexist_aluno.title('')
            janela_notexist_aluno.geometry('200x50+120+120')
            #texto
            texto_notexist_aluno = tk.Label(janela_notexist_aluno, text='Este aluno não está na turma!')
            texto_notexist_aluno.pack()
            #botao de fechar
            btn_fechar_addaluno = tk.Button(janela_notexist_aluno, text='Fechar', command=fechar_janela_notexist_aluno)
            btn_fechar_addaluno.pack()
        #retira oque foi escrito nos entry.
        input_aluno_removido.delete(0, tk.END)
    
    #cria a janela do menu
    global janela_menu_turma1
    janela_menu_turma1 = tk.Toplevel(janela_login)
    janela_menu_turma1.title('')
    janela_menu_turma1.geometry('222x480+100+100')
    #ao fechar a janela do menu mostra a janela de login novamente e destroi a janela do menu.
    def comando_ao_fechar():
        janela_login.deiconify()
        janela_menu_turma1.destroy()
    #maneira de ativar um comando ao fechar uma janela (delete window)
    janela_menu_turma1.protocol("WM_DELETE_WINDOW", comando_ao_fechar)
    
    #frame para adicionar aluno
    frame_add_aluno = tk.LabelFrame(janela_menu_turma1, text='Adicionar Aluno')
    frame_add_aluno.pack()
    #texto nome
    texto_nome = tk.Label(frame_add_aluno, text='Nome do novo aluno:')
    texto_nome.pack()
    #entry nome
    input_aluno = tk.Entry(frame_add_aluno)
    input_aluno.pack()
    #texto idade
    texto_idade = tk.Label(frame_add_aluno, text='Idade do novo aluno:')
    texto_idade.pack(padx=40)
    #entry idade
    input_idade = tk.Entry(frame_add_aluno)
    input_idade.pack()
    #texto cpf
    texto_cpf = tk.Label(frame_add_aluno, text='Cpf do novo aluno:')
    texto_cpf.pack()
    #entry cpf
    input_cpf = tk.Entry(frame_add_aluno)
    input_cpf.pack()
    #texto telefone
    texto_telefone = tk.Label(frame_add_aluno, text='Telefone do novo aluno:')
    texto_telefone.pack()
    #entry telefone
    input_telefone = tk.Entry(frame_add_aluno)
    input_telefone.pack()
    #texto cidade
    texto_cidade = tk.Label(frame_add_aluno, text='Cidade do novo aluno:')
    texto_cidade.pack()
    #entry cidade
    input_cidade = tk.Entry(frame_add_aluno)
    input_cidade.pack(pady=5)
    #botao novo aluno
    btn_adicionar_novo_aluno = tk.Button(frame_add_aluno, text='Adicionar aluno a turma', command=adicionar_novo_aluno)
    btn_adicionar_novo_aluno.pack(pady=5)

    #frame para remover aluno
    frame_remover_aluno = tk.LabelFrame(janela_menu_turma1, text='Remover Aluno')
    frame_remover_aluno.pack()
    #texto aluno removido
    texto_aluno_removido = tk.Label(frame_remover_aluno, text='Nome do aluno a ser removido:')
    texto_aluno_removido.pack()
    #entry aluno removido
    input_aluno_removido = tk.Entry(frame_remover_aluno)
    input_aluno_removido.pack(pady=5, padx=40)
    #botão aluno removido
    btn_remove_member = tk.Button(frame_remover_aluno, text='Remover aluno da turma', command=remover_aluno)
    btn_remove_member.pack(pady=5)

    #frame para ver dados
    frame_ver_dados = tk.LabelFrame(janela_menu_turma1, text='Dataframe')
    frame_ver_dados.pack()
    #texto ver dataframe
    texto_ver_dados = tk.Label(frame_ver_dados, text='Dados sobre a turma:')
    texto_ver_dados.pack(pady=5,padx=40)
    #botão para ver dados
    btn_ver_turma1 = tk.Button(frame_ver_dados, text='Dados', command=ver_turma1)
    btn_ver_turma1.pack()

#criando a tela de login
janela_login = tk.Tk()
janela_login.title('')
#define a dimensão da janela
janela_login.geometry("250x500+100+100")
#colocando uma imagem na interface de login
imagem = PhotoImage(file="D:\Documentos\Fábio Python\Trabalho Final\logopython.png")
#redimensionando a imagem
fator_redimensionamento = 1
imagem_redimensionada = imagem.subsample(fator_redimensionamento)
#criando um rótulo para exibir a imagem
label_imagem = tk.Label(janela_login, image=imagem_redimensionada)
label_imagem.pack()




#criando frame analizador
frame_view_open = tk.LabelFrame(janela_login, text='Analizador')
frame_view_open.place(height=100, width=180, rely=0.25, relx=0.5, anchor='center')
#botao criar turma
btn_criar_turma = tk.Button(frame_view_open, text='Criar Turma', command=criar_turma) #command criar_turma
btn_criar_turma.pack()
#botao ver turmas
botao_ver_turmas = tk.Button(frame_view_open, text='Ver Turmas', command=ver_turmas) #command ver_turmas
botao_ver_turmas.pack(pady=5)

#criando frame editor
frame_open = tk.LabelFrame(janela_login, text='Editor de Turmas')
frame_open.place(height=250, width=180, rely=0.65, relx=0.5, anchor='center')
#criando entry
entry_arquivo = tk.Entry(frame_open)
entry_arquivo.pack(pady=5)
#botao slct arqv
btn_selecionar_arquivo = tk.Button(frame_open, text='Selecionar Arquivo', command=selecionar_arquivo)
btn_selecionar_arquivo.pack(pady=5)
#texto user
texto_inserir_usuario = tk.Label(frame_open, text='Insira seu nome de usuário:')
texto_inserir_usuario.pack()
#campo user
inserir_usuario = tk.Entry(frame_open)
inserir_usuario.pack(pady=5)
#texto senha
texto_inserir_senha = tk.Label(frame_open, text='Insira a sua senha de acesso:')
texto_inserir_senha.pack()
#campo senha
inserir_senha = tk.Entry(frame_open, show='*')
inserir_senha.pack(pady=5)
#botão login
botao_login = tk.Button(frame_open, text='Entrar', command=verificar_senha)
botao_login.pack(pady=5)
#quando a senha estiver errada, aparecer esse erro
label_erro = tk.Label(frame_open, text='', fg='red')
label_erro.pack()

#colocando a janela em loop pra manter ela aberta.
janela_login.mainloop()