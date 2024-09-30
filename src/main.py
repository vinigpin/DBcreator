import pdfplumber
import pyodbc as bd
import getpass 
import tkinter as tk
from tkinter import filedialog
import os

root = tk.Tk()
root.withdraw()
root.attributes('-topmost', True)

def continuar():
    input('[enter] para continuar: ')
    os.system('cls')

def draw_line():
    print('\033[1;36m-----------------------------------------------\033[1;39m')

def find_pdf():
    print('Selecione o arquivo PDF para leitura: ')
    arq = filedialog.askopenfilename(
        title='Selecione o arquivo PDF para leitura',
        filetypes=(('Arquivo PDF', '*.pdf'), ('Outros arquivos', '*.*')),
        multiple = False)
    print('lendo arquivo selecionado...')
    draw_line()
    return arq

def connect_BD():
    global connect
    
    user = input('Digite seu usuário do banco de dados: ')
    password = getpass.getpass('Digite sua senha do banco de dados: ')
    
    try:
        connect = bd.connect(driver = '{SQL server}',
                         server = 'regulus.cotuca.unicamp.br',
                         database = user,
                         uid = user,
                         pwd = password)
        print('\033[1;32mbanco de dados conectado\033[1;39m')
        draw_line()
        return True
    except:
        print('\033[1;31mErro de conexão com o banco de dados!\nVerifique se a VPN está ligada e se vc prencheu as informações corretamente\033[1;39m')
        draw_line()
        return False
    
    
def create_tables(start, end):
    cont_table = 0
    cont_columns = 0
    cont_rows = 0
    with pdfplumber.open(find_pdf()) as pdf:
        table_name = '!'
        cursor = connect.cursor()
        for page_num in range(start-1, end+1):
            pg = page_num
            first_page = pdf.pages[pg]
            text = first_page.extract_text()
            columns_name = []
            ant_line = '!'
            
            rows = text.split('\n')
            for line in rows:
                words = line.split(' ')
                try:
                    int(words[0])
                except:
                    if line.upper() == line and line[0] != '-':
                        table_name = line
                        table_exist = False
                    elif ant_line[0] == '-' and line[0] != '-':
                        columns_name = words
                        cont_table += 1
                        continuar()
                        print(f'{cont_table}ª tabela: {table_name}')
                        print(f'Colunas:   {columns_name}\n')

                        comand = f'create table {table_name} ( '
                        for column in columns_name:
                            cont_columns += 1
                            type = input(f'Digite o tipo da coluna "{column}": ')
                            key = input('Conteudo é uma PK ou FK: ').lower()
                            comand += f'{column} {type}'
                            if key == 'pk':
                                comand += ' primary key'
                            if key == 'fk':
                                reference = input('Digite a referencia: ')
                                comand += f' foreign key references {reference}'
                            print()
                            comand += ', '
                        comand += ')'
                        print(f'\033[1;36m{comand}\033[1;39m')
                        try:
                            cursor.execute(comand)
                            table_exist = True
                            print('\033[1;32mTabela criada com sucesso!\033[1;39m')
                        except:
                            print('\033[1;31mErro na criação da tabela!\033[1;39m')
                        cursor.commit()
                        print('\n')
                else:
                    if table_exist:
                        cont_rows += 1
                        if table_name == 'PESSOAS':
                            PESSOAS_name = words[:]
                            PESSOAS_name[0] = ''
                            PESSOAS_name[-1] = ''
                            comand = f"insert into {table_name} values ({words[0]}, '{(' '.join(PESSOAS_name)).strip()}', '{words[-1]}')"
                            
                            try:
                                cursor.execute(comand)
                                print(f'\033[1;32m{comand}\nNova linha adicionada a tabela {table_name}!\033[1;39m')
                            except:
                                print(f'\033[1;31m{comand}\nErro ao adicionar linha!\033[1;39m')
                            print()
                            cursor.commit()
                        elif table_name == 'NASCIMENTOS':
                            
                            local = words[:]
                            local[0] = ''
                            local[1] = ''
                            local[2] = ''
                            local[-3] = ''
                            local[-2] = ''
                            local[-1] = ''
                            
                            if words[1] != 'NULL':
                                datetime = words[1] + 'T' + words[2]
                                comand = f"insert into {table_name} values ({words[0]}, '{datetime}', '{(' '.join(local)).strip()}', {words[-3]}, {words[-2]}, {words[-1]})"
                            else:
                                datetime = 'NULL'
                                comand = f"insert into {table_name} values ({words[0]}, {datetime}, '{(' '.join(local)).strip()}', {words[-3]}, {words[-2]}, {words[-1]})"
                            
                            try:
                                cursor.execute(comand)
                                print(f'\033[1;32m{comand}\nNova linha adicionada a tabela {table_name}!\033[1;39m')
                            except:
                                print(f'\033[1;31m{comand}\nErro ao adicionar linha!\033[1;39m')
                            print()
                            cursor.commit()
                            
                        elif table_name == 'MORTES':
                            
                            local = words[:]
                            local[0] = ''
                            local[1] = ''
                            local[2] = ''
                            local[-1] = ''
                            
                            if words[1] != 'NULL':
                                datetime = words[1] + 'T' + words[2]
                                comand = f"insert into {table_name} values ({words[0]}, '{datetime}', '{(' '.join(local)).strip()}', {words[-1]})"
                            else:
                                datetime = 'NULL' 
                                comand = f"insert into {table_name} values ({words[0]}, {datetime}, '{(' '.join(local)).strip()}', {words[-1]})"
                            
                            try:
                                cursor.execute(comand)
                                print(f'\033[1;32m{comand}\nNova linha adicionada a tabela {table_name}!\033[1;39m')
                            except:
                                print(f'\033[1;31m{comand}\nErro ao adicionar linha!\033[1;39m')
                            print()
                            cursor.commit()
                            
                        elif table_name == 'CASAMENTOS':
                            if words[1] != 'NULL':
                                datetime = words[1] + 'T' + words[2]
                                comand = f"insert into {table_name} values ({words[0]}, '{datetime}', "
                            else:
                                datetime = 'NULL'
                                comand = f"insert into {table_name} values ({words[0]}, {datetime}, "
                            
                            index_local = 0
                            index_nome = 0
                            c = 0
                            for i in words:
                                try:
                                    int(i)
                                except:
                                    pass
                                else:
                                    c += 1
                                
                                    if c == 2:
                                        index_local = words.index(i)
                                    elif c == 3:
                                        index_nome = words.index(i)
                            
                            local = ''
                            for c in range(3, index_local):
                                local += words[c] + ' '
                                
                            nome = ''
                            for c in range(index_nome+1, len(words)):
                                nome += words[c] + ' '
                                
                            comand += f"'{local}', {words[index_local]}, {words[index_nome]}, "
                            if nome == 'NULL':
                                nome = 'NULL'
                                comand += f'{nome} )'
                            else:
                                comand += f"'{nome}')"
                            try:
                                cursor.execute(comand)
                                print(f'\033[1;32m{comand}\nNova linha adicionada a tabela {table_name}!\033[1;39m')
                            except:
                                print(f'\033[1;31m{comand}\nErro ao adicionar linha!\033[1;39m')
                            print()
                            cursor.commit()
                        
                        
                ant_line = line

    print('\033[1;36m')
    draw_line()
    print(f'tabelas lidas: {cont_table}')
    print(f'colunas lidas: {cont_columns}')
    print(f'linhas lidas:  {cont_rows}')
    draw_line()
    print(f'Programa encerrado!')
    draw_line()
        

def main():
    print('\033[1;36m-----------------------------------------------')
    print('PDF -> BANCO DE DADOS')
    print('-----------------------------------------------\033[1;39m')                
    if connect_BD():
        create_tables(54, 59)


if __name__ == '__main__':
    main()

