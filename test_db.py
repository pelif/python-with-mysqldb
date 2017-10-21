#enconding utf-8

import MySQLdb
import os
import xlwt

con = MySQLdb.connect('172.17.0.2','root','root_mysql')
con.select_db('test01')
cursor = con.cursor()

#LIMPAR TELA
def limpar_tela():
    os.system('clear')
    

#IMPRIMIR MENU
def imprimir_menu():
    print("=======================DIGITE UMA OPCAO========================")
    
    print("""1 - Inserir
    \n2 - Alterar
    \n3 - Consultar
    \n4 - Listar todos
    \n5 - Excluir
    \n6 - Gerar Relatorio""")
    
    opcao = str(input('\n'))
    if opcao in '1':
        inserir_pessoa()
    elif opcao in '2':
        alterar_pessoa()
    elif opcao in '3':
        consultar_pessoa()
    elif opcao in '4':
        consultar_todos()        
    elif opcao in '5':
        excluir_pessoa()   
    elif opcao in '6':
        gerar_relatorio()    
    

#INSERCAO DE DADOS
def inserir_pessoa():
    nome = str(raw_input('Digite o nome:'))
    dt_nasc = str(raw_input('Digite a data de nascimento: formato(dd/mm/yyyy)'))
    data_nasc = dt_nasc[6:10]+"-"+dt_nasc[3:5]+"-"+dt_nasc[0:2]
        
    telefone = str(raw_input('Digite o Telefone: (11)988776677'))
    email = str(raw_input('Digite o E-mail: '))
    cpf = str(raw_input('Digite o cpf: (sem porntos e tra\xc3os)'))
        
    try:
        cursor.execute("""INSERT INTO pessoas 
        (nome,data_nasc,telefone,email,cpf) VALUES 
        (%s,%s,%s,%s,%s)""",(nome,data_nasc,telefone,email,cpf))
        con.commit()
    except:
        con.rollback        
        
    limpar_tela()
    imprimir_menu()    
        
#ALTERACAO DE DADOS
def alterar_pessoa():
    cpf = str(raw_input('Digite o CPF da pessoa: '))
    
    try:
        cursor.execute("SELECT nome FROM pessoas where cpf in ("+cpf+")")
        con.commit()
        rs = cursor.fetchone()
        print("Voce vai alterar dados de {}".format(rs[0]))
        
        nome = str(raw_input('Digite o nome:'))
        dt_nasc = str(raw_input('Digite a data de nascimento: formato(dd/mm/yyyy)'))
        data_nasc = dt_nasc[6:10]+"-"+dt_nasc[3:5]+"-"+dt_nasc[0:2]
        
        telefone = str(raw_input('Digite o Telefone: (11)988776677'))
        email = str(raw_input('Digite o E-mail: '))
                
        sql = """UPDATE pessoas set nome = %s,
        data_nasc = %s, 
        telefone = %s,
        email = %s        
        WHERE cpf = %s
        """
        
        try:
            cursor.execute(sql, (nome, data_nasc, telefone, email, cpf))
            con.commit()
            print("{} Alterado com sucesso".format(rs[0]))
        except:
            print("Nao foi possivel alterar")
            
    except:
        print("Nao selecionou!")  
    
    limpar_tela()
    imprimir_menu() 
        
        
#CONSULTAR REGISTRO
def consultar_pessoa():
    cpf = str(raw_input('Digite o Cpf da Pessoa: '))
    
    try:
        cursor.execute("SELECT nome,data_nasc,telefone,email,cpf FROM pessoas where cpf in ("+cpf+")")
        con.commit()
        rs = cursor.fetchone()    
        print(rs[0])
    except:
        print("Nao encontrou pessoa com este CPF!")
          
        
#EXCLUSAO DE DADOS
def excluir_pessoa():
    cpf = str(raw_input('Digite o Cpf da Pessoa: '))
    
    cursor.execute("SELECT nome FROM pessoas where cpf in ("+cpf+")")
    con.commit()
    rs = cursor.fetchone()
    
    try:
        cursor.execute("DELETE FROM pessoas where cpf = "+cpf+"")
        con.commit()
        print("Voce excluiu o cadastro de {}".format(rs[0]))
    except:
        print("Nao foi possivel excluir a pessoa!")
        

#CONSULTAR A TODOS
def consultar_todos():
    try:
        cursor.execute("SELECT nome,data_nasc,telefone,email,cpf FROM pessoas")
        con.commit()
        result = cursor.fetchall()
        for row in result:
            print("{}\n{}\n{}\n{}\n{}\n".format(row[0],row[1],row[2],row[3],row[4]))
            print("="*50)
    except:
        print("Nao ha cadastros de pessoas para listar!")        
                
                
#GERAR RELATORIO
def gerar_relatorio():
    
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet(u'Pessoas')
    
    worksheet.write(0,0,u'Nome')
    worksheet.write(0,1,u'Data Nascimento')
    worksheet.write(0,2,u'Telefone')
    worksheet.write(0,3,u'E-mail')
    worksheet.write(0,4,u'CPF')    
    
    try:
        cursor.execute("SELECT nome,data_nasc,telefone,email,cpf FROM pessoas")
        con.commit()
        result = cursor.fetchall()
        x=1
        for row in result:
            #print("{}\n{}\n{}\n{}\n{}\n".format(row[0],row[1],row[2],row[3],row[4]))
#             worksheet.write(x, 0, row[0])
#             worksheet.write(x, 1, row[1], 
#                 style=xlwt.easyxf(num_format_str='dd/mm/yyyy'))
#             worksheet.write(x, 2, row[2])
#             worksheet.write(x, 3, row[3])
#             worksheet.write(x, 4, row[4],
#                 style=xlwt.easyxf(num_format_str='###.###.###-##'))
            
            row_nome = str(row[0])
            row_dt_nasc = str(row[1][0:2]+"/"+row[1][3:5]+"/"+row[1][6:10])
            row_tel = str(row[2])
            row_email = str(row[3])
            row_cpf = str(row[4])            
                
            worksheet.write(x,0,row_nome)
            worksheet.write(x,1,row_dt_nasc)
            worksheet.write(x,2,row_tel)
            worksheet.write(x,3,row_email)
            worksheet.write(x,4,row_cpf)           
            x = x+1
             
        workbook.save('/home/felipe/pessoas.xls')    
        print("Relatorio gerado com sucesso!")
    except:                
        print("Nao ha registros para gerar relatorio")
        
       
imprimir_menu()