from pydoc import safeimport
import pandas as pd
import numpy as np
from tqdm import tqdm
from openpyxl import load_workbook
from tkinter import Button, filedialog
import tkinter as tk
import time


# pyinstaller --onefile --noconsole .\nome.py
# pyinstaller --onefile --console ./zRateioSAPCromex.py


print('\n\033[1;32mSoftware by Anderson Marinho \033[m')
print('\033[1;32mVersão 1.0 \033[m')
print('\n')

while True:
    if input('Deseja continuar? [S/N] ').upper() == 'S':
        root= tk.Tk()
        root.title('Import Base_Saude Excel')
        cenario = tk.Canvas(root, width = 300, height = 300, bg = 'lightsteelblue')
        cenario.pack()

        # Abrindo base de arquivos saúde
        def getExcel():
            global base_saude
            import_caminho = filedialog.askopenfilename()
            base_saude = import_caminho
            botao_excel.destroy()
            root.destroy()

            base_saude = pd.read_excel(base_saude, sheet_name='Fat1')
            print(base_saude.head(5))

        botao_excel = tk.Button(text='Import Base_Saude Excel', command=getExcel, bg='green', fg='white', font=('helvetica', 10, 'bold'))
        cenario.create_window(150, 150, window=botao_excel)
        root.mainloop()

        time.sleep(1)

        root= tk.Tk()
        root.title('Import BaseAtivos Excel')
        cenario = tk.Canvas(root, width = 300, height = 300, bg = 'lightsteelblue')
        cenario.pack()

        # Abrindo base de arquivos saúde
        def getExcel1():
            global base_ativos
            import_caminho = filedialog.askopenfilename()
            base_ativos = import_caminho
            botao_excel.destroy()
            root.destroy()

            base_ativos = pd.read_excel(base_ativos, sheet_name='SRA')
            print(base_ativos.head(5))

        botao_excel = tk.Button(text='Import BaseAtivos Excel', command=getExcel1, bg='green', fg='white', font=('helvetica', 10, 'bold'))
        cenario.create_window(150, 150, window=botao_excel)
        root.mainloop()

        print('\n\033[1;33mDigitar as colunas para iniciar o processo de Rateio \033[m')


        #  switch case com letras do alfaberto
        letras = {'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5, 
        'F': 6, 'G': 7, 'H': 8, 'I': 9, 'J': 10, 'K': 11, 
        'L': 12, 'M': 13, 'N': 14, 'O': 15, 'P': 16, 
        'Q': 17, 'R': 18, 'S': 19, 'T': 20, 'U': 21, 
        'V': 22, 'W': 23, 'X': 24, 'Y': 25, 'Z': 26}

        # ##################
        letra = input('\33[1;32mDigite a letra da coluna com Nome Associado: \33[m').upper()    #05 Nome Associado
        while True:
            try:
                if letra in letras:
                    simbolo = int(letras[letra]) - 1
                    coluna3 = simbolo
                    break
                else:
                    print('\33[1;31mLetra inválida\33[m')
                    letra = input('\33[1;32mDigite a letra da coluna com Nome Associado: \33[m').upper()
            except:
                print('Letra inválida')
                letra = input('\33[1;32mDigite APENAS letra: \33[m').upper()
                break

        # ##################
        letra = input('\33[1;32mDigite a coluna CPF: \33[m').upper()    #06 CPF
        while True:
            try:
                if letra in letras:
                    simbolo = int(letras[letra]) - 1
                    coluna2 = simbolo
                    break
                else:
                    print('\33[1;31mLetra inválida\33[m')
                    letra = input('\33[1;32mDigite a coluna CPF: \33[m').upper()
            except:
                print('Letra inválida')
                letra = input('\33[1;32mDigite APENAS letra: \33[m').upper()
                break

        # ##################
        letra = input('\33[1;32mDigite a coluna Nome Titular: \33[m').upper()   #07 Nome Titular
        while True:
            try:
                if letra in letras:
                    simbolo = int(letras[letra]) - 1
                    coluna1 = simbolo
                    break
                else:
                    print('\33[1;31mLetra inválida\33[m')
                    letra = input('\33[1;32mDigite a coluna Nome Titular: \33[m').upper()
            except:
                print('Letra inválida')
                letra = input('\33[1;32mDigite APENAS letra: \33[m').upper()
                break

        # ##################
        letra = input('\33[1;32mDigite a coluna Grau Parentesco: \33[m').upper()   #11 Grau Parentesco
        while True:
            try:
                if letra in letras:
                    simbolo = int(letras[letra]) - 1
                    coluna4 = simbolo
                    break
                else:
                    print('\33[1;31mLetra inválida\33[m')
                    letra = input('\33[1;32mDigite a coluna Grau Parentesco: \33[m').upper()
            except:
                print('Letra inválida')
                letra = input('\33[1;32mDigite APENAS letra: \33[m').upper()
                break

        # ##################
        letra = input('\33[1;32mDigite a coluna VL BENEFICIO: \33[m').upper()   #22 Soma de VL BENEFICIO
        while True:
            try:
                if letra in letras:
                    simbolo = int(letras[letra]) - 1
                    coluna5 = simbolo
                    break
                else:
                    print('\33[1;31mLetra inválida\33[m')
                    letra = input('\33[1;32mDigite a coluna VL BENEFICIO: \33[m').upper()
            except:
                print('Letra inválida')
                letra = input('\33[1;32mDigite APENAS letra: \33[m').upper()
                break


        contaContabil = input('\33[1;32mDigite a conta contábil: \33[m')  #Conta Contábil
        while True:
            try:
                if contaContabil.isnumeric():
                    ContaRateio = int(contaContabil)
                    break
                else:
                    print('\33[1;31mConta contábil inválida, contém letras ou espaço\33[m')
                    contaContabil = input('\33[1;32mDigite a conta contábil: \33[m')
            except:
                print('Conta contábil inválida')
                contaContabil = input('\33[1;32mDigite a conta contábil novamente: \33[m')

                break
        Descricao = input('\33[1;32mDigite a descrição da conta contábil: \33[m')  #Descrição da Conta Contábil
        while True:
            try:
                if all(c.isalpha() or c.isspace() for c in Descricao):
                    DescricaoTxt = str(Descricao)
                    break
                else:
                    print('\33[1;31mDescrição da conta contábil inválida, contém números ou espaço\33[m')
                    Descricao = input('\33[1;32mDigite a descrição da conta contábil: \33[m')
            except:
                print('Descrição da conta contábil inválida')
                Descricao = input('\33[1;32mDigite a descrição da conta contábil novamente: \33[m')
                break


        refMes = input('\33[1;32mDigite o ano/mês do rateio [AAAA/MM]: \33[m')  #Descrição MêS/Ano
        refMesTxt = str(refMes)


        print('\n')

        saude = base_saude.iloc[:, [coluna1, coluna2, coluna3, coluna4, coluna5]]
        saude.columns.values[0] = 'Nome Titular'
        saude.columns.values[1] = 'CPF do Titular'
        saude.columns.values[2] = 'Nome Associado'
        saude.columns.values[3] = 'Grau Parentesco'
        saude.columns.values[4] = 'Soma de VL BENEFICIO'


        print(saude.head(5))

        ativos = base_ativos.iloc[6:, [8,4,17,5,6,48,3,2,40,22,1]]
        ativos.columns.values[0] = 'CPF do Titular'
        ativos.columns.values[1] = 'CCUSTO'
        ativos.columns.values[2] = 'STATUS'
        ativos.columns.values[3] = 'Admissão'
        ativos.columns.values[4] = 'Demissão'
        ativos.columns.values[5] = 'DIRETORIA'
        ativos.columns.values[6] = 'Nome TMF'
        ativos.columns.values[7] = 'MAT'
        ativos.columns.values[8] = 'Aviso Demissão'
        ativos.columns.values[9] = 'TIPO'
        ativos.columns.values[10] = 'FILIAL'

        print(ativos.head(5))

        ativos_saude = pd.merge(saude, ativos, how='left', on=['CPF do Titular'])
        ativos_saude['Soma de VL BENEFICIO'] = ativos_saude['Soma de VL BENEFICIO'].astype(float)*100
        ativos_saude['CPF do Titular'] = ativos_saude['CPF do Titular'].astype(str)
        ativos_saude['Admissão'] = ativos_saude['Admissão'].astype('datetime64')
        ativos_saude['Admissão'] = ativos_saude['Admissão'].dt.strftime('%d/%m/%Y')
        ativos_saude['Admissão'] = ativos_saude['Admissão'].astype(str)
        ativos_saude=ativos_saude.assign(ContaRazao=ContaRateio)
        ativos_saude=ativos_saude.assign(Tipo='DB')
        ativos_saude=ativos_saude.assign(Descricao=DescricaoTxt)
        ativos_saude=ativos_saude.assign(AnoMes=refMesTxt)

        print(ativos_saude.head(5))
        # Termino da junção das bases

        input('Pressione ENTER para continuar ')

        root= tk.Tk()
        root.title('Salvar Base Excel')
        cenario = tk.Canvas(root, width = 300, height = 300, bg = 'lightsteelblue')
        cenario.pack()

        # Salvar arquivos Rateio
        def getExcel3():
            save = filedialog.asksaveasfilename(defaultextension='.*', initialdir='C:/', title='Rateio.xlsx', filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))
            baseDadosOut = ativos_saude.to_excel(save, sheet_name='Rateio', index=False)
            baseDadosOut = pd.DataFrame(np.random.randint(0, int(1e8), (10000, 1000)))
            tqdm.pandas()
            baseDadosOut.groupby(0).progress_apply(lambda x: x**2)
            botao_excel.destroy()
            root.destroy()

        botao_excel = tk.Button(text='Salvar Base conferência Rateio Excel', command=getExcel3, bg='green', fg='white', font=('helvetica', 10, 'bold'))
        cenario.create_window(150, 150, window=botao_excel)
        root.mainloop()



        # Pivot Table

        somatorio = pd.pivot_table(ativos_saude, index=['CCUSTO', 'ContaRazao','Tipo', 'Descricao', 'AnoMes'], 
        values=['Soma de VL BENEFICIO'], aggfunc=np.sum, margins=True, margins_name='30355')
        rateioSoma = somatorio
        rateioSoma['Soma de VL BENEFICIO'] = rateioSoma['Soma de VL BENEFICIO'].astype(int)
        print(rateioSoma.head(5))



        root= tk.Tk()
        root.title('Salvar PivotTable Rateio Excel')
        cenario = tk.Canvas(root, width = 300, height = 300, bg = 'lightsteelblue')
        cenario.pack()

        # Salvar arquivos Rateio
        def getExcel3():
            save = filedialog.asksaveasfilename(defaultextension='.*', initialdir='C:/', title='RateioPivot.xlsx', filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))
            baseDadosOut = rateioSoma.to_excel(save, sheet_name='RateioPivot', index=True)
            baseDadosOut = pd.DataFrame(np.random.randint(0, int(1e8), (10000, 1000)))
            tqdm.pandas()
            baseDadosOut.groupby(0).progress_apply(lambda x: x**2)
            botao_excel.destroy()
            root.destroy()

        botao_excel = tk.Button(text='Salvar PivotTable Rateio Excel', command=getExcel3, bg='green', fg='white', font=('helvetica', 10, 'bold'))
        cenario.create_window(150, 150, window=botao_excel)
        root.mainloop()

        print('\033[1;33mProcesso Gerador planilha Rateio e Pivot Table \033[m')
        print('\n')
        input('Pressione ENTER para continuar ')

        # pyinstaller --onefile --console .\docusign4.py

        print('\n\033[1;32mSoftware by Anderson Marinho \033[m')
        print('\033[1;32mVersão 1.0 \033[m')
        print('\n')

        print("|   --> Robot ENCONTROU esses dados para fazer ação <--                   ")
        print('|   --> Siga as instruções <--                   ')
        print('\n\033[1;33mSelecione o arquivo de base de dados na tela tk.\033[m')

        root= tk.Tk()
        root.title('Import Pivot Table Excel')
        canvas1 = tk.Canvas(root, width = 300, height = 300, bg = 'lightsteelblue')
        canvas1.pack()


        def getExcel ():
            global nome_arquivo, sheet_selecionada
            import_file_path = filedialog.askopenfilename()
            nome_arquivo = import_file_path
            planilha_aberta = nome_arquivo
            print(nome_arquivo)
            browseButton_Excel.destroy()
            root.destroy()
            time.sleep(1)    

            planilha_aberta = load_workbook(filename=nome_arquivo)
            sheet_selecionada = planilha_aberta['RateioPivot']

        browseButton_Excel = tk.Button(text='Import Pivot Table Excel', command=getExcel, bg='green', fg='white', font=('helvetica', 10, 'bold'))
        canvas1.create_window(150, 150, window=browseButton_Excel)
        root.mainloop()


        campo0 = 'L'
        campo2 = '0'
        # campo80L = str(input('\33[1;32mDigite o mês e ano [AAAAMM]: \33[m'))
        campo81L = str(input('\33[1;32mDigite o ano [AA] com dois dígitos apenas: \33[m'))
        while True:
            try:
                if campo81L.isnumeric() and len(campo81L) == 2:
                    campo81L = str(campo81L)
                    break
                else:
                    print('\33[1;31mDigite apenas números e com dois dígitos\33[m')
                    campo81L = str(input('\33[1;32mDigite o ano [AA] com dois dígitos apenas: \33[m'))

            except ValueError:
                print('\33[1;31mDigite um número inteiro válido.\33[m')
                campo81L = str(input('\33[1;32mDigite o ano [AA] com dois dígitos apenas: \33[m'))
                break

        campo83L = str(input('\33[1;32mDigite o mês [MM] com dois dígitos apenas: \33[m'))
        while True:
            try:
                if campo83L.isnumeric() and len(campo83L) == 2:
                    campo83L = str(campo83L)
                    break
                else:
                    print('\33[1;31mDigite apenas números e com dois dígitos\33[m')
                    campo83L = str(input('\33[1;32mDigite o mês [MM] com dois dígitos apenas: \33[m'))

            except ValueError:
                print('\33[1;31mDigite um número inteiro válido.\33[m')
                campo83L = str(input('\33[1;32mDigite o mês [MM] com dois dígitos apenas: \33[m'))
                break

        campo82L = str(input('\33[1;32mDigite o dia [DD] com dois dígitos apenas: \33[m'))
        while True:
            try:
                if campo82L.isnumeric() and len(campo82L) == 2:
                    campo82L = str(campo82L)
                    break
                else:
                    print('\33[1;31mDigite apenas números e com dois dígitos\33[m')
                    campo82L = str(input('\33[1;32mDigite o dia [DD] com dois dígitos apenas: \33[m'))

            except ValueError:
                print('\33[1;31mDigite um número inteiro válido.\33[m')
                campo82L = str(input('\33[1;32mDigite o dia [DD] com dois dígitos apenas: \33[m'))
                break

        campo7 = " "
        campo11 = 'FP'
        cabecalho = (campo0 + campo2*5 + campo82L + campo83L +campo81L + campo2*18 + campo7*4 + campo2*22 + campo11 + campo7*2 + campo83L)

        nomeFile = str(input('\33[1;33mDigite o nome do arquivo de saída: \33[m'))
        with open(nomeFile + str('.txt'), 'w') as arquivo:
            arquivo.write(str(cabecalho) + '\n')
            print(cabecalho)


        for linha in range(2, len(sheet_selecionada['A']) + 1):

            ContacontabilNr = sheet_selecionada['B' +'%s' % linha].value
            CentroCustoNr = sheet_selecionada['A' + '%s' % linha].value
            tipoCReBD = sheet_selecionada['C' + '%s' % linha].value
            if tipoCReBD == 'DB':
                tipoCReBD = 'DB'
            else:
                tipoCReBD = 'CR'
            descricaoNr = sheet_selecionada['D' + '%s' % linha].value
            anoMerNr = sheet_selecionada['E' + '%s' % linha].value
            valorNr = sheet_selecionada['F' + '%s' % linha].value

            # lista = str(ContacontabilNr)+str(CentroCustoNr)+str(tipoCReBD)+str(descricaoNr)+str(anoMerNr)+str(valorNr)
            # print(lista)   



            campo0 = 'L'
            campo1 = "E"
            campo2 = '0'
            def campo3():
                campo3 = "DB" or "CR"
                if campo3 == str(tipoCReBD):
                    campo3 = "01"
                else:
                    campo3 = "02"
                return campo3
            campo3 = campo3()
            campo4 = str(ContacontabilNr).rjust(10, '0')   # Conta Contábil
            campo5 = str(CentroCustoNr).rjust(5, '0')    # Centro de Custo
            campo6 = "MPF"
            campo7 = " "
            campo80 = str(campo83L+campo81L)    # Ano e Mês
            campo81 = campo81L
            campo82 = campo82L
            campo9 = str(valorNr).rjust(20, '0')   # Valor
            campo10 = (str(descricaoNr)+str(' ')+str(anoMerNr)).ljust(30, ' ')  # Descrição
            campo11 = 'FP'
            

            
            linha = [campo1 + campo2*4 + campo3 + campo4 + campo5 + campo6 + campo7 + campo80 + campo9 + campo10]
            
            # print(linha)

            with open(nomeFile + str('.txt'), 'a') as arquivo:
                for item in linha:
                    arquivo.write(str(item) + '\n')
                    print(item)

    else:
        # input('Deseja continuar? [S/N] ').upper() == 'S'
        print('Programa finalizado.')
        break