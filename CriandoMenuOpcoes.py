# Menu opções
# https://www.youtube.com/watch?v=KowQ_UIMuI8
from time import sleep
print('\n\33[34m *Software by Anderson Marinho*\33[m')
print('\33[32m -Menu de opções:\33[m')
print('\33[32m -Olá, seja bem vindo ao menu de opções!\33[m')

def leiaInt(msg):
    while True:
        try:
            n = int(input(msg))
        except (ValueError, TypeError):
            print('\33[31mERRO! Por favor, digite um número inteiro válido.\33[m')
            continue
        except (KeyboardInterrupt):
            print('\33[31mO usuário preferiu não digitar um número.\33[m')
            return 0
        else:
            return n
def leiaFloat(msg):
    while True:
        try:
            n = float(input(msg))
        except (ValueError, TypeError):
            print('\33[31mERRO! Por favor, digite um número real válido.\33[m')
            continue
        except (KeyboardInterrupt):
            print('\33[31mO usuário preferiu não digitar um número.\33[m')
            return 0
        else:
            return n


n1 = leiaFloat('\n\33[33m -Digite 1º valor para calcular:\33[m ')
n2 = leiaFloat('\33[33m -Digite 2º valor para calcular:\33[m ')

opcao = 0
while opcao != 5:
    print('''\n\33[32m
    [ 1 ] Somar
    [ 2 ] Multiplicar
    [ 3 ] Maior
    [ 4 ] Novos números
    [ 5 ] Sair do programa\33[m''')
    opcao = int(input('\33[33m -Digite a opção desejada:\33[m '))
    if opcao == 1:
        soma = n1 + n2
        print('')
        print('\33[30m -A soma entre {} + {} é igual a {}.\33[m'.format(n1, n2, soma))
    elif opcao == 2:
        produto = n1 * n2
        print('')
        print('\33[31m -O produto entre {} x {} é igual a {}.\33[m'.format(n1, n2, produto))
    elif opcao == 3:
        if n1 > n2:
            maior = n1
        else:
            maior = n2
        print('')
        print('\33[31m -Entre {} e {} o maior valor é {}.\33[m'.format(n1, n2, maior))
    elif opcao == 4:
        print('')
        print('\33[30m -Informe os números novamente:\33[m')
        n1 = leiaFloat('\n\33[33m -Digite 1º valor para calcular:\33[m ')
        n2 = leiaFloat('\33[33m -Digite 2º valor para calcular:\33[m ')

    elif opcao == 5:
        print('')
        print('\33[30m -Finalizando...\33[m')
    else:
        print('')
        print('\33[31m -Opção inválida! Tente novamente.\33[m')
    print('')
    print('=-=' * 10)
    sleep(2)
    input('\33[33m -Pressione ENTER para continuar...\33[m')
print('\n\33[30m -Programa finalizado!\33[m')
# Fim do programa