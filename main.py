import funcoes_excel as func_e


def main():
    loop = True
    nome_pasta_trabalho = ''
    nome_planilha = ''
    acao = ''
    while loop:
        acao = input('\nSelecione o que deseja fazer.\n\
1 - Criar pasta de trabalho.\n\
2 - Abrir pasta de trabalho.\n\
3 - Selecionar planilha ativa.\n\
4 - Criar planilha.\n\
5 - Mudar nome da planilha.\n\
6 - Imprimir dados da planilha.\n\
7 - Sair.\n\
')
        match acao:
            case '1':
                aux1 = input('\nNome da nova pasta de trabalho: ')
                func_e.criar_pasta_trabalho(aux1)
            case '2':
                aux1 = input('\nNome da pasta de trabalho: ')
                nome_pasta_trabalho = func_e.abrir_pasta_trabalho(aux1)
            case '3':
                if nome_pasta_trabalho == '':
                    aux1 = input('\nNome da pasta de trabalho: ')
                    nome_pasta_trabalho = func_e.abrir_pasta_trabalho(aux1)
                    print(nome_pasta_trabalho.active)
                    print(nome_pasta_trabalho.sheetnames)
                    aux2 = input('\nSelecione a planilha ativa (Deixe em branco para selecionar a planilha ativa): ')
                    nome_planilha = func_e.planilha_ativa(nome_pasta_trabalho, aux2)
                    print(nome_planilha)
                else:
                    print(nome_pasta_trabalho.active)
                    print(nome_pasta_trabalho.sheetnames)
                    aux2 = input('\nSelecione a planilha ativa (Deixe em branco para selecionar a planilha ativa): ')
                    nome_planilha = func_e.planilha_ativa(nome_pasta_trabalho, aux2)
                    print(nome_planilha)
            case '4':
                if nome_pasta_trabalho == '':
                    nome_pasta_trabalho = input('\nNome da pasta de trabalho: ')
                    nome_planilha = input('\nNome da planilha: ')
                    func_e.criar_planilha(nome_pasta_trabalho, nome_planilha)
                else:
                    nome_planilha = input('\nNome da planilha: ')
                    func_e.criar_planilha(nome_pasta_trabalho, nome_planilha)
            case '5':
                if nome_pasta_trabalho == '':
                    aux1 = input('\nNome da pasta de trabalho: ')
                    nome_pasta_trabalho = func_e.abrir_pasta_trabalho(aux1)
                    print(nome_pasta_trabalho.active)
                    print(nome_pasta_trabalho.sheetnames)
                    aux2 = input('\nSelecione a planilha ativa (Deixe em branco para selecionar a planilha ativa): ')
                    nome_planilha = func_e.planilha_ativa(nome_pasta_trabalho, aux2)
                    print(nome_planilha)
                    linha = input('Qual linha: ')
                    func_e.criar_cabo_mao(nome_pasta_trabalho, nome_planilha, linha)
                else:
                    print(nome_pasta_trabalho.active)
                    print(nome_pasta_trabalho.sheetnames)
                    aux2 = input('\nSelecione a planilha ativa (Deixe em branco para selecionar a planilha ativa): ')
                    nome_planilha = func_e.planilha_ativa(nome_pasta_trabalho, aux2)
                    print(nome_planilha)
                    linha = input('Qual linha: ')
                    func_e.criar_cabo_mao(nome_pasta_trabalho, nome_planilha, linha)
            case '6':
                continue
            case _:
                loop = False
                # break
        # pasta_trabalho = func_e.abrir_pasta_trabalho('excel/Lista_cabos_mao')
        # planilha = func_e.planilha_ativa(pasta_trabalho, '')
        # func_e.imprime(1, 30, 1, 11, planilha)
        # print(pasta_trabalho.sheetnames)
        # planilha = func_e.planilha_ativa(pasta_trabalho, 'CABOS ENTRADA')
        # func_e.imprime(1, 30, 1, 11, planilha)
        # func_e.criar_planilha('excel/Lista_cabos_mao', pasta_trabalho, 'teste2')
        # func_e.mudar_nome_planilha('excel/Lista_cabos_mao', pasta_trabalho, 'teste2', 'pp tiltado')
        # print(pasta_trabalho.sheetnames)


if __name__ == '__main__':
    main()
