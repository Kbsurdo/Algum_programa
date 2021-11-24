from copy import copy
import funcoes_excel as fe

wb = fe.criar_pasta_trabalho()
print(f"{fe.planilha_ativa(wb)}")
fe.mudar_nome_planilha(wb, 'Sem ideia de nome')
print(f"{fe.planilha_ativa(wb)}")
