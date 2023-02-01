# Funções para atribuir processos aos Promotor com atribuições para atuar no feito.

def indica_promotor(n_controle):
    """ Retorna o nome do PJ responsável pelo caso, de acordo com o número de controle do processo.
        Considera as atribuições dos PJ atuantes na 4a. Vara Criminal de Piracicaba """
    # Regra de distribuição:
    # 4º PROMOTOR DE JUSTIÇA: feitos de finais 17, 19, 21, 23, 25, 27 e 29 da 4ª Vara Criminal;
    # 6º PROMOTOR DE JUSTIÇA: feitos de finais 31, 33, 35, 37, 39, 41 e 43 da 4ª Vara Criminal;
    # 7º PROMOTOR DE JUSTIÇA: feitos de finais 45, 47, 49, 51, 53, 55 e 57 da 4ª Vara Criminal;
    # 9° PROMOTOR DE JUSTIÇA: feitos de finais 59, 61, 63, 65, 67, 69 e 71 da 4ª Vara Criminal;
    # 10° PROMOTOR DE JUSTIÇA:feitos de finais 73, 75, 77, 79, 81, 83 e 85 da 4ª Vara Criminal;
    # 11° PROMOTOR DE JUSTIÇA:feitos de finais 87, 89, 91, 93, 95, 97 e 99 da 4ª Vara Criminal;
    # 17° PROMOTOR DE JUSTIÇA:feitos de finais pares e finais 01, 03, 05, 07, 09, 11, 13 e 15 da 4ª Vara Criminal.
    
    atribuicoes = {
        "pj04": [17, 19, 21, 23, 25, 27, 29], 
        "pj06": [31, 33, 35, 37, 39, 41, 43], 
        "pj07": [45, 47, 49, 51, 53, 55, 57], 
        "pj09": [59, 61, 63, 65, 67, 69, 71], 
        "pj10":[73, 75, 77, 79, 81, 83, 85], 
        "pj11":[87, 89, 91, 93, 95, 97, 99]
        } # O pj17 se relaciona a todos os demais finais.
   
    final = int(n_controle[-2:]) # últimos 2 caracteres do número de controle informado
   
    for chave in atribuicoes:
        if final in atribuicoes[chave]:
            return chave
    return "pj17"

# def indica_promotor(n_controle):
#     """ Retorna o nome do PJ responsável pelo caso, de acordo com o número de controle do processo.
#         Considera as atribuições dos PJ atuantes na 1a. Vara Criminal de Piracicaba """
#     # Regra de distribuição:
#     # 6º PROMOTOR DE JUSTIÇA: Feitos de finais 1, 2, 3, 9 e 0 da 1ª Vara Criminal;
#     # 11º PROMOTOR DE JUSTIÇA:Feitos de finais 4, 5, 6, 7 e 8 da 1ª Vara Criminal;

#     atribuicoes = {
#         "pj06": [1, 2, 3, 9, 0], 
#         } # O pj17 se relaciona a todos os demais finais.

#     final = int(n_controle[-1:]) # último caracter do número de controle informado

#     for chave in atribuicoes:
#         if final in atribuicoes[chave]:
#             return chave
#     return "pj11"


# Seleção de arquivo com o Tkinter

import tkinter as tk
from tkinter import filedialog

def seleciona_arquivo(title):
    root = tk.Tk()
    root.withdraw() # Esconde a janela root
    file_path = filedialog.askopenfilename(initialdir = "/", title = title)
    return file_path