import xlrd, xlwt, re
from striprtf.striprtf import rtf_to_text

relatorio_rtf = "Fevereiro (Pauta Analítica).rtf"
relatorio_xls = "Fevereiro (Pauta sintética - Com número de controle).xls"

def indica_promotor(n_controle):
    """ Retorna o nome do PJ responsável pelo caso, de acordo com o número de controle do processo. """
    # Regra de distribuição:
    # 4º PROMOTOR DE JUSTIÇA:Feitos de finais 17, 19, 21, 23, 25, 27 e 29 da 4ª Vara Criminal;
    # 6º PROMOTOR DE JUSTIÇA:feitos de finais 31, 33, 35, 37, 39, 41 e 43 da 4ª Vara Criminal;
    # 7º PROMOTOR DE JUSTIÇA:feitos de finais 45, 47, 49, 51, 53, 55 e 57, da 4ª Vara Criminal;
    # 9° PROMOTOR DE JUSTIÇA:feitos de finais 59, 61, 63, 65, 67, 69 e 71 da 4ª Vara Criminal;
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
   
    final = int(n_controle[-2:]) # 2 últimos caracteres da string informada
   
    for chave in atribuicoes:
        if final in atribuicoes[chave]:
            return chave
    return "pj17"

# Leitura do arquivo texto rtf

with open(relatorio_rtf, encoding="latin-1") as f:
  conteudo = f.read()
  texto = rtf_to_text(conteudo)

# Leitura da planilha xls 

pasta_excel = xlrd.open_workbook(relatorio_xls, encoding_override="latin-1")
planilha = pasta_excel.sheet_by_index(0)

# Obtendo dados da planilha
# Relacionando em dicionário o número do processo ao número de controle bem como ao PJ com atribuições para atuar no feito

processo_controle = {} 

for row_index in range(planilha.nrows):
  controle = planilha.row_values(row_index)[9:11][1]
  if (controle != 'Número de controle') & (len(controle)>10):
    elemento = planilha.row_values(row_index)[9:11]
    processo_controle[f'{elemento[0]}'] = [f'{elemento[1]}', f'{indica_promotor(controle)[-2:]} PJ'] 

print('Os números de processo foram relacionados com os de com controle e PJs com atribuições para atuar no feito.')

# processo_controle = {'1501514-91.2022.8.26.0599': ['2022/001339', '06 PJ'], '1501535-67.2022.8.26.0599': ['2022/001332', '17 PJ']}

# Procura audiências no arquivo rtf com RegEx e divide a string

padrao_data_hora = re.compile(r'\d{2}/\d{2}/\d{2} \d{2}:\d{2}')

def acrescentar_caractere(match):
    return "-xx-" + match.group()

texto_atualizado = re.sub(padrao_data_hora, acrescentar_caractere, texto)
lista_audiencias = texto_atualizado.split("-xx-")

print(f'Foram encontradas {len(lista_audiencias)} audiências na Pauta Analítica (arquivo "rtf").')

# Cria o dicionário de audiências com o número de controle e PJ com atribuições para atuar no feito

padrao_processo = re.compile(r'\d{7}-\d{2}.\d{4}.8.26.\d{4}')
pauta_audiencias = {}

for audiencia in lista_audiencias:
  data_hora = re.search(padrao_data_hora, audiencia)
  processo = re.search(padrao_processo, audiencia)
  try:
    data_audiencia = data_hora.group()
    n_controle = processo_controle[processo.group()][0]
    promotor = processo_controle[processo.group()][1]
    pauta_audiencias[processo.group()] = [data_audiencia, n_controle, audiencia, promotor]
  except:
    pass

audiencias = []
for aud in pauta_audiencias.values():
  audiencias.append(aud)

# Gravação da planilha Excel

workbook = xlwt.Workbook()
sheet = workbook.add_sheet("Audiências")
cabecalho = ["Data_hora", "N_controle", "Informacoes", "Promotor"]

for col, value in enumerate(cabecalho):
    sheet.write(0, col, value)

for row, row_data in enumerate(audiencias, start=1):
    for col, value in enumerate(row_data):
        sheet.write(row, col, value)

workbook.save("audiencias.xls")  