import xlrd, xlwt, re
from striprtf.striprtf import rtf_to_text
from util import indica_promotor, seleciona_arquivo

relatorio_rtf = seleciona_arquivo('Selecione a Pauta Analítica (arquivo rtf)')
relatorio_xls = seleciona_arquivo('Selecione a Pauta Sintética (arquivo xls)')

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
    data = data_audiencia.split()[0]
    horario = data_audiencia.split()[1]
    n_controle = processo_controle[processo.group()][0]
    promotor = processo_controle[processo.group()][1]
    pauta_audiencias[processo.group()] = [data, horario, n_controle, audiencia, promotor]
  except:
    pass

audiencias = []
for aud in pauta_audiencias.values():
  audiencias.append(aud)

# Gravação da planilha Excel

workbook = xlwt.Workbook()
sheet = workbook.add_sheet("Audiências")
cabecalho = ["Data", "horario", "N_controle", "Informacoes", "Promotor"]

for col, value in enumerate(cabecalho):
    sheet.write(0, col, value)

for row, row_data in enumerate(audiencias, start=1):
    for col, value in enumerate(row_data):
        sheet.write(row, col, value)

workbook.save("audiencias.xls")  

print('Programa concluído!')