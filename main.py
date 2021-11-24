import pandas as pd
from pandas.core.frame import DataFrame

WB_PATH = './data/planilha-para-resolucao-de-exercicios.xlsx'

def main(SHEET):
   df = pd.read_excel(WB_PATH, sheet_name=SHEET)
   df = df.dropna(axis='columns', how='all')
   df = df.dropna(how='all')
   df.columns = df.iloc[0]

   #Dados da Aeroporto
   if SHEET == 'Aeroporto':
      df = df.drop('TOTAL', axis='columns')
      df = df.reset_index(drop=True)

      drop_index = []
      for index, row in df.iterrows():

         if isinstance(row['ANO'], str):
            drop_index.append(index)

         if isinstance(row['MES'], str):
            drop_index.append(index)

         if isinstance(row['NUMERO DE POUSOS'], str):
            drop_index.append(index)

         if isinstance(row['NUMERO DE DECOLAGEM'], str):
            drop_index.append(index)

         if isinstance(row['NUMERO DE PASSAGEIROS EMBARQUE'], str):
            drop_index.append(index)

         if isinstance(row['NUMERO DE PASSAGEIROS DESEMBARQUE'], str):
            drop_index.append(index)

      drop_index = list(dict.fromkeys(drop_index))
      df = df.drop(drop_index)
      df = df.reset_index(drop=True)

      for index, row in df.iterrows():
         year_col = 'ANO'
         year_cell = row[year_col]

         if pd.isna(year_cell):
            df.loc[index, year_col] = df.loc[index + 1, year_col]

      df = df.fillna(0)

      df.to_excel('./data/Aeroporto - dados tratados.xlsx', sheet_name='Aeroporto', index=False)
      print("Planilha gerada com sucesso!")

   #Dados do Rodoviária
   elif SHEET == 'Rodoviaria':
      df = df.drop('TOTAL', axis='columns')
      df = df.reset_index(drop=True)

      drop_index = []
      for index, row in df.iterrows():

         if isinstance(row['ANO'], str):
            drop_index.append(index)

         if isinstance(row['MES'], str):
            drop_index.append(index)

         if isinstance(row['NUMERO DE SAIDA DE ONIBUS'], str):
            drop_index.append(index)

         if isinstance(row['NUMERO DE CHEGADA DE ONIBUS'], str):
            drop_index.append(index)

         if isinstance(row['NUMERO DE PASSAGEIROS EMBARQUE'], str):
            drop_index.append(index)

         if isinstance(row['NUMERO DE PASSAGEIROS DESEMBARQUE'], str):
            drop_index.append(index)

      drop_index = list(dict.fromkeys(drop_index))
      df = df.drop(drop_index)
      df = df.reset_index(drop=True)

      for index, row in df.iterrows():
         year_col = 'ANO'
         year_cell = row[year_col]

         if pd.isna(year_cell):
            df.loc[index, year_col] = df.loc[index + 1, year_col]

      df = df.fillna(0)

      df.to_excel('./data/Rodoviaria - dados tratados.xlsx', sheet_name='Rodoviaria', index=False)
      print("Planilha gerada com sucesso!")

   #Dados do Apresentação Secretário
   elif SHEET == 'Apresentacao Secretario':   
      df = df.drop('TOTAL', axis='columns')
      df = df.reset_index(drop=True)

      drop_index = []
      for index, row in df.iterrows():

         if isinstance(row['ANO'], str):
            drop_index.append(index)

         if isinstance(row['MES'], str):
            drop_index.append(index)

         if isinstance(row['NUMERO DE VOOS REGULARES'], str):
            drop_index.append(index)

         if isinstance(row['NUMERO DE CHARTER'], str):
            drop_index.append(index)

      drop_index = list(dict.fromkeys(drop_index))
      df = df.drop(drop_index)
      df = df.reset_index(drop=True)

      for index, row in df.iterrows():
         year_col = 'ANO'
         year_cell = row[year_col]

         if pd.isna(year_cell):
            df.loc[index, year_col] = df.loc[index + 1, year_col]

      df = df.fillna(0)

      df.to_excel('./data/Apresentacao Secretario - dados tratados.xlsx', sheet_name='Apresentacao Secretario', index=False)
      print("Planilha gerada com sucesso!")


print('Selecione uma planilha abaixo:')
print('1 - Aeroporto')
print('2 - Rodoviária')
print('3 - Apresentação Secretário')
selected_option = int(input('Digite o número que gostaria de tratar: '))
if selected_option == 1:
   valid = True
   selected_sheet = 'Aeroporto'
elif selected_option == 2:
   valid = True
   selected_sheet = 'Rodoviaria'
elif selected_option == 3:
   valid = True
   selected_sheet = 'Apresentacao Secretario'
else:
   print('Opção invalida!')
   valid = False

if valid:
   main(selected_sheet)