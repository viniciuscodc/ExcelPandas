import pandas as pd
from openpyxl import load_workbook

def write_excel(filename,sheetname,dataframe):
    with pd.ExcelWriter(filename, engine='openpyxl', mode='a') as writer: 
        workBook = writer.book
        try:
            workBook.remove(workBook[sheetname])
        except:
            print("Worksheet does not exist")
        finally:
            dataframe.to_excel(writer, sheet_name=sheetname,index=False)
            writer.save()

df = pd.read_excel('Dados.xlsx', sheet_name = 'Plan1')

df = df[df.Operação != 'BONIFICACAO']

df = df[df['Família'].str.contains("sticks|palitos", case=False, na=False)]

#na ignores nan values
df = df[df['Família'].str.contains("sticks|palitos", case=False, na=False)]

df_sum = df.groupby(['Fantasia Distribuidor','Vendedor Novo'])['Valor SellOut'].transform('sum')

df['Total'] = df_sum

#new_df = df_filtered.drop_duplicates(subset=['Fantasia Distribuidor'])
df = df.drop_duplicates(subset=['Vendedor Novo'])

del df['Valor SellOut']
del df['Razão Distribuidor']
del df['Família']

print(df)

write_excel('Dinamica.xlsx','Planilha1',df)
