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

df1 = pd.read_excel('Dinamica.xlsx', sheet_name = 'Planilha1')

df2 = pd.read_excel('Dinamica.xlsx', sheet_name = 'Planilha4')

df1['Soma'] = df1.groupby(['Fantasia Distribuidor'])['Total'].transform('sum')

df1['Valor'] = 0

df1['Meta Mensal'] = 0

for index1,row1 in df1.iterrows():
    for index2,row2 in df2.iterrows(): 
        if(row1['Fantasia Distribuidor'] == row2['Distribuidora']):
            df1.loc[index1,'Valor'] = row2['nome']
    df1.loc[index1,'Meta Mensal'] = row1['Total']*(df1.loc[index1,'Valor']/row1['Soma']) 

distribuidores = []

for i in range(len(df2)): 
    distribuidores.insert(len(distribuidores),df2.iloc[i,0])

dist_df = pd.DataFrame()  
dist_df['Distribuidores'] = distribuidores

df1 = df1.groupby('Fantasia Distribuidor').apply(pd.DataFrame.sort_values, 'Meta Mensal', ascending=False)

del df1['Soma']
del df1['Valor']
del df1['Total']

pd.set_option('display.max_rows', len(df1)) 
print(df1)


