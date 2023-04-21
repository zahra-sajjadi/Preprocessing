import pandas as pd
import xlsxwriter

df = pd.read_excel('Output of the previous phase.xlsx')

wb = xlsxwriter.Workbook('output.xlsx')
ws = wb.add_worksheet()

ws.write(0, 0, 'auth')
ws.write(0, 1, 'interest')
ws.write(0, 2, 'num')

r = 1

for i in range(df.shape[0]):
    row = df.iloc[i]
    

    if row[1] in ['data base', 'Data mining','Information retrieval','Artificial intelligence']:
        ws.write(r, 0, row[0])
        ws.write(r, 1, row[1])
        ws.write(r, 2, row[2])

        r += 1

wb.close()
