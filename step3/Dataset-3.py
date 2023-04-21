import pandas as pd
import xlsxwriter

df = pd.read_excel('Book1-Full-DBLP(4 Interest).xlsx', header=None)
df.columns = ['A']

wb = xlsxwriter.Workbook('result-1000.xlsx')
ws = wb.add_worksheet()

ws.write(0, 0, 'auth1')
ws.write(0, 1, 'auth2')
ws.write(0, 2, 'num')

row = 1

for i in range(5):
    df2 = pd.read_excel('AMiner-Coauthor.xlsx', sheet_name=i)

    for j in range(df2.shape[0]):
        isfound = False

        auth1 = df2.iat[j, 0]
        auth2 = df2.iat[j, 1]
        num = df2.iat[j, 2]

        if not((df.query('A == @auth1')).empty) and not((df.query('A == @auth2')).empty):
            isfound = True

        if isfound:
            ws.write(row, 0, auth1)
            ws.write(row, 1, auth2)
            ws.write(row, 2, num)

            row += 1

            if row == 1048576:
                ws = wb.add_worksheet()

                ws.write(0, 0, 'auth1')
                ws.write(0, 1, 'auth2')
                ws.write(0, 2, 'num')

                row = 1

wb.close()
