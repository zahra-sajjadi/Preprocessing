import pandas as pd
import xlsxwriter

df = pd.read_excel('newoutput.xlsx')
df2 = pd.read_excel('result-1000.xlsx')

s1 = set(df2['auth1'])
s2= set(df2['auth2'])
s3 = set(df['auth'])

s4 = s1.union(s2)
s5 = s3.difference(s4)
##print(s5)

authors = list(s4)

authors.sort()

counter = 1

for item in authors:
    df2.replace(item, counter, inplace=True)
    df.replace(item, counter, inplace=True)
    
    counter += 1

authors = list(s5)

authors.sort()

for item in authors:
    df.replace(item, counter, inplace=True)
    
    counter += 1

df.sort_values(by=['auth'], inplace=True)
df2.sort_values(by=['auth1'], inplace=True)

wb = xlsxwriter.Workbook('New-OutPut3.xlsx')
ws = wb.add_worksheet()

ws.write(0, 0, 'auth')
ws.write(0, 1, 'interest')

r = 1

for i in range(df.shape[0]):
    row = df.iloc[i]

    ws.write(r, 0, row[0])
    ws.write(r, 1, row[1])

    r += 1

wb.close()

wb = xlsxwriter.Workbook('result2.xlsx')
ws = wb.add_worksheet()

ws.write(0, 0, 'auth1')
ws.write(0, 1, 'auth2')
ws.write(0, 2, 'num')

r = 1

for i in range(df2.shape[0]):
    row = df2.iloc[i]

    ws.write(r, 0, row[0])
    ws.write(r, 1, row[1])
    ws.write(r, 2, row[2])

    r += 1

wb.close()
