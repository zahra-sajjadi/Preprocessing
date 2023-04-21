import xlsxwriter
import re

file_ = open(r"result.txt")

filename_ = input("enter excel file name: ")
workbook = xlsxwriter.Workbook(f'{filename_}'+'.xlsx')

sheet = workbook.add_worksheet()

name = ''
data = []
i = 0
index = ''

for lines in file_:
    if '#index' in lines:
        l = len(lines)
        index = str(lines[7:l-1])
        print(index)
        

    if '#n' in lines:
        l = len(lines)
        name = str(lines[2:l-2])

    if '#t' in lines:
        data = lines[2:].split(';')
        for interests in data:
            interests = ''.join(interests.split())
            sheet.write(i, 0, index)
            sheet.write(i, 1, name)
            sheet.write(i, 2, interests)
            i += 1
    # for interests in data:
    #         interests = ''.join(interests.split())
    #         sheet.write(i, 0, name)
    #         sheet.write(i, 1, interests)
    #         i += 1


workbook.close()
