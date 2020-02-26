# Rubens_COOL
from openpyxl import Workbook
from openpyxl.styles import Font
wb = Workbook()
ws = wb.active


def lineLetter(line):
    a = [('А', 0), ('Б', 0), ('В', 0), ('Г', 0), ('Д', 0), ('Е', 0), ('Ё', 0), ('Ж', 0), ('З', 0), ('И', 0), ('Й', 0), ('К', 0), ('Л', 0), ('М', 0), ('Н', 0), ('О', 0), ('П', 0), ('Р', 0), ('С', 0), ('Т', 0), ('У', 0), ('Ф', 0), ('Х', 0), ('Ц', 0), ('Ч', 0), ('Ш', 0), ('Щ', 0), ('Ъ', 0), ('Ы', 0), ('Ь', 0), ('Э', 0), ('Ю', 0), ('Я', 0), ('.', 0), (',', 0), (' ', 0)]
    for letter in line:
        if letter != '\n':
            print(letter, end="")
            for i in range(len(a)):
                if letter == a[i][0] or letter == a[i][0].swapcase():
                    a[i] = (letter.upper(), a[i][1] + 1)
                    break
    return a


fin = open('text.txt', 'r')
lengthF = len(fin.read())
fin = open('text.txt', 'r')
lengthF = lengthF // 24
for numLine in range(1, 24):
    cntLater = [0] * 36
    if numLine == 23:
        line = fin.read()
    else:
        line = fin.read(lengthF)
        while line[-1] != ' ':
            line = line + fin.read(1)
    print(numLine, end=". ")
    cntLater = lineLetter(line)
    print()
    print(cntLater)
    ws.cell(row=numLine+1, column=1, value=numLine).font = Font(bold=True, color='DC143C')
    for i in range(len(cntLater)):
        ws.cell(row=numLine+1, column=i + 2, value=cntLater[i][1])
for i in range(36):
    ws.cell(row=1, column=i+2, value=cntLater[i][0]).font = Font(bold=True, color='DC143C')
wb.save('result.xlsx')
input('нажмите Enter для выхода...')
