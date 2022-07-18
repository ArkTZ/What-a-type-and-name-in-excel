import os
import openpyxl
from pathlib import Path


def excel_res():
    path = '.'
    book = openpyxl.Workbook()
    sheet = book.active
    sheet['A1'] = 'Номер строки'
    sheet['B1'] = 'Папка в которой лежит файл'
    sheet['C1'] = 'Название файла'
    sheet['D1'] = 'Расширение файла'

    filelist = []

    for root, dirs, files in os.walk(path):
        for file in files:
            filelist.append(os.path.join(root, file))

    row = 2
    try:
        for name in filelist:
            if not name.startswith('.\.'):
                sheet[row][0].value = row - 1
                if os.path.dirname(name) == ".":
                    sheet[row][1].value = "testtask"
                else:
                    sheet[row][1].value = str(os.path.split(os.path.dirname(name))[-1])
                try:
                    index = os.path.basename(name).index('.')
                    sheet[row][2].value = str(os.path.basename(name)[:index])
                except:
                    sheet[row][2].value = str(os.path.basename(name))
                if Path(name).suffix != '':
                    sheet[row][3].value = str(Path(name).suffix)
                else:
                    sheet[row][3].value = "-"

                print(name)
                print(Path(name).suffix)
                row += 1
    except PermissionError:
        print('Не забудь выйти из Экселя')
    book.save("result.xlsx")


excel_res()
