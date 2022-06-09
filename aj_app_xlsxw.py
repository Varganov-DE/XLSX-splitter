from appJar import gui
from pathlib import Path
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side, Protection
from os.path import join, abspath

class NotAllData(Exception): # проверка наличия книги и данных в ней
    pass

data_path = ("Спецификация ОВ.xlsx") # относительный путь
data_path = abspath(data_path) # абсолютный путь

# задаём параметры работы с файлом:
wb = load_workbook(filename = data_path, data_only = True, read_only = True)

wsn = wb.sheetnames # присваивает список листов в книге
print(wsn)

wsdata = None

for i in wsn:
    if wb[i]['D1'].value == 'Код, оборудования, изделия, материала': #проверка значения в ячейке D1
        wsdata = i
if wsdata == None:
    raise NotAllData('Нет данных в указанном столбце')

# запись в переменную shapka названий всех столбцов:

ws = wb[wsdata]
shapka = [cell.value for cell in next(
    ws.iter_rows(min_row=1, min_col=1, max_row=1, max_col=ws.max_column))]

# Цикл сбора всех данных в промаркированных строках, и подсчёт кол-ва, по каждому маркеру:

mandata = {} # создаём пустой словарь

# бежим по всем ячейкам в каждой строке начиная со второй(min_row=2) и с первой колонки(min_col=1), пока в них есть данные(max_row=ws.max_row, max_col=ws.max_column) читаем их значения и записываем в кортеж row: 
for row in ws.iter_rows(min_row=2, min_col=1, max_row=ws.max_row,
                        max_col=ws.max_column): 
    
    if len(row) > 0:
        marker = row[5].value # записываем в переменную значение ячейки в 4-м столбце(3-м по индексу)
        if marker is not None: # если данные есть в ячейке, то:
            markerdata = [cell.value for cell in row] # записываем в переменную список из значений ячеек в строке
            
            if marker not in mandata: # если в словаре нет ещё маркера, то:
                mandata[marker] = []
            mandata[marker].append(markerdata) # добавляем в словарь к каждому индексу(маркеру), значение(список из значений ячеек в этой строке)

for marker in mandata: # для каждого индекса(маркера) в словаре:
    print(f'По маркеру {marker}, количество позиций: {len(mandata[marker])}') # выводим длинну кортежа(кол-во строк)

wb.close # закрываем рабочую книгу

# Создание отдельной книги xlsx под каждый маркер материалов

for marker in mandata: # для каждого индекса словаря
    exname, *_ = marker.split()
    wb = Workbook()
    ws = wb.active
    ws.title = "Лист1"

    ws.append(shapka) # Добавляем шапку
    for row in mandata[marker]:  #  для каждого индекса(маркера):
        ws.append(row) # добавляем список

    # создайм файл и записываем в него данные из словаря:

    exfilname = join('.', 'Data', ('Заявка ' + exname + '.xlsx')) # прописываем путь и название сохраняемого файла
    exfilname = abspath(exfilname)
    print(exfilname)

    wb.save(exfilname) # сохраняем файл
    wb.close # закрываем файл

    # копируем стиль ячеек из исходного документа в новый:



print('\nДанные по маркеру материалов обработаны')
print('Заявки созданы')