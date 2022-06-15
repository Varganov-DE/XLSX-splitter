from appJar import gui
from pathlib import Path
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side, Protection
from os.path import join, abspath
 
# Определение всех необходимых для обработки файла функций
 
def split_pages(data_path, name_of_marker_cell):
    class NotAllData(Exception): 
        pass

    # Аргумент №1: data_path - название файла и путь к нему

    data_path = input("Введите название файла(вида: название.xlsx):  ")
    #data_path = ("Спецификация ОВ.xlsx") # относительный путь
    #data_path = abspath(data_path) # абсолютный путь

    # задаём параметры работы с файлом:
    wb = load_workbook(filename = data_path, data_only = True, read_only = True)

    wsn = wb.sheetnames # присваивает список листов в книге
    print(f"В файле \"{data_path}\", есть листы: {wsn}.")
    wsdata = None

    # Аргумент №2: name_of_marker_cell - координаты ячейки с маркером

    name_of_marker_cell = input("Введите координаты ячейки с маркером(напр E1): ")
    #name_of_marker_cell = 'E1'

    number_cell = name_of_marker_cell
    number_cell = number_cell[:-1]
    number_cell = ord(number_cell.lower()) - 97

    for i in wsn:
        if wb[i][name_of_marker_cell].value != None: #проверка есть ли значение в указанной ячейке с маркером
            wsdata = i
    if wsdata == None:
        raise NotAllData('Нет данных в указанной колонке')

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
            marker = row[number_cell].value # записываем в переменную значение ячейки в 4-м столбце(3-м по индексу)
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

    input('\nНажмите ENTER, что-бы закрыть окно.')
    
    
    """ Проверяем, если введенные пользователем значения являются правильными.
 
    Аргументы:
        input_file: Исходный PDF файл
        output_dir: Директория для хранения готового файла
        range: File Строка, содержащая число копируемых страниц: : 1-3,4
        file_name: Имя вывода готового PDF файла
 
    Возвращает:
        True, если ошибка и False, если нет
        Список сообщений об ошибке
    """
    errors = False
    error_msgs = []
 
    # Проверяет, выбран ли xlsx файл
    if Path(data_path).suffix.upper() != ".xlsx":
        errors = True
        error_msgs.append("Please select a PDF input file")
 
    # Проверяет действительный каталог
    if not(Path(output_dir)).exists():
        errors = True
        error_msgs.append("Please Select a valid output directory")
 
    # Проверяет название файла
    if len(data_path) < 1:
        errors = True
        error_msgs.append("Please enter a file name")
 
    return(errors, error_msgs)
 
 
def press(button):
    """ Выполняет нажатие кнопки
 
    Аргументы:
        button: название кнопки. Используем названия Выполнить или Выход
    """
    if button == "Process":
        src_file = app.getEntry("Input_File")
        dest_dir = app.getEntry("Output_Directory")
        page_range = app.getEntry("Page_Ranges")
        out_file = app.getEntry("Output_name")
        errors, error_msg = validate_inputs(src_file, dest_dir, page_range, out_file)
        if errors:
            app.errorBox("Error", "\n".join(error_msg), parent=None)
        else:
            split_pages(src_file, page_range, Path(dest_dir, out_file))
    else:
        app.stop()
 
 
# Создать окно пользовательского интерфейса
app = gui("PDF Splitter", useTtk=True)
app.setTtkTheme("default")
app.setSize(500, 200)
 
# Добавить интерактивные компоненты
app.addLabel("Choose Source PDF File")
app.addFileEntry("Input_File")
 
app.addLabel("Select Output Directory")
app.addDirectoryEntry("Output_Directory")
 
app.addLabel("Output file name")
app.addEntry("Output_name")
 
app.addLabel("Page Ranges: 1,3,4-10")
app.addEntry("Page_Ranges")
 
# Связать кнопки с функцией под названием press
app.addButtons(["Process", "Quit"], press)
 
# Запуск интерфейса
app.go()