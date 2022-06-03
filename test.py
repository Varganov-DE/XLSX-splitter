import openpyxl
import os

wb = openpyxl.load_workbook('Specification_test.xlsx')
type(wb)

os.chdir(path = 'e:\site\coding') #прописываем путь к папке в которой лежит файл
wb_new = openpyxl.load_workbook('Specification_test_1.xlsx')