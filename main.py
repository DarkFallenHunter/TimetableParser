import os
import openpyxl as opx
from pprint import pprint
from modules.parser import TimetableParser
from modules.db import TimetableDb


PROJ_DIR = os.path.dirname(os.path.abspath(__file__))
XLSX_DIR = os.path.join(PROJ_DIR, '..', 'xlsx')
TEACHERS_LIST = [
    'Аншина М.Л.',
    'Баранюк В.В.',
    'Басок Б.М.',
    'Бескин А.Л.',
    'Бирюкова А.А.',
    'Володина А.М.',
    'Воронцов Ю.А.',
    'Гданский Н.И.',
    'Головин С.А.',
    'Григорьев В.К.',
    'Грушицын А.С.',
    'Гусев К.В.',
    'Данилкин Ф.А.',
    'Зубарев И.В.',
    'Копылова А.В.',
    'Леонтьев А.С.',
    'Лоскутников О.В.',
    'Миронов А.Н.',
    'Михайлова Е.К.',
    'Овчинников М.А.',
    'Петренко А.А.',
    'Рысин М.Л.',
    'Синицын И.В.',
    'Скворцова Л.А.',
    'Смольянинова В.А.',
    'Сыромятников В.П.',
    'Филатов А.С.',
    'Борзых Н.Ю.',
    'Иванов М.Е.',
    'Степанов П.С.',
    'Макаревич А.Д.',
    'Коновалова С.В.',
    'Хлебникова В.Л.',
    'Демидова Л.А.',
    'Панов А.В.'
]

DB_PARAMS = {
    'user': 'postgres',
    'password': '123123',
    'dbname': 'mosit',
    'schema': 'timetable'
}


def main():
    xlsx_files = os.listdir(XLSX_DIR)
    workbooks = []

    print(xlsx_files)
    for file_name in xlsx_files:
        file_path = os.path.join(XLSX_DIR, file_name)
        if os.path.isfile(file_path) and \
           file_name.endswith('.xlsx') and \
           not file_name.startswith('~$'):
            print("Загрузка книги {}...".format(file_name))
            workbooks.append(opx.load_workbook(file_path))

    parser = TimetableParser(opx.cell.MergedCell, TEACHERS_LIST)
    db = TimetableDb(DB_PARAMS)

    for index, workbook in enumerate(workbooks):
        print('Обработка книги {}...'.format(xlsx_files[index]))
        result = parser.parse_sheet(workbook['Лист1'])
        pprint(result)
        print('Загрузка данных из книги {} в бд...'.format(xlsx_files[index]))
        db.insert_classes(result)
        print('Обработка книги {} завершена.'.format(xlsx_files[index]))


if __name__ == "__main__":
    main()
