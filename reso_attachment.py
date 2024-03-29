import docx2txt
import xlsxwriter
import xlwt
from os import listdir
from os import chdir
from parse import striprtf


def add_first_row(_ws):
    _ws.write(0, 0, 'SURNAME')
    _ws.write(0, 1, 'FIRST_NAME')
    _ws.write(0, 2, 'SEC_NAME')
    _ws.write(0, 3, 'SEX')
    _ws.write(0, 4, 'DATE_BIRTH')
    _ws.write(0, 5, 'POLICY')
    _ws.write(0, 6, 'YEAR')
    _ws.write(0, 7, 'DATE_FRM')
    _ws.write(0, 8, 'DATE_TO')
    _ws.write(0, 9, 'DATE_CNCL')


def write_to_excel(scroll, _ws, _start_period, _finish_period, row, n, start_col=0):
    print('выполняется запись в exel, строка = ', row)
    name_people = scroll.get('fio')[n].split()
    print(name_people)
    try:
        if (name_people[2][-1].lower() == 'ч'):
            gender = 0
        else:
            gender = 1
    except Exception:
        print('товарищ без отчества, пол определить не можем - пусть будет мальчиком')
        gender = 0

    _ws.write(row, start_col, name_people[0])
    _ws.write(row, start_col + 1, name_people[1])
    try:
        _ws.write(row, start_col + 2, name_people[2])
    except Exception:
        print("У вот него -> ", name_people, " нет отчества!")

    _ws.write(row, start_col + 3, gender)
    _ws.write(row, start_col + 4, str(scroll.get('birth')[n]))
    _ws.write(row, start_col + 5, scroll.get('polis')[n])
    _ws.write(row, start_col + 6, scroll.get('birth')[n].split('.')[2])
    _ws.write(row, start_col + 7, _start_period)

    _ws.write(row, start_col + 8, _finish_period)
    _ws.write(row, start_col + 9, _finish_period)
    print('     записали ', scroll.get('fio')[n], '\n')


def my_file(ws, _file_name, _row):

    # Открытие файла от страховой
    my_text = docx2txt.process(_file_name)
    my_text = my_text.split('\n')
    #print("Содержимое файла:")
    #print(my_text)

    # нашли период страхования
    for idx, u in enumerate(my_text):
        if "Период страхования:" in u:
            people = {'polis': [], 'fio': [], 'birth': []}
            print('обрабатываем таблицу')
            start_period = u.split()[2]
            finish_period = u.split()[4]
            print('Период действия: ', start_period, finish_period)
            continue

        # нашли начало списка с фамилиями
        if "ФИО застрахованного" in u:
            st_tab = idx + 8  # начало таблицы
            people = {'polis': [], 'fio': [], 'birth': []}
            i = 0
            while (i < 2000):
                try:
                    int(my_text[st_tab + i * 12])
                except:
                    # print('заканчивается на номере', num)
                    break
                people['polis'].append(my_text[st_tab + 2 + i * 12])
                people['fio'].append(my_text[st_tab + 4 + i * 12])
                people['birth'].append(my_text[st_tab + 6 + i * 12])
                i += 1

            print(people)
            n = 0

            print('Записываем ', i, 'людей')
            while n < i:
                write_to_excel(people, ws, start_period, finish_period, _row, n, start_col)
                _row += 1
                n += 1
    return _row

chdir("C:\\Users\\baa\\PycharmProjects\\convert_lists_from_insurance\\DOCX\\Прикрепить")
print('\n ***Обработка прикреплений*** \n')
for name in listdir():
    print(name)

# Создали файл
people = {'polis': [], 'fio': [], 'birth': []}
wb = xlwt.Workbook()
ws = wb.add_sheet("Sheet1")   # создали лист
add_first_row(ws)             # и для листа создали "шапку"

_row = 1  # строка
start_col = 0

for name in listdir():
    print("...")
    print("Обработка файла ", name)
    print("...")
    _row = my_file(ws, name, _row)
        #print("последняя строка = ", _row)
    wb.save("RESO_.xls")
