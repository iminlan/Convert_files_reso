import docx2txt
import xlsxwriter
import xlwt
from os import listdir
from os import chdir
from parse import striprtf

print('\n ***Обработка откреплений*** \n')
print('Список файлов в директории:')
chdir("C:\\Users\\baa\\PycharmProjects\\convert_lists_from_insurance\\DOCX\\Закрыть")

for name in listdir():
    print(name)


def add_first_row(_ws):
    _ws.write(0, 0, 'POLICY')
    _ws.write(0, 1, 'SURNAME')
    _ws.write(0, 2, 'FIRST_NAME')
    _ws.write(0, 3, 'SEC_NAME')
    _ws.write(0, 4, 'SEX')
    _ws.write(0, 5, 'DATE_BIRTH')
    _ws.write(0, 6, 'DATE_TO')
    _ws.write(0, 7, 'DATE_CNCL')
    _ws.write(0, 8, 'DATE_FRM')


def write_to_excel(scroll, _ws, _finish_period, row, n, start_col=0):
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

    _ws.write(row, start_col, scroll.get('policy')[n])
    _ws.write(row, start_col + 1, name_people[0])
    _ws.write(row, start_col + 2, name_people[1])
    try:
        _ws.write(row, start_col + 3, name_people[2])
    except Exception:
        print("У вот него -> ", name_people, " нет отчества!")

    _ws.write(row, start_col + 4, gender)
    _ws.write(row, start_col + 5, str(scroll.get('birth')[n]))
    _ws.write(row, start_col + 6, _finish_period)
    _ws.write(row, start_col + 7, _finish_period)
    print('     записали ', scroll.get('fio')[n], '\n')


def my_file(ws, _file_name, _row):

    # Открытие файла от страховой
    my_text = docx2txt.process(_file_name)
    my_text = my_text.split('\n')
    #print("Содержимое файла:")
    #print(my_text)

    # нашли период страхования
    for idx, u in enumerate(my_text):
        if "Дата открепления с:" in u:
            people = {'policy': [], 'fio': [], 'birth': []}
            print('обрабатываем таблицу')
            finish_period = u.split()[3]
            print('Дата открепления с: ', finish_period)
            continue

        # нашли начало списка с фамилиями
        if "ФИО застрахованного" in u:
            st_tab = idx + 4  # начало таблицы
            #print(my_text[st_tab])
            #print(my_text[st_tab + 8])
            people = {'policy': [], 'fio': [], 'birth': []}
            i = 0
            while (i < 2000):
                try:
                    num = int(my_text[st_tab + i * 8])
                except:
                    #print('заканчивается на номере', num)
                    break
                people['policy'].append(my_text[st_tab + 2 + i * 8])
                people['fio'].append(my_text[st_tab + 4 + i * 8])
                people['birth'].append(my_text[st_tab + 6 + i * 8])
                i += 1

            print(people)
            n = 0

            print('Записываем ', i, 'людей')
            while n < i:
                write_to_excel(people, ws, finish_period, _row, n, start_col)
                _row += 1
                n += 1
    return _row

# Создали файл
people = {'policy': [], 'fio': [], 'birth': []}
wb = xlwt.Workbook()
ws = wb.add_sheet("Sheet1")   # создали лист
add_first_row(ws)             # и для листа создали "шапку"
print('Создали файл загрузки DETACH_reso_.xls')

_row = 1  # строка
start_col = 0


for name in listdir():
    print("...")
    print("Обработка файла ", name)
    print("...")
    _row = my_file(ws, name, _row)
    #print("последняя строка = ", _row)
    wb.save("DETACH_reso_.xls")

