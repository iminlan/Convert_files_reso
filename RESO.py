from os import listdir, chdir, path, remove
from shutil import move, rmtree
import RESO_attachment
import RESO_detachment

chdir("C:\\Users\\baa\\PycharmProjects\\convert_lists_from_insurance\\DOCX")

# копируем файлы в папку DOCX
move('Прикрепить\\RESO_.xls', 'RESO_.xls')
move('Закрыть\\DETACH_reso_.xls', 'DETACH_reso_.xls')

# почистили папки для дальнейшего использования

dir = "C:\\Users\\baa\\PycharmProjects\\convert_lists_from_insurance\\DOCX\\Прикрепить"

for name in listdir(dir):
    filepath = path.join(dir, name)
    try:
        rmtree(filepath)
    except OSError:
        remove(filepath)

dir = "C:\\Users\\baa\\PycharmProjects\\convert_lists_from_insurance\\DOCX\\Закрыть"

for name in listdir(dir):
    filepath = path.join(dir, name)
    try:
        rmtree(filepath)
    except OSError:
        remove(filepath)
