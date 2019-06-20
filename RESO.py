from os import listdir, chdir, path, remove
from shutil import move, rmtree
import RESO_attachment
import RESO_detachment

chdir("C:\\Users\\baa\\PycharmProjects\\convert_lists_from_insurance\\DOCX")

# копируем файлы в папку DOCX
# и чистим папки для дальнейшего использования

try:
    move('Прикрепить\\RESO_.xls', 'RESO_.xls')
    dir = "Прикрепить"
    for name in listdir(dir):
        filepath = path.join(dir, name)
        try:
            rmtree(filepath)
        except OSError:
            remove(filepath)
except FileNotFoundError:
    print("не найдено файлов прикрепления")


try:
    move('Закрыть\\DETACH_reso_.xls', 'DETACH_reso_.xls')
    dir = "Закрыть"
    for name in listdir(dir):
        filepath = path.join(dir, name)
        try:
            rmtree(filepath)
        except OSError:
            remove(filepath)
except FileNotFoundError:
    print("не найдено файлов отрепления")
