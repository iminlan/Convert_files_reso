from os import chdir
from shutil import move
import RESO_attachment
import RESO_detachment

chdir("C:\\Users\\baa\\PycharmProjects\\convert_lists_from_insurance\\DOCX")

# копируем файлы в папку DOCX
move('Прикрепить\\RESO_.xls', 'RESO_.xls')
move('Закрыть\\DETACH_reso_.xls', 'DETACH_reso_.xls')
