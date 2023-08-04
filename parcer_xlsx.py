from openpyxl import load_workbook, Workbook
from kivymd.toast import toast


def ParcerXlsxData(file1, file2):
    # toast('Открываю файлы')
    wb_file1 = load_workbook(file1)
    wb_file2 = load_workbook(file2)
    # toast('Создаю новый файл')
    # new_file = Workbook()
    # ws = new_file.active
    # toast('Извлекаю данные')
    FIO = []
    Subject = {}
    Evaluations = []
    # toast('Заношу в файл')
    headers = []
    # toast('Сохраняю файл')
    # new_file.save('Сведения студентов из ФРДО.xlsx')


if __name__ == "__main__":
    ParcerXlsxData()
