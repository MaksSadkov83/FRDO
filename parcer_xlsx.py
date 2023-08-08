import os.path

from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Border, Side
from openpyxl.utils.cell import get_column_letter
from kivymd.toast import toast


def ParcerXlsxData(file1="Docs/--2023 (6).xlsx", file2="Docs/Оценки.xlsx", spec="Информационные системы и программирование"):

    # toast('Открываю файлы')
    wb_file1 = load_workbook(file1, read_only=True, data_only=True)
    wb_file2 = load_workbook(file2, read_only=True, data_only=True)
    ws_f1 = wb_file1.active
    ws_f2 = wb_file2.active

    # toast('Создаю новый файл')
    new_file = Workbook()
    ws = new_file.active

    # toast('Извлекаю данные')
    FIO = {}
    Subject = []
    Subjetc_hour = []
    data_student = []
    for row in range(6, ws_f2.max_row + 1):
        fio = ws_f2.cell(row=row, column=1).value
        eval = []
        for column in range(4, ws_f2.max_column + 1):
            eval.append(ws_f2.cell(row=row, column=column).value)
        FIO[fio] = eval

    for column in range(4, ws_f2.max_column + 1):
        if ws_f2.cell(row=5, column=column).value != "в том числе":
            Subject.append(ws_f2.cell(row=5, column=column).value)
            Subjetc_hour.append(ws_f2.cell(row=4, column=column).value)

    print(Subject)
    print(Subjetc_hour)
    print(FIO)


    # toast('Заношу в файл')
    # Высота строк
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 30
    ws.row_dimensions[3].height = 30
    ws.row_dimensions[4].height = 30

    # ЗАголовки таблицы
    headers = [["ФАМИЛИЯ", "ИМЯ",  "ОТЧЕСТВО",
               "ДАТА РОЖДЕНИЯ",	"ГОД ОКОНЧАНИЯ",  "КОД СПЕЦИАЛЬНОСТИ",
               "СПЕЦИАЛЬНОСТЬ",	"ДАТА ВЫДАЧИ",	"Квалификация",	"ПОЛ",	"СНИЛС",
               "НА ЕПГУ",	"Гражданство",	"На бланке",	"Основание выдачи документа",
               "Основание приема на обучение",	"Адрес электронной почты (email)",	"Председатель Государственной экзаменационной комиссии",
               "Предыдущий документ об образовании или об образовании и о квалификации",
               "Вид документа",	"Решение Государственной экзаменационной комиссии",
               "Срок освоения образовательной программы по очной форме обучения"
                ]]

    # Добавление заголовков
    for row in headers:
        ws.append(row)

    # Объединение ячеек
    ws.merge_cells('W1:DM1')
    ws.merge_cells('DO1:DP1')

    # Добавление значений после объеденной фячейки и в
    ws['W1'].value = "Сведения о содержании и результатах освоения образовательной программы среднего профессионального образования [ Наименование учебных предметов, курсов, дисциплин (модулей), практик | Общее количество часов | Оценка ]"
    ws['DN1'].value = "Курсовые работы (проекты) [ Курсовые работы (проекты) | Оценка ]"
    ws['DO1'].value = "Дополнительные сведения [ Содержание дополнительных сведений ]"
    ws['DQ1'].value = "ПОДПИСИ"
    ws['CY3'].value = "Практики"
    ws['CY4'].value = "+в том числе:"
    ws['DL3'].value = "Государственная итоговая аттестация"
    ws['DL4'].value = "+в том числе:"
    ws['DQ3'].value = "Ректор"

    # Назначение ширины столбцов
    for i in range(1, ws.max_column + 1):
        letter = get_column_letter(i)
        ws.column_dimensions[letter].width = 35

    # Выбор заголовков
    INFO_1 = ws['A1:G2']
    INFO_2 = ws['H1:V2']
    INFO_3 = ws['W1']
    INFO_4 = ws['DO1:DP2']
    INFO_5 = ws['DQ1:DQ2']
    INFO_6 = ws['CY3:DK4']
    INFO_7 = ws['DL3:DN4']
    thins = Side(border_style="medium", color="211c16")
    double = Side(border_style="medium", color="211c16")

    # Изменение цвета зоголовков
    for row in INFO_1:
        for cell in row:
            cell.fill = PatternFill('solid', fgColor='2e75b6')
            cell.border = Border(top=double, bottom=double, left=thins, right=thins)

    for row in INFO_2:
        for cell in row:
            cell.fill = PatternFill('solid', fgColor='92d050')
            cell.border = Border(top=double, bottom=double, left=thins, right=thins)

    INFO_3.fill = PatternFill('solid', fgColor='ffe699')
    INFO_3.border = Border(top=double, bottom=double, left=thins, right=thins)

    for row in ws['W2:DN2']:
        for cell in row:
            cell.fill = PatternFill('solid', fgColor='ffe699')
            cell.border = Border(top=double, bottom=double, left=thins, right=thins)

    ws['DN1'].fill = PatternFill('solid', fgColor='ffe699')
    ws['DN1'].border = Border(top=double, bottom=double, left=thins, right=thins)

    for row in INFO_4:
        for cell in row:
            cell.fill = PatternFill('solid', fgColor='b4c7e7')
            cell.border = Border(top=double, bottom=double, left=thins, right=thins)

    for row in INFO_5:
        for cell in row:
            cell.fill = PatternFill('solid', fgColor='a7074b')
            cell.border = Border(top=double, bottom=double, left=thins, right=thins)

    for row in INFO_6:
        for cell in row:
            cell.fill = PatternFill('solid', fgColor='ffcccc')
            cell.border = Border(top=double, bottom=double, left=thins, right=thins)

    for row in INFO_7:
        for cell in row:
            cell.fill = PatternFill('solid', fgColor='ffffcc')
            cell.border = Border(top=double, bottom=double, left=thins, right=thins)

    # toast('Сохраняю файл')
    path = os.path.join(f'C:\\Users\\{os.getlogin()}\\Documents\\Сведения студентов из ФРДО {spec}.xlsx')
    new_file.save(path)
    return path


if __name__ == "__main__":
    ParcerXlsxData()
