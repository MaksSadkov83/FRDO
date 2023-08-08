import os.path

from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Border, Side
from openpyxl.utils.cell import get_column_letter, column_index_from_string
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
    # Парсинг файла с оценками
    FIO = {}
    Subject = []
    Subjetc_hour = []

    for row in range(6, ws_f2.max_row + 1):
        fio = ws_f2.cell(row=row, column=1).value
        eval = []
        for column in range(4, ws_f2.max_column + 1):
            eval.append(ws_f2.cell(row=row, column=column).value)
        FIO[fio] = eval

    for column in range(4, ws_f2.max_column + 1):
        Subject.append(ws_f2.cell(row=5, column=column).value)
        Subjetc_hour.append(ws_f2.cell(row=4, column=column).value)

    # парсинг файла с данными о студентах
    data_student = []
    for row in range(2, ws_f1.max_row + 1):
        fio = \
            f"{ws_f1.cell(row=row, column=column_index_from_string('S')).value} " \
            f"{ws_f1.cell(row=row, column=column_index_from_string('T')).value} " \
            f"{ws_f1.cell(row=row, column=column_index_from_string('U')).value}"
        if fio in FIO.keys():
            student = {
                'surname': ws_f1.cell(row=row, column=column_index_from_string('S')).value,
                'name': ws_f1.cell(row=row, column=column_index_from_string('T')).value,
                'patronymic': ws_f1.cell(row=row, column=column_index_from_string('U')).value,
                'year_references': ws_f1.cell(row=row, column=column_index_from_string('Q')).value,
                'speciality_code': ws_f1.cell(row=row, column=column_index_from_string('L')).value,
                'specialty': ws_f1.cell(row=row, column=column_index_from_string('M')).value,
                'data_references': ws_f1.cell(row=row, column=column_index_from_string('J')).value,
                'qualification': ws_f1.cell(row=row, column=column_index_from_string('N')).value,
                'gender': list(ws_f1.cell(row=row, column=column_index_from_string('W')).value)[0],
                'SNILS': ws_f1.cell(row=row, column=column_index_from_string('X')).value,
                'on_EPGU': "Да",
                'citizenship': "RU",
                'on_blank': "Да",
                'base_references': "",
                'basis_acceptance': "",
                'email': "",
                'chair_gec': "",
                'previous_document_education': f"Аттестат о среднем образовании, {ws_f1.cell(row=row, column=column_index_from_string('P')).value}",
                'document_view': "",
                'solution_gec': "",
                'term_accumulation': f"{ws_f1.cell(row=row, column=column_index_from_string('R')).value} года",
                'rektor': "Данилова Оксана Вячеславовна"
            }
            print(student)
            data_student.append(student)
            break

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

    # Занесение предметов в таблицу
    sub = [j for j in Subject if (j != 'в том числе:' and 'аттестация' not in j
                and 'экзамен' not in j and 'рактика' not in j and
                'курсовая' not in j)]
    theme_praktik = [j for j in Subject if ('рактика' in j and j != "Практика")]
    name_kurs_job = [j for j in Subject if ('курсовая'in j)]

    for i in range(0, len(sub)):
        ws.cell(row=3, column=column_index_from_string('W') + i).value = sub[i]

    for i in range(0, len(theme_praktik)):
        ws.cell(row=4, column=column_index_from_string('CZ') + i).value = theme_praktik[i]

    ws['DN3'].value = name_kurs_job[0]

    # toast('Сохраняю файл')
    path = os.path.join(f'C:\\Users\\{os.getlogin()}\\Documents\\Сведения студентов из ФРДО {spec}.xlsx')
    new_file.save(path)
    return path


if __name__ == "__main__":
    ParcerXlsxData()
