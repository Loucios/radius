from classes import Connection, Event, Style, TSO
from dataclasses import asdict
from docx import Document
from docx.shared import Cm
from openpyxl import load_workbook
from progress.bar import Bar


def get_connections(ws):
    connections = [Connection()]
    for i in range(7):
        connection = Connection(
            id=ws['A' + str(i + 11)].value,
            title=ws['B' + str(i + 11)].value,
            units=ws['C' + str(i + 11)].value,
            value=ws['D' + str(i + 11)].value,
        )
        connections.append(connection)
    return connections


def get_events(ws):
    events = [Event()]
    for i in range(3):
        event = Event(
            id=ws['A' + str(20 + i)].value,
            title=ws['B' + str(20 + i)].value,
            diameter=ws['C' + str(20 + i)].value,
            length=ws['D' + str(20 + i)].value,
            capex=ws['E' + str(20 + i)].value,
        )
        events.append(event)
    return events


def get_tsos(ws, name):
    names = {
        'ГУП "ТЭК СПб"': 75,
        'ПАО "ТГК-1"': 119,
        'ООО "Петербургтеплоэнерго"': 75,
        'ООО "Теплоэнерго"': 75,
        'ОАО "НПО ЦКТИ"': 100
    }

    tsos = [TSO()]
    for i in range(names[name]):
        tso = TSO(
            id=ws['A' + str(25 + i)].value,
            title=ws['B' + str(25 + i)].value,
            units=ws['C' + str(25 + i)].value,
            old_nvv=ws['D' + str(25 + i)].value,
            delta_nvv=ws['E' + str(25 + i)].value,
            new_nvv=ws['F' + str(25 + i)].value,
        )
        tsos.append(tso)
    return tsos


def create_table(tbl_data, mydoc, widths, table_number=1, table_name='',
                 appendix_number='', style=Style()):

    def set_col_widths(table, widths, table_number):
        # print(table_number)
        for row in table.rows:
            for idx, width in enumerate(widths):
                row.cells[idx].width = width
                row.cells[idx].paragraphs[0].style = style.table_txt_style

    mydoc.add_paragraph('', style=style.txt_style)

    if table_name == '':
        table_name = f'Таблица {appendix_number}{table_number}'
    table_name = f'Таблица {appendix_number}{table_number} - {table_name}'
    mydoc.add_paragraph(table_name, style=style.table_name_style)

    cols_number = len(asdict(tbl_data[0]).values())
    rows_number = len(tbl_data)
    table = mydoc.add_table(rows=rows_number, cols=cols_number)
    table.autofit = False
    for row in range(rows_number):
        row_data = asdict(tbl_data[row]).values()
        for key, value in enumerate(row_data):
            table.cell(row, key).paragraphs[0].add_run(value)

    table.style = style.table_style
    set_col_widths(table, widths, table_number)
    mydoc.add_paragraph('', style=style.txt_style)


def create_block(mydoc, ws, table_number=1, appendix_number='',
                 style=Style()):
    # Добавляем заголовок
    length = len(mydoc.paragraphs)
    mydoc.paragraphs[length - 1].style = '_1.'
    mydoc.paragraphs[length - 1].add_run(ws['D6'].value)

    #########################################################
    # Формируем блок таблицы: абзац перед ней и таблица после
    connections = get_connections(ws)
    mydoc.add_paragraph(
        f'В настоящем разделе рассматривается целесообразность '
        f'подключения к источнику тепловой энергии {connections[3].value} '
        f'следующей территории {ws["D7"].value}: '
        f'{ws["D6"].value}. '
        f'В таблице {appendix_number}{table_number} приведены показатели '
        f'тепловой нагрузки рассматриваемого потребителя, а также '
        f'наименования ТСО, участвующих в подключении. '
        f'Приведен вывод о целесообразности рассматриваемоего подключения '
        f'на основе выполненных расчетов.',
        style=style.txt_style
    )

    # Задаем параметры таблицы
    table_name = (
        'Тепловая нагрузка перспективного потребителя, '
        'источник тепловой энергии и ТСО, участвующие в подключении'
    )
    widths = (Cm(1.49), Cm(4.75), Cm(1.75), Cm(8.49))
    create_table(
        connections, mydoc, widths, table_number, table_name, appendix_number
    )
    table_number += 1

    #########################################################
    # Формируем блок таблицы: абзац перед ней и таблица после
    events = get_events(ws)
    mydoc.add_paragraph(
        f'Произведена оценка необходимых капитальных затрат '
        f'для подключения рассматриваемоего потребителя к источнику '
        f'тепловой энергии {connections[3].value} '
        f'(таблица {appendix_number}{table_number}).',
        style=style.txt_style
    )

    # Задаем параметры таблицы
    table_name = (
        'Основные мероприятия и объемы капитальных затрат, '
        'необходиые для рассматриваемого подключения'
    )
    widths = (Cm(1.24), Cm(5.50), Cm(1.75), Cm(2.50), Cm(5.49))
    create_table(
        events, mydoc, widths, table_number, table_name, appendix_number
    )
    table_number += 1

    #########################################################
    # Формируем блок таблицы: абзац перед ней и таблица после
    tsos = get_tsos(ws, ws['D16'].value)
    mydoc.add_paragraph(
        f'Произведен расчет изменения НВВ с целью определения '
        f'целесобразности подключения рассматриваемой территории '
        f'(таблица {appendix_number}{table_number}).',
        style=style.txt_style
    )

    # Задаем параметры таблицы
    table_name = 'Расчет изменения НВВ после предлагаемого подключения'
    widths = (Cm(1.5), Cm(8.0), Cm(1.75), Cm(1.75), Cm(1.75), Cm(1.75))
    create_table(
        tsos, mydoc, widths, table_number, table_name, appendix_number
    )
    table_number += 1


def main():
    print('Загружаем Excel')
    wb = load_workbook(filename='RET.xlsx', data_only=True)
    chapters_number = 5  # wb['Результат']['A1'].value
    # style = Style(txt_style='Обычный')

    books_number = 1
    appendix_number = 'Д'
    tables_number = 1
    bar = Bar('Создаем Word', max=chapters_number)  # Индикатор выполнения
    for j in range(1, chapters_number + 1):
        ws = wb[str(j)]
        # Разбиваем на книги
        if j % 125 == 0 or j == 1:
            mydoc = Document('my_doc.docx')
        # Создаем повторяющийся блок документа
        create_block(mydoc, ws, tables_number, appendix_number)
        # Разбиваем на книги
        if j % 124 == 0 or j == chapters_number:
            mydoc.save(f'mydoc{books_number}.docx')
            books_number += 1
            tables_number = 1
        bar.next()
    bar.finish()


if __name__ == '__main__':
    main()
