from classes import Connection, Event
from docx import Document
from docx.shared import Cm
from openpyxl import load_workbook
from progress.bar import Bar


def set_col_widths(table, widths):
    for row in table.rows:
        for idx, width in enumerate(widths):
            row.cells[idx].width = width
            row.cells[idx].paragraphs[0].style = '_Обычный_табл_10пт_по центру'


def get_connections(ws):
    connections = []
    for i in range(7):
        connection = Connection(
            id=str(ws['A' + str(i + 11)].value),
            title=ws['B' + str(i + 11)].value,
            units=ws['C' + str(i + 11)].value,
            input_value=ws['D' + str(i + 11)].value,
        )
        connections.append(connection)

    for i in range(2):
        connection = Connection(
            id=str(ws['A' + str(6 + i)].value),
            title=ws['B' + str(6 + i)].value,
            units=ws['C' + str(6 + i)].value,
            input_value=ws['D' + str(6 + i)].value,
        )
        connections.append(connection)

    return connections


def get_events(ws):
    events = []
    for i in range(3):
        event = Event(
            id=str(ws['A' + str(20 + i)].value),
            title=ws['B' + str(20 + i)].value,
            diameter=str(ws['C' + str(20 + i)].value),
            length=str(ws['D' + str(20 + i)].value),
            capex=str(ws['F' + str(20 + i)].value),
        )
    events.append(event)

    return events


def create_table_2(events, mydoc, j):
    print(events)
    mydoc.add_paragraph('', style='_Обычный')
    mydoc.add_paragraph(
        f'Таблица Д{j} - Основные мероприятия и объемы капитальных затрат, '
        f'необходиые для рассматриваемого подключения',
        style='_Подпись таблицы'
    )

    table = mydoc.add_table(rows=1, cols=5)
    table.autofit = False
    table.cell(0, 0).paragraphs[0].add_run('№ п/п')
    table.cell(0, 1).paragraphs[0].add_run('Наименование мероприятия')
    table.cell(0, 2).paragraphs[0].add_run('Диаметр, мм')
    table.cell(0, 3).paragraphs[0].add_run('Протяженность, м')
    table.cell(0, 4).paragraphs[0].add_run(
        'Капитальные затраты в ценах 2021 года, млн руб. без НДС'
    )

    for i in range(3):
        row_cells = table.add_row().cells
        # print(events[i].id)
        row_cells[0].paragraphs[0].add_run(events[i].id)
        row_cells[1].paragraphs[0].add_run(events[i].title)
        row_cells[2].paragraphs[0].add_run(events[i].diameter)
        row_cells[3].paragraphs[0].add_run(events[i].length)
        row_cells[4].paragraphs[0].add_run(events[i].capex)

    widths = (Cm(1.06), Cm(5.44), Cm(1.75), Cm(3.5), Cm(5.0))
    set_col_widths(table, widths)
    table.style = 'Table Grid'
    mydoc.add_paragraph('', style='_Обычный')


def create_table_1(connections, mydoc, j):

    mydoc.add_paragraph('', style='_Обычный')
    mydoc.add_paragraph(
        f'Таблица Д{j} - Тепловая нагрузка перспективного потребителя, '
        f'источник тепловой энергии и ТСО, участвующие в подключении',
        style='_Подпись таблицы'
    )

    table = mydoc.add_table(rows=1, cols=4)
    table.autofit = False
    table.cell(0, 0).paragraphs[0].add_run('№ п/п')
    table.cell(0, 1).paragraphs[0].add_run('Наименование показателя')
    table.cell(0, 2).paragraphs[0].add_run('Ед. изм.')
    table.cell(0, 3).paragraphs[0].add_run('Значения показателя')

    for i in range(7):
        row_cells = table.add_row().cells
        row_cells[0].paragraphs[0].add_run(connections[i].id)
        row_cells[1].paragraphs[0].add_run(connections[i].title)
        row_cells[2].paragraphs[0].add_run(connections[i].units)
        row_cells[3].paragraphs[0].add_run(connections[i].value)

    widths = (Cm(1.49), Cm(4.75), Cm(1.75), Cm(8.49))
    set_col_widths(table, widths)
    table.style = 'Table Grid'
    mydoc.add_paragraph('', style='_Обычный')


def main():
    print('Загружаем Excel')
    wb = load_workbook(filename='RET.xlsx', data_only=True)
    chapters_number = wb['Результат']['A1'].value

    books_number = 1
    bar = Bar('Создаем Word', max=chapters_number, suffix='%(percent)d%%')
    tables_number = 0
    for j in range(1, chapters_number):
        if j % 125 == 0 or j == 1:
            mydoc = Document('my_doc.docx')

        ws = wb[str(j)]
        connections = get_connections(ws)
        events = get_events(ws)

        paragraphs = mydoc.paragraphs
        length = len(paragraphs)
        paragraphs[length - 1].style = '_1.'
        paragraphs[length - 1].add_run(connections[7].value)

        tables_number += 1
        mydoc.add_paragraph(
            f'В настоящем разделе рассматривается целесообразность '
            f'подключения к источнику тепловой энергии {connections[2].value} '
            f'следующей территории: {connections[7].value} '
            f'{connections[8].value}. '
            f'В таблице Д{tables_number} приведены показатели тепловой '
            f'нагрузки рассматриваемого потребителя, а также наименования ТСО,'
            f' участвующих в подключении. '
            f'Приведен вывод о целесообразности рассматриваемоего подключения '
            f'на основе выполненных расчетов.', style='_Обычный')

        create_table_1(connections, mydoc, tables_number)

        tables_number += 1
        mydoc.add_paragraph(
            f'Произведеная оценка необходимых капитальных затрат '
            f'для подключения рассматриваемоего потребителя к источнику '
            f'тепловой энергии {connections[2].value} '
            f'(таблица Д{tables_number}).'
        )

        create_table_2(events, mydoc, tables_number)

        if j % 124 == 0 or j == chapters_number:
            mydoc.save(f'mydoc{books_number}.docx')
            books_number += 1
        bar.next()
    bar.finish()


if __name__ == '__main__':
    main()
