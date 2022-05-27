from classes import Connection, Event, Style, Table, TSO
from docx import Document
from docx.shared import Cm
from openpyxl import load_workbook
from progress.bar import Bar


def get_connections(wb, j):

    '''
    connections = [Connection()]
    for i in range(7):
        connection = Connection(
            id=ws['A' + str(i + 11)].value,
            title=ws['B' + str(i + 11)].value,
            units=ws['C' + str(i + 11)].value,
            value=ws['D' + str(i + 11)].value,
        )
        connections.append(connection)
    return connections'''

    rng = wb.defined_names.get('Table1', scope=wb.sheetnames.index(str(j)))
    rng_dict = dict(rng.destinations)
    dest = rng_dict[str(j)]
    private_range = wb[str(j)][dest]

    tbl_data = []
    for row in private_range:
        cell_values = []
        for cell in row:
            cell_values.append(cell.value)
        tbl_data.append(Connection(*cell_values))

    # print(tbl_data)
    return tbl_data


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


def create_block(mydoc, ws, j, table_number=1, appendix_number='',
                 style=Style()):
    # Добавляем заголовок
    length = len(mydoc.paragraphs)
    mydoc.paragraphs[length - 1].style = '_1.'
    mydoc.paragraphs[length - 1].add_run(ws['D6'].value)

    #########################################################
    # Формируем блок таблицы: абзац перед ней и таблица после
    connections = get_connections(wb, j)
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
    connections_table = Table(
        connections, widths, table_number, table_name, appendix_number
    )
    connections_table.create_table(mydoc)
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
    events_table = Table(
        events, widths, table_number, table_name, appendix_number
    )
    events_table.create_table(mydoc)
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
    events_table = Table(
        tsos, widths, table_number, table_name, appendix_number
    )
    events_table.create_table(mydoc)
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
        create_block(mydoc, ws, j, tables_number, appendix_number)
        # Разбиваем на книги
        if j % 124 == 0 or j == chapters_number:
            mydoc.save(f'mydoc{books_number}.docx')
            books_number += 1
            tables_number = 1
        bar.next()
    bar.finish()


if __name__ == '__main__':
    main()
