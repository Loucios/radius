from classes import Style, Table, dataclasses_list
from docx import Document
from docx.shared import Cm
from openpyxl import load_workbook
from progress.bar import Bar


def get_connections(wb, j, table_name):

    rng = wb.defined_names.get(table_name, scope=wb.sheetnames.index(str(j)))
    rng_dict = dict(rng.destinations)
    dest = rng_dict[str(j)]
    private_range = wb[str(j)][dest]

    tbl_data = []
    for row in private_range:
        cell_values = []
        for cell in row:
            cell_values.append(cell.value)
        tbl_data.append(dataclasses_list[table_name](*cell_values))

    return tbl_data


def create_block(mydoc, wb, j, table_number=1, appendix_number='',
                 style=Style()):
    ws = wb[str(j)]
    # Добавляем заголовок
    length = len(mydoc.paragraphs)
    mydoc.paragraphs[length - 1].style = '_1.'
    mydoc.paragraphs[length - 1].add_run(ws['D6'].value)

    #########################################################
    # Формируем блок таблицы: абзац перед ней и таблица после
    connections = get_connections(wb, j, 'Table1')
    mydoc.add_paragraph(
        f'В настоящем разделе рассматривается целесообразность '
        f'подключения к источнику тепловой энергии {connections[3].value} '
        f'следующей территории{ws["D7"].value}: '
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
    events = get_connections(wb, j, 'table2')
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
    tsos = get_connections(wb, j, 'Table3')
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

    return table_number


def main():
    print('Загружаем Excel')
    wb = load_workbook(filename='RET_3.1.xlsm', data_only=True)
    chapters_number = wb['Результат']['A1'].value

    books_number = 1
    appendix_number = 'A'
    table_number = 1
    bar = Bar('Создаем Word', max=chapters_number)  # Индикатор выполнения
    for j in range(1, chapters_number + 1):
        # Разбиваем на книги
        if j % 125 == 0 or j == 1:
            mydoc = Document('my_doc.docx')
        # Создаем повторяющийся блок документа
        table_number = create_block(mydoc, wb, j, table_number,
                                    appendix_number)
        # Разбиваем на книги
        if j % 124 == 0 or j == chapters_number:
            mydoc.save(f'mydoc{books_number}.docx')
            books_number += 1
            table_number = 1
        bar.next()
    bar.finish()


if __name__ == '__main__':
    main()
