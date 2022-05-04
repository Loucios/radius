from dataclasses import asdict, dataclass


@dataclass
class Connection:
    id: str = '№ п/п'
    title: str = 'Наименование мероприятия'
    units: str = 'Ед. изм.'
    value: str = 'Значения показателя'

    def __post_init__(self):
        self.id = str(self.id)
        self.title = str(self.title)
        self.units = str(self.units)

        if self.value is None:
            self.value = ''
        elif isinstance(self.value, float) or isinstance(self.value, int):
            self.value = format(
                self.value, "6,.2f"
            ).replace(",", " ").replace(".", ",")
        else:
            self.value = str(self.value)


@dataclass
class Event:
    id: str = '№ п/п'
    title: str = 'Наименование мероприятия'
    diameter: str = 'Диаметр, мм'
    length: str = 'Протяженность, м'
    capex: str = 'Капитальные затраты в ценах 2021 года, млн руб. без НДС'

    def __post_init__(self):
        self.id = str(self.id)
        self.title = str(self.title)

        if self.diameter is None:
            self.diameter = ''
        else:
            self.diameter = str(self.diameter)

        if isinstance(self.capex, float) or isinstance(self.capex, int):
            self.capex = format(
                self.capex, "6,.2f"
            ).replace(",", " ").replace(".", ",")
        else:
            self.capex = str(self.capex)

        if isinstance(self.length, float) or isinstance(self.length, int):
            self.length = format(
                self.length, "6,.1f"
            ).replace(",", " ").replace(".", ",")
        else:
            self.length = str(self.length)


@dataclass
class TSO:
    id: str = '№ п/п'
    title: str = 'Наименование показателя'
    units: str = 'Ед. изм.'
    old_nvv: str = 'НВВ'
    delta_nvv: str = 'Изменение НВВ'
    new_nvv: str = 'НВВ после мероприятий'

    def __post_init__(self):
        self.id = str(self.id)
        self.title = str(self.title)

        if self.units is None:
            self.units = ''
        else:
            self.units = str(self.units)

        if self.old_nvv is None:
            self.old_nvv = ''
        elif isinstance(self.old_nvv, float) or isinstance(self.old_nvv, int):
            self.old_nvv = format(
                self.old_nvv, "6,.2f"
            ).replace(",", " ").replace(".", ",")
        else:
            self.old_nvv = str(self.old_nvv)

        if self.delta_nvv is None:
            self.delta_nvv = ''
        elif isinstance(self.delta_nvv, float) or isinstance(self.delta_nvv,
                                                             int):
            self.delta_nvv = format(
                self.delta_nvv, "6,.2f"
            ).replace(",", " ").replace(".", ",")
        else:
            self.delta_nvv = str(self.delta_nvv)

        if self.new_nvv is None:
            self.new_nvv = ''
        elif isinstance(self.new_nvv, float) or isinstance(self.new_nvv, int):
            self.new_nvv = format(
                self.new_nvv, "6,.2f"
            ).replace(",", " ").replace(".", ",")
        else:
            self.new_nvv = str(self.new_nvv)


@dataclass
class Style:
    txt_style: str = '_Обычный'
    table_style: str = 'Table Grid'
    table_txt_style: str = '_Обычный_табл_10пт_по центру'
    table_name_style: str = '_Подпись таблицы'


class Table:
    def __init__(self, tbl_data, widths, table_number=1, table_name='',
                 appendix_number='', style=Style()):
        self.data = tbl_data
        self.style = style
        self.cols_number = len(asdict(tbl_data[0]).values())
        self.rows_number = len(tbl_data)
        self.widths = widths

        if table_name == '':
            self.table_name = f'Таблица {appendix_number}{table_number}'
        else:
            self.table_name = (
                f'Таблица {appendix_number}{table_number} - '
                f'{table_name}'
            )

    def __set_col_widths(self):
        for row in self.data.rows:
            for idx, width in enumerate(self.widths):
                row.cells[idx].width = width
                row.cells[idx].paragraphs[0].style = (
                    self.style.table_txt_style
                )

    def create_table(self, mydoc):
        mydoc.add_paragraph('', style=self.style.txt_style)
        mydoc.add_paragraph(self.table_name, style=self.style.table_name_style)

        table = mydoc.add_table(rows=self.rows_number, cols=self.cols_number)
        table.autofit = False
        for row in range(self.rows_number):
            row_data = asdict(self.data[row]).values()
            for key, value in enumerate(row_data):
                table.cell(row, key).paragraphs[0].add_run(value)
        table.style = self.style.table_style

        self.__set_col_widths(table, self.widths, self.table_number)
        mydoc.add_paragraph('', style=self.style.txt_style)
