from dataclasses import dataclass, InitVar, field


@dataclass
class Connection:
    id: str
    title: str
    units: str
    value: str = field(init=False)
    input_value: InitVar(float)

    def __post_init__(self, input_value):
        if input_value is None:
            self.value = ''
        elif isinstance(input_value, float) or isinstance(input_value, int):
            self.value = str(round(float(input_value), 2))
        else:
            self.value = input_value


@dataclass
class Event:
    id: str
    title: str
    diameter: str
    length: str
    capex: str = field(init=False)
    input_capex: InitVar(float)

    def __post_init__(self, input_capex):
        if isinstance(input_capex, float) or isinstance(input_capex, int):
            self.capex = str(round(float(input_capex), 1))
        else:
            self.capex = input_capex


@dataclass
class TSO:
    id: str = '№ п/п'
    title: str = 'Наименование показателя'
    units: str = 'Ед. изм.'
    old_nvv: str = field(init=False)
    delta_nvv: str = field(init=False)
    new_nvv: str = field(init=False)
    input_old_nvv: InitVar(float) = 'НВВ'
    input_delta_nvv: InitVar(float) = 'Изменение НВВ'
    input_new_nvv: InitVar(float) = 'НВВ после мероприятий'

    def __post_init__(self, input_old_nvv, input_delta_nvv, input_new_nvv):
        if isinstance(input_old_nvv, float) or isinstance(input_old_nvv, int):
            self.old_nvv = str(round(float(input_old_nvv), 2))
        else:
            self.old_nvv = input_old_nvv

        if isinstance(input_delta_nvv, float) or isinstance(input_delta_nvv,
                                                            int):
            self.delta_nvv = str(round(float(input_delta_nvv), 2))
        else:
            self.delta_nvv = input_delta_nvv

        if isinstance(input_new_nvv, float) or isinstance(input_new_nvv, int):
            self.new_nvv = str(round(float(input_new_nvv), 2))
        else:
            self.new_nvv = input_new_nvv
