from dataclasses import dataclass, InitVar, field


@dataclass
class Connection:
    id: str
    title: str
    units: str
    value: str = field(init=False)
    input_value: InitVar(float)

    def __post_init__(self, input_value):
        if isinstance(input_value, float) or isinstance(input_value, int):
            self.value = str(round(float(input_value), 2))
        else:
            self.value = input_value


@dataclass
class Event:
    id: str
    title: str
    diameter: str
    length: str
    capex: str
