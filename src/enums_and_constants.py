import enum
import math


class Colors(enum.Enum):
    PARAMETER_COLOR = "FF00B0F0"
    VALUE_COLOR = "FF00B050"
    FORMULA_COLOR = "FFFF0000"
    DRIVED_COLOR = "5"


PARAMETER_PATTERN = r'[A-Z]+\d{1,5}'

CONVENTIONS = {
    'MAX': 'max',
    'MIN': 'min',
    '=': '',
    'SQRT': 'sqrt',
    'SUM': 'sum',
    '^': '**'
}


ALLOWED_NAMES = {
    k: v for k, v in math.__dict__.items() if not k.startswith("__")
}
