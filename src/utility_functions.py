import re
from src.enums_and_constants import (CONVENTIONS,
                                     PARAMETER_PATTERN,
                                     ALLOWED_NAMES,
                                     Colors,
                                     PARAMETER_PATTERN)

# convert execl functions into python math functions


def converter(val):
    pattern = '|'.join(map(re.escape, CONVENTIONS.keys()))
    return re.sub(pattern, lambda x: CONVENTIONS[x.group()], str(val))

# convert H22 into parameter name


def convert_execl_formul_to_parameter(val: str, data):

    def finder_pos(x):
        x = x.group()
        print(x)
        for section in data:
            if data[section]['values'].get(x, None) != None:
                print(data[section]['values'][x]['name'])
                return data[section]['values'][x]['name']
            elif data[section]['formula'].get(x, None) != None:
                return data[section]['formula'][x]['name']

    return re.sub(PARAMETER_PATTERN, finder_pos, val)


# calculate the value of the formulas in dp bottom-up manner
def calculate_the_value(val_str: str, sheet):

    def recursive_value(formul, params_dic={}):
        params = re.findall(PARAMETER_PATTERN, formul)
        if params_dic.get(formul, None) != None:
            return sheet[params_dic[formul]]
        try:
            for param in params:
                if sheet[param] != None and sheet[param].font.color.index == Colors.FORMULA_COLOR.value:
                    params_dic[param] = recursive_value(
                        converter(sheet[param].value), params_dic)
                else:
                    params_dic[param] = sheet[param].value

            return eval(formul, params_dic, ALLOWED_NAMES)
        except:
            pass

    return recursive_value(val_str)
