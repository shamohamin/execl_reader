import re
import sys
from src.enums_and_constants import (CONVENTIONS,
                                     PARAMETER_PATTERN,
                                     ALLOWED_NAMES,
                                     Colors,
                                     PARAMETER_PATTERN)

# convert execl functions into python math functions
 

def converter(val):
    pattern = '|'.join(map(re.escape, CONVENTIONS.keys()))
    converted_str = re.sub(pattern, lambda x: CONVENTIONS[x.group()], str(val))
    
    if re.search(r"\([A-Z]+\d{1,5}\:[A-Z]+\d{1,5}\)", converted_str) is not None:
        pos = list(re.finditer(PARAMETER_PATTERN, converted_str))
        first_nr = int(re.sub(r'[A-Z]+', "", pos[0].group()))
        last_nr = int(re.sub(r'[A-Z]+', "", pos[-1].group()))
        alpha = re.search(r'[A-Z]+', pos[0].group()).group()
        
        temp = converted_str[:pos[0].start()]
        temp += '['
        for i in range(int(first_nr), int(last_nr) + 1):    
            temp += alpha + str(i)
            if i != last_nr:
                temp += ','
        temp += ']'
        temp += converted_str[pos[-1].end():] 
        del first_nr, last_nr, alpha
        return temp
    
    return converted_str

# convert H22 into parameter name


def convert_execl_formul_to_parameter(val: str, data):

    def finder_pos(x):
        tmp = x.group()
        for section in data:
            if data[section]['values'].get(tmp, None) != None:
                return data[section]['values'][tmp]['name']
            elif data[section]['formula'].get(tmp, None) != None:
                return data[section]['formula'][tmp]['name']

    return re.sub(PARAMETER_PATTERN, finder_pos, val)


# calculate the value of the formulas in dp bottom-up manner
def calculate_the_value(val_str: str, sheet):
    print(val_str)

    def recursive_value(formul, params_dic={}):
        params = re.findall(PARAMETER_PATTERN, formul)
        if params_dic.get(formul, None) != None:
            return params_dic[formul]
        try:

            for param in params:
                if sheet[param].value != None:
                    val = str(converter(str(sheet[param].value)))
                    if sheet[param].font.color.index == Colors.FORMULA_COLOR.value:
                        params_dic[param] = handling_error(recursive_value(
                            val, params_dic
                        ))
                    elif sheet[param].font.color.index == Colors.VALUE_COLOR.value and \
                            re.match(PARAMETER_PATTERN, str(val)) is not None:
                        params_dic[param] = handling_error(recursive_value(
                            val, params_dic
                        ))
                    else:
                        params_dic[param] = handling_error(val)

            return eval(formul, params_dic, ALLOWED_NAMES)
        except Exception as ex:
            print(params)
            print("formul: ", formul, ex.args[0])
            # sys.exit(1)

    return recursive_value(val_str, {})


def handling_error(val):
    try:
        return float(val)
    except:
        return val
