from PyQt5.QtWidgets import (QGridLayout, QGroupBox,
                             QHBoxLayout,
                             QPushButton, QScrollArea, QTabWidget, QVBoxLayout, QWidget)
from src.excel_reader import initialing_parse_excel
from src.custom_widget import CustomLabel, CustomInput
from src.utility_functions import calculate_the_value, convert_execl_formul_to_parameter


class Tab(QWidget):
    def __init__(self, file_path, *args) -> None:
        super().__init__(*args)
        self.file_path = file_path
        self.sheet, self.data, self.workbook = initialing_parse_excel(
            self.file_path)


class ExeclTab(Tab):
    def __init__(self, filepath, *args) -> None:
        super().__init__(filepath, *args)

        self.main_layout = QVBoxLayout()
        self.parameter_tab = ParamTab(self.data, self.sheet, self)
        self.output_tab = OutputTab(self.data, self.sheet, self)

        self.childrent_tabs = QTabWidget()
        self.childrent_tabs.addTab(self.parameter_tab, 'Parameter Tab')
        self.childrent_tabs.addTab(self.output_tab, 'Formula Tab')
        self.main_layout.addWidget(self.childrent_tabs)

        self.setLayout(self.main_layout)

    def save(self):
        self.output_tab.saveChanges()


class ParamTab(QWidget):
    def __init__(self, data, sheet, parrentTab: ExeclTab, *args) -> None:
        super().__init__(*args)

        self.data, self.sheet, self.parrentTab = data, sheet, parrentTab
        self.mainLayout, self.scrol = QVBoxLayout(), QScrollArea()
        self.initUi()

        self.scrol.setWidget(self.formGroupBox)
        self.scrol.setWidgetResizable(True)
        self.mainLayout.addWidget(self.scrol)
        self.setLayout(self.mainLayout)

    def initUi(self):
        def inputCall(inp, section, key):
            return lambda x: self.handelInput(x, inp, section, key)

        self.formGroupBox = QGroupBox('Parameters')
        self.vb = QVBoxLayout()
        if isinstance(self.data, dict):
            for section in self.data.keys():
                sectionGroupBox = QGroupBox(section)
                sectionGrid = QGridLayout()
                sectionVbox = QVBoxLayout()
                keys = list(self.data[section]['values'].keys())
                if len(keys) != 0:
                    for i in range(0, len(keys), 2):
                        inp = CustomInput(
                            str(self.data[section]['values'][keys[i]]['value']))
                        inp.textChanged.connect(
                            inputCall(inp, section, keys[i]))
                        sectionGrid.addWidget(CustomLabel(
                            text=str(self.data[section]['values'][keys[i]]['name'])), i, 0)
                        sectionGrid.addWidget(inp, i, 1)
                        if i + 1 < len(keys):
                            inp2 = CustomInput(
                                str(self.data[section]['values']
                                    [keys[i + 1]]['value'])
                            )
                            inp2.textChanged.connect(
                                inputCall(inp, section, keys[i]))
                            sectionGrid.addWidget(
                                CustomLabel(str(self.data[section]['values'][keys[i+1]]['name'])), i, 2)
                            sectionGrid.addWidget(inp2, i, 3)

                sectionVbox.addLayout(sectionGrid)
                btn1 = QPushButton('save')
                btn1.clicked.connect(self.save)
                btn2 = QPushButton('reset')
                style = """
                    background-color: #275efe;
                    border-radius: 5px;
                    width: 100px;
                    padding: 2px;
                """
                btn1.setStyleSheet(style)
                btn2.setStyleSheet(style)
                hb = QHBoxLayout()
                hb.addStretch(1)
                hb.addWidget(btn1)
                hb.addWidget(btn2)
                sectionVbox.addLayout(hb)

                sectionGroupBox.setLayout(sectionVbox)
                self.vb.addWidget(sectionGroupBox)

        self.formGroupBox.setLayout(self.vb)

    def reset(self, x):
        pass

    def save(self, _):
        self.parrentTab.save()

    def handelInput(self, x, qinput, section, key):
        """
        docstring
        """
        self.data[section]['values'][key]['value'] = x
        self.sheet[key] = x


class OutputTab(QWidget):
    def __init__(self, data: dict, sheet, parrentTab: ExeclTab, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)

        self.data, self.sheet, parrentTab = data, sheet, parrentTab
        self.inputs = []
        self.mainLayout, self.scrol = QVBoxLayout(), QScrollArea()
        self.initUi()

        self.scrol.setWidget(self.formGroupBox)
        self.scrol.setWidgetResizable(True)
        self.mainLayout.addWidget(self.scrol)
        self.setLayout(self.mainLayout)

    def initUi(self):
        self.formGroupBox = QGroupBox('Formulas')
        self.vb = QVBoxLayout()

        for section in self.data.keys():
            sectionGroupBox = QGroupBox(section)
            sectionGrid = QGridLayout()

            for index, key in enumerate(self.data[section]['formula'].keys()):
                (str(self.data[section]['formula'][key]['value']))
                inp = CustomInput(
                    convert_execl_formul_to_parameter(
                        str(self.data[section]['formula']
                            [key]['value']), self.data
                    )
                )
                inp1 = CustomInput(
                    str(self.data[section]['formula'][key]['calculated_value']))
                self.inputs.append(inp1)
                label = CustomLabel(
                    text=str(self.data[section]['formula'][key]['name']))

                sectionGrid.addWidget(label, index, 0)
                sectionGrid.addWidget(inp, index, 1)
                sectionGrid.addWidget(inp1, index, 2)

            sectionGroupBox.setLayout(sectionGrid)
            self.vb.addWidget(sectionGroupBox)

        self.formGroupBox.setLayout(self.vb)

    def saveChanges(self):
        counter = 0
        for section in self.data.keys():
            keys = list(self.data[section]['formula'].keys())

            for j, key in enumerate(keys):
                try:
                    calculated_val = calculate_the_value(
                        self.data[section]['formula'][key]['value'], self.sheet)
                    self.data[section]['formula'][key]['calculated_value'] = calculated_val
                    self.inputs[counter + j].setText(str(calculated_val))
                except Exception as ex:
                    print(ex)

            counter += len(keys)
            del keys
