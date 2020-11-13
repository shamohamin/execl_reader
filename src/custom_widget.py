from PyQt5.QtWidgets import QLabel, QLineEdit
from PyQt5.QtGui import QFont


class CustomLabel(QLabel):
    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)
        self.setFont(QFont('Arial', 15))
        self.setStyleSheet("""
                            color: #fff;
                            background: transparent;
                            border-radius: 5px;
                            padding: 4px;
                           """)
        
class CustomInput(QLineEdit):
    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)
        
        self.setStyleSheet("""
                           padding: 4px;
                           border: 2px solid gray;
                           color: #fff;
                            border-radius: 10px;
                            
                            selection-background-color: darkgray;
                           """)
