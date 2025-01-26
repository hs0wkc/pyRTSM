from os.path import basename

from PyQt6.QtCore import QFile
from PyQt6.QtWidgets import QApplication, QFrame, QSpacerItem, QSizePolicy, QVBoxLayout, QTextBrowser

_module_name = f'\033[31;1m{basename(__file__):25}:: \033[{0}m'
VSpacer = QSpacerItem(5, 0, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding)

class VLine(QFrame):
    # a simple VLine, like the one you get from designer
    def __init__(self):
        super(VLine, self).__init__()
        self.setFrameShape(QFrame.Shape.VLine)
        self.setFrameShadow(QFrame.Shadow.Sunken)

class HLine(QFrame):
    # a simple VLine, like the one you get from designer
    def __init__(self, height=3):
        super(HLine, self).__init__()
        self.setFrameShape(QFrame.Shape.HLine)
        self.setFrameShadow(QFrame.Shadow.Sunken)
        self.setFixedHeight(height)

def loadStylesheet(filename: str):
    """
    Loads an qss stylesheet to the current QApplication instance

    :param filename: Filename of qss stylesheet
    :type filename: str
    """
    # print(_module_name, 'STYLE loading:', filename)
    file = QFile(filename)
    file.open(QFile.OpenModeFlag.ReadOnly | QFile.OpenModeFlag.Text)
    stylesheet = file.readAll()
    QApplication.instance().setStyleSheet(str(stylesheet, encoding='utf-8'))


