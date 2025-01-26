from os.path import basename
from PyQt6.QtCore import Qt, QRegularExpression
from PyQt6.QtGui import QRegularExpressionValidator, QColor, QPen
from PyQt6.QtWidgets import QTableWidget, QTableWidgetItem, QLineEdit, QStyledItemDelegate

_module_name = f'\033[31;1m{basename(__file__):25}:: \033[{0}m'
hourlyHeader = ['0','1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','22','23']

class NumericDelegate(QStyledItemDelegate):

    def createEditor(self, parent, option, index):
        editor = super(NumericDelegate, self).createEditor(parent, option, index)
        if isinstance(editor, QLineEdit):
            reg_ex = QRegularExpression("[0-9]+.?[0-9]{,2}")
            validator = QRegularExpressionValidator(reg_ex, editor)
            editor.setValidator(validator)
        return editor

    def paint(self,painter,option,index):
        painter.setPen(QPen(QColor('red' if int(index.data()) > 0 else 'gray')))
        painter.drawText(option.rect, Qt.AlignmentFlag.AlignCenter, index.data())

class numericHourlyTable(QTableWidget):
    def __init__(self, columns, columnWidth=20):
        super().__init__(1, columns)
        self.columnWidth = columnWidth
        self.setItemDelegate(NumericDelegate(self))
        self.horizontalHeader().setMinimumSectionSize(10)
        self.verticalHeader().setVisible(False)
        self.setValue([0,0,0,0,0,0,0,0,10,10,10,10,0,10,10,10,10,0,0,0,0,0,0,0])
        self.setHorizontalHeaderLabels(hourlyHeader)

    def setValue(self, v:list):
        for column in range(self.columnCount()):
            item = QTableWidgetItem(str(v[column]))
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.setItem(0, column, item)
            self.setColumnWidth(column, self.columnWidth)

    def value(self)->list:
        try:
            return [int(self.item(0, i).text()) for i in range(self.columnCount())]
        except Exception as e:
            print(_module_name, 'Error:', e)
            return []
        # rawlist = [self.item(0, i).text() for i in range(self.columnCount())]
        # returnlist = []
        # for i in rawlist:
        #     if i == '': returnlist.append(0)
        #     else: returnlist.append(int(i))
        # return returnlist

class checkBoxHourlyTable(QTableWidget):
    def __init__(self, columns, columnWidth=20):
        super().__init__(1, columns)
        self.columnWidth = columnWidth
        self.horizontalHeader().setMinimumSectionSize(10)
        self.verticalHeader().setVisible(False)
        self.setValue([0,0,0,0,0,0,0,0,1,1,1,1,0,1,1,1,1,0,0,0,0,0,0,0])
        self.setHorizontalHeaderLabels(hourlyHeader)

    def setValue(self, v:list):
        for column in range(self.columnCount()):
            item = QTableWidgetItem()
            item.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
            item.setCheckState(Qt.CheckState.Checked if v[column] == 1 else Qt.CheckState.Unchecked)
            self.setItem(0, column, item)
            self.setColumnWidth(column, self.columnWidth)

    def value(self)->list:
        return [1 if self.item(0, i).checkState() == Qt.CheckState.Checked else 0 for i in range(self.columnCount())]
        # return [self.item(0, i).checkState() == Qt.CheckState.Checked for i in range(self.columnCount())]
