from PyQt6.QtCore import Qt, QPoint, QLine
from PyQt6.QtGui import QPixmap, QPainter, QColor, QBrush, QPolygon
from PyQt6.QtWidgets import QLabel

class SideView(QLabel):
    def __init__(self, w, h, background= Qt.GlobalColor.gray ,parent=None):
        super().__init__(parent)

        self.canvas = QPixmap(w, h)
        self.setPixmap(self.canvas)
        self.background = background
        self.by0, self.by1 = 60, 80
        self.ry0, self.ry1 = 83, 105

    def setFloor(self, floorType:int):
        match floorType:
            case 0:
                self.by0, self.by1 = 65, 80
                self.ry0, self.ry1 = 83, 105
            case 1:
                self.by0, self.by1 = 15, 30
                self.ry0, self.ry1 = 33, 55
            case 2:
                self.by0, self.by1 = 15, 30
                self.ry0, self.ry1 = 58, 80
            case 3:
                self.by0, self.by1 = 15, 30
                self.ry0, self.ry1 = 83, 105
        self.draw()

    def draw(self):
        self.canvas.fill(self.background)
        painter = QPainter(self.canvas)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        # Draw Building
        points = QPolygon([QPoint( 85, self.by0),
                           QPoint( 65, self.by1),
                           QPoint( 65, 105),
                           QPoint(105, 105),
                           QPoint(105, self.by1)])
        painter.drawPolygon(points)

        # Draw Room
        painter.setBrush(Qt.GlobalColor.green)
        points = QPolygon([QPoint( 65, self.ry0),
                           QPoint( 65, self.ry1),
                           QPoint(105, self.ry1),
                           QPoint(105, self.ry0)])
        painter.drawPolygon(points)

        # Draw Sun
        painter.setPen(Qt.GlobalColor.red)
        painter.setBrush(Qt.GlobalColor.red)
        painter.drawEllipse(20, 20, 20, 20)
        painter.drawLines([QLine(30, 15, 30, 45),
                           QLine(15, 30, 45, 30),
                           QLine(15, 15, 45, 45),
                           QLine(15, 45, 45, 15)])
        painter.end()
        self.setPixmap(self.canvas)

class TopView(QLabel):
    def __init__(self, w, h, background= Qt.GlobalColor.gray ,parent=None):
        super().__init__(parent)

        self.canvas = QPixmap(w, h)
        self.setPixmap(self.canvas)
        self.background = background
        self.internalColor = Qt.GlobalColor.black
        self.windowBrush = QBrush(QColor('#5000ffff'))
        self.compassColor = QColor('#55aa00')
        self.brushWall_A = QBrush(background)
        self.brushWall_B = QBrush(background)
        self.brushWall_C = QBrush(background)
        self.brushWall_D = QBrush(background)
        self.windowWall_A = True
        self.windowWall_B = True
        self.windowWall_C = True
        self.windowWall_D = True
        self._length = 0
        self._width = 0
        self._rotate = 0

    def setWallType(self, wall:str, isExternal:bool):
        brush = QBrush(self.background if isExternal else self.internalColor)
        match wall:
            case 'Wall_A':
                self.brushWall_A = brush
            case 'Wall_B':
                self.brushWall_B = brush
            case 'Wall_C':
                self.brushWall_C = brush
            case 'Wall_D':
                self.brushWall_D = brush
        self.draw()

    def setWallWindow(self, wall:str, isHasWindow:bool):
        match wall:
            case 'Wall_A':
                self.windowWall_A = isHasWindow
            case 'Wall_B':
                self.windowWall_B = isHasWindow
            case 'Wall_C':
                self.windowWall_C = isHasWindow
            case 'Wall_D':
                self.windowWall_D = isHasWindow
        self.draw()

    def setLength(self, length: float):
        self._length = length
        self.draw()

    def setWidth(self, width: float):
        self._width = width
        self.draw()

    def setRotate(self, angle: int):
        self._rotate = angle
        self.draw()

    def draw(self):
        self.canvas.fill(self.background)
        painter = QPainter(self.canvas)
        # painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        # painter.setRenderHint(QPainter.RenderHint.TextAntialiasing)

        # Draw Walls
        painter.setBrush(self.brushWall_A)
        painter.drawRect(50, 30, 150, 5)
        painter.setBrush(self.brushWall_B)
        painter.drawRect(50, 35, 5, 70)
        painter.setBrush(self.brushWall_C)
        painter.drawRect(50, 105, 150, 5)
        painter.setBrush(self.brushWall_D)
        painter.drawRect(195, 35, 5, 70)

        # Draw Windows
        painter.setBrush(self.windowBrush)
        if self.windowWall_A:
            painter.drawRect(100, 30, 50, 5)
        if self.windowWall_B:
            painter.drawRect(50, 55, 5, 30)
        if self.windowWall_C:
            painter.drawRect(100, 105, 50, 5)
        if self.windowWall_D:
            painter.drawRect(195, 55, 5, 30)

        # Darw Compass
        painter.save()
        painter.setBrush(Qt.BrushStyle.NoBrush)
        painter.drawEllipse(113, 58, 24, 24)
        compass = QPolygon([QPoint(12, 20), QPoint(0, 6), QPoint(-12, 20), QPoint(0, -20)])
        painter.translate(125, 70)
        painter.rotate(self._rotate)
        painter.setBrush(self.compassColor)
        painter.drawPolygon(compass)
        painter.restore()

        # Draw Labels
        painter.drawText(100, 13, 50, 15, Qt.AlignmentFlag.AlignHCenter, 'A')
        painter.drawText(38, 75, 'B')
        painter.drawText(100, 110, 50, 15, Qt.AlignmentFlag.AlignHCenter, 'C')
        painter.drawText(205, 75, 'D')
        painter.setPen(QColor('red'))
        painter.drawLines([QLine(50, 115, 50, 135),
                           QLine(50, 130, 95, 130),
                           QLine(155, 130, 200, 130),
                           QLine(200, 115, 200, 135)])
        painter.drawLines([QLine(205, 30, 240, 30),
                           QLine(235, 30, 235, 60),
                           QLine(235, 80, 235, 110),
                           QLine(205, 110, 240, 110)])
        painter.drawText(100, 123, 50, 15, Qt.AlignmentFlag.AlignHCenter, f'({self._length:.2f} m)')
        painter.drawText(220, 75, f'({self._width:.2f} m)')

        painter.end()
        self.setPixmap(self.canvas)