from PyQt6.QtCore import Qt, QLineF
from PyQt6.QtGui import QPen, QPainterPath, QColor, QBrush
from PyQt6.QtWidgets import QGraphicsPathItem


class ShadingDrawer:
    def __init__(self, lengths, ends, diameter_offsets, diameter_hole, x_offset, scene=None, reverse=False):
        self.horizontal_offset = 30
        self.reverse = reverse
        self.x_offset = x_offset
        self.diameter_hole = diameter_hole
        self.lengths = lengths
        self.scene = scene
        self.ends = [ends[0]] + [max(ends[i] - ends[i - 1], 0) for i in range(1, len(ends))]
        self.rectangles = [[end, max(self.horizontal_offset, offset)] for end, offset in
                           zip(self.ends, diameter_offsets)]

        self.offsets = [min(self.horizontal_offset, offset) for offset in diameter_offsets]
        self.rectangles[0][0] += 5

        if reverse:
            for i in range(0, len(self.offsets)):
                self.offsets[i] *= -1
                self.rectangles[i][1] *= -1

        self.pen_hole = QPen(Qt.GlobalColor.black, 2)

    def build_curve(self, path, index, x, y):

        # Текущие размеры и отступы блока
        height, width = self.rectangles[index]
        offset = self.offsets[index]

        x += offset

        # Начальная точка в верхнем левом углу блока
        if index == 0:
            path.moveTo(x, y)
        # Слева направо
        path.lineTo(x, y)
        # Сверху-вниз линия
        path.lineTo(x, y + height)

        if index < len(self.rectangles) - 1:
            self.build_curve(path, index + 1, x, y + height)

        # Линия вверх
        path.lineTo(x + width, y + height)
        path.lineTo(x + width, y)
        # Влево
        path.lineTo(x + width - offset, y)

        self.scene.addLine(QLineF(x + width, y + height, x + width, y), self.pen_hole)

        if index == 0:
            # Верхняя линия
            path.lineTo(x, y)
            self.scene.addLine(QLineF(x + width, y, x, y), self.pen_hole)
        else:
            self.scene.addLine(QLineF(x + width, y, x + width - offset, y), self.pen_hole)

    def draw_curve(self):
        path = QPainterPath()
        self.build_curve(path, 0,
                         self.x_offset - self.horizontal_offset if not self.reverse else self.x_offset + self.horizontal_offset,
                         self.ends[0] - self.lengths[0])  # Начальная точка
        path.closeSubpath()

        # Добавление пути на сцену
        path_item = QGraphicsPathItem(path)
        # Установка прозрачного пера
        transparent_pen = QPen(QColor(0, 0, 0, 0))  # Черный цвет с альфа-прозрачностью 50
        path_item.setPen(transparent_pen)
        # Установка штриховки для заливки
        hatch_brush = QBrush(Qt.BrushStyle.DiagCrossPattern)
        path_item.setBrush(hatch_brush)
        self.scene.addItem(path_item)
