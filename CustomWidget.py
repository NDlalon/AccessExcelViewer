import sys
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QListWidgetItem, QFrame,QSizePolicy
)
from PyQt6.QtGui import QFont
from PyQt6.QtCore import pyqtSignal,Qt

class ClickableWidget(QWidget):
    clicked = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)

    def mousePressEvent(self, event):
        self.clicked.emit()
        super().mousePressEvent(event)

#可顯示擴展內容詳細資料
class ExpandableItem(QWidget):
    def __init__(self, title: str, details: str, list_item: QListWidgetItem):
        super().__init__()
        self.is_expanded = False
        self.list_item = list_item

        self.main_layout = QVBoxLayout()
        self.main_layout.setContentsMargins(5, 5, 5, 5)
        self.setLayout(self.main_layout)

        #使用 ClickableWidget 當作 header 容器
        header_widget = ClickableWidget()
        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(0, 0, 0, 0)
        header_widget.setLayout(header_layout)
        header_widget.clicked.connect(self.toggle_expand)

        self.title_label = QLabel(title)
        self.expand_button = QPushButton("▼")
        self.expand_button.setFixedWidth(30)
        self.expand_button.clicked.connect(self.toggle_expand)

        header_layout.addWidget(self.title_label)
        header_layout.addStretch()
        header_layout.addWidget(self.expand_button)

        self.main_layout.addWidget(header_widget)

        # 詳細內容部分
        self.detail_widget = QFrame()
        self.detail_widget.setFrameShape(QFrame.Shape.StyledPanel)
        self.detail_layout = QVBoxLayout()
        self.detail_layout.setContentsMargins(10, 5, 10, 5)
        self.detail_widget.setLayout(self.detail_layout)

        self.detail_label = QLabel(details)
        self.detail_label.setWordWrap(True)
        self.detail_layout.addWidget(self.detail_label)

        self.detail_widget.setVisible(False)
        self.main_layout.addWidget(self.detail_widget)

    def toggle_expand(self):
        self.is_expanded = not self.is_expanded
        self.detail_widget.setVisible(self.is_expanded)
        self.expand_button.setText("▲" if self.is_expanded else "▼")
        self.list_item.setSizeHint(self.sizeHint())

class CalendarItem(QWidget):
    def __init__(self,date,info):
        super().__init__()

        self.mainLayout=QVBoxLayout()
        self.mainLayout.setContentsMargins(0,0,0,0)
        self.mainLayout.setSpacing(0)
        self.setLayout(self.mainLayout)

        #使用 ClickableWidget 當作 header 容器
        self.headerLayout=QVBoxLayout()
        self.clickableWidget=ClickableWidget()
        self.headerLayout.setSpacing(0)
        self.clickableWidget.setLayout(self.headerLayout)

        #設定日期標籤
        self.dateLabel=QLabel(parent=self.clickableWidget)
        self.dateLabel.setSizePolicy(QSizePolicy.Policy.Preferred,QSizePolicy.Policy.Maximum)
        self.dateLabel.setText(date)
        self.dateLabel.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.dateLabel.setStyleSheet("""
            color:rgb(255, 255, 255);
            background-color: rgb(200, 50, 50);
            border-top-right-radius:5px;
            border-top-left-radius:5px;
        """)
        self.headerLayout.addWidget(self.dateLabel)

        #設定內容標籤
        self.infoLabel=QLabel(parent=self.clickableWidget)
        self.infoLabel.setSizePolicy(QSizePolicy.Policy.Preferred,QSizePolicy.Policy.Preferred)
        self.infoLabel.setText(info)
        self.infoLabel.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.infoLabel.setStyleSheet("""
            color:rgb(0, 0, 0);
            background-color: rgb(255, 255, 255);
            border-bottom-right-radius:5px;
            border-bottom-left-radius:5px;
        """)
        self.headerLayout.addWidget(self.infoLabel)

        self.mainLayout.addWidget(self.clickableWidget)

    def updateFontSize(self):
        #設定日期字形
        dateFont_size = int(self.width() * 0.05)    # 根據視窗寬度動態計算字體大小，取寬度的 5%
        dateFont_size = max(10, dateFont_size) # 設置字體最小10
        #設定內容字形
        infoFont_size = int(self.width() * 0.05)    # 根據視窗寬度動態計算字體大小，取寬度的 5%
        infoFont_size = max(5, dateFont_size) # 設置字體最小5

        font = QFont()

        font.setPointSize(dateFont_size)
        self.dateLabel.setFont(font)
        font.setPointSize(infoFont_size)
        self.infoLabel.setFont(font)

    def resizeEvent(self, event):
        # 當視窗大小改變時，更新字體大小
        self.updateFontSize()
        super().resizeEvent(event)
