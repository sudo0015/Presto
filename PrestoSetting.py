# -*- coding: utf-8 -*-

import os
import sys
import darkdetect
import subprocess
import portalocker
import PrestoResource
from enum import Enum
from typing import Union
from PrestoConfig import cfg, BufSize, VERSION
from pygetwindow import getWindowsWithTitle as GetWindow
from PyQt5.QtCore import Qt, pyqtSignal, QTimer, QThread, QRectF, QEasingCurve, QEvent, QUrl, QDir
from PyQt5.QtGui import QColor, QIcon, QPainter, QTextCursor, QPainterPath, QCursor, QDesktopServices
from PyQt5.QtWidgets import QFrame, QApplication, QWidget, QHBoxLayout, QFileDialog, QLabel, QVBoxLayout, QGridLayout, \
    QPushButton, QTextEdit, QSizePolicy, QLineEdit, QSpinBox, QScrollArea, QScroller, QAction
from qfluentwidgets import NavigationItemPosition, SubtitleLabel, MessageBox, ExpandLayout, SettingCardGroup, \
    RadioButton, ExpandSettingCard, ComboBox, SwitchButton, IndicatorPosition, qconfig, isDarkTheme, ConfigItem, \
    OptionsConfigItem, FluentStyleSheet, HyperlinkButton, Slider, IconWidget, drawIcon, setThemeColor, MessageBoxBase, \
    SmoothScrollDelegate, setFont, themeColor, setTheme, Theme, qrouter, NavigationBar, SplashScreen, InfoBarIcon, \
    PushButton, TextWrap, TransparentToolButton, NavigationBarPushButton, TransparentDropDownPushButton, LineEdit, \
    ToolTipFilter, ToolTipPosition, FluentFontIconBase, BodyLabel, InfoBar, InfoBarPosition, ExpandGroupSettingCard, \
    CardWidget
from qfluentwidgets.components.widgets.line_edit import LineEditButton
from qfluentwidgets.components.widgets.menu import MenuAnimationType, RoundMenu, CheckableMenu, MenuIndicatorType
from qfluentwidgets.components.widgets.spin_box import SpinButton, SpinIcon
from qfluentwidgets.window.fluent_window import FluentWindowBase
from qframelesswindow.titlebar import MinimizeButton, CloseButton, MaximizeButton
from qframelesswindow import TitleBarButton
from qframelesswindow.utils import startSystemMove


class Mutex:
    def __init__(self):
        self.file = None

    def __enter__(self):
        self.file = open('PrestoSetting.lockfile', 'w')
        try:
            portalocker.lock(self.file, portalocker.LOCK_EX | portalocker.LOCK_NB)
        except portalocker.AlreadyLocked:
            try:
                window = GetWindow("Presto 设置")[0]
                if window.isMinimized:
                    window.restore()
                window.activate()
            except:
                pass
            sys.exit()

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.file:
            portalocker.unlock(self.file)
            self.file.close()
            os.remove('PrestoSetting.lockfile')


class FluentFontIcon(FluentFontIconBase):

    def path(self, theme=Theme.AUTO):
        return "Font/SegoeIcons.ttf"


class SmoothScrollArea(QScrollArea):

    def __init__(self, parent=None):
        super().__init__(parent)
        self.viewport().setAttribute(Qt.WA_AcceptTouchEvents, True)
        self.delegate = SmoothScrollDelegate(self, True)
        QScroller.grabGesture(self.viewport(), QScroller.TouchGesture)

    def setScrollAnimation(self, orient, duration, easing=QEasingCurve.OutCubic):
        bar = self.delegate.hScrollBar if orient == Qt.Horizontal else self.delegate.vScrollBar
        bar.setScrollAnimation(duration, easing)

    def enableTransparentBackground(self):
        self.setStyleSheet("QScrollArea{border: none; background: transparent}")

        if self.widget():
            self.widget().setStyleSheet("QWidget{background: transparent}")


class EditMenu(RoundMenu):

    def createActions(self):
        self.cutAct = QAction(
            FluentFontIcon("\ue8c6").icon(),
            "剪切",
            self,
            shortcut="Ctrl+X",
            triggered=self.parent().cut,
        )
        self.copyAct = QAction(
            FluentFontIcon("\ue8c8").icon(),
            "复制",
            self,
            shortcut="Ctrl+C",
            triggered=self.parent().copy,
        )
        self.pasteAct = QAction(
            FluentFontIcon("\ue77f").icon(),
            "粘贴",
            self,
            shortcut="Ctrl+V",
            triggered=self.parent().paste,
        )
        self.cancelAct = QAction(
            FluentFontIcon("\ue7a7").icon(),
            "撤销",
            self,
            shortcut="Ctrl+Z",
            triggered=self.parent().undo,
        )
        self.selectAllAct = QAction(
            "全选",
            self,
            shortcut="Ctrl+A",
            triggered=self.parent().selectAll
        )
        self.action_list = [self.cutAct, self.copyAct, self.pasteAct, self.cancelAct, self.selectAllAct]

    def _parentText(self):
        raise NotImplementedError

    def _parentSelectedText(self):
        raise NotImplementedError

    def exec(self, pos, ani=True, aniType=MenuAnimationType.DROP_DOWN):
        self.clear()
        self.createActions()

        if QApplication.clipboard().mimeData().hasText():
            if self._parentText():
                if self._parentSelectedText():
                    if self.parent().isReadOnly():
                        self.addActions([self.copyAct, self.selectAllAct])
                    else:
                        self.addActions(self.action_list)
                else:
                    if self.parent().isReadOnly():
                        self.addAction(self.selectAllAct)
                    else:
                        self.addActions(self.action_list[2:])
            elif not self.parent().isReadOnly():
                self.addAction(self.pasteAct)
            else:
                return
        else:
            if not self._parentText():
                return

            if self._parentSelectedText():
                if self.parent().isReadOnly():
                    self.addActions([self.copyAct, self.selectAllAct])
                else:
                    self.addActions(self.action_list[:2] + self.action_list[3:])
            else:
                if self.parent().isReadOnly():
                    self.addAction(self.selectAllAct)
                else:
                    self.addActions(self.action_list[3:])

        super().exec(pos, ani, aniType)


class LineEditMenu(EditMenu):

    def __init__(self, parent: QLineEdit):
        super().__init__("", parent)
        self.selectionStart = parent.selectionStart()
        self.selectionLength = parent.selectionLength()

    def _onItemClicked(self, item):
        if self.selectionStart >= 0:
            self.parent().setSelection(self.selectionStart, self.selectionLength)

        super()._onItemClicked(item)

    def _parentText(self):
        return self.parent().text()

    def _parentSelectedText(self):
        return self.parent().selectedText()

    def exec(self, pos, ani=True, aniType=MenuAnimationType.DROP_DOWN):
        return super().exec(pos, ani, aniType)


class CustomLineEdit(LineEdit):
    acceptSignal = pyqtSignal(bool)
    showInfoSignal = pyqtSignal(bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.isAccept = False
        self.acceptButton = LineEditButton(FluentFontIcon("\ue8fb"), self)
        self.cancelButton = LineEditButton(FluentFontIcon("\ue711"), self)
        self.acceptButton.clicked.connect(self.accept)
        self.cancelButton.clicked.connect(self.cancel)

        self.hBoxLayout.addWidget(self.acceptButton, 0, Qt.AlignRight)
        self.hBoxLayout.addWidget(self.cancelButton, 0, Qt.AlignRight)
        self.setClearButtonEnabled(False)
        self.setTextMargins(0, 0, 59, 0)

    def validate(self, s: str) -> bool:
        if len(s) > 11:
            return False
        for char in s:
            if not ('a' <= char <= 'z' or '0' <= char <= '9' or char == '.' or char == '_'):
                return False
        return True

    def accept(self):
        text = self.text().strip().lower()
        if text and self.validate(text):
            self.isAccept = True
            self.acceptSignal.emit(True)
        else:
            self.isAccept = False
            self.showInfoSignal.emit(True)
            return

    def cancel(self):
        self.isAccept = False
        self.acceptSignal.emit(False)

    def setClearButtonEnabled(self, enable: bool):
        self._isClearButtonEnabled = enable
        self.setTextMargins(0, 0, 28 * enable + 30, 0)


class SpinBoxBase:

    def __init__(self, parent=None):
        super().__init__(parent=parent)
        self.hBoxLayout = QHBoxLayout(self)

        self.setProperty('transparent', True)
        FluentStyleSheet.SPIN_BOX.apply(self)
        self.setButtonSymbols(QSpinBox.NoButtons)
        self.setFixedHeight(33)
        setFont(self)

        self.setAttribute(Qt.WA_MacShowFocusRect, False)
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self._showContextMenu)

    def setReadOnly(self, isReadOnly: bool):
        super().setReadOnly(isReadOnly)
        self.setSymbolVisible(not isReadOnly)

    def setSymbolVisible(self, isVisible: bool):
        self.setProperty("symbolVisible", isVisible)
        self.setStyle(QApplication.style())

    def _showContextMenu(self, pos):
        menu = LineEditMenu(self.lineEdit())
        menu.exec_(self.mapToGlobal(pos))

    def _drawBorderBottom(self):
        if not self.hasFocus():
            return

        painter = QPainter(self)
        painter.setRenderHints(QPainter.Antialiasing)
        painter.setPen(Qt.NoPen)

        path = QPainterPath()
        w, h = self.width(), self.height()
        path.addRoundedRect(QRectF(0, h - 10, w, 10), 5, 5)

        rectPath = QPainterPath()
        rectPath.addRect(0, h - 10, w, 8)
        path = path.subtracted(rectPath)

        painter.fillPath(path, themeColor())

    def paintEvent(self, e):
        super().paintEvent(e)
        self._drawBorderBottom()


class InlineSpinBoxBase(SpinBoxBase):

    def __init__(self, parent=None):
        super().__init__(parent)
        self.upButton = SpinButton(SpinIcon.UP, self)
        self.downButton = SpinButton(SpinIcon.DOWN, self)

        self.hBoxLayout.setContentsMargins(0, 4, 4, 4)
        self.hBoxLayout.setSpacing(5)
        self.hBoxLayout.addWidget(self.upButton, 0, Qt.AlignRight)
        self.hBoxLayout.addWidget(self.downButton, 0, Qt.AlignRight)
        self.hBoxLayout.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

        self.upButton.clicked.connect(self.stepUp)
        self.downButton.clicked.connect(self.stepDown)

    def setSymbolVisible(self, isVisible: bool):
        super().setSymbolVisible(isVisible)
        self.upButton.setVisible(isVisible)
        self.downButton.setVisible(isVisible)

    def setAccelerated(self, on: bool):
        super().setAccelerated(on)
        self.upButton.setAutoRepeat(on)
        self.downButton.setAutoRepeat(on)


class SpinBox(InlineSpinBoxBase, QSpinBox):
    """ Spin box """


class TextEditMenu(EditMenu):
    def __init__(self, parent: QTextEdit):
        super().__init__("", parent)
        cursor = parent.textCursor()
        self.selectionStart = cursor.selectionStart()
        self.selectionLength = cursor.selectionEnd() - self.selectionStart + 1

    def _parentText(self):
        return self.parent().toPlainText()

    def _parentSelectedText(self):
        return self.parent().textCursor().selectedText()

    def _onItemClicked(self, item):
        if self.selectionStart >= 0:
            cursor = self.parent().textCursor()
            cursor.setPosition(self.selectionStart)
            cursor.movePosition(QTextCursor.Right, QTextCursor.KeepAnchor, self.selectionLength)

        super()._onItemClicked(item)

    def exec(self, pos, ani=True, aniType=MenuAnimationType.DROP_DOWN):
        return super().exec(pos, ani, aniType)


class SettingIconWidget(IconWidget):

    def paintEvent(self, e):
        painter = QPainter(self)

        if not self.isEnabled():
            painter.setOpacity(0.36)

        painter.setRenderHints(QPainter.Antialiasing | QPainter.SmoothPixmapTransform)
        drawIcon(self._icon, painter, self.rect())


class SettingCard(QFrame):
    def __init__(self, icon: Union[str, QIcon], title, content=None, parent=None):
        super().__init__(parent=parent)
        self.iconLabel = SettingIconWidget(icon, self)
        self.titleLabel = QLabel(title, self)
        self.contentLabel = QLabel(content or '', self)
        self.hBoxLayout = QHBoxLayout(self)
        self.vBoxLayout = QVBoxLayout()

        if not content:
            self.contentLabel.hide()

        self.setFixedHeight(70 if content else 50)
        self.iconLabel.setFixedSize(16, 16)

        self.hBoxLayout.setSpacing(0)
        self.hBoxLayout.setContentsMargins(16, 0, 0, 0)
        self.hBoxLayout.setAlignment(Qt.AlignVCenter)
        self.vBoxLayout.setSpacing(0)
        self.vBoxLayout.setContentsMargins(0, 0, 0, 0)
        self.vBoxLayout.setAlignment(Qt.AlignVCenter)

        self.hBoxLayout.addWidget(self.iconLabel, 0, Qt.AlignLeft)
        self.hBoxLayout.addSpacing(16)

        self.hBoxLayout.addLayout(self.vBoxLayout)
        self.vBoxLayout.addWidget(self.titleLabel, 0, Qt.AlignLeft)
        self.vBoxLayout.addWidget(self.contentLabel, 0, Qt.AlignLeft)

        self.hBoxLayout.addSpacing(16)
        self.hBoxLayout.addStretch(1)

        self.contentLabel.setObjectName('contentLabel')
        FluentStyleSheet.SETTING_CARD.apply(self)

    def setTitle(self, title: str):
        self.titleLabel.setText(title)

    def setContent(self, content: str):
        self.contentLabel.setText(content)
        self.contentLabel.setVisible(bool(content))

    def setValue(self, value):
        pass

    def setIconSize(self, width: int, height: int):
        self.iconLabel.setFixedSize(width, height)

    def paintEvent(self, e):
        painter = QPainter(self)
        painter.setRenderHints(QPainter.Antialiasing)

        if isDarkTheme():
            painter.setBrush(QColor(255, 255, 255, 13))
            painter.setPen(QColor(0, 0, 0, 50))
        else:
            painter.setBrush(QColor(255, 255, 255, 170))
            painter.setPen(QColor(0, 0, 0, 19))

        painter.drawRoundedRect(self.rect().adjusted(1, 1, -1, -1), 6, 6)


class SwitchSettingCard(SettingCard):
    checkedChanged = pyqtSignal(bool)

    def __init__(self, icon: Union[str, QIcon], title, content=None, configItem: ConfigItem = None, parent=None):
        super().__init__(icon, title, content, parent)
        self.configItem = configItem
        self.switchButton = SwitchButton('关', self, IndicatorPosition.RIGHT)

        if configItem:
            self.setValue(qconfig.get(configItem))
            configItem.valueChanged.connect(self.setValue)

        self.hBoxLayout.addWidget(self.switchButton, 0, Qt.AlignRight)
        self.hBoxLayout.addSpacing(16)

        self.switchButton.checkedChanged.connect(self.__onCheckedChanged)

    def __onCheckedChanged(self, isChecked: bool):
        self.setValue(isChecked)
        self.checkedChanged.emit(isChecked)

    def setValue(self, isChecked: bool):
        if self.configItem:
            qconfig.set(self.configItem, isChecked)

        self.switchButton.setChecked(isChecked)
        self.switchButton.setText('开' if isChecked else '关')

    def setChecked(self, isChecked: bool):
        self.setValue(isChecked)

    def isChecked(self):
        return self.switchButton.isChecked()


class RangeSettingCard(SettingCard):
    valueChanged = pyqtSignal(int)

    def __init__(self, configItem, icon: Union[str, QIcon], title, content=None, parent=None):
        super().__init__(icon, title, content, parent)
        self.configItem = configItem
        self.slider = Slider(Qt.Horizontal, self)
        self.valueLabel = QLabel(self)
        self.slider.setMinimumWidth(200)

        self.slider.setSingleStep(1)
        self.slider.setRange(*configItem.range)
        self.slider.setValue(configItem.value)
        self.valueLabel.setText(str(configItem.value / 10) + ' s')

        self.hBoxLayout.addStretch(1)
        self.hBoxLayout.addWidget(self.valueLabel, 0, Qt.AlignRight)
        self.hBoxLayout.addSpacing(6)
        self.hBoxLayout.addWidget(self.slider, 0, Qt.AlignRight)
        self.hBoxLayout.addSpacing(16)

        self.valueLabel.setObjectName('valueLabel')
        configItem.valueChanged.connect(self.setValue)
        self.slider.valueChanged.connect(self.__onValueChanged)

    def __onValueChanged(self, value: int):
        self.setValue(value)
        self.valueChanged.emit(value)

    def setValue(self, value):
        qconfig.set(self.configItem, value)
        self.valueLabel.setText(str(value / 10) + ' s')
        self.valueLabel.adjustSize()
        self.slider.setValue(value)


class PushSettingCard(SettingCard):
    clicked = pyqtSignal()

    def __init__(self, text, icon: Union[str, QIcon], title, content=None, parent=None):
        super().__init__(icon, title, content, parent)
        self.button = QPushButton(text, self)
        self.hBoxLayout.addWidget(self.button, 0, Qt.AlignRight)
        self.hBoxLayout.addSpacing(16)
        self.button.clicked.connect(self.clicked)


class PrimaryPushSettingCard(PushSettingCard):
    def __init__(self, text, icon, title, content=None, parent=None):
        super().__init__(text, icon, title, content, parent)
        self.button.setObjectName('primaryButton')


class SpinBoxSettingCard(SettingCard):
    valueChanged = pyqtSignal(int)

    def __init__(self, configItem: ConfigItem, icon: Union[str, QIcon], title, content=None, parent=None):
        super().__init__(icon, title, content, parent)
        self.configItem = configItem
        self.spinBox = SpinBox(self)
        self.spinBox.setFixedWidth(130)
        self.spinBox.setAccelerated(True)
        self.spinBox.setMaximum(5)
        self.spinBox.setMinimum(1)
        self.spinBox.setValue(configItem.value)

        self.hBoxLayout.addStretch(1)
        self.hBoxLayout.addWidget(self.spinBox, 0, Qt.AlignRight)
        self.hBoxLayout.addSpacing(16)

        configItem.valueChanged.connect(self.setValue)
        self.spinBox.valueChanged.connect(self.__onValueChanged)

    def __onValueChanged(self, value: int):
        self.setValue(value)
        self.valueChanged.emit(value)

    def setValue(self, value):
        qconfig.set(self.configItem, value)
        self.spinBox.setValue(value)


class ComboBoxSettingCard(SettingCard):
    def __init__(self, configItem: OptionsConfigItem, icon: Union[str, QIcon], title, content=None, texts=None, parent=None):
        super().__init__(icon, title, content, parent)
        self.configItem = configItem
        self.comboBox = ComboBox(self)
        self.hBoxLayout.addWidget(self.comboBox, 0, Qt.AlignRight)
        self.hBoxLayout.addSpacing(16)

        self.optionToText = {o: t for o, t in zip(configItem.options, texts)}
        for text, option in zip(texts, configItem.options):
            self.comboBox.addItem(text, userData=option)

        self.comboBox.setCurrentText(self.optionToText[qconfig.get(configItem)])
        self.comboBox.currentIndexChanged.connect(self._onCurrentIndexChanged)
        configItem.valueChanged.connect(self.setValue)

    def _onCurrentIndexChanged(self, index: int):

        qconfig.set(self.configItem, self.comboBox.itemData(index))

    def setValue(self, value):
        if value not in self.optionToText:
            return

        self.comboBox.setCurrentText(self.optionToText[value])
        qconfig.set(self.configItem, value)


class BufSizePreviewCard(CardWidget):

    def __init__(self, bufSize, parent=None):
        super().__init__(parent)
        self.bufSize = bufSize
        self.name = bufSize.value

        self.radioButton = RadioButton(self.name, self)

        self.hBoxLayout = QHBoxLayout(self)
        self.hBoxLayout.setContentsMargins(20, 6, 20, 6)
        self.hBoxLayout.addWidget(self.radioButton)

        self.setFixedSize(180, 40)
        self.setBorderRadius(8)

        self.radioButton.setAttribute(Qt.WA_TransparentForMouseEvents, True)

    def setSelected(self, selected: bool):
        self.radioButton.setChecked(selected)

    def isSelected(self) -> bool:
        return self.radioButton.isChecked()


class BufSizeSettingCard(ExpandGroupSettingCard):

    def __init__(self, configItem, icon: Union[str, QIcon], title, content=None, parent=None):
        super().__init__(icon, title, content, parent=parent)
        self.configItem = configItem
        self.choiceLabel = QLabel(self)
        self.bufSizeCards = []

        self.bufSizeContainer = QWidget(self.view)
        self.containerLayout = QHBoxLayout(self.bufSizeContainer)
        self.bufSizeWidget = QWidget(self.bufSizeContainer)
        self.bufSizeLayout = QGridLayout(self.bufSizeWidget)

        self.__initWidget()

    def __initWidget(self):
        self.__initLayout()
        currentValue = qconfig.get(self.configItem)
        for card in self.bufSizeCards:
            if card.bufSize == currentValue:
                card.setSelected(True)
                self.choiceLabel.setText(card.name)
                self.choiceLabel.adjustSize()
                break

        self.choiceLabel.setObjectName("titleLabel")
        if darkdetect.isDark():
            self.choiceLabel.setStyleSheet(
                "font: 13px 'Segoe UI', 'Microsoft YaHei', 'PingFang SC'; padding: 0; border: none; background-color: transparent; color: white;")
        else:
            self.choiceLabel.setStyleSheet(
                "font: 13px 'Segoe UI', 'Microsoft YaHei', 'PingFang SC'; padding: 0; border: none; background-color: transparent; color: black;")

    def __initLayout(self):
        self.addWidget(self.choiceLabel)

        self.bufSizeLayout.setSpacing(8)
        self.bufSizeLayout.setContentsMargins(48, 14, 44, 14)
        for i, bufSize in enumerate(self.configItem.options):
            card = BufSizePreviewCard(bufSize, self.bufSizeWidget)
            card.clicked.connect(lambda b=bufSize: self.__onBufSizeClicked(b))
            self.bufSizeCards.append(card)
            self.bufSizeLayout.addWidget(card, i // 3, i % 3)

        self.bufSizeLayout.setSizeConstraint(QGridLayout.SetFixedSize)

        self.containerLayout.setContentsMargins(0, 0, 0, 0)
        self.containerLayout.addStretch(1)
        self.containerLayout.addWidget(self.bufSizeWidget)
        self.containerLayout.addStretch(1)

        self.viewLayout.setSpacing(0)
        self.viewLayout.setContentsMargins(0, 0, 0, 0)
        self.addGroupWidget(self.bufSizeContainer)

    def _clearAllSelection(self):
        for card in self.bufSizeCards:
            card.setSelected(False)

    def setValue(self, value):
        qconfig.set(self.configItem, value)
        for card in self.bufSizeCards:
            if card.bufSize == value:
                self._clearAllSelection()
                card.setSelected(True)
                self.choiceLabel.setText(card.name)
                self.choiceLabel.adjustSize()
                break

    def __onBufSizeClicked(self, bufSize):
        if self.choiceLabel.text() == bufSize.value:
            return
        self._clearAllSelection()
        for card in self.bufSizeCards:
            if card.bufSize == bufSize:
                card.setSelected(True)
                self.choiceLabel.setText(card.name)
                self.choiceLabel.adjustSize()
                break
        qconfig.set(self.configItem, bufSize)


class SizeFilterItem(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent=parent)
        self.hBoxLayout = QHBoxLayout(self)
        self.titleLabel = QLabel("最大文件大小", self)
        if darkdetect.isDark():
            self.titleLabel.setStyleSheet(
                "font: 13px 'Segoe UI', 'Microsoft YaHei', 'PingFang SC'; padding: 0; border: none; background-color: transparent; color: white;")
        else:
            self.titleLabel.setStyleSheet(
                "font: 13px 'Segoe UI', 'Microsoft YaHei', 'PingFang SC'; padding: 0; border: none; background-color: transparent; color: black;")

        self.spinBox = SpinBox(self)
        self.spinBox.setFixedWidth(130)
        self.spinBox.setAccelerated(True)
        self.spinBox.setMinimum(1)
        self.spinBox.setMaximum(1024)
        self.spinBox.setValue(cfg.SizeFilterValue.value)
        self.spinBox.valueChanged.connect(self.setValue)

        self.KbAction = QAction("KB", self)
        self.MbAction = QAction("MB", self)
        self.GbAction = QAction("GB", self)
        self.KbAction.setCheckable(True)
        self.MbAction.setCheckable(True)
        self.GbAction.setCheckable(True)
        self.KbAction.triggered.connect(self.onKbAction)
        self.MbAction.triggered.connect(self.onMbAction)
        self.GbAction.triggered.connect(self.onGbAction)
        self.unitMenu = CheckableMenu(indicatorType=MenuIndicatorType.RADIO)
        self.unitMenu.addAction(self.KbAction)
        self.unitMenu.addAction(self.MbAction)
        self.unitMenu.addAction(self.GbAction)
        self.updateStatus()
        self.unitBtn = TransparentDropDownPushButton(self)
        self.unitBtn.setMenu(self.unitMenu)
        self.unitBtn.setText(cfg.SizeFilterUnit.value)

        self.setFixedHeight(68)
        self.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Fixed)
        self.hBoxLayout.setContentsMargins(48, 0, 32, 0)
        self.hBoxLayout.addWidget(self.titleLabel, 0, Qt.AlignLeft)
        self.hBoxLayout.addSpacing(16)
        self.hBoxLayout.addStretch(1)
        self.hBoxLayout.addWidget(self.spinBox, 0, Qt.AlignRight)
        self.hBoxLayout.addWidget(self.unitBtn, 0, Qt.AlignRight)
        self.hBoxLayout.setAlignment(Qt.AlignVCenter)

    def setValue(self):
        cfg.set(cfg.SizeFilterValue, self.spinBox.value())

    def updateStatus(self):
        if cfg.SizeFilterUnit.value == 'KB':
            self.unitMenu.actions()[0].setChecked(True)
        elif cfg.SizeFilterUnit.value == 'MB':
            self.unitMenu.actions()[1].setChecked(True)
        elif cfg.SizeFilterUnit.value == 'GB':
            self.unitMenu.actions()[2].setChecked(True)

    def onUnitBtn(self):
        self.unitMenu.exec(QCursor.pos())

    def onKbAction(self):
        self.unitMenu.actions()[0].setChecked(True)
        self.unitMenu.actions()[1].setChecked(False)
        self.unitMenu.actions()[2].setChecked(False)
        self.unitBtn.setText('KB')
        cfg.set(cfg.SizeFilterUnit, 'KB')

    def onMbAction(self):
        self.unitMenu.actions()[0].setChecked(False)
        self.unitMenu.actions()[1].setChecked(True)
        self.unitMenu.actions()[2].setChecked(False)
        self.unitBtn.setText('MB')
        cfg.set(cfg.SizeFilterUnit, 'MB')

    def onGbAction(self):
        self.unitMenu.actions()[0].setChecked(False)
        self.unitMenu.actions()[1].setChecked(False)
        self.unitMenu.actions()[2].setChecked(True)
        self.unitBtn.setText('GB')
        cfg.set(cfg.SizeFilterUnit, 'GB')


class SizeFilterSettingCard(ExpandSettingCard):

    def __init__(self, title: str, content: str = None, parent=None):
        super().__init__(FluentFontIcon("\ue71c"), title, content, parent)
        self.switchBtn = SwitchButton(self, IndicatorPosition.RIGHT)
        self.switchBtn.setChecked(cfg.IsSizeFilter.value)
        self.switchBtn.setText('开' if cfg.IsSizeFilter.value else '关')
        self.sizeFilterItem = SizeFilterItem(self)

        self.__initWidget()

    def __initWidget(self):
        self.addWidget(self.switchBtn)

        self.viewLayout.setSpacing(0)
        self.viewLayout.setAlignment(Qt.AlignTop)
        self.viewLayout.setContentsMargins(0, 0, 0, 0)
        self.viewLayout.addWidget(self.sizeFilterItem)

        self.switchBtn.checkedChanged.connect(self.onSwitchBtn)

    def onSwitchBtn(self):
        self.setValue(self.switchBtn.isChecked())

    def setValue(self, isChecked: bool):
        self.switchBtn.setText('开' if isChecked else '关')
        cfg.set(cfg.IsSizeFilter, isChecked)

    def setChecked(self, isChecked: bool):
        self.switchBtn.setChecked(isChecked)
        self.setValue(isChecked)



class TypeFilterModeItem(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent=parent)
        self.hBoxLayout = QHBoxLayout(self)
        self.titleLabel = QLabel("过滤模式", self)
        if darkdetect.isDark():
            self.titleLabel.setStyleSheet(
                "font: 14px 'Segoe UI', 'Microsoft YaHei', 'PingFang SC'; padding: 0; border: none; background-color: transparent; color: white;")
        else:
            self.titleLabel.setStyleSheet(
                "font: 14px 'Segoe UI', 'Microsoft YaHei', 'PingFang SC'; padding: 0; border: none; background-color: transparent; color: black;")

        self.comboItems = ["排除", "包含"]
        self.currentIndex = 0 if cfg.TypeFilterMode.value == "Exclude" else 1
        self.modeComboBox = ComboBox(self)
        self.modeComboBox.setPlaceholderText("选择过滤模式")
        self.modeComboBox.addItems(self.comboItems)
        self.modeComboBox.setCurrentIndex(self.currentIndex)
        self.modeComboBox.currentIndexChanged.connect(self.onModeChanged)

        self.setFixedHeight(68)
        self.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Fixed)
        self.hBoxLayout.setContentsMargins(48, 0, 32, 0)
        self.hBoxLayout.addWidget(self.titleLabel, 0, Qt.AlignLeft)
        self.hBoxLayout.addSpacing(16)
        self.hBoxLayout.addStretch(1)
        self.hBoxLayout.addWidget(self.modeComboBox, 0, Qt.AlignRight)
        self.hBoxLayout.setAlignment(Qt.AlignVCenter)

    def onModeChanged(self):
        if self.modeComboBox.currentIndex() == 0:
            cfg.set(cfg.TypeFilterMode, "Exclude")
        elif self.modeComboBox.currentIndex() == 1:
            cfg.set(cfg.TypeFilterMode, "Include")


class TypeFilterItem(QWidget):

    def __init__(self, configItem: ConfigItem, icon: Union[str, QIcon], title, content=None, parent=None):
        super().__init__(parent=parent)
        self.config = configItem
        self.iconLabel = SettingIconWidget(icon, self)
        self.titleLabel = QLabel(title, self)
        self.contentLabel = QLabel(content, self)
        self.setFixedHeight(70)
        self.iconLabel.setFixedSize(16, 16)
        if darkdetect.isDark():
            self.titleLabel.setStyleSheet(
                "font: 14px 'Segoe UI', 'Microsoft YaHei', 'PingFang SC'; padding: 0; border: none; background-color: transparent; color: white;")
            self.contentLabel.setStyleSheet(
                "font: 11px 'Segoe UI', 'Microsoft YaHei', 'PingFang SC'; padding: 0; border: none; background-color: transparent; color: rgb(208, 208, 208);")
        else:
            self.titleLabel.setStyleSheet(
                "font: 14px 'Segoe UI', 'Microsoft YaHei', 'PingFang SC'; padding: 0; border: none; background-color: transparent; color: black;")
            self.contentLabel.setStyleSheet(
                "font: 11px 'Segoe UI', 'Microsoft YaHei', 'PingFang SC'; padding: 0; border: none; background-color: transparent; color: rgb(96, 96, 96);")

        self.switchBtn = SwitchButton(self, IndicatorPosition.RIGHT)
        self.switchBtn.setChecked(self.config.value)
        self.switchBtn.setOnText('')
        self.switchBtn.setOffText('')
        self.switchBtn.checkedChanged.connect(self.onSwitchBtnChanged)

        self.hBoxLayout = QHBoxLayout(self)
        self.vBoxLayout = QVBoxLayout()
        self.hBoxLayout.setSpacing(0)
        self.hBoxLayout.setContentsMargins(48, 0, 32, 0)
        self.hBoxLayout.setAlignment(Qt.AlignVCenter)
        self.vBoxLayout.setSpacing(0)
        self.vBoxLayout.setContentsMargins(0, 0, 0, 0)
        self.vBoxLayout.setAlignment(Qt.AlignVCenter)

        self.hBoxLayout.addWidget(self.iconLabel, 0, Qt.AlignLeft)
        self.hBoxLayout.addSpacing(16)

        self.hBoxLayout.addLayout(self.vBoxLayout)
        self.vBoxLayout.addWidget(self.titleLabel, 0, Qt.AlignLeft)
        self.vBoxLayout.addWidget(self.contentLabel, 0, Qt.AlignLeft)

        self.hBoxLayout.addSpacing(16)
        self.hBoxLayout.addStretch(1)
        self.hBoxLayout.addWidget(self.switchBtn, 0, Qt.AlignRight)

    def onSwitchBtnChanged(self):
        cfg.set(self.config, self.switchBtn.isChecked())


class CustomTypeFilterItem(QWidget):

    def __init__(self, parent=None):
        super().__init__(parent=parent)
        self.config = cfg.IsCustomType
        self.iconLabel = QLabel(self)
        self.titleLabel = QLabel("自定义", self)
        self.contentLabel = QLabel(self)
        if cfg.CustomType.value:
            self.contentLabel.setText(cfg.CustomType.value)
            self.contentLabel.setHidden(False)
        else:
            self.contentLabel.setHidden(True)
        self.setFixedHeight(70)
        self.iconLabel.setFixedSize(16, 16)
        if darkdetect.isDark():
            self.titleLabel.setStyleSheet(
                "font: 14px 'Segoe UI', 'Microsoft YaHei', 'PingFang SC'; padding: 0; border: none; background-color: transparent; color: white;")
            self.contentLabel.setStyleSheet(
                "font: 11px 'Segoe UI', 'Microsoft YaHei', 'PingFang SC'; padding: 0; border: none; background-color: transparent; color: rgb(208, 208, 208);")
        else:
            self.titleLabel.setStyleSheet(
                "font: 14px 'Segoe UI', 'Microsoft YaHei', 'PingFang SC'; padding: 0; border: none; background-color: transparent; color: black;")
            self.contentLabel.setStyleSheet(
                "font: 11px 'Segoe UI', 'Microsoft YaHei', 'PingFang SC'; padding: 0; border: none; background-color: transparent; color: rgb(96, 96, 96);")

        self.switchBtn = SwitchButton(self, IndicatorPosition.RIGHT)
        self.switchBtn.setChecked(self.config.value)
        self.switchBtn.setOnText('')
        self.switchBtn.setOffText('')
        self.switchBtn.checkedChanged.connect(self.onSwitchBtnChanged)

        self.editBtn = HyperlinkButton(self)
        self.editBtn.setIcon(FluentFontIcon("\ue70f"))
        self.editBtn.setText("编辑")
        self.editBtn.clicked.connect(self.onEditBtn)

        self.hBoxLayout = QHBoxLayout(self)
        self.vBoxLayout = QVBoxLayout()
        self.hBoxLayout.setSpacing(0)
        self.hBoxLayout.setContentsMargins(48, 0, 32, 0)
        self.hBoxLayout.setAlignment(Qt.AlignVCenter)
        self.vBoxLayout.setSpacing(0)
        self.vBoxLayout.setContentsMargins(0, 0, 0, 0)
        self.vBoxLayout.setAlignment(Qt.AlignVCenter)

        self.hBoxLayout.addWidget(self.iconLabel, 0, Qt.AlignLeft)
        self.hBoxLayout.addSpacing(16)

        self.hBoxLayout.addLayout(self.vBoxLayout)
        self.vBoxLayout.addWidget(self.titleLabel, 0, Qt.AlignLeft)
        self.vBoxLayout.addWidget(self.contentLabel, 0, Qt.AlignLeft)

        self.hBoxLayout.addSpacing(16)
        self.hBoxLayout.addStretch(1)
        self.hBoxLayout.addWidget(self.editBtn, 0, Qt.AlignRight)
        self.hBoxLayout.addWidget(self.switchBtn, 0, Qt.AlignRight)

    def onSwitchBtnChanged(self):
        cfg.set(self.config, self.switchBtn.isChecked())

    def onEditBtn(self):
        w = CustomTypeMessageBox(self.window())
        if w.exec():
            text = w.getType()
            self.contentLabel.setText(text)
            cfg.set(cfg.CustomType, text)
            if text:
                self.contentLabel.setHidden(False)
            else:
                self.contentLabel.setHidden(True)


class TypeFilterSettingCard(ExpandSettingCard):

    def __init__(self, title: str, content: str = None, parent=None):
        super().__init__(FluentFontIcon("\uea69"), title, content, parent)
        self.switchBtn = SwitchButton(self, IndicatorPosition.RIGHT)
        self.switchBtn.setChecked(cfg.IsTypeFilter.value)
        self.switchBtn.setText('开' if cfg.IsTypeFilter.value else '关')

        self.typeFilterModeItem = TypeFilterModeItem(self)
        self.documentItem = TypeFilterItem(
            cfg.IsDocument,
            FluentFontIcon("\ue8a5"),
            "文档",
            ".pdf, .doc, .docx, .ppt, .pptx, .txt"
        )
        self.pictureItem = TypeFilterItem(
            cfg.IsPicture,
            FluentFontIcon("\ue91b"),
            "图片",
            ".jpg, .jpeg, .png, .gif, .bmp, .svg, .webp, .avif"
        )
        self.audioItem = TypeFilterItem(
            cfg.IsAudio,
            FluentFontIcon("\ue8d6"),
            "音频",
            ".mp3, .wav, .flac, .ape, .acc, .ogg, .wma"
        )
        self.videoItem = TypeFilterItem(
            cfg.IsVideo,
            FluentFontIcon("\ue8b2"),
            "视频",
            ".mp4, .avi, .mov, .wmv, .mkv"
        )
        self.applicationItem = TypeFilterItem(
            cfg.IsApplication,
            FluentFontIcon("\uecaa"),
            "应用",
            ".exe, .dll, .msi, .bat, .cmd"
        )
        self.zipFileItem = TypeFilterItem(
            cfg.IsZipFile,
            FluentFontIcon("\uf012"),
            "压缩文件",
            ".zip, .rar, .7z, .iso"
        )
        self.customTypeItem = CustomTypeFilterItem(self)

        self.__initWidget()

    def __initWidget(self):
        self.addWidget(self.switchBtn)

        self.viewLayout.setSpacing(0)
        self.viewLayout.setAlignment(Qt.AlignTop)
        self.viewLayout.setContentsMargins(0, 0, 0, 0)

        self.viewLayout.addWidget(self.typeFilterModeItem)
        self.viewLayout.addWidget(self.documentItem)
        self.viewLayout.addWidget(self.pictureItem)
        self.viewLayout.addWidget(self.audioItem)
        self.viewLayout.addWidget(self.videoItem)
        self.viewLayout.addWidget(self.applicationItem)
        self.viewLayout.addWidget(self.zipFileItem)
        self.viewLayout.addWidget(self.customTypeItem)

        self.switchBtn.checkedChanged.connect(self.onSwitchBtn)

    def onSwitchBtn(self):
        self.setValue(self.switchBtn.isChecked())

    def setValue(self, isChecked: bool):
        self.switchBtn.setText('开' if isChecked else '关')
        cfg.set(cfg.IsTypeFilter, isChecked)

    def setChecked(self, isChecked: bool):
        self.switchBtn.setChecked(isChecked)
        self.setValue(isChecked)


class FolderItem(QWidget):
    def __init__(self, title: str, configItem: ConfigItem = None, parent=None):
        super().__init__(parent=parent)
        self.config = configItem
        self.titleLabel = QLabel(title, self)
        self.contentLabel = QLabel(self.config.value, self)
        self.setFixedHeight(56)
        if darkdetect.isDark():
            self.titleLabel.setStyleSheet("font: 14px 'Segoe UI', 'Microsoft YaHei', 'PingFang SC'; padding: 0; border: none; background-color: transparent; color: white;")
            self.contentLabel.setStyleSheet("font: 11px 'Segoe UI', 'Microsoft YaHei', 'PingFang SC'; padding: 0; border: none; background-color: transparent; color: rgb(208, 208, 208);")
        else:
            self.titleLabel.setStyleSheet("font: 14px 'Segoe UI', 'Microsoft YaHei', 'PingFang SC'; padding: 0; border: none; background-color: transparent; color: black;")
            self.contentLabel.setStyleSheet("font: 11px 'Segoe UI', 'Microsoft YaHei', 'PingFang SC'; padding: 0; border: none; background-color: transparent; color: rgb(96, 96, 96);")

        self.changeBtn = HyperlinkButton(self)
        self.changeBtn.setText("更改")
        self.changeBtn.clicked.connect(self.showFolderDialog)

        self.hBoxLayout = QHBoxLayout(self)
        self.vBoxLayout = QVBoxLayout()
        self.hBoxLayout.setSpacing(0)
        self.hBoxLayout.setContentsMargins(36, 0, 32, 0)
        self.hBoxLayout.setAlignment(Qt.AlignVCenter)
        self.vBoxLayout.setSpacing(0)
        self.vBoxLayout.setContentsMargins(0, 0, 0, 0)
        self.vBoxLayout.setAlignment(Qt.AlignVCenter)
        self.hBoxLayout.addSpacing(16)

        self.hBoxLayout.addLayout(self.vBoxLayout)
        self.vBoxLayout.addWidget(self.titleLabel, 0, Qt.AlignLeft)
        self.vBoxLayout.addWidget(self.contentLabel, 0, Qt.AlignLeft)

        self.hBoxLayout.addSpacing(16)
        self.hBoxLayout.addStretch(1)
        self.hBoxLayout.addWidget(self.changeBtn, 0, Qt.AlignRight)

    def showFolderDialog(self):
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹", cfg.sourceFolder.value)
        if not folder:
            return
        self.contentLabel.setText(folder)
        cfg.set(self.config, folder)


class CustomFolderListSettingCard(ExpandSettingCard):
    folderChanged = pyqtSignal(list)

    def __init__(self, title: str, content: str = None, parent=None):
        super().__init__(FluentFontIcon("\ue8f4"), title, content, parent)
        self.__initWidget()

    def __initWidget(self):
        self.viewLayout.setSpacing(0)
        self.viewLayout.setAlignment(Qt.AlignTop)
        self.viewLayout.setContentsMargins(0, 0, 0, 0)

        self.yuwenItem = FolderItem("语文", cfg.yuwenFolder, self.view)
        self.shuxueItem = FolderItem("数学", cfg.shuxueFolder, self.view)
        self.yingyuItem = FolderItem("英语", cfg.yingyuFolder, self.view)
        self.wuliItem = FolderItem("物理", cfg.wuliFolder, self.view)
        self.huaxueItem = FolderItem("化学", cfg.huaxueFolder, self.view)
        self.shengwuItem = FolderItem("生物", cfg.shengwuFolder, self.view)
        self.zhengzhiItem = FolderItem("政治", cfg.zhengzhiFolder, self.view)
        self.lishiItem = FolderItem("历史", cfg.lishiFolder, self.view)
        self.diliItem = FolderItem("地理", cfg.diliFolder, self.view)
        self.jishuItem = FolderItem("技术", cfg.jishuFolder, self.view)
        self.ziliaoItem = FolderItem("资料", cfg.ziliaoFolder, self.view)

        self.viewLayout.addWidget(self.yuwenItem)
        self.viewLayout.addWidget(self.shuxueItem)
        self.viewLayout.addWidget(self.yingyuItem)
        self.viewLayout.addWidget(self.wuliItem)
        self.viewLayout.addWidget(self.huaxueItem)
        self.viewLayout.addWidget(self.shengwuItem)
        self.viewLayout.addWidget(self.zhengzhiItem)
        self.viewLayout.addWidget(self.lishiItem)
        self.viewLayout.addWidget(self.diliItem)
        self.viewLayout.addWidget(self.jishuItem)
        self.viewLayout.addWidget(self.ziliaoItem)

        self._adjustViewSize()

    def updateContent(self):
        self.yuwenItem.contentLabel.setText(cfg.yuwenFolder.value)
        self.shuxueItem.contentLabel.setText(cfg.shuxueFolder.value)
        self.yingyuItem.contentLabel.setText(cfg.yingyuFolder.value)
        self.wuliItem.contentLabel.setText(cfg.wuliFolder.value)
        self.huaxueItem.contentLabel.setText(cfg.huaxueFolder.value)
        self.shengwuItem.contentLabel.setText(cfg.shengwuFolder.value)
        self.zhengzhiItem.contentLabel.setText(cfg.zhengzhiFolder.value)
        self.lishiItem.contentLabel.setText(cfg.lishiFolder.value)
        self.diliItem.contentLabel.setText(cfg.diliFolder.value)
        self.jishuItem.contentLabel.setText(cfg.jishuFolder.value)
        self.ziliaoItem.contentLabel.setText(cfg.ziliaoFolder.value)


class InfoIconWidget(QWidget):

    def __init__(self, icon: InfoBarIcon, parent=None):
        super().__init__(parent=parent)
        self.setFixedSize(36, 36)
        self.icon = icon

    def paintEvent(self, e):
        painter = QPainter(self)
        painter.setRenderHints(QPainter.Antialiasing | QPainter.SmoothPixmapTransform)

        rect = QRectF(10, 10, 15, 15)
        if self.icon:
            drawIcon(self.icon, painter, rect)


class InformationBar(QFrame):

    def __init__(self, title: str, content: str, parent=None):
        super().__init__(parent=parent)
        self.title = title
        self.content = content
        self.icon = InfoBarIcon.INFORMATION

        self.titleLabel = QLabel(self)
        self.contentLabel = QLabel(self)
        self.titleLabel.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.contentLabel.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.iconWidget = InfoIconWidget(self.icon)
        self.helpBtn = PushButton("详情", self)
        self.helpBtn.clicked.connect(self.onHelpAction)

        self.hBoxLayout = QHBoxLayout(self)
        self.textLayout = QHBoxLayout()
        self.widgetLayout = QHBoxLayout()

        self.lightBackgroundColor = QColor(211, 231, 247)
        self.darkBackgroundColor = QColor(52, 66, 77)

        self.__setQss()
        self.__initLayout()

    def onHelpAction(self):
        if os.path.exists(os.path.abspath("./Doc/PrestoHelp.html")):
            os.startfile(os.path.abspath("./Doc/PrestoHelp.html"))
        else:
            QDesktopServices.openUrl(QUrl("https://sudo0015.github.io/post/Presto%20-bang-zhu.html"))

    def __initLayout(self):
        self.hBoxLayout.setContentsMargins(6, 6, 6, 6)
        self.hBoxLayout.setSizeConstraint(QVBoxLayout.SetMinimumSize)
        self.textLayout.setSizeConstraint(QHBoxLayout.SetMinimumSize)
        self.textLayout.setAlignment(Qt.AlignTop)
        self.textLayout.setContentsMargins(1, 8, 0, 8)

        self.hBoxLayout.setSpacing(0)
        self.textLayout.setSpacing(5)
        self.hBoxLayout.addWidget(self.iconWidget, 0, Qt.AlignTop | Qt.AlignLeft)

        self.titleLabel.setVisible(bool(self.title))
        self.contentLabel.setVisible(bool(self.content))
        self.textLayout.addWidget(self.titleLabel, 1, Qt.AlignLeft | Qt.AlignTop)
        self.textLayout.addSpacing(7)
        self.textLayout.addWidget(self.contentLabel, 1, Qt.AlignLeft | Qt.AlignTop)
        self.widgetLayout.addWidget(self.helpBtn)

        self.hBoxLayout.addLayout(self.textLayout)
        self.hBoxLayout.addLayout(self.widgetLayout)
        self.widgetLayout.setSpacing(10)
        self.hBoxLayout.addSpacing(12)

        self._adjustText()

    def __setQss(self):
        self.titleLabel.setObjectName('titleLabel')
        self.contentLabel.setObjectName('contentLabel')
        if isinstance(self.icon, Enum):
            self.setProperty('type', self.icon.value)

        FluentStyleSheet.INFO_BAR.apply(self)

    def _adjustText(self):
        w = 900 if not self.parent() else (self.parent().width() - 50)
        self.titleLabel.setText(TextWrap.wrap(self.title, 80, False)[0])
        self.contentLabel.setText(TextWrap.wrap(self.content, 80, False)[0])
        self.adjustSize()

    def addWidget(self, widget: QWidget, stretch=0):
        self.widgetLayout.addSpacing(6)
        self.widgetLayout.addWidget(widget, stretch, Qt.AlignLeft | Qt.AlignTop)

    def eventFilter(self, obj, e: QEvent):
        if obj is self.parent():
            if e.type() in [QEvent.Resize, QEvent.WindowStateChange]:
                self._adjustText()

        return super().eventFilter(obj, e)

    def paintEvent(self, e):
        super().paintEvent(e)
        if self.lightBackgroundColor is None:
            return

        painter = QPainter(self)
        painter.setRenderHints(QPainter.Antialiasing)
        painter.setPen(Qt.NoPen)

        if isDarkTheme():
            painter.setBrush(self.darkBackgroundColor)
        else:
            painter.setBrush(self.lightBackgroundColor)

        rect = self.rect().adjusted(1, 1, -1, -1)
        painter.drawRoundedRect(rect, 6, 6)


class ClearCache(QThread):
    isFinished = pyqtSignal(bool)

    def __init__(self):
        super(ClearCache, self).__init__()

    def run(self):
        if os.path.exists("./Log"):
            try:
                for item in os.listdir("./Log"):
                    item_path = os.path.join("./Log", item)
                    if os.path.isfile(item_path):
                        os.remove(item_path)
            except:
                pass
        if os.path.exists("FastCopy2.ini"):
            try:
                os.remove("FastCopy2.ini")
            except:
                pass
        self.isFinished.emit(True)


class SettingInterface(SmoothScrollArea):
    sourceFolderChanged = pyqtSignal(list)

    def __init__(self, parent=None):
        super().__init__(parent=parent)
        self.scrollWidget = QWidget()
        self.stateTooltip = None
        self.expandLayout = ExpandLayout(self.scrollWidget)
        self.enableTransparentBackground()
        self.settingLabel = QLabel("设置", self)
        if darkdetect.isDark():
            self.scrollWidget.setStyleSheet("background-color: rgba(39, 39, 39, 0);")
            self.settingLabel.setStyleSheet("font: 33px 'Microsoft YaHei Light'; background-color: transparent; color: white;")
        else:
            self.scrollWidget.setStyleSheet("background-color: rgba(249, 249, 249, 0);")
            self.settingLabel.setStyleSheet("font: 33px 'Microsoft YaHei Light'; background-color: transparent;")

        self.sourceGroup = SettingCardGroup('源', self.scrollWidget)
        self.actGroup = SettingCardGroup('行为', self.scrollWidget)
        self.performanceGroup = SettingCardGroup('性能', self.scrollWidget)
        self.filterGroup = SettingCardGroup('文件过滤', self.scrollWidget)
        self.storageGroup = SettingCardGroup('存储', self.scrollWidget)
        self.advanceGroup = SettingCardGroup('高级', self.scrollWidget)
        self.aboutGroup = SettingCardGroup('关于', self.scrollWidget)
        self.optionSourceCard = ComboBoxSettingCard(
            cfg.IsSourceCloud,
            FluentFontIcon("\ue8b7"),
            '源文件夹',
            '选择源文件夹来源',
            texts=['云上春晖', '自定义'],
            parent=self.sourceGroup
        )
        self.autoRunCard = SwitchSettingCard(
            FluentFontIcon("\ue7e8"),
            "开机时启动",
            "",
            configItem=cfg.AutoRun,
            parent=self.actGroup
        )
        self.notifyCard = SwitchSettingCard(
            FluentFontIcon("\uea8f"),
            "完成后通知",
            "",
            configItem=cfg.Notify,
            parent=self.actGroup
        )
        self.cloudCard = PushSettingCard(
            '选择文件夹',
            FluentFontIcon("\ue753"),
            "云上春晖",
            cfg.get(cfg.sourceFolder),
            self.sourceGroup
        )
        self.customFolderCard = CustomFolderListSettingCard(
            "自定义",
            "展开选项卡以设置",
            parent=self.sourceGroup)
        self.scanCycleCard = RangeSettingCard(
            cfg.ScanCycle,
            FluentFontIcon("\ue916"),
            '扫描周期',
            parent=self.performanceGroup
        )
        self.concurrentProcessCard = SpinBoxSettingCard(
            cfg.ConcurrentProcess,
            FluentFontIcon("\ue8e4"),
            '并行进程数',
            parent=self.performanceGroup
        )
        self.bufSizeCard = BufSizeSettingCard(
            cfg.BufSize,
            FluentFontIcon("\ueb05"),
            '缓冲区大小',
            parent=self.performanceGroup
        )
        self.infoBar = InformationBar(
            title="",
            content="以下选项会对所有任务产生直接而现实的影响",
            parent=self.filterGroup
        )
        self.isSkipEmptyDirCard = SwitchSettingCard(
            FluentFontIcon("\uecc9"),
            "跳过空文件夹",
            "不复制空文件夹及过滤后为空的文件夹",
            configItem=cfg.IsSkipEmptyDir,
            parent=self.filterGroup
        )
        self.sizeFilterCard = SizeFilterSettingCard(
            title='大小过滤',
            content='过滤大于指定大小的文件',
            parent=self.filterGroup
        )
        self.typeFilterCard = TypeFilterSettingCard(
            title='类型过滤',
            content='排除或包含指定类型的文件',
            parent=self.filterGroup
        )
        self.clearCard = PushSettingCard(
            '清除',
            FluentFontIcon("\uea99"),
            '清除缓存',
            self.getSize(),
            self.storageGroup
        )
        self.recoverCard = PushSettingCard(
            '恢复',
            FluentFontIcon("\uebc4"),
            '恢复默认设置',
            '重置所有设置为初始值',
            self.advanceGroup
        )
        self.devCard = PushSettingCard(
            '打开',
            FluentFontIcon("\uec7a"),
            '开发者选项',
            '打开配置文件',
            self.advanceGroup
        )
        self.helpCard = PrimaryPushSettingCard(
            '转到帮助',
            FluentFontIcon("\uea6b"),
            '帮助',
            '提示与常见问题',
            self.aboutGroup
        )
        self.aboutESCard = PrimaryPushSettingCard(
            '检查更新',
            QIcon(':/icon.png'),
            'Presto',
            VERSION,
            self.aboutGroup
        )
        self.feedbackCard = PrimaryPushSettingCard(
            '提供反馈',
            FluentFontIcon("\ued15"),
            '反馈',
            '报告问题或提出建议',
            self.aboutGroup
        )
        self.__initWidget()

    def __initWidget(self):
        self.resize(500, 800)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setViewportMargins(0, 60, 0, 5)
        self.setWidget(self.scrollWidget)
        self.setWidgetResizable(True)
        self.__initLayout()
        self.__connectSignalToSlot()
        if self.optionSourceCard.comboBox.text() == "云上春晖":
            self.cloudCard.setVisible(True)
            self.customFolderCard.setVisible(False)
            self.sourceGroup.adjustSize()
        else:
            self.cloudCard.setVisible(False)
            self.customFolderCard.setVisible(True)
            self.sourceGroup.adjustSize()

    def __initLayout(self):
        self.settingLabel.move(30, 10)

        self.sourceGroup.addSettingCard(self.optionSourceCard)
        self.sourceGroup.addSettingCard(self.cloudCard)
        self.sourceGroup.addSettingCard(self.customFolderCard)
        self.actGroup.addSettingCard(self.autoRunCard)
        self.actGroup.addSettingCard(self.notifyCard)
        self.performanceGroup.addSettingCard(self.scanCycleCard)
        self.performanceGroup.addSettingCard(self.concurrentProcessCard)
        self.performanceGroup.addSettingCard(self.bufSizeCard)
        self.filterGroup.addSettingCard(self.infoBar)
        self.filterGroup.addSettingCard(self.isSkipEmptyDirCard)
        self.filterGroup.addSettingCard(self.sizeFilterCard)
        self.filterGroup.addSettingCard(self.typeFilterCard)
        self.storageGroup.addSettingCard(self.clearCard)
        self.advanceGroup.addSettingCard(self.recoverCard)
        self.advanceGroup.addSettingCard(self.devCard)
        self.aboutGroup.addSettingCard(self.aboutESCard)
        self.aboutGroup.addSettingCard(self.helpCard)
        self.aboutGroup.addSettingCard(self.feedbackCard)
        self.expandLayout.setSpacing(28)
        self.expandLayout.setContentsMargins(25, 20, 25, 20)
        self.expandLayout.addWidget(self.sourceGroup)
        self.expandLayout.addWidget(self.actGroup)
        self.expandLayout.addWidget(self.performanceGroup)
        self.expandLayout.addWidget(self.filterGroup)
        self.expandLayout.addWidget(self.storageGroup)
        self.expandLayout.addWidget(self.advanceGroup)
        self.expandLayout.addWidget(self.aboutGroup)

    def getSize(self):
        try:
            size = os.path.getsize("FastCopy2.ini")
        except:
            size = 0
        if os.path.exists("./Log"):
            for root, dirs, files in os.walk("./Log"):
                try:
                    size += sum([os.path.getsize(os.path.join(root, name)) for name in files])
                except:
                    pass
        kbSize = float(size / 1024)
        if kbSize >= 1024 * 1024:
            return str(round(kbSize / 1024 / 1024, 1)) + ' GB'
        elif kbSize >= 1024:
            return str(round(kbSize / 1024, 1)) + ' MB'
        else:
            return str(int(kbSize)) + ' KB'

    def onOptionSourceCard(self):
        if self.optionSourceCard.comboBox.text() == "云上春晖":
            self.cloudCard.setVisible(True)
            self.customFolderCard.setVisible(False)
            self.sourceGroup.adjustSize()

            folder = cfg.sourceFolder.value
            cfg.set(cfg.yuwenFolder, os.path.join(folder, '语文'))
            cfg.set(cfg.shuxueFolder, os.path.join(folder, '数学'))
            cfg.set(cfg.yingyuFolder, os.path.join(folder, '英语'))
            cfg.set(cfg.wuliFolder, os.path.join(folder, '物理'))
            cfg.set(cfg.huaxueFolder, os.path.join(folder, '化学'))
            cfg.set(cfg.shengwuFolder, os.path.join(folder, '生物'))
            cfg.set(cfg.zhengzhiFolder, os.path.join(folder, '政治'))
            cfg.set(cfg.lishiFolder, os.path.join(folder, '历史'))
            cfg.set(cfg.diliFolder, os.path.join(folder, '地理'))
            cfg.set(cfg.jishuFolder, os.path.join(folder, '技术'))
            cfg.set(cfg.ziliaoFolder, os.path.join(folder, '资料'))
            self.customFolderCard.updateContent()
        else:
            self.cloudCard.setVisible(False)
            self.customFolderCard.setVisible(True)
            self.sourceGroup.adjustSize()

    def onCloudCard(self):
        try:
            cmd = ['ping', '-n', '1', '-w', '1000', '10.181.201.188']
            result = subprocess.run(
                cmd,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            if result.returncode == 0:
                self.showFolderDialog(True)
            else:
                self.showFolderDialog(False)
        except:
            self.showFolderDialog(False)

    def showFolderDialog(self, exists):
        if exists:
            folder = QFileDialog.getExistingDirectory(self, "选择文件夹", r"\\10.181.201.188\云上春晖")
        else:
            folder = QFileDialog.getExistingDirectory(self, "选择文件夹", QDir.homePath())
        if not folder or cfg.get(cfg.sourceFolder) == folder:
            return

        self.cloudCard.setContent(folder)
        cfg.set(cfg.sourceFolder, folder)
        cfg.set(cfg.yuwenFolder, os.path.join(folder, '语文'))
        cfg.set(cfg.shuxueFolder, os.path.join(folder, '数学'))
        cfg.set(cfg.yingyuFolder, os.path.join(folder, '英语'))
        cfg.set(cfg.wuliFolder, os.path.join(folder, '物理'))
        cfg.set(cfg.huaxueFolder, os.path.join(folder, '化学'))
        cfg.set(cfg.shengwuFolder, os.path.join(folder, '生物'))
        cfg.set(cfg.zhengzhiFolder, os.path.join(folder, '政治'))
        cfg.set(cfg.lishiFolder, os.path.join(folder, '历史'))
        cfg.set(cfg.diliFolder, os.path.join(folder, '地理'))
        cfg.set(cfg.jishuFolder, os.path.join(folder, '技术'))
        cfg.set(cfg.ziliaoFolder, os.path.join(folder, '资料'))
        self.customFolderCard.updateContent()

    def onClearFinished(self):
        self.clearCacheThread.exit(0)
        self.clearCard.contentLabel.setText(self.getSize())
        self.clearCard.button.setText('已清除')
        QTimer.singleShot(2000, lambda: self.clearCard.button.setText('清除'))
        self.clearCard.button.setDisabled(False)

    def clearCache(self):
        w = MessageBox(
            '清除缓存',
            '缓存包含日志文件。点击确定以继续。',
            self.window()
        )
        w.yesButton.setText('确定')
        w.cancelButton.setText('取消')
        if w.exec():
            self.clearCard.button.setText('清除中')
            self.clearCard.button.setDisabled(True)
            self.clearCacheThread = ClearCache()
            self.clearCacheThread.start()
            self.clearCacheThread.isFinished.connect(self.onClearFinished)

    def recoverConfig(self):
        w = MessageBox(
            '恢复默认设置',
            '点击确定以重置所有设置为默认值。',
            self.window()
        )
        w.yesButton.setText('确定')
        w.cancelButton.setText('取消')
        if w.exec():
            self.autoRunCard.setChecked(True)
            self.notifyCard.setChecked(True)
            self.scanCycleCard.setValue(10)
            self.concurrentProcessCard.setValue(3)
            self.bufSizeCard.setValue(BufSize._256)
            self.isSkipEmptyDirCard.setChecked(False)
            self.sizeFilterCard.setChecked(False)
            self.typeFilterCard.setChecked(False)

    def openConfig(self):
        w = MessageBox(
            '打开配置文件',
            '随意修改配置文件可能导致程序异常。点击确定以继续。',
            self.window()
        )
        w.yesButton.setText('确定')
        w.cancelButton.setText('取消')
        if w.exec():
            os.startfile(os.path.join(os.path.expanduser('~'), '.Presto', 'config', 'config.json'))

    def onHelpAction(self):
        if os.path.exists(os.path.abspath("./Doc/PrestoHelp.html")):
            os.startfile(os.path.abspath("./Doc/PrestoHelp.html"))
        else:
            QDesktopServices.openUrl(QUrl("https://sudo0015.github.io/post/Presto%20-bang-zhu.html"))

    def __connectSignalToSlot(self):
        self.optionSourceCard.comboBox.currentTextChanged.connect(self.onOptionSourceCard)
        self.cloudCard.clicked.connect(self.onCloudCard)
        self.clearCard.clicked.connect(self.clearCache)
        self.recoverCard.clicked.connect(self.recoverConfig)
        self.devCard.clicked.connect(self.openConfig)
        self.helpCard.clicked.connect(self.onHelpAction)
        self.feedbackCard.clicked.connect(lambda: QDesktopServices.openUrl(QUrl("https://github.com/sudo0015/Presto/issues")))


class CustomTypeMessageBoxItem(QWidget):
    deleteSignal = pyqtSignal(bool)

    def __init__(self, CustomType: str, parent=None):
        super().__init__(parent=parent)
        self.customType = CustomType
        self.hBoxLayout = QHBoxLayout(self)
        self.titleLabel = QLabel(self.customType, self)
        if darkdetect.isDark():
            self.titleLabel.setStyleSheet(
                "font: 14px 'Segoe UI', 'Microsoft YaHei', 'PingFang SC'; padding: 0; border: none; background-color: transparent; color: white;")
        else:
            self.titleLabel.setStyleSheet(
                "font: 14px 'Segoe UI', 'Microsoft YaHei', 'PingFang SC'; padding: 0; border: none; background-color: transparent; color: black;")

        self.lineEdit = CustomLineEdit(self)
        self.lineEdit.setFixedWidth(200)
        self.lineEdit.setHidden(True)
        self.lineEdit.acceptSignal.connect(self.onAcceptSignal)

        self.editBtn = TransparentToolButton(self)
        self.editBtn.setIcon(FluentFontIcon("\ue70f"))
        self.editBtn.setToolTip("编辑")
        self.editBtn.installEventFilter(ToolTipFilter(self.editBtn, 0, ToolTipPosition.BOTTOM))
        self.editBtn.clicked.connect(self.edit)

        self.deleteBtn = TransparentToolButton(self)
        self.deleteBtn.setIcon(FluentFontIcon("\ue74d"))
        self.deleteBtn.setToolTip("删除")
        self.deleteBtn.installEventFilter(ToolTipFilter(self.deleteBtn, 0, ToolTipPosition.BOTTOM))
        self.deleteBtn.clicked.connect(self.delete)

        self.setFixedHeight(48)
        self.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Fixed)
        self.hBoxLayout.setContentsMargins(24, 0, 18, 0)
        self.hBoxLayout.addWidget(self.titleLabel, 0, Qt.AlignLeft)
        self.hBoxLayout.addWidget(self.lineEdit, 0, Qt.AlignLeft)
        self.hBoxLayout.addSpacing(16)
        self.hBoxLayout.addStretch(1)
        self.hBoxLayout.addWidget(self.editBtn, 0, Qt.AlignRight)
        self.hBoxLayout.addWidget(self.deleteBtn, 0, Qt.AlignRight)
        self.hBoxLayout.setAlignment(Qt.AlignVCenter)

    def paintEvent(self, e):
        painter = QPainter(self)
        painter.setRenderHints(QPainter.Antialiasing)

        if isDarkTheme():
            painter.setBrush(QColor(255, 255, 255, 13))
            painter.setPen(QColor(0, 0, 0, 50))
        else:
            painter.setBrush(QColor(255, 255, 255, 170))
            painter.setPen(QColor(0, 0, 0, 19))

        painter.drawRoundedRect(self.rect().adjusted(1, 1, -1, -1), 6, 6)

    def onAcceptSignal(self):
        self.editBtn.setEnabled(True)
        if self.lineEdit.isAccept:
            self.titleLabel.setHidden(False)
            self.lineEdit.setHidden(True)

            text = self.lineEdit.text().strip().lower()
            self.customType = '.' + text if text[0] != '.' else text
            self.titleLabel.setText(self.customType)
        else:
            self.titleLabel.setHidden(False)
            self.lineEdit.setHidden(True)
            if self.customType:
                self.titleLabel.setText(self.customType)
            else:
                self.delete()

    def edit(self):
        self.editBtn.setEnabled(False)
        self.lineEdit.setText(self.customType)
        self.titleLabel.setHidden(True)
        self.lineEdit.setHidden(False)
        self.lineEdit.setFocus()

    def delete(self):
        self.deleteSignal.emit(True)


class CustomTypeMessageBox(MessageBoxBase):

    def __init__(self, parent=None):
        super().__init__(parent)
        self.titleLabel = SubtitleLabel('自定义文件类型', self)
        self.countLabel = BodyLabel("0 个项目", self)
        self.addBtn = HyperlinkButton(self)
        self.addBtn.setText("添加文件类型")
        self.addBtn.setIcon(FluentFontIcon("\ue710"))
        self.addBtn.setFixedWidth(148)
        self.addBtn.clicked.connect(self.onAddBtn)
        self.topLayout = QHBoxLayout(self)
        self.topLayout.setContentsMargins(14, 0, 10, 0)
        self.topLayout.addWidget(self.countLabel)
        self.topLayout.addWidget(self.addBtn)

        self.scrollContent = QWidget()
        self.scrollLayout = QVBoxLayout(self.scrollContent)
        self.scrollArea = SmoothScrollArea(self)
        self.scrollArea.setWidget(self.scrollContent)
        self.scrollArea.setWidgetResizable(True)
        self.scrollArea.enableTransparentBackground()
        self.scrollLayout.setContentsMargins(10, 0, 16, 0)
        self.scrollLayout.setSpacing(8)
        self.scrollLayout.addStretch(1)
        self.frames = []
        self.types = [x.strip() for x in cfg.CustomType.value.split(",")][::-1]

        self.emptyLabel = QLabel('没有要显示的内容', self.scrollContent)
        self.emptyLabel.setAlignment(Qt.AlignCenter)
        if darkdetect.isDark():
            self.emptyLabel.setStyleSheet("font: 14px 'Segoe UI', 'Microsoft YaHei', 'PingFang SC'; color: #999999;")
        else:
            self.emptyLabel.setStyleSheet("font: 14px 'Segoe UI', 'Microsoft YaHei', 'PingFang SC'; color: #777777;")
        self.emptyLabel.setHidden(True)
        self.scrollLayout.addWidget(self.emptyLabel, 0, Qt.AlignCenter)

        for i in self.types:
            if i:
                item = CustomTypeMessageBoxItem(i, self.scrollContent)
                item.deleteSignal.connect(lambda _, frame=item: self.removeFrame(frame))
                item.lineEdit.showInfoSignal.connect(self.showInfo)
                self.addFrame(item)

        self.viewLayout.addWidget(self.titleLabel)
        self.viewLayout.addLayout(self.topLayout)
        self.viewLayout.addWidget(self.scrollArea, 1)

        self.yesButton.setText('保存')
        self.cancelButton.setText('取消')

        self.widget.setMinimumSize(500, 420)
        self.updateLayout()

    def onAddBtn(self):
        item = CustomTypeMessageBoxItem("", self.scrollContent)
        item.deleteSignal.connect(lambda _, frame=item: self.removeFrame(frame))
        item.lineEdit.showInfoSignal.connect(self.showInfo)
        item.edit()
        self.addFrame(item)
        self.updateLayout()

    def addFrame(self, frame: QFrame = None) -> QFrame:
        if frame is None:
            frame = QFrame(self.scrollContent)

        self.scrollLayout.insertWidget(0, frame)
        self.frames.append(frame)
        self.updateLayout()

        return frame

    def removeFrame(self, frame: QFrame):
        if frame in self.frames:
            self.scrollLayout.removeWidget(frame)
            frame.setParent(None)
            self.frames.remove(frame)

            self.updateLayout()

    def updateLayout(self):
        self.countLabel.setText(f"{len(self.frames)} 个项目")

        self.scrollContent.adjustSize()
        self.scrollArea.updateGeometry()

        if not self.frames:
            self.emptyLabel.setHidden(False)
            self.emptyLabel.setFixedSize(self.scrollContent.sizeHint().width(), 200)
        else:
            self.emptyLabel.setHidden(True)

    def showInfo(self):
        InfoBar.error(
            title='无效的文件类型',
            content='',
            orient=Qt.Horizontal,
            isClosable=True,
            position=InfoBarPosition.TOP,
            duration=1000,
            parent=self
        )

    def getType(self):
        types = []
        for item in self.frames[::-1]:
            if item.customType:
                types.append(item.customType)
        return ', '.join(types)


class WelcomeMessageBox(MessageBoxBase):

    def __init__(self, parent=None):
        super().__init__(parent)
        self.sourceFolder = ''
        self.titleLabel = SubtitleLabel("欢迎使用 Presto", self)
        self.contentLabel = BodyLabel(self)
        self.contentLabel.setText(
            "Presto 是一款用于文件同步/复制的桌面应用。为方便同学们拷课件而设计。具有简洁、快速、高效、可靠等特性。\n"
            "\n设置云上春晖班级文件夹以继续。"
        )
        self.contentLabel.setWordWrap(True)
        self.cloudCard = PushSettingCard(
            '选择文件夹',
            FluentFontIcon("\ue753"),
            "云上春晖",
            "未设置",
            self
        )
        self.cloudCard.clicked.connect(self.onCloudCard)
        self.helpBtn = HyperlinkButton(self)
        self.helpBtn.setIcon(FluentFontIcon("\uea6b"))
        self.helpBtn.setText("查看完整帮助")
        self.helpBtn.clicked.connect(self.onHelpBtn)

        self.vBoxLayout = QVBoxLayout(self)
        self.vBoxLayout.setContentsMargins(16, 0, 16, 0)
        self.vBoxLayout.setSpacing(18)
        self.vBoxLayout.addWidget(self.contentLabel)
        self.vBoxLayout.addWidget(self.cloudCard)
        self.vBoxLayout.addWidget(self.helpBtn, alignment=Qt.AlignHCenter)

        self.viewLayout.addWidget(self.titleLabel)
        self.viewLayout.addLayout(self.vBoxLayout)
        self.viewLayout.addStretch()

        self.yesButton.setText("保存")
        self.cancelButton.setText("跳过")

        self.widget.setMinimumSize(500, 380)

    def onCloudCard(self):
        if os.path.exists(r"\\10.181.201.188\云上春晖"):
            folder = QFileDialog.getExistingDirectory(self, "选择文件夹", r"\\10.181.201.188\云上春晖")
        else:
            folder = QFileDialog.getExistingDirectory(self, "选择文件夹", QDir.homePath())
        if not folder:
            return

        self.cloudCard.setContent(folder)
        self.sourceFolder = folder
        cfg.set(cfg.sourceFolder, folder)
        cfg.set(cfg.yuwenFolder, os.path.join(folder, '语文'))
        cfg.set(cfg.shuxueFolder, os.path.join(folder, '数学'))
        cfg.set(cfg.yingyuFolder, os.path.join(folder, '英语'))
        cfg.set(cfg.wuliFolder, os.path.join(folder, '物理'))
        cfg.set(cfg.huaxueFolder, os.path.join(folder, '化学'))
        cfg.set(cfg.shengwuFolder, os.path.join(folder, '生物'))
        cfg.set(cfg.zhengzhiFolder, os.path.join(folder, '政治'))
        cfg.set(cfg.lishiFolder, os.path.join(folder, '历史'))
        cfg.set(cfg.diliFolder, os.path.join(folder, '地理'))
        cfg.set(cfg.jishuFolder, os.path.join(folder, '技术'))
        cfg.set(cfg.ziliaoFolder, os.path.join(folder, '资料'))

    def onHelpBtn(self):
        if os.path.exists(os.path.abspath("./Doc/PrestoHelp.html")):
            os.startfile(os.path.abspath("./Doc/PrestoHelp.html"))
        else:
            QDesktopServices.openUrl(QUrl("https://sudo0015.github.io/post/Presto%20-bang-zhu.html"))


class TitleBarBase(QWidget):

    def __init__(self, parent):
        super().__init__(parent)
        self.minBtn = MinimizeButton(parent=self)
        self.closeBtn = CloseButton(parent=self)
        self.maxBtn = MaximizeButton(parent=self)

        self._isDoubleClickEnabled = True

        self.resize(200, 32)
        self.setFixedHeight(32)

        self.minBtn.clicked.connect(self.window().showMinimized)
        self.maxBtn.clicked.connect(self.__toggleMaxState)
        self.closeBtn.clicked.connect(self.window().close)

        self.window().installEventFilter(self)

    def eventFilter(self, obj, e):
        if obj is self.window():
            if e.type() == QEvent.WindowStateChange:
                self.maxBtn.setMaxState(self.window().isMaximized())
                return False

        return super().eventFilter(obj, e)

    def mouseDoubleClickEvent(self, event):
        if event.button() != Qt.LeftButton or not self._isDoubleClickEnabled:
            return

        self.__toggleMaxState()

    def mouseMoveEvent(self, e):
        if sys.platform != "win32" or not self.canDrag(e.pos()):
            return

        startSystemMove(self.window(), e.globalPos())

    def mousePressEvent(self, e):
        if sys.platform == "win32" or not self.canDrag(e.pos()):
            return

        startSystemMove(self.window(), e.globalPos())

    def __toggleMaxState(self):
        if self.window().isMaximized():
            self.window().showNormal()
        else:
            self.window().showMaximized()

        if sys.platform == "win32":
            from qframelesswindow.utils.win32_utils import releaseMouseLeftButton
            releaseMouseLeftButton(self.window().winId())

    def _isDragRegion(self, pos):
        width = 0
        for button in self.findChildren(TitleBarButton):
            if button.isVisible():
                width += button.width()

        return 0 < pos.x() < self.width() - width

    def _hasButtonPressed(self):
        return any(btn.isPressed() for btn in self.findChildren(TitleBarButton))

    def canDrag(self, pos):
        return self._isDragRegion(pos) and not self._hasButtonPressed()

    def setDoubleClickEnabled(self, isEnabled):
        self._isDoubleClickEnabled = isEnabled


class TitleBar(TitleBarBase):

    def __init__(self, parent):
        super().__init__(parent)
        self.hBoxLayout = QHBoxLayout(self)

        self.hBoxLayout.setSpacing(0)
        self.hBoxLayout.setContentsMargins(0, 0, 0, 0)
        self.hBoxLayout.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
        self.hBoxLayout.addStretch(1)
        self.hBoxLayout.addWidget(self.minBtn, 0, Qt.AlignRight)
        self.hBoxLayout.addWidget(self.maxBtn, 0, Qt.AlignRight)
        self.hBoxLayout.addWidget(self.closeBtn, 0, Qt.AlignRight)


class FluentTitleBar(TitleBar):

    def __init__(self, parent):
        super().__init__(parent)
        self.setFixedHeight(48)
        self.hBoxLayout.removeWidget(self.minBtn)
        self.hBoxLayout.removeWidget(self.maxBtn)
        self.hBoxLayout.removeWidget(self.closeBtn)

        self.iconLabel = QLabel(self)
        self.iconLabel.setFixedSize(18, 18)
        self.hBoxLayout.insertWidget(0, self.iconLabel, 0, Qt.AlignLeft | Qt.AlignVCenter)
        self.window().windowIconChanged.connect(self.setIcon)

        self.titleLabel = QLabel(self)
        self.hBoxLayout.insertWidget(1, self.titleLabel, 0, Qt.AlignLeft | Qt.AlignVCenter)
        self.titleLabel.setObjectName('titleLabel')
        self.window().windowTitleChanged.connect(self.setTitle)

        self.vBoxLayout = QVBoxLayout()
        self.buttonLayout = QHBoxLayout()
        self.buttonLayout.setSpacing(0)
        self.buttonLayout.setContentsMargins(0, 0, 0, 0)
        self.buttonLayout.setAlignment(Qt.AlignTop)
        self.buttonLayout.addWidget(self.minBtn)
        self.buttonLayout.addWidget(self.maxBtn)
        self.buttonLayout.addWidget(self.closeBtn)
        self.vBoxLayout.addLayout(self.buttonLayout)
        self.vBoxLayout.addStretch(1)
        self.hBoxLayout.addLayout(self.vBoxLayout, 0)

        FluentStyleSheet.FLUENT_WINDOW.apply(self)

    def setTitle(self, title):
        self.titleLabel.setText(title)
        self.titleLabel.adjustSize()

    def setIcon(self, icon):
        self.iconLabel.setPixmap(QIcon(icon).pixmap(18, 18))


class MSFluentTitleBar(FluentTitleBar):

    def __init__(self, parent):
        super().__init__(parent)
        self.hBoxLayout.insertSpacing(0, 20)
        self.hBoxLayout.insertSpacing(2, 2)


class MSFluentWindow(FluentWindowBase):

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setTitleBar(MSFluentTitleBar(self))

        self.navigationInterface = NavigationBar(self)

        self.hBoxLayout.setContentsMargins(0, 48, 0, 0)
        self.hBoxLayout.addWidget(self.navigationInterface)
        self.hBoxLayout.addWidget(self.stackedWidget, 1)

        self.titleBar.raise_()
        self.titleBar.setAttribute(Qt.WA_StyledBackground)

    def addSubInterface(self, interface: QWidget, icon: Union[QIcon, str], text: str,
                        selectedIcon=None, position=NavigationItemPosition.TOP,
                        isTransparent=False) -> NavigationBarPushButton:
        if not interface.objectName():
            sys.exit()

        interface.setProperty("isStackedTransparent", isTransparent)
        self.stackedWidget.addWidget(interface)

        routeKey = interface.objectName()
        item = self.navigationInterface.addItem(
            routeKey=routeKey,
            icon=icon,
            text=text,
            onClick=lambda: self.switchTo(interface),
            selectedIcon=selectedIcon,
            position=position
        )

        if self.stackedWidget.count() == 1:
            self.stackedWidget.currentChanged.connect(self._onCurrentInterfaceChanged)
            self.navigationInterface.setCurrentItem(routeKey)
            qrouter.setDefaultRouteKey(self.stackedWidget, routeKey)

        self._updateStackedBackground()

        return item


class Main(MSFluentWindow):

    def __init__(self, parent=None):
        super().__init__(parent)
        if isDarkTheme():
            setThemeColor(QColor(113, 89, 249))
        else:
            setThemeColor(QColor(90, 51, 174))
        self.resize(800, 600)
        self.setMinimumSize(640, 550)
        self.setWindowTitle('Presto 设置')
        self.setWindowIcon(QIcon(':/icon.png'))
        self.titleBar.raise_()
        self.move(
            QApplication.screens()[0].size().width() // 2 - self.width() // 2,
            QApplication.screens()[0].size().height() // 2 - self.height() // 2
        )

        self.splashScreen = SplashScreen(self.windowIcon(), self)
        self.splashScreen.raise_()

        self.settingInterface = SettingInterface(self)
        self.settingInterface.setObjectName('settingInterface')
        self.addSubInterface(
            self.settingInterface,
            FluentFontIcon("\ue713"),
            '设置',
            FluentFontIcon("\uf8b0")
        )
        self.navigationInterface.addItem(
            routeKey='Log',
            icon=FluentFontIcon("\ue8a5"),
            text='日志',
            onClick=self.onLogBtn,
            selectable=False,
            position=NavigationItemPosition.TOP
        )
        self.navigationInterface.addItem(
            routeKey='Help',
            icon=FluentFontIcon("\uea6b"),
            text='帮助',
            onClick=self.onHelpBtn,
            selectable=False,
            position=NavigationItemPosition.BOTTOM
        )
        self.navigationInterface.setCurrentItem(self.settingInterface.objectName())

        self.splashScreen.finish()
        QTimer.singleShot(300, self.checkFirstRun)

    def checkFirstRun(self):
        if not os.path.exists(os.path.join(os.path.expanduser('~'), '.Presto', 'config', 'config.json')):
            w = WelcomeMessageBox(self.window())
            if w.exec():
                if w.sourceFolder:
                    self.settingInterface.cloudCard.setContent(w.sourceFolder)
                    self.settingInterface.customFolderCard.updateContent()

    def onHelpBtn(self):
        if os.path.exists(os.path.abspath("./Doc/PrestoHelp.html")):
            os.startfile(os.path.abspath("./Doc/PrestoHelp.html"))
        else:
            QDesktopServices.openUrl(QUrl("https://sudo0015.github.io/post/Presto%20-bang-zhu.html"))

    def onLogBtn(self):
        try:
            os.startfile('Log')
        except:
            pass


if __name__ == '__main__':
    with Mutex():
        QApplication.setHighDpiScaleFactorRoundingPolicy(
            Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps)

        if darkdetect.isDark():
            setTheme(Theme.DARK)
        else:
            setTheme(Theme.LIGHT)

        app = QApplication(sys.argv)
        w = Main()
        w.show()
        app.exec()
