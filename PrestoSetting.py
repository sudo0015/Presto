# -*- coding: utf-8 -*-

import os
import sys
import darkdetect
import subprocess
import portalocker
import PrestoResource
from typing import Union
from webbrowser import open as WebOpen
from PrestoConfig import cfg, BufSize, VERSION, YEAR
from pygetwindow import getWindowsWithTitle as GetWindow
from PyQt5.QtCore import Qt, pyqtSignal, QTimer, QThread, QRectF, QEasingCurve, QEvent
from PyQt5.QtGui import QColor, QIcon, QPainter, QTextCursor, QPainterPath
from PyQt5.QtWidgets import QFrame, QApplication, QWidget, QHBoxLayout, QFileDialog, QLabel, QVBoxLayout, \
    QPushButton, QButtonGroup, QTextBrowser, QTextEdit, QSizePolicy, QLineEdit, QSpinBox, QScrollArea, \
    QScroller, QAction
from qfluentwidgets import NavigationItemPosition, SubtitleLabel, MessageBox, ExpandLayout, \
    SettingCardGroup, RadioButton, ExpandSettingCard, ComboBox, SwitchButton, IndicatorPosition, qconfig, \
    isDarkTheme, ConfigItem, OptionsConfigItem, FluentStyleSheet, HyperlinkButton, Slider, IconWidget, drawIcon, \
    setThemeColor, ImageLabel, MessageBoxBase, SmoothScrollDelegate, setFont, themeColor, setTheme, Theme, qrouter, \
    NavigationBar, NavigationBarPushButton, BodyLabel, InfoBadge, SplashScreen
from qfluentwidgets.components.widgets.line_edit import EditLayer
from qfluentwidgets.components.widgets.menu import MenuAnimationType, RoundMenu
from qfluentwidgets.components.widgets.spin_box import SpinButton, SpinIcon
from qfluentwidgets.window.fluent_window import FluentWindowBase
from qframelesswindow.titlebar import MinimizeButton, CloseButton, MaximizeButton
from qframelesswindow import TitleBarButton
from qframelesswindow.utils import startSystemMove
from qfluentwidgets import FluentIcon as FIF


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


class SmoothScrollArea(QScrollArea):

    def __init__(self, parent=None):
        super().__init__(parent)
        self.delegate = SmoothScrollDelegate(self, True)
        QScroller.grabGesture(self.viewport(), QScroller.TouchGesture)

    def setScrollAnimation(self, orient, duration, easing=QEasingCurve.OutCubic):
        """ set scroll animation

        Parameters
        ----------
        orient: Orient
            scroll orientation

        duration: int
            scroll duration

        easing: QEasingCurve
            animation type
        """
        bar = self.delegate.hScrollBar if orient == Qt.Horizontal else self.delegate.vScrollBar
        bar.setScrollAnimation(duration, easing)

    def enableTransparentBackground(self):
        self.setStyleSheet("QScrollArea{border: none; background: transparent}")

        if self.widget():
            self.widget().setStyleSheet("QWidget{background: transparent}")


class EditMenu(RoundMenu):
    """ Edit menu """

    def createActions(self):
        self.cutAct = QAction(
            FIF.CUT.icon(),
            "剪切",
            self,
            shortcut="Ctrl+X",
            triggered=self.parent().cut,
        )
        self.copyAct = QAction(
            FIF.COPY.icon(),
            "复制",
            self,
            shortcut="Ctrl+C",
            triggered=self.parent().copy,
        )
        self.pasteAct = QAction(
            FIF.PASTE.icon(),
            "粘贴",
            self,
            shortcut="Ctrl+V",
            triggered=self.parent().paste,
        )
        self.cancelAct = QAction(
            FIF.CANCEL.icon(),
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
                    self.addActions(
                        self.action_list[:2] + self.action_list[3:])
            else:
                if self.parent().isReadOnly():
                    self.addAction(self.selectAllAct)
                else:
                    self.addActions(self.action_list[3:])

        super().exec(pos, ani, aniType)


class LineEditMenu(EditMenu):
    """ Line edit menu """

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


class SpinBoxBase:
    """ Spin box ui """

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
        """ set whether the spin symbol is visible """
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
        path.addRoundedRect(QRectF(0, h-10, w, 10), 5, 5)

        rectPath = QPainterPath()
        rectPath.addRect(0, h-10, w, 8)
        path = path.subtracted(rectPath)

        painter.fillPath(path, themeColor())

    def paintEvent(self, e):
        super().paintEvent(e)
        self._drawBorderBottom()


class InlineSpinBoxBase(SpinBoxBase):
    """ Inline spin box base """

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


class TextBrowser(QTextBrowser):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.layer = EditLayer(self)
        self.scrollDelegate = SmoothScrollDelegate(self)
        FluentStyleSheet.LINE_EDIT.apply(self)
        setFont(self)

    def contextMenuEvent(self, e):
        menu = TextEditMenu(self)
        menu.exec(e.globalPos())


class SettingIconWidget(IconWidget):

    def paintEvent(self, e):
        painter = QPainter(self)

        if not self.isEnabled():
            painter.setOpacity(0.36)

        painter.setRenderHints(QPainter.Antialiasing | QPainter.SmoothPixmapTransform)
        drawIcon(self._icon, painter, self.rect())


class SettingCard(QFrame):
    def __init__(self, icon: Union[str, QIcon, FIF], title, content=None, parent=None):
        """
        Parameters
        ----------
        icon: str | QIcon | FluentIconBase
            the icon to be drawn

        title: str
            the title of card

        content: str
            the content of card

        parent: QWidget
            parent widget
        """
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

        # initialize layout
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

    def __init__(self, icon: Union[str, QIcon, FIF], title, content=None,
                 configItem: ConfigItem = None, parent=None):
        """
        Parameters
        ----------
        icon: str | QIcon | FluentIconBase
            the icon to be drawn

        title: str
            the title of card

        content: str
            the content of card

        configItem: ConfigItem
            configuration item operated by the card

        parent: QWidget
            parent widget
        """
        super().__init__(icon, title, content, parent)
        self.configItem = configItem
        self.switchButton = SwitchButton('关', self, IndicatorPosition.RIGHT)

        if configItem:
            self.setValue(qconfig.get(configItem))
            configItem.valueChanged.connect(self.setValue)

        # add switch button to layout
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

    def __init__(self, configItem, icon: Union[str, QIcon, FIF], title, content=None, parent=None):
        """
        Parameters
        ----------
        configItem: RangeConfigItem
            configuration item operated by the card

        icon: str | QIcon | FluentIconBase
            the icon to be drawn

        title: str
            the title of card

        content: str
            the content of card

        parent: QWidget
            parent widget
        """
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

    def __init__(self, text, icon: Union[str, QIcon, FIF], title, content=None, parent=None):
        """
        Parameters
        ----------
        text: str
            the text of push button

        icon: str | QIcon | FluentIconBase
            the icon to be drawn

        title: str
            the title of card

        content: str
            the content of card

        parent: QWidget
            parent widget
        """
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

    def __init__(self, configItem: ConfigItem, icon: Union[str, QIcon, FIF], title, content=None, parent=None):
        """
        Parameters
        ----------
        configItem: ConfigItem
            configuration item operated by the card

        icon: str | QIcon | FluentIconBase
            the icon to be drawn

        title: str
            the title of card

        content: str
            the content of card

        parent: QWidget
            parent widget
        """
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
    def __init__(self, configItem: OptionsConfigItem, icon: Union[str, QIcon, FIF], title, content=None, texts=None, parent=None):
        """
        Parameters
        ----------
        configItem: OptionsConfigItem
            configuration item operated by the card

        icon: str | QIcon | FluentIconBase
            the icon to be drawn

        title: str
            the title of card

        content: str
            the content of card

        texts: List[str]
            the text of items

        parent: QWidget
            parent widget
        """
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


class OptionsSettingCard(ExpandSettingCard):
    optionChanged = pyqtSignal(OptionsConfigItem)

    def __init__(self, configItem, icon: Union[str, QIcon, FIF], title, content=None, texts=None, parent=None):
        """
        Parameters
        ----------
        configItem: OptionsConfigItem
            options config item

        icon: str | QIcon | FluentIconBase
            the icon to be drawn

        title: str
            the title of setting card

        content: str
            the content of setting card

        texts: List[str]
            the texts of radio buttons

        parent: QWidget
            parent window
        """
        super().__init__(icon, title, content, parent)
        self.texts = texts or []
        self.configItem = configItem
        self.configName = configItem.name
        self.choiceLabel = QLabel(self)
        self.buttonGroup = QButtonGroup(self)

        self.addWidget(self.choiceLabel)

        self.viewLayout.setSpacing(19)
        self.viewLayout.setContentsMargins(48, 18, 0, 18)
        for text, option in zip(texts, configItem.options):
            button = RadioButton(text, self.view)
            self.buttonGroup.addButton(button)
            self.viewLayout.addWidget(button)
            button.setProperty(self.configName, option)

        self._adjustViewSize()
        self.setValue(qconfig.get(self.configItem))
        configItem.valueChanged.connect(self.setValue)
        self.buttonGroup.buttonClicked.connect(self.__onButtonClicked)
        if darkdetect.isDark():
            self.choiceLabel.setStyleSheet("font: 13px 'Segoe UI', 'Microsoft YaHei', 'PingFang SC'; padding: 0; border: none; background-color: transparent; color: white;")
        else:
            self.choiceLabel.setStyleSheet("font: 13px 'Segoe UI', 'Microsoft YaHei', 'PingFang SC'; padding: 0; border: none; background-color: transparent; color: black;")

    def __onButtonClicked(self, button: RadioButton):
        if button.text() == self.choiceLabel.text():
            return

        value = button.property(self.configName)
        qconfig.set(self.configItem, value)

        self.choiceLabel.setText(button.text())
        self.choiceLabel.adjustSize()
        self.optionChanged.emit(self.configItem)

    def setValue(self, value):
        qconfig.set(self.configItem, value)

        for button in self.buttonGroup.buttons():
            isChecked = button.property(self.configName) == value
            button.setChecked(isChecked)

            if isChecked:
                self.choiceLabel.setText(button.text())


class FolderItem(QWidget):
    def __init__(self, folder: str, parent=None):
        super().__init__(parent=parent)
        self.hBoxLayout = QHBoxLayout(self)
        self.badge = InfoBadge.custom(folder[:2], '#9a9a9a', '#8e8e8e')
        self.folderLabel = BodyLabel(folder[4:], self)
        if darkdetect.isDark():
            self.folderLabel.setStyleSheet("font: 13px 'Segoe UI', 'Microsoft YaHei', 'PingFang SC'; padding: 0; border: none; background-color: transparent; color: white;")
        else:
            self.folderLabel.setStyleSheet("font: 13px 'Segoe UI', 'Microsoft YaHei', 'PingFang SC'; padding: 0; border: none; background-color: transparent; color: black;")

        self.changeButton = HyperlinkButton(self)
        self.changeButton.setText("更改")

        self.setFixedHeight(53)
        self.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Fixed)
        self.hBoxLayout.setContentsMargins(48, 0, 60, 0)
        self.hBoxLayout.addWidget(self.badge, 0, Qt.AlignLeft)
        self.hBoxLayout.addWidget(self.folderLabel, 0, Qt.AlignLeft)
        self.hBoxLayout.addSpacing(16)
        self.hBoxLayout.addStretch(1)
        self.hBoxLayout.addWidget(self.changeButton, 0, Qt.AlignRight)
        self.hBoxLayout.setAlignment(Qt.AlignVCenter)

    def setFolder(self, text):
        self.folderLabel.setText(text[4:])
        if darkdetect.isDark():
            self.folderLabel.setStyleSheet("font: 13px 'Segoe UI', 'Microsoft YaHei', 'PingFang SC'; padding: 0; border: none; background-color: transparent; color: white;")
        else:
            self.folderLabel.setStyleSheet("font: 13px 'Segoe UI', 'Microsoft YaHei', 'PingFang SC'; padding: 0; border: none; background-color: transparent; color: black;")


class CustomFolderListSettingCard(ExpandSettingCard):
    folderChanged = pyqtSignal(list)

    def __init__(self, title: str, content: str = None, directory=cfg.sourceFolder.value, parent=None):
        """
        Parameters
        ----------
        title: str
            the title of card

        content: str
            the content of card

        directory: str
            working directory of file dialog

        parent: QWidget
            parent widget
        """
        super().__init__(FIF.FOLDER_ADD, title, content, parent)
        self._dialogDirectory = directory
        self.__initWidget()

    def __initWidget(self):
        self.viewLayout.setSpacing(0)
        self.viewLayout.setAlignment(Qt.AlignTop)
        self.viewLayout.setContentsMargins(0, 0, 0, 0)

        self.yuwenItem = FolderItem("语文: " + cfg.yuwenFolder.value, self.view)
        self.shuxueItem = FolderItem("数学: " + cfg.shuxueFolder.value, self.view)
        self.yingyuItem = FolderItem("英语: " + cfg.yingyuFolder.value, self.view)
        self.wuliItem = FolderItem("物理: " + cfg.wuliFolder.value, self.view)
        self.huaxueItem = FolderItem("化学: " + cfg.huaxueFolder.value, self.view)
        self.shengwuItem = FolderItem("生物: " + cfg.shengwuFolder.value, self.view)
        self.zhengzhiItem = FolderItem("政治: " + cfg.zhengzhiFolder.value, self.view)
        self.lishiItem = FolderItem("历史: " + cfg.lishiFolder.value, self.view)
        self.diliItem = FolderItem("地理: " + cfg.diliFolder.value, self.view)
        self.jishuItem = FolderItem("技术: " + cfg.jishuFolder.value, self.view)
        self.ziliaoItem = FolderItem("资料: " + cfg.ziliaoFolder.value, self.view)
        self.yuwenItem.changeButton.clicked.connect(lambda: self.showFolderDialog(1))
        self.shuxueItem.changeButton.clicked.connect(lambda: self.showFolderDialog(2))
        self.yingyuItem.changeButton.clicked.connect(lambda: self.showFolderDialog(3))
        self.wuliItem.changeButton.clicked.connect(lambda: self.showFolderDialog(4))
        self.huaxueItem.changeButton.clicked.connect(lambda: self.showFolderDialog(5))
        self.shengwuItem.changeButton.clicked.connect(lambda: self.showFolderDialog(6))
        self.zhengzhiItem.changeButton.clicked.connect(lambda: self.showFolderDialog(7))
        self.lishiItem.changeButton.clicked.connect(lambda: self.showFolderDialog(8))
        self.diliItem.changeButton.clicked.connect(lambda: self.showFolderDialog(9))
        self.jishuItem.changeButton.clicked.connect(lambda: self.showFolderDialog(10))
        self.ziliaoItem.changeButton.clicked.connect(lambda: self.showFolderDialog(11))

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

    def showFolderDialog(self, index):
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹", self._dialogDirectory)
        if not folder:
            return

        if index==1:
            cfg.set(cfg.yuwenFolder, folder)
            self.yuwenItem.setFolder("语文: " + cfg.yuwenFolder.value)
        elif index==2:
            cfg.set(cfg.shuxueFolder, folder)
            self.shuxueItem.setFolder("数学: " + cfg.shuxueFolder.value)
        elif index==3:
            cfg.set(cfg.yingyuFolder, folder)
            self.yingyuItem.setFolder("英语: " + cfg.yingyuFolder.value)
        elif index==4:
            cfg.set(cfg.wuliFolder, folder)
            self.wuliItem.setFolder("物理: " + cfg.wuliFolder.value)
        elif index==5:
            cfg.set(cfg.huaxueFolder, folder)
            self.huaxueItem.setFolder("化学: " + cfg.huaxueFolder.value)
        elif index==6:
            cfg.set(cfg.shengwuFolder, folder)
            self.shengwuItem.setFolder("生物: " + cfg.shengwuFolder.value)
        elif index==7:
            cfg.set(cfg.zhengzhiFolder, folder)
            self.zhengzhiItem.setFolder("政治: " + cfg.zhengzhiFolder.value)
        elif index==8:
            cfg.set(cfg.lishiFolder, folder)
            self.lishiItem.setFolder("历史: " + cfg.lishiFolder.value)
        elif index==9:
            cfg.set(cfg.lishiFolder, folder)
            self.diliItem.setFolder("地理: " + cfg.diliFolder.value)
        elif index==10:
            cfg.set(cfg.jishuFolder, folder)
            self.jishuItem.setFolder("技术: " + cfg.jishuFolder.value)
        else:
            cfg.set(cfg.ziliaoFolder, folder)
            self.ziliaoItem.setFolder("资料: " + cfg.zilaioFolder.value)

    def updateContent(self):
        self.yuwenItem.setFolder("语文: " + cfg.yuwenFolder.value)
        self.shuxueItem.setFolder("数学: " + cfg.shuxueFolder.value)
        self.yingyuItem.setFolder("英语: " + cfg.yingyuFolder.value)
        self.wuliItem.setFolder("物理: " + cfg.wuliFolder.value)
        self.huaxueItem.setFolder("化学: " + cfg.huaxueFolder.value)
        self.shengwuItem.setFolder("生物: " + cfg.shengwuFolder.value)
        self.zhengzhiItem.setFolder("政治: " + cfg.zhengzhiFolder.value)
        self.lishiItem.setFolder("历史: " + cfg.lishiFolder.value)
        self.diliItem.setFolder("地理: " + cfg.diliFolder.value)
        self.jishuItem.setFolder("技术: " + cfg.jishuFolder.value)
        self.ziliaoItem.setFolder("资料: " + cfg.ziliaoFolder.value)


class ClearCache(QThread):
    isFinished = pyqtSignal(bool)
    def __init__(self):
        super(ClearCache, self).__init__()

    def run(self):
        if os.path.exists('./Log'):
            subprocess.call("del /s /q Log", shell=True)
        if os.path.exists('FastCopy2.ini'):
            subprocess.call("del /q FastCopy2.ini", shell=True)
        self.isFinished.emit(True)


class HomeInterface(SmoothScrollArea):
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
        self.storageGroup = SettingCardGroup('存储', self.scrollWidget)
        self.advanceGroup = SettingCardGroup('高级', self.scrollWidget)
        self.optionSourceCard = ComboBoxSettingCard(
            cfg.IsSourceCloud,
            FIF.FOLDER,
            '源文件夹',
            '选择源文件夹来源',
            texts=['云上春晖', '自定义'],
            parent=self.sourceGroup)
        self.autoRunCard = SwitchSettingCard(
            FIF.POWER_BUTTON,
            "开机时启动",
            "",
            configItem=cfg.AutoRun,
            parent=self.actGroup)
        self.notifyCard = SwitchSettingCard(
            FIF.RINGER,
            "完成后通知",
            "",
            configItem=cfg.Notify,
            parent=self.actGroup)
        self.cloudCard = PushSettingCard(
            '选择文件夹',
            FIF.CLOUD,
            "云上春晖",
            cfg.get(cfg.sourceFolder),
            self.sourceGroup)
        self.customFolderCard = CustomFolderListSettingCard(
            "自定义",
            "展开选项卡以设置",
            directory=cfg.sourceFolder.value,
            parent=self.sourceGroup)
        self.scanCycleCard = RangeSettingCard(
            cfg.ScanCycle,
            FIF.STOP_WATCH,
            '扫描周期',
            parent=self.performanceGroup)
        self.concurrentProcessCard = SpinBoxSettingCard(
            cfg.ConcurrentProcess,
            FIF.ALIGNMENT,
            '并行进程数',
            parent=self.performanceGroup)
        self.bufSizeCard = OptionsSettingCard(
            cfg.BufSize,
            FIF.PIE_SINGLE,
            '缓冲区大小',
            texts=['32 MB', '64 MB', '128 MB', '256 MB', '512 MB', '1 GB'],
            parent=self.performanceGroup)
        self.clearCard = PushSettingCard(
            '清除',
            FIF.BROOM,
            '清除缓存',
            self.getSize(),
            self.storageGroup)
        self.recoverCard = PushSettingCard(
            '恢复',
            FIF.CLEAR_SELECTION,
            '恢复默认设置',
            '重置所有参数为初始值',
            self.advanceGroup)
        self.devCard = PushSettingCard(
            '打开',
            FIF.DEVELOPER_TOOLS,
            '开发者选项',
            '打开配置文件',
            self.advanceGroup)
        self.helpCard = PrimaryPushSettingCard(
            '转到帮助',
            FIF.HELP,
            '帮助',
            '提示与常见问题',
            self.advanceGroup)
        self.__initWidget()

    def __initWidget(self):
        self.resize(1000, 800)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setViewportMargins(0, 80, 0, 20)
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
        self.settingLabel.move(50, 20)

        self.sourceGroup.addSettingCard(self.optionSourceCard)
        self.sourceGroup.addSettingCard(self.cloudCard)
        self.sourceGroup.addSettingCard(self.customFolderCard)
        self.actGroup.addSettingCard(self.autoRunCard)
        self.actGroup.addSettingCard(self.notifyCard)
        self.performanceGroup.addSettingCard(self.scanCycleCard)
        self.performanceGroup.addSettingCard(self.concurrentProcessCard)
        self.performanceGroup.addSettingCard(self.bufSizeCard)
        self.storageGroup.addSettingCard(self.clearCard)
        self.advanceGroup.addSettingCard(self.recoverCard)
        self.advanceGroup.addSettingCard(self.devCard)
        self.advanceGroup.addSettingCard(self.helpCard)
        self.expandLayout.setSpacing(28)
        self.expandLayout.setContentsMargins(60, 10, 60, 0)
        self.expandLayout.addWidget(self.sourceGroup)
        self.expandLayout.addWidget(self.actGroup)
        self.expandLayout.addWidget(self.performanceGroup)
        self.expandLayout.addWidget(self.storageGroup)
        self.expandLayout.addWidget(self.advanceGroup)

    def getSize(self):
        try:
            size = os.path.getsize('FastCopy2.ini')
        except:
            size = 0
        if os.path.exists('./Log'):
            for root, dirs, files in os.walk('./Log'):
                try:
                    size += sum([os.path.getsize(os.path.join(root, name)) for name in files])
                except:
                    pass
        kbSize = float(size / 1024)
        if kbSize >= 1024*1024:
            return str(round(kbSize/1024/1024, 1)) + ' GB'
        elif kbSize >= 1024:
            return str(round(kbSize/1024, 1)) + ' MB'
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

    def __onCloudCardClicked(self):
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹", "./")
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

    def clearFinished(self):
        self.clearCacheThread.exit(0)
        self.clearCard.contentLabel.setText(self.getSize())
        self.clearCard.button.setText('已清除')
        QTimer.singleShot(2000, lambda: self.clearCard.button.setText('清除'))
        self.clearCard.button.setDisabled(False)

    def clearCache(self):
        w = MessageBox(
            '清除缓存',
            '缓存包含日志文件，是否清除？',
            self.window())
        w.yesButton.setText('确定')
        w.cancelButton.setText('取消')
        if w.exec():
            self.clearCard.button.setText('清除中')
            self.clearCard.button.setDisabled(True)
            self.clearCacheThread = ClearCache()
            self.clearCacheThread.start()
            self.clearCacheThread.isFinished.connect(self.clearFinished)

    def recoverConfig(self):
        w = MessageBox(
            '恢复默认设置',
            '是否要重置所有设置？',
            self.window())
        w.yesButton.setText('确定')
        w.cancelButton.setText('取消')
        if w.exec():
            self.autoRunCard.setChecked(True)
            self.notifyCard.setChecked(True)
            self.scanCycleCard.setValue(10)
            self.concurrentProcessCard.setValue(3)
            self.bufSizeCard.setValue(BufSize._256)

    def openConfig(self):
        w = MessageBox(
            '打开配置文件',
            '即将打开配置文件，请谨慎操作。',
            self.window())
        w.yesButton.setText('确定')
        w.cancelButton.setText('取消')
        if w.exec():
            os.startfile(os.path.join(os.path.expanduser('~'), '.Presto', 'config', 'config.json'))

    def onHelpAction(self):
        if os.path.exists(os.path.abspath("./Doc/PrestoHelp.html")):
            os.startfile(os.path.abspath("./Doc/PrestoHelp.html"))
        else:
            WebOpen("https://sudo0015.github.io/post/Presto%20-bang-zhu.html")

    def __connectSignalToSlot(self):
        self.optionSourceCard.comboBox.currentTextChanged.connect(self.onOptionSourceCard)
        self.cloudCard.clicked.connect(self.__onCloudCardClicked)
        self.clearCard.clicked.connect(self.clearCache)
        self.recoverCard.clicked.connect(self.recoverConfig)
        self.devCard.clicked.connect(self.openConfig)
        self.helpCard.clicked.connect(self.onHelpAction)


class DetailMessageBox(MessageBoxBase):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.titleLabel = SubtitleLabel('关于 Presto', self)
        self.textBox = TextBrowser(self)
        self.textBox.setText(
            f'Presto v{VERSION}\nCopyright © {YEAR} BUG STUDIO\n\n' +
            'MIT License\nPermission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:\n' +
            'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.\n' +
            'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.\n'
        )

        self.githubBtn = HyperlinkButton(self)
        self.websiteBtn = HyperlinkButton(self)
        self.onlineDocBtn = HyperlinkButton(self)
        self.githubBtn.setText('源代码')
        self.websiteBtn.setText('网站主页')
        self.onlineDocBtn.setText('在线文档')
        self.githubBtn.setIcon(FIF.GITHUB)
        self.websiteBtn.setIcon(FIF.GLOBE)
        self.onlineDocBtn.setIcon(FIF.DOCUMENT)
        self.githubBtn.clicked.connect(lambda: WebOpen("https://github.com/sudo0015/Presto"))
        self.websiteBtn.clicked.connect(lambda: WebOpen("https://sudo0015.github.io/"))
        self.onlineDocBtn.clicked.connect(lambda: WebOpen("https://sudo0015.github.io/post/Presto%20-bang-zhu.html"))

        self.btnLayout = QHBoxLayout(self)
        self.btnLayout.addWidget(self.githubBtn)
        self.btnLayout.addWidget(self.websiteBtn)
        self.btnLayout.addWidget(self.onlineDocBtn)

        self.viewLayout.addWidget(self.titleLabel)
        self.viewLayout.addWidget(self.textBox)
        self.viewLayout.addLayout(self.btnLayout)

        self.yesButton.setText('确定')
        self.hideCancelButton()

        self.widget.setMinimumWidth(350)


class AboutInterface(SmoothScrollArea):
    def __init__(self, parent=None):
        super().__init__(parent=parent)
        self.scrollWidget = QWidget()
        self.stateTooltip = None
        self.expandLayout = ExpandLayout(self.scrollWidget)
        self.enableTransparentBackground()
        if darkdetect.isDark():
            self.scrollWidget.setStyleSheet("background-color: rgba(39, 39, 39, 0);")
        else:
            self.scrollWidget.setStyleSheet("background-color: rgba(249, 249, 249, 0);")

        self.aboutGroup = SettingCardGroup('', self.scrollWidget)
        self.aboutESCard = PushSettingCard(
            '详细信息',
            FIF.INFO,
            '关于 Presto',
            f'版本 {VERSION}',
            self.aboutGroup)
        self.aboutBSCard = PushSettingCard(
            '了解更多',
            FIF.PEOPLE,
            '关于作者',
            'BUG STUDIO',
            self.aboutGroup)
        self.helpCard = PrimaryPushSettingCard(
            '转到帮助',
            FIF.HELP,
            '帮助',
            '提示与常见问题',
            self.aboutGroup)
        self.feedbackCard = PrimaryPushSettingCard(
            '提供反馈',
            FIF.FEEDBACK,
            '反馈',
            '报告问题或提出建议',
            self.aboutGroup)
        self.__initWidget()

    def __initWidget(self):
        self.resize(1000, 800)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setViewportMargins(0, 0, 0, 0)
        self.setWidget(self.scrollWidget)
        self.setWidgetResizable(True)

        self.imgLabel = ImageLabel(self)
        if isDarkTheme():
            self.imgLabel.setImage(':/BannerDark.png')
        else:
            self.imgLabel.setImage(':/BannerLight.png')
        self.imgLabel.setFixedSize(401, 150)

        self.__initLayout()
        self.__connectSignalToSlot()

    def __initLayout(self):

        self.aboutGroup.addSettingCard(self.aboutESCard)
        self.aboutGroup.addSettingCard(self.aboutBSCard)
        self.aboutGroup.addSettingCard(self.helpCard)
        self.aboutGroup.addSettingCard(self.feedbackCard)

        self.expandLayout.setContentsMargins(60, 10, 60, 0)
        self.expandLayout.addWidget(self.imgLabel)
        self.expandLayout.addWidget(self.aboutGroup)

    def onAboutESCardClicked(self):
        w = DetailMessageBox(self.window())
        if w.exec():
            pass

    def onHelpAction(self):
        if os.path.exists(os.path.abspath("./Doc/PrestoHelp.html")):
            os.startfile(os.path.abspath("./Doc/PrestoHelp.html"))
        else:
            WebOpen("https://sudo0015.github.io/post/Presto%20-bang-zhu.html")

    def onAboutBSAction(self):
        if os.path.exists(os.path.abspath("./Doc/AboutBugStudio.html")):
            os.startfile(os.path.abspath("./Doc/AboutBugStudio.html"))
        else:
            WebOpen("https://sudo0015.github.io/post/guan-yu-%20BUG%20STUDIO.html")

    def __connectSignalToSlot(self):
        self.aboutESCard.clicked.connect(self.onAboutESCardClicked)
        self.aboutBSCard.clicked.connect(self.onAboutBSAction)
        self.helpCard.clicked.connect(self.onHelpAction)
        self.feedbackCard.clicked.connect(lambda: WebOpen("https://github.com/sudo0015/Presto/issues"))


class TitleBarBase(QWidget):
    """ Title bar base class """

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
        """ Toggles the maximization state of the window """
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
        """ Toggles the maximization state of the window and change icon """
        if self.window().isMaximized():
            self.window().showNormal()
        else:
            self.window().showMaximized()

        if sys.platform == "win32":
            from qframelesswindow.utils.win32_utils import releaseMouseLeftButton
            releaseMouseLeftButton(self.window().winId())

    def _isDragRegion(self, pos):
        """ Check whether the position belongs to the area where dragging is allowed """
        width = 0
        for button in self.findChildren(TitleBarButton):
            if button.isVisible():
                width += button.width()

        return 0 < pos.x() < self.width() - width

    def _hasButtonPressed(self):
        """ whether any button is pressed """
        return any(btn.isPressed() for btn in self.findChildren(TitleBarButton))

    def canDrag(self, pos):
        """ whether the position is draggable """
        return self._isDragRegion(pos) and not self._hasButtonPressed()

    def setDoubleClickEnabled(self, isEnabled):
        """ whether to switch window maximization status when double clicked

        Parameters
        ----------
        isEnabled: bool
            whether to enable double click
        """
        self._isDoubleClickEnabled = isEnabled


class TitleBar(TitleBarBase):
    """ Title bar with minimize, maximum and close button """

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
    """ Fluent title bar"""

    def __init__(self, parent):
        super().__init__(parent)
        self.setFixedHeight(48)
        self.hBoxLayout.removeWidget(self.minBtn)
        self.hBoxLayout.removeWidget(self.maxBtn)
        self.hBoxLayout.removeWidget(self.closeBtn)

        # add window icon
        self.iconLabel = QLabel(self)
        self.iconLabel.setFixedSize(18, 18)
        self.hBoxLayout.insertWidget(0, self.iconLabel, 0, Qt.AlignLeft | Qt.AlignVCenter)
        self.window().windowIconChanged.connect(self.setIcon)

        # add title label
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
    """ Fluent window in Microsoft Store style """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setTitleBar(MSFluentTitleBar(self))

        self.navigationInterface = NavigationBar(self)

        self.hBoxLayout.setContentsMargins(0, 48, 0, 0)
        self.hBoxLayout.addWidget(self.navigationInterface)
        self.hBoxLayout.addWidget(self.stackedWidget, 1)

        self.titleBar.raise_()
        self.titleBar.setAttribute(Qt.WA_StyledBackground)

    def addSubInterface(self, interface: QWidget, icon: Union[FIF, QIcon, str], text: str,
                        selectedIcon=None, position=NavigationItemPosition.TOP, isTransparent=False) -> NavigationBarPushButton:
        """ add sub interface, the object name of `interface` should be set already
        before calling this method

        Parameters
        ----------
        interface: QWidget
            the subinterface to be added

        icon: FluentIconBase | QIcon | str
            the icon of navigation item

        text: str
            the text of navigation item

        selectedIcon: str | QIcon | FluentIconBase
            the icon of navigation item in selected state

        position: NavigationItemPosition
            the position of navigation item
        """
        if not interface.objectName():
            raise ValueError("The object name of `interface` can't be empty string.")

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
        setThemeColor(QColor(113, 89, 249))
        self.resize(800, 600)
        self.setWindowTitle('Presto 设置')
        self.setWindowIcon(QIcon(':/icon.png'))
        self.titleBar.raise_()
        desktop = QApplication.screens()[0].size()
        self.move(desktop.width() // 2 - self.width() // 2, desktop.height() // 2 - self.height() // 2)

        self.splashScreen = SplashScreen(self.windowIcon(), self)
        self.splashScreen.raise_()

        self.homeInterface = HomeInterface(self)
        self.aboutInterface = AboutInterface(self)
        self.homeInterface.setObjectName('homeInterface')
        self.aboutInterface.setObjectName('aboutInterface')
        self.addSubInterface(self.homeInterface, FIF.HOME, '设置', FIF.HOME_FILL)
        self.navigationInterface.addItem(
            routeKey='Log',
            icon=FIF.DOCUMENT,
            text='日志',
            onClick=self.onLogBtn,
            selectable=False,
            position=NavigationItemPosition.TOP, )
        self.navigationInterface.addItem(
            routeKey='Help',
            icon=FIF.HELP,
            text='帮助',
            onClick=self.onHelpBtn,
            selectable=False,
            position=NavigationItemPosition.BOTTOM, )
        self.addSubInterface(self.aboutInterface, FIF.INFO, '关于', FIF.INFO, NavigationItemPosition.BOTTOM)
        self.navigationInterface.setCurrentItem(self.homeInterface.objectName())

        self.splashScreen.finish()

    def onHelpBtn(self):
        if os.path.exists(os.path.abspath("./Doc/PrestoHelp.html")):
            os.startfile(os.path.abspath("./Doc/PrestoHelp.html"))
        else:
            WebOpen("https://sudo0015.github.io/post/Presto%20-bang-zhu.html")

    def onLogBtn(self):
        try:
            os.startfile('Log')
        except:
            pass


if __name__ == '__main__':
    with Mutex():
        QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
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
