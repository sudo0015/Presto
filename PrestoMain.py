# -*- coding: utf-8 -*-

import os
import sys
import subprocess
import darkdetect
import PrestoResource
from psutil import Process
from random import randint
from PrestoConfig import cfg
from winotify import Notification, audio
from win32api import GetVolumeInformation
from PyQt5.QtGui import QIcon, QColor, QPainter
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QEvent, QTimer, QPropertyAnimation
from PyQt5.QtWidgets import QApplication, QHBoxLayout, QVBoxLayout, QLabel, QWidget, QGridLayout, QFrame, \
    QTableWidgetItem, QGraphicsOpacityEffect
from PyQt5.QtWinExtras import QWinTaskbarButton
from qfluentwidgets import setTheme, Theme, BodyLabel, isDarkTheme, PushButton, SubtitleLabel, ProgressBar, \
    InfoBar, InfoBarIcon, InfoBarPosition, IndeterminateProgressBar, setThemeColor, PrimaryPushButton, \
    TableWidget, IndeterminateProgressRing, TextWrap
from qfluentwidgets.components.widgets.info_bar import InfoIconWidget, InfoBarManager
from qframelesswindow.titlebar import MinimizeButton, CloseButton, MaximizeButton
from qframelesswindow import TitleBarButton
from qframelesswindow.utils import startSystemMove
from qframelesswindow import FramelessDialog
from qfluentwidgets.common.style_sheet import FluentStyleSheet
from qfluentwidgets import FluentIcon as FIF


if sys.platform == 'win32' and sys.getwindowsversion().build >= 22000:
    from FramelessWindow import AcrylicWindow as Window
else:
    from FramelessWindow import WindowsFramelessWindow as Window


class TitleBarBase(QWidget):
    """ Title bar base class """

    def __init__(self, parent):
        super().__init__(parent)
        self.minBtn = MinimizeButton(parent=self)
        self.closeBtn = CloseButton(parent=self)
        self.maxBtn = MaximizeButton(parent=self)

        self._isDoubleClickEnabled = True

        self.resize(200, 45)
        self.setFixedHeight(45)

        self.minBtn.clicked.connect(self.window().showMinimized)
        self.maxBtn.clicked.connect(self.__toggleMaxState)

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
    def __init__(self, parent):
        super().__init__(parent)
        self.hBoxLayout = QHBoxLayout(self)

        # add buttons to layout
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
        self.setFixedHeight(45)
        self.hBoxLayout.removeWidget(self.minBtn)
        self.hBoxLayout.removeWidget(self.maxBtn)
        self.hBoxLayout.removeWidget(self.closeBtn)
        self.maxBtn.setVisible(False)
        self.titleLabel = QLabel(self)
        self.titleLabel.setText(self.getDriveName() + ' (' + drive + ')' + ' - Presto')
        self.titleLabel.setObjectName('titleLabel')
        self.vBoxLayout = QVBoxLayout()
        self.buttonLayout = QHBoxLayout()
        self.buttonLayout.setSpacing(0)
        self.buttonLayout.setContentsMargins(0, 0, 0, 0)
        self.buttonLayout.setAlignment(Qt.AlignTop)
        self.buttonLayout.addWidget(self.minBtn)
        self.buttonLayout.addWidget(self.closeBtn)
        self.titleLayout = QHBoxLayout()
        self.titleLayout.setContentsMargins(0, 8, 0, 0)
        self.titleLayout.addWidget(self.titleLabel)
        self.titleLayout.setAlignment(Qt.AlignTop)
        self.vBoxLayout.addLayout(self.buttonLayout)
        self.vBoxLayout.addStretch(1)
        self.hBoxLayout.addLayout(self.titleLayout)
        self.hBoxLayout.addStretch(50)
        self.hBoxLayout.addLayout(self.vBoxLayout)
        FluentStyleSheet.FLUENT_WINDOW.apply(self)

    def getDriveName(self):
        try:
            if GetVolumeInformation(drive)[0] != '':
                return GetVolumeInformation(drive)[0]
            else:
                return "U盘"
        except:
            sys.exit()


class MicaWindow(Window):

    def __init__(self):
        super().__init__()
        self.setTitleBar(FluentTitleBar(self))
        if sys.platform == 'win32' and sys.getwindowsversion().build >= 22000:
            self.windowEffect.setMicaEffect(self.winId(), isDarkTheme())


class EjectUsbInfoBar(QFrame):

    closedSignal = pyqtSignal()
    _desktopView = None

    def __init__(self, title: str, position=InfoBarPosition.BOTTOM, parent=None):
        super().__init__(parent=parent)
        self.title = title
        self.position = position

        self.titleLabel = QLabel(self)
        self.progressRing = IndeterminateProgressRing()
        self.progressRing.setFixedSize(18, 36)
        self.progressRing.setStrokeWidth(3)

        self.hBoxLayout = QHBoxLayout(self)

        self.opacityEffect = QGraphicsOpacityEffect(self)
        self.opacityAni = QPropertyAnimation(self.opacityEffect, b'opacity', self)

        self.lightBackgroundColor = QColor(244, 244, 244)
        self.darkBackgroundColor = QColor(39, 39, 39)

        self.__initWidget()

    def __initWidget(self):
        self.opacityEffect.setOpacity(1)
        self.setGraphicsEffect(self.opacityEffect)

        self.setQss()
        self.setLayout(False)

    def setLayout(self, isStatusChanged):
        if isStatusChanged:
            self.hBoxLayout.setContentsMargins(6, 6, 16, 6)
            self.hBoxLayout.setSizeConstraint(QVBoxLayout.SetMinimumSize)
            self.hBoxLayout.setSpacing(6)
            self.hBoxLayout.addWidget(self.iconWidget)
        else:
            self.hBoxLayout.setContentsMargins(16, 6, 16, 6)
            self.hBoxLayout.setSizeConstraint(QVBoxLayout.SetMinimumSize)
            self.hBoxLayout.setSpacing(15)
            self.hBoxLayout.addWidget(self.progressRing)
        self.hBoxLayout.addWidget(self.titleLabel)
        self.adjustText()

    def setQss(self):
        self.titleLabel.setObjectName('titleLabel')
        FluentStyleSheet.INFO_BAR.apply(self)

    def adjustText(self):
        w = 900 if not self.parent() else (self.parent().width() - 50)
        chars = max(min(w / 10, 120), 30)
        self.titleLabel.setText(TextWrap.wrap(self.title, chars, False)[0])
        self.adjustSize()

    def setCustomBackgroundColor(self, light, dark):
        self.lightBackgroundColor = QColor(light)
        self.darkBackgroundColor = QColor(dark)
        self.update()

    def statusChanged(self, isSuccess):
        self.hBoxLayout.removeWidget(self.progressRing)
        self.progressRing.deleteLater()
        icon = InfoBarIcon.SUCCESS if isSuccess else InfoBarIcon.ERROR
        self.iconWidget = InfoIconWidget(icon)
        if isSuccess:
            self.title = "U盘已成功退出"
            self.setCustomBackgroundColor("#dff6dd", "#393d1b")
        else:
            self.title = "U盘退出失败"
            self.setCustomBackgroundColor("#fde7e9", "#442726")
        self.setLayout(True)

    def showEvent(self, e):
        super().showEvent(e)

        if self.position != InfoBarPosition.NONE:
            manager = InfoBarManager.make(self.position)
            manager.add(self)

        if self.parent():
            self.parent().installEventFilter(self)

    def eventFilter(self, obj, e: QEvent):
        if obj is self.parent():
            if e.type() in [QEvent.Resize, QEvent.WindowStateChange]:
                self.adjustText()
        return super().eventFilter(obj, e)

    def showEvent(self, e):
        self.adjustText()
        super().showEvent(e)

        if self.position != InfoBarPosition.NONE:
            manager = InfoBarManager.make(self.position)
            manager.add(self)

        if self.parent():
            self.parent().installEventFilter(self)

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


class UiDetailDialog:
    def __init__(self, *args, **kwargs):
        pass

    def _setUpUi(self, title, mode, subject, parent):
        self.titleLabel = QLabel(title, parent)
        self.tableView = TableWidget(parent)
        self.tableView.setBorderVisible(True)
        self.tableView.setBorderRadius(8)
        self.tableView.setWordWrap(False)
        self.tableView.setRowCount(6)
        self.tableView.setColumnCount(2)
        self.tableView.verticalHeader().hide()
        self.tableView.setEnabled(False)
        self.tableView.setHorizontalHeaderLabels(['选项', '参数'])
        self.tableView.setItem(0, 0, QTableWidgetItem('目标驱动器'))
        self.tableView.setItem(0, 1, QTableWidgetItem(drive + '\\'))
        self.tableView.setItem(1, 0, QTableWidgetItem('模式'))
        self.tableView.setItem(1, 1, QTableWidgetItem(mode))
        self.tableView.setItem(2, 0, QTableWidgetItem('学科'))
        self.tableView.setItem(2, 1, QTableWidgetItem(subject[:len(subject)-2]))
        self.tableView.setItem(3, 0, QTableWidgetItem('命令行选项'))
        self.tableView.setItem(3, 1, QTableWidgetItem(commandOption))
        self.tableView.setItem(4, 0, QTableWidgetItem('缓冲区大小'))
        self.tableView.setItem(4, 1, QTableWidgetItem(buf + ' MB'))
        self.tableView.setItem(5, 0, QTableWidgetItem('并行进程数'))
        self.tableView.setItem(5, 1, QTableWidgetItem(str(concurrentProcess)))

        self.tableView.resizeColumnsToContents()
        self.titleLabel.setVisible(False)

        self.buttonGroup = QFrame(parent)
        self.yesButton = PrimaryPushButton('确定', self.buttonGroup)

        self.vBoxLayout = QVBoxLayout(parent)
        self.textLayout = QVBoxLayout()
        self.buttonLayout = QHBoxLayout(self.buttonGroup)

        self.__initWidget()

    def __initWidget(self):
        self.__setQss()
        self.__initLayout()

        self.yesButton.setAttribute(Qt.WA_LayoutUsesWidgetRect)
        self.yesButton.setFixedWidth(150)
        self.yesButton.setFocus()
        self.buttonGroup.setFixedHeight(81)

        self.yesButton.clicked.connect(self.close)

    def __initLayout(self):
        self.vBoxLayout.setSpacing(0)
        self.vBoxLayout.setContentsMargins(0, 0, 0, 0)
        self.vBoxLayout.addLayout(self.textLayout, 1)
        self.vBoxLayout.addWidget(self.buttonGroup, 0, Qt.AlignBottom)
        self.vBoxLayout.setSizeConstraint(QVBoxLayout.SetMinimumSize)

        self.textLayout.setSpacing(15)
        self.textLayout.setContentsMargins(24, 24, 24, 24)
        self.textLayout.addWidget(self.titleLabel)
        self.textLayout.addWidget(self.tableView)

        self.buttonLayout.setContentsMargins(24, 24, 24, 24)
        self.buttonLayout.addWidget(self.yesButton, 1, Qt.AlignRight)

    def __setQss(self):
        self.titleLabel.setObjectName("titleLabel")
        self.tableView.setObjectName("tableWidget")
        self.buttonGroup.setObjectName('buttonGroup')
        FluentStyleSheet.DIALOG.apply(self)
        self.yesButton.adjustSize()


class DetailDialog(FramelessDialog, UiDetailDialog):
    def __init__(self, title: str, mode: str, subject: str, parent=None):
        super().__init__(parent=parent)
        self._setUpUi(title, mode, subject, self)
        self.windowTitleLabel = QLabel(title, self)
        self.setResizeEnabled(False)
        self.resize(530, 450)
        self.titleBar.hide()

        self.vBoxLayout.insertWidget(0, self.windowTitleLabel, 0, Qt.AlignTop)
        self.windowTitleLabel.setObjectName('windowTitleLabel')
        FluentStyleSheet.DIALOG.apply(self)
        self.windowTitleLabel.setVisible(False)
        self.setFixedSize(self.size())

    def setTitleBarVisible(self, isVisible: bool):
        self.windowTitleLabel.setVisible(isVisible)


class UiErrorDialog:

    yesSignal = pyqtSignal()
    cancelSignal = pyqtSignal()

    def _setUpUi(self, title, content, parent):
        self.content = content
        self.titleLabel = QLabel(title, parent)
        self.contentLabel = BodyLabel(content, parent)

        self.buttonGroup = QFrame(parent)
        self.yesButton = PrimaryPushButton("确定", self.buttonGroup)
        self.yesButton.setFixedWidth(140)

        self.vBoxLayout = QVBoxLayout(parent)
        self.textLayout = QVBoxLayout()
        self.buttonLayout = QHBoxLayout(self.buttonGroup)

        self.__initWidget()

    def __initWidget(self):
        self.__setQss()
        self.__initLayout()

        self.yesButton.setAttribute(Qt.WA_LayoutUsesWidgetRect)

        self.yesButton.setFocus()
        self.buttonGroup.setFixedHeight(81)

        self.contentLabel.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._adjustText()

        self.yesButton.clicked.connect(self.__onYesButtonClicked)

    def _adjustText(self):
        if self.isWindow():
            if self.parent():
                w = max(self.titleLabel.width(), self.parent().width())
                chars = max(min(w / 9, 140), 30)
            else:
                chars = 100
        else:
            w = max(self.titleLabel.width(), self.window().width())
            chars = max(min(w / 9, 100), 30)

        self.contentLabel.setText(TextWrap.wrap(self.content, chars, False)[0])

    def __initLayout(self):
        self.vBoxLayout.setSpacing(0)
        self.vBoxLayout.setContentsMargins(0, 0, 0, 0)
        self.vBoxLayout.addLayout(self.textLayout, 1)
        self.vBoxLayout.addWidget(self.buttonGroup, 0, Qt.AlignBottom)
        self.vBoxLayout.setSizeConstraint(QVBoxLayout.SetMinimumSize)

        self.textLayout.setSpacing(12)
        self.textLayout.setContentsMargins(24, 24, 24, 24)
        self.textLayout.addWidget(self.titleLabel, 0, Qt.AlignTop)
        self.textLayout.addWidget(self.contentLabel, 0, Qt.AlignTop)

        self.buttonLayout.setSpacing(12)
        self.buttonLayout.setContentsMargins(24, 24, 24, 24)
        self.buttonLayout.addWidget(self.yesButton, 1, Qt.AlignRight | Qt.AlignVCenter)

    def __onYesButtonClicked(self):
        self.accept()
        self.yesSignal.emit()

    def __setQss(self):
        self.titleLabel.setObjectName("titleLabel")
        self.contentLabel.setObjectName("contentLabel")
        self.buttonGroup.setObjectName('buttonGroup')

        FluentStyleSheet.DIALOG.apply(self)
        FluentStyleSheet.DIALOG.apply(self.contentLabel)

        self.yesButton.adjustSize()


class ErrorDialog(FramelessDialog, UiErrorDialog):

    yesSignal = pyqtSignal()
    cancelSignal = pyqtSignal()

    def __init__(self, title: str, content: str, parent=None):
        super().__init__(parent=parent)
        self._setUpUi(title, content, self)

        self.windowTitleLabel = QLabel("Presto", self)

        self.setResizeEnabled(False)
        self.resize(240, 192)
        self.titleBar.hide()

        self.vBoxLayout.insertWidget(0, self.windowTitleLabel, 0, Qt.AlignTop)
        self.windowTitleLabel.setObjectName('windowTitleLabel')
        FluentStyleSheet.DIALOG.apply(self)
        self.setFixedSize(self.size())


class DeleteThread(QThread):
    deleteFinished = pyqtSignal(bool)

    def __init__(self, parent=None):
        super().__init__(parent=parent)
        self.isRunning = True
        self.process = None

    def stop(self):
        self.isRunning = False

    def run(self):
        if self.isRunning:
            args = "fcp.exe /cmd=delete " + f"/bufsize={buf} /log=FALSE " + f'/force_start={concurrentProcess} "{destFolder}"'
            if os.path.exists('fcp.exe'):
                self.process = subprocess.Popen(args, shell=True)
                self.process.wait()
            else:
                w = ErrorDialog("错误", "核心文件缺失，请尝试重新安装。Presto 将退出。")
                w.yesButton.setText("确定")
                if w.exec():
                    sys.exit()

            self.deleteFinished.emit(True)
            return
        else:
            self.process.terminate()
            return


class SyncThread(QThread):
    valueChange = pyqtSignal(int)

    def __init__(self, parent=None):
        super().__init__(parent=parent)
        self.is_paused = bool(0)
        self.progress_value = int(0)
        self.isRunning = True
        self.process = None

    def stop(self):
        self.isRunning = False

    def run(self):
        global currentTask
        while self.isRunning:
            if taskList or currentTask:
                if currentTask == 0:
                    currentTask = taskList.pop(0)

                if currentTask == 1:
                    currentFolder = cfg.yuwenFolder.value
                elif currentTask == 2:
                    currentFolder = cfg.shuxueFolder.value
                elif currentTask == 3:
                    currentFolder = cfg.yingyuFolder.value
                elif currentTask == 4:
                    currentFolder = cfg.wuliFolder.value
                elif currentTask == 5:
                    currentFolder = cfg.huaxueFolder.value
                elif currentTask == 6:
                    currentFolder = cfg.shengwuFolder.value
                elif currentTask == 7:
                    currentFolder = cfg.zhengzhiFolder.value
                elif currentTask == 8:
                    currentFolder = cfg.lishiFolder.value
                elif currentTask == 9:
                    currentFolder = cfg.diliFolder.value
                elif currentTask == 10:
                    currentFolder = cfg.jishuFolder.value
                elif currentTask == 11:
                    currentFolder = cfg.ziliaoFolder.value

                args = f'fcp.exe /cmd=sync /log=FALSE /no_confirm_stop /error_stop=FALSE /bufsize={buf} /force_start={concurrentProcess} {commandOption} "' + currentFolder.replace('/', '\\') + f'" /to="{destFolder}"'
                if os.path.exists('fcp.exe'):
                    self.process = subprocess.Popen(args, shell=True)
                    self.process.wait()
                else:
                    w = ErrorDialog("错误", "核心文件缺失，请尝试重新安装。Presto 将退出。")
                    w.yesButton.setText("确定")
                    if w.exec():
                        sys.exit()

                self.progress_value = int((taskNum - len(taskList)) / taskNum * 100)
                self.valueChange.emit(self.progress_value)

                if taskList:
                    currentTask = taskList.pop(0)
                else:
                    currentTask = 0
            else:
                self.progress_value = -1
                self.valueChange.emit(self.progress_value)
                return
        else:
            self.process.terminate()
            return


class EjectThread(QThread):
    ejectFinished = pyqtSignal(bool)

    def __init__(self, parent=None):
        super().__init__(parent=parent)
        self.exitCode = -1
        self.isRunning = True

    def stop(self):
        self.isRunning = False

    def run(self):
        if self.isRunning:
            for i in range(2):
                if os.path.exists('RemoveDrive.exe'):
                    self.exitCode = subprocess.call(["RemoveDrive.exe", drive, '-f'], shell=True)
                else:
                    w = ErrorDialog("错误", "核心文件缺失，请尝试重新安装。Presto 将退出。")
                    w.yesButton.setText("确定")
                    if w.exec():
                        sys.exit()

                if self.exitCode == 0:
                    self.ejectFinished.emit(True)
                    break
            else:
                self.ejectFinished.emit(False)
            return
        else:
            self.terminate()
            return


class MainWindow(MicaWindow):

    def __init__(self):
        super().__init__()
        if isDarkTheme():
            setThemeColor(QColor(113, 89, 249))
        else:
            setThemeColor(QColor(90, 51, 174))
        self.driveName = self.getDriveName() + ' (' + drive + ')'

        self.resize(550, 150)
        self.setWindowTitle(self.driveName + ' - Presto')
        self.setWindowIcon(QIcon(':/icon.png'))
        self.setFixedHeight(150)
        self.move(QApplication.screens()[0].size().width() // 2 - self.width() // 2 - randint(0, 50),
                  QApplication.screens()[0].size().height() // 2 - self.height() // 2 - randint(0, 50))

        self.displayText = {1:"同步 (默认)", 2:"同步 (低占用)", 3:"复制 (最近文件)", 4:"复制 (从时间戳)"}[mode]
        if mode == 3 or mode == 4:
            self.displayText += ' - '
            self.displayText += "删除原有文件" if isDelete else "保留原有文件"
        self.subject = ""
        for i in taskList:
            self.subject += {1: '语文', 2: '数学', 3: '英语', 4: '物理', 5: '化学', 6: '生物', 7: '政治', 8: '历史', 9: '地理', 10: '技术', 11: '资料'}[i] + '; '

        self.titleBar.closeBtn.clicked.connect(self.onCancelBtn)

        self.isPrepare = True
        self.isProgress = True
        self.isPaused = False
        self.syncThreadRunning = False
        self.deleteThreadRunning = False
        self.ejectThreadRunning = False

        self.mainLayout = QVBoxLayout(self)
        self.topLayout = QGridLayout(self)
        self.bottomLayout = QHBoxLayout(self)
        self.mainLayout.setContentsMargins(0, 0, 0, 0)
        self.topLayout.setContentsMargins(20, 0, 20, 0)
        self.bottomLayout.setContentsMargins(20, 5, 20, 20)

        self.statusLabel = SubtitleLabel(self)
        self.detailLabel = BodyLabel(self)
        self.detailLabel.setText(self.displayText)
        self.detailLabel.setTextColor(QColor(114, 114, 114))
        self.showDetailBtn = PrimaryPushButton(FIF.CHECKBOX, '选项', self)
        self.cancelBtn = PushButton(FIF.CLOSE, '取消', self)
        self.pauseBtn = PushButton(FIF.PAUSE, '暂停', self)
        self.showDetailBtn.clicked.connect(self.onShowDetailBtn)
        self.cancelBtn.clicked.connect(self.onCancelBtn)
        self.pauseBtn.clicked.connect(self.onPauseBtn)
        self.topLayout.addWidget(self.statusLabel, 0, 0)
        self.topLayout.addWidget(self.detailLabel, 1, 0)
        self.topLayout.setColumnStretch(0, 1)
        self.topLayout.addWidget(self.showDetailBtn, 0, 1, 2, 1)
        self.topLayout.addWidget(self.pauseBtn, 0, 2, 2, 1)
        self.topLayout.addWidget(self.cancelBtn, 0, 3, 2, 1)

        self.inProgressBar = IndeterminateProgressBar(self)
        self.progressBar = ProgressBar(self)
        self.spaceLabel = QLabel(self)
        self.spaceLabel.setFixedWidth(5)
        self.progressLabel = BodyLabel(self)
        self.progressBar.setVisible(False)
        self.spaceLabel.setVisible(False)
        self.progressLabel.setVisible(False)
        self.bottomLayout.addWidget(self.inProgressBar)

        self.timeLeft = 9
        self.finishBtn = PrimaryPushButton(FIF.ACCEPT, f'完成({self.timeLeft})', self)
        self.finishBtn.setVisible(False)
        self.viewBtn = PushButton(FIF.FOLDER, '查看', self)
        self.viewBtn.setVisible(False)
        self.ejectBtn = PushButton(FIF.EMBED, '退出U盘', self)
        self.ejectBtn.setVisible(False)
        self.finishBtn.clicked.connect(self.Quit)
        self.viewBtn.clicked.connect(self.onViewBtn)
        self.ejectBtn.clicked.connect(self.onEjectBtn)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.updateTime)

        self.mainLayout.addWidget(self.titleBar)
        self.mainLayout.addLayout(self.topLayout)
        self.mainLayout.addStretch(1)
        self.mainLayout.addLayout(self.bottomLayout)

        self.deleteThread = DeleteThread()
        self.syncThread = SyncThread()
        self.ejectThread = EjectThread()
        if isDelete:
            self.statusLabel.setText("正在删除原有文件")
            self.setupDeleteThread()
            self.startDeleteThread()
        else:
            self.statusLabel.setText("准备中")
            self.setupSyncThread()
            self.startSyncThread()

        self.taskbarButton = QWinTaskbarButton(self)
        self.taskbarProgress = self.taskbarButton.progress()
        self.taskbarProgress.setRange(0, 0)

        self.opacityAni = QPropertyAnimation(self, b'windowOpacity', self)
        self.opacityAni.setDuration(1000)
        self.opacityAni.setStartValue(1)
        self.opacityAni.setEndValue(0)
        self.opacityAni.finished.connect(self.Quit)

    def showEvent(self, event):
        super(Window, self).showEvent(event)
        if not self.taskbarButton.window():
            self.taskbarButton.setWindow(self.windowHandle())
            self.taskbarProgress.show()

    def closeEvent(self, event):
        super(Window, self).closeEvent(event)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.update()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            startSystemMove(self, event.globalPos())
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        super().mouseMoveEvent(event)
        self.update()

    def setupDeleteThread(self):
        self.deleteThread.deleteFinished.connect(self.deleteThreadFinished)
        self.deleteThreadRunning = True

    def startDeleteThread(self):
        if self.deleteThreadRunning:
            self.deleteThread.start()
        else:
            self.setupDeleteThread()
            self.deleteThread.start()

    def deleteThreadFinished(self):
        self.deleteThread.quit()
        self.deleteThreadRunning = False
        self.setupSyncThread()
        self.startSyncThread()
        self.statusLabel.setText("准备中")

    def setupSyncThread(self):
        self.syncThread.valueChange.connect(self.setSyncValue)
        self.syncThreadRunning = True

    def startSyncThread(self):
        if self.syncThreadRunning:
            self.syncThread.start()
        else:
            self.setupSyncThread()
            self.syncThread.start()

    def syncThreadFinished(self):
        self.isProgress = False
        try:
            self.bottomLayout.removeWidget(self.progressBar)
        except:
            pass
        try:
            self.bottomLayout.removeWidget(self.inProgressBar)
        except:
            pass
        try:
            self.bottomLayout.removeWidget(self.progressLabel)
        except:
            pass
        try:
            self.topLayout.removeWidget(self.showDetailBtn)
        except:
            pass
        try:
            self.topLayout.removeWidget(self.cancelBtn)
        except:
            pass
        try:
            self.topLayout.removeWidget(self.pauseBtn)
        except:
            pass
        self.progressBar.deleteLater()
        self.inProgressBar.deleteLater()
        self.progressLabel.deleteLater()
        self.bottomLayout.deleteLater()
        self.showDetailBtn.deleteLater()
        self.cancelBtn.deleteLater()
        self.pauseBtn.deleteLater()

        self.statusLabel.setText({1:"同步完成", 2:"同步完成", 3:"复制完成", 4:"复制完成"}[mode])
        self.detailLabel.setText(self.driveName)
        self.finishBtn.setVisible(True)
        self.viewBtn.setVisible(True)
        self.ejectBtn.setVisible(True)
        self.topLayout.addWidget(self.finishBtn, 0, 1, 2, 1)
        self.topLayout.addWidget(self.viewBtn, 0, 2, 2, 1)
        self.topLayout.addWidget(self.ejectBtn, 0, 3, 2, 1)

        self.timer.start(1000)

    def setSyncValue(self):
        if self.syncThread.progress_value == -1:
            self.syncThread.terminate()
            self.syncThreadRunning = False
            self.taskbarProgress.setVisible(False)
            if cfg.Notify.value:
                if mode == 1 or mode == 2:
                    title = "同步完成"
                else:
                    title = "复制完成"
                toast = Notification(app_id="Presto", title=title, msg=self.driveName, duration="short")
                toast.set_audio(audio.Default, loop=False)
                toast.show()
            self.syncThreadFinished()
        else:
            if self.isPrepare:
                if mode == 1 or mode == 2:
                    self.statusLabel.setText("正在同步")
                else:
                    self.statusLabel.setText("正在复制")
                self.bottomLayout.removeWidget(self.inProgressBar)
                self.inProgressBar.setVisible(False)
                self.progressBar.setVisible(True)
                self.spaceLabel.setVisible(True)
                self.progressLabel.setVisible(True)
                self.isPrepare = False
                self.taskbarProgress.setRange(0, 100)

            self.taskbarProgress.setValue(self.syncThread.progress_value)
            self.progressBar.setValue(self.syncThread.progress_value)
            self.progressLabel.setText(str(self.syncThread.progress_value) + '%')
            self.bottomLayout.addWidget(self.progressBar)
            self.bottomLayout.addWidget(self.spaceLabel)
            self.bottomLayout.addWidget(self.progressLabel)

    def setupEjectThread(self):
        self.ejectThread.ejectFinished.connect(self.ejectThreadFinished)
        self.ejectThreadRunning = True

    def startEjectThread(self):
        if self.ejectThreadRunning:
            self.ejectThread.start()
        else:
            self.setupEjectThread()
            self.ejectThread.start()

    def ejectThreadFinished(self):
        self.ejectThread.isRunning = False
        self.ejectThread.terminate()
        self.ejectThreadRunning = False

        if self.ejectThread.exitCode == 0:
            self.ejectInfoBar.statusChanged(True)
            QTimer.singleShot(1000, self.opacityAni.start)
        else:
            self.ejectInfoBar.statusChanged(False)
            QTimer.singleShot(1000, self.opacityAni.start)

    def killSubprocess(self):
        children = Process(os.getpid()).children(recursive=True)
        for child in children:
            try:
                child.terminate()
            except:
                pass
        return

    def stopThread(self):
        self.close()
        self.progressBar.pause()
        self.inProgressBar.pause()
        self.taskbarProgress.stop()

        self.deleteThread.stop()
        self.syncThread.stop()
        self.deleteThreadRunning = False
        self.syncThreadRunning = False
        self.deleteThread.quit()
        self.syncThread.quit()

        self.killSubprocess()
        sys.exit()

    def getDriveName(self):
        try:
            if GetVolumeInformation(drive)[0] != '':
                return GetVolumeInformation(drive)[0]
            else:
                return "U盘"
        except:
            sys.exit()

    def onCancelBtn(self):
        if self.isProgress:
            yesBtn = PushButton('确定')
            yesBtn.clicked.connect(self.stopThread)
            w = InfoBar(icon=InfoBarIcon.WARNING,
                        title='取消此次任务',
                        content='',
                        orient=Qt.Horizontal,
                        isClosable=True,
                        position=InfoBarPosition.BOTTOM,
                        duration=-1,
                        parent=self)
            w.addWidget(yesBtn)
            w.show()
        else:
            self.close()
            sys.exit()

    def onShowDetailBtn(self):
        w = DetailDialog('Presto 选项', self.displayText, self.subject, self)
        w.setTitleBarVisible(False)
        if w.exec():
            pass

    def onPauseBtn(self):
        if self.isPaused:
            """resume"""
            self.pauseBtn.setText("暂停")
            self.pauseBtn.setIcon(FIF.PAUSE)
            self.isPaused = False
            self.statusLabel.setText("准备中")
            if self.deleteThreadRunning:
                self.statusLabel.setText("正在删除原有文件")
                self.inProgressBar.setPaused(False)
                self.progressBar.setPaused(False)
                self.taskbarProgress.setPaused(False)
                self.setupDeleteThread()
                self.startDeleteThread()
            elif self.syncThreadRunning:
                self.statusLabel.setText("正在同步")
                self.inProgressBar.setPaused(False)
                self.progressBar.setPaused(False)
                self.taskbarProgress.setPaused(False)
                self.setupSyncThread()
                self.startSyncThread()
        else:
            """pause"""
            self.pauseBtn.setText("继续")
            self.pauseBtn.setIcon(FIF.PLAY)
            self.isPaused = True
            self.statusLabel.setText("已暂停")
            if self.deleteThreadRunning:
                self.inProgressBar.setPaused(True)
                self.progressBar.setPaused(True)
                self.taskbarProgress.setPaused(True)
                self.deleteThread.isRunning = False
                self.deleteThread.terminate()
                self.killSubprocess()
            elif self.syncThreadRunning:
                self.inProgressBar.setPaused(True)
                self.progressBar.setPaused(True)
                self.taskbarProgress.setPaused(True)
                self.syncThread.isRunning = True
                self.syncThread.terminate()
                self.killSubprocess()

    def Quit(self):
        self.close()
        sys.exit()

    def onViewBtn(self):
        os.startfile(destFolder)
        self.close()
        sys.exit()

    def onEjectBtn(self):
        self.timer.stop()
        self.statusLabel.setDisabled(True)
        self.detailLabel.setDisabled(True)
        self.ejectBtn.setDisabled(True)
        self.finishBtn.setDisabled(True)
        self.viewBtn.setDisabled(True)

        self.ejectInfoBar = EjectUsbInfoBar(
            title='正在退出U盘',
            position=InfoBarPosition.BOTTOM,
            parent=self.window()
        )
        self.ejectInfoBar.show()

        self.setupEjectThread()
        self.startEjectThread()

    def updateTime(self):
        self.timeLeft -= 1
        if self.timeLeft == 0:
            self.opacityAni.start()
        else:
            self.finishBtn.setText(f'完成 ({self.timeLeft})')


if __name__ == '__main__':

    """
    args
    1           drive
    2 - 12      subjects
    13          mode{1:"sync(default)", 2:"sync(low)", 3:"copy(lately)", 4:"copy(from_date)"}
    14          isDelete
    15          commandOption
    """

    try:
        drive = sys.argv[1][0] + ':'
        taskList = []
        for i in range(2, 13):
            if sys.argv[i] == '1':
                taskList.append(i - 1)
        taskNum = len(taskList)
        buf = str(cfg.BufSize.value)[9:]
        concurrentProcess = cfg.ConcurrentProcess.value
        sourceFolder = os.path.normpath(cfg.sourceFolder.value)
        destFolder = drive + '\\' + os.path.basename(sourceFolder) + '\\'
        mode = int(sys.argv[13])
        isDelete = False if sys.argv[14] == 'False' else True

        commandOption = sys.argv[15]
        if cfg.IsSkipEmptyDir.value:
            commandOption = commandOption + " /skip_empty_dir"
        else:
            commandOption = commandOption + " /skip_empty_dir=FALSE"
        if cfg.IsSizeFilter.value:
            commandOption = commandOption + f' /max_size="{cfg.SizeFilterValue.value}{cfg.SizeFilterUnit.value[0]}"'

        currentTask = 0
    except:
        sys.exit()

    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps)
    if darkdetect.isDark():
        setTheme(Theme.DARK)
    else:
        setTheme(Theme.LIGHT)
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    app.exec()
