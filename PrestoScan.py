# -*- coding: utf-8 -*-

import os
import sys
import time
import subprocess
import darkdetect
import portalocker
import PrestoResource
from PrestoConfig import cfg
from psutil import disk_partitions
from webbrowser import open as WebOpen
from win32api import GetVolumeInformation
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QPropertyAnimation, QEvent, QRunnable, QThreadPool, QObject
from PyQt5.QtGui import QIcon, QCursor, QColor, QPainter
from PyQt5.QtWidgets import QApplication, QSystemTrayIcon, QMainWindow, QAction, QFrame, QLabel, QVBoxLayout, \
    QGraphicsOpacityEffect, QGridLayout
from qfluentwidgets import RoundMenu, setTheme, Theme, InfoBarPosition, IndeterminateProgressRing, FluentStyleSheet, \
    InfoBarIcon, isDarkTheme, setThemeColor
from qfluentwidgets.components.widgets.info_bar import InfoIconWidget, InfoBarManager, InfoBar
from qfluentwidgets import FluentIcon as FIF


class Mutex:
    def __init__(self):
        self.file = None

    def __enter__(self):
        self.file = open('PrestoScan.lockfile', 'w')
        try:
            portalocker.lock(self.file, portalocker.LOCK_EX | portalocker.LOCK_NB)
        except portalocker.AlreadyLocked:
            self.file.close()
            sys.exit()

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.file:
            try:
                portalocker.unlock(self.file)
            except:
                pass
            self.file.close()
            try:
                os.remove('PrestoScan.lockfile')
            except:
                pass


class EjectUsbInfoBar(QFrame):

    closedSignal = pyqtSignal()
    _desktopView = None

    def __init__(self, title: str, drive: str, position, parent=None):
        super().__init__(parent=parent)
        self.title = title
        self.content = self.getDriveName(drive) + f" ({drive})"
        self.position = position

        self.titleLabel = QLabel(self)
        self.contentLabel = QLabel(self)
        self.progressRing = IndeterminateProgressRing()
        self.progressRing.setFixedSize(18, 36)
        self.progressRing.setStrokeWidth(3)

        self.layout = QGridLayout(self)
        self.layout.setHorizontalSpacing(0)

        self.opacityEffect = QGraphicsOpacityEffect(self)
        self.opacityAni = QPropertyAnimation(self.opacityEffect, b'opacity', self)

        self.lightBackgroundColor = QColor(244, 244, 244)
        self.darkBackgroundColor = QColor(39, 39, 39)

        self.__initWidget()
        self.startEjectThread(drive)

    def __initWidget(self):
        self.opacityEffect.setOpacity(1)
        self.setGraphicsEffect(self.opacityEffect)

        self.setFixedSize(200, 85)
        self.setQss()
        self.setLayout(False)

    def setLayout(self, isStatusChanged):
        if isStatusChanged:
            self.layout.setContentsMargins(6, 6, 24, 16)
            self.layout.setSizeConstraint(QVBoxLayout.SetMinimumSize)
            self.layout.setHorizontalSpacing(6)
            self.layout.addWidget(self.iconWidget, 0, 0, 1, 1)
        else:
            self.layout.setContentsMargins(16, 6, 24, 16)
            self.layout.setSizeConstraint(QVBoxLayout.SetMinimumSize)
            self.layout.setHorizontalSpacing(15)
            self.layout.addWidget(self.progressRing, 0, 0, 1, 1)
        self.layout.addWidget(self.titleLabel, 0, 1, 1, 1)
        self.layout.addWidget(self.contentLabel, 1, 1, 1, 1)
        self.adjustText()

    def setQss(self):
        self.titleLabel.setObjectName('titleLabel')
        self.contentLabel.setObjectName('contentLabel')

        FluentStyleSheet.INFO_BAR.apply(self)

    def fadeOut(self):
        self.opacityAni.setDuration(3000)
        self.opacityAni.setStartValue(1)
        self.opacityAni.setEndValue(0)
        self.opacityAni.finished.connect(self.opacityAniFinished)
        self.opacityAni.start()

    def opacityAniFinished(self):
        self.manager.remove(self)
        self.close()

    def adjustText(self):
        self.titleLabel.setText(self.title)
        self.contentLabel.setText(self.content)

    def setCustomBackgroundColor(self, light, dark):
        self.lightBackgroundColor = QColor(light)
        self.darkBackgroundColor = QColor(dark)
        self.update()

    def ejectFinished(self):
        isSuccess = self.task.signal.exitCode == 0

        self.layout.removeWidget(self.progressRing)
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
        self.fadeOut()

    def eventFilter(self, obj, e: QEvent):
        if obj is self.parent():
            if e.type() in [QEvent.Resize, QEvent.WindowStateChange]:
                self.adjustText()
        return super().eventFilter(obj, e)

    def showEvent(self, e):
        self.adjustText()
        super().showEvent(e)

        if self.position != InfoBarPosition.NONE:
            self.manager = InfoBarManager.make(self.position)
            self.manager.add(self)

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

    def getDriveName(self, drive):
        try:
            if GetVolumeInformation(drive)[0] != '':
                return GetVolumeInformation(drive)[0]
            else:
                return "未知"
        except:
            return

    def startEjectThread(self, drive):
        self.task = EjectRunnable(drive)
        self.task.signal.ejectFinished.connect(self.ejectFinished)
        QThreadPool.globalInstance().start(self.task)


class ScanThread(QThread):

    listChanged = pyqtSignal(bool)

    def __init__(self, parent=None):
        super().__init__(parent=parent)
        self.drive = ''
        self.isRunning = True

        self.cycle = cfg.ScanCycle.value / 10
        self.local_device = []
        self.local_letter = []
        self.local_number = 0
        self.mobile_device = []
        self.mobile_letter = []
        self.mobile_number = 0
        self.now_number = 0
        self.before_number = self.update()
        self.before_letter = self.local_letter + self.mobile_letter

    def stop(self):
        self.isRunning = False

    def update(self):
        tmp_local_device, tmp_local_letter = [], []
        tmp_mobile_device, tmp_mobile_letter = [], []
        tmp_local_number, tmp_mobile_number = 0, 0
        try:
            part = disk_partitions()
        except:
            pass
        else:
            for i in range(len(part)):
                tmplist = part[i].opts.split(",")
                if len(tmplist) > 1:
                    if tmplist[1] == "fixed":
                        tmp_local_number = tmp_local_number + 1
                        tmp_local_letter.append(part[i].device[:2])
                        tmp_local_device.append(part[i])
                    else:
                        tmp_mobile_number = tmp_mobile_number + 1
                        tmp_mobile_letter.append(part[i].device[:2])
                        tmp_mobile_device.append(part[i])
            self.local_device, self.local_letter = tmp_local_device[:], tmp_local_letter[:]
            self.mobile_device, self.mobile_letter = tmp_mobile_device[:], tmp_mobile_letter[:]
            self.local_number, self.mobile_number = tmp_local_number, tmp_mobile_number
        return len(part)

    def run(self):
        while self.isRunning:
            self.now_number = self.update()
            if self.now_number > self.before_number and len(set(self.local_letter + self.mobile_letter).difference(set(self.before_letter))) == 1:
                self.drive = ''.join(set(self.local_letter + self.mobile_letter).difference(set(self.before_letter)))
                subprocess.call(["PrestoUsbService.exe ", self.drive], shell=True)
                self.before_number = self.now_number
                self.before_device = self.local_device + self.mobile_letter
                self.before_letter = self.local_letter + self.mobile_letter
                self.listChanged.emit(True)
            elif self.now_number < self.before_number:
                self.before_number = self.now_number
                self.before_letter = self.local_letter + self.mobile_letter
                self.listChanged.emit(True)
            time.sleep(self.cycle)
        else:
            return


class EjectSignal(QObject):
    ejectFinished = pyqtSignal(bool)
    exitCode = -1


class EjectRunnable(QRunnable):
    def __init__(self, drive):
        super().__init__()
        self.drive = drive
        self.signal = EjectSignal()
        self.signal.moveToThread(QApplication.instance().thread())
        self.setAutoDelete(True)

    def run(self):
        exit_code = -1
        for i in range(2):
            exit_code = subprocess.call(["RemoveDrive.exe", self.drive, '-f'], shell=True)
            if exit_code == 0:
                break

        self.signal.exitCode = exit_code
        self.signal.ejectFinished.emit(True)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Presto")
        self.setWindowIcon(QIcon(":/icon.png"))
        self.resize(400, 300)


class TrayApp:
    def __init__(self):
        setThemeColor(QColor(113, 89, 249))
        self.app = QApplication(sys.argv)
        self.main_window = MainWindow()
        self._restore_action = QAction()
        self._quit_action = QAction()
        self._tray_icon_menu = RoundMenu("", self.main_window)
        self.exeSubMenu = RoundMenu("快速启动", self._tray_icon_menu)
        self.exeSubMenu.setIcon(FIF.SEND)
        self.ejectSubMenu = RoundMenu("退出U盘", self._tray_icon_menu)
        self.ejectSubMenu.setIcon(FIF.EMBED)

        self.tray_icon = QSystemTrayIcon()
        self.tray_icon.setIcon(QIcon(":/icon.png"))
        self.tray_icon.setToolTip("Presto U盘扫描")
        self.createActions()
        self.createTrayIcon()
        self.tray_icon.show()
        self.tray_icon.activated.connect(self.trayIconActivated)

        self.onStartCheck()
        self.updateDriveActions()

        self.scanThread = ScanThread()
        self.scanThread.listChanged.connect(self.updateDriveActions)
        self.scanThread.start()

    def trayIconActivated(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.Trigger or reason == QSystemTrayIcon.ActivationReason.Context:
            self._tray_icon_menu.exec(QCursor.pos())

    def createActions(self):
        self._launcher_action = QAction(FIF.APPLICATION.icon(), "启动台")
        self._launcher_action.triggered.connect(lambda: subprocess.Popen("PrestoLauncher.exe", shell=True))
        self._setting_action = QAction(FIF.SETTING.icon(), "设置")
        self._setting_action.triggered.connect(lambda: subprocess.Popen("PrestoSetting.exe", shell=True))
        self._help_action = QAction(FIF.HELP.icon(), "帮助")
        self._help_action.triggered.connect(self.onHelpAction)

        self._quit_action = QAction(FIF.POWER_BUTTON.icon(), "退出")
        self._quit_action.triggered.connect(self.quit)

    def createTrayIcon(self):
        self._tray_icon_menu.addAction(self._launcher_action)
        self._tray_icon_menu.addAction(self._setting_action)
        self._tray_icon_menu.addAction(self._help_action)
        self._tray_icon_menu.addSeparator()
        self._tray_icon_menu.addMenu(self.exeSubMenu)
        self._tray_icon_menu.addMenu(self.ejectSubMenu)
        self._tray_icon_menu.addSeparator()
        self._tray_icon_menu.addAction(self._quit_action)
        self.tray_icon.show()

    def quit(self):
        self.tray_icon.hide()
        self.tray_icon.deleteLater()
        self.scanThread.stop()
        self.scanThread.wait()
        QApplication.quit()

    def onStartCheck(self):
        part = disk_partitions()
        for i in range(len(part)):
            tmplist = part[i].opts.split(",")
            if len(tmplist) > 1 and tmplist[1] == "removable":
                subprocess.call(["PrestoUsbService.exe ", str(part[i].device[:2])], shell=True)

    def onHelpAction(self):
        if os.path.exists(os.path.abspath("./Doc/PrestoHelp.html")):
            os.startfile(os.path.abspath("./Doc/PrestoHelp.html"))
        else:
            WebOpen("https://sudo0015.github.io/post/Presto%20-bang-zhu.html")

    def updateDriveActions(self):
        for action in self.exeSubMenu.actions():
            self.exeSubMenu.removeAction(action)
            action.deleteLater()
        for action in self.ejectSubMenu.actions():
            self.ejectSubMenu.removeAction(action)
            action.deleteLater()

        drives = ["D:", "E:", "F:", "G:", "H:", "I:"]
        for drive in drives:
            drive_name = self.getDriveName(drive)
            if drive_name:
                exe_action = QAction(f"{drive_name} ({drive})", self.exeSubMenu)
                eject_action = QAction(f"{drive_name} ({drive})", self.ejectSubMenu)
                exe_action.triggered.connect(lambda _, d=drive: subprocess.Popen(["PrestoUsbService.exe", d], shell=True))
                eject_action.triggered.connect(lambda _, d=drive: self.ejectDrive(d))

                self.exeSubMenu.addAction(exe_action)
                self.ejectSubMenu.addAction(eject_action)

        if not self.exeSubMenu.actions():
            exeActionNull = QAction("(无驱动器)", self.exeSubMenu)
            ejectActionNull = QAction("(无驱动器)", self.ejectSubMenu)
            exeActionNull.setDisabled(True)
            ejectActionNull.setDisabled(True)
            self.exeSubMenu.addAction(exeActionNull)
            self.ejectSubMenu.addAction(ejectActionNull)

    def getDriveName(self, drive):
        try:
            if GetVolumeInformation(drive)[0] != '':
                return GetVolumeInformation(drive)[0]
            else:
                return "未知"
        except:
            return ""

    def ejectDrive(self, drive):
        w = EjectUsbInfoBar(
            title='正在退出U盘',
            drive=drive,
            position=InfoBarPosition.BOTTOM_RIGHT,
            parent=InfoBar.desktopView()
        )
        w.show()

    def run(self):
        sys.exit(self.app.exec())


if __name__ == "__main__":
    if (len(sys.argv) == 2 and sys.argv[1] == '--force-start') or cfg.AutoRun.value:
        with Mutex():
            QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
            QApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
            QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps)
            if darkdetect.isDark():
                setTheme(Theme.DARK)
            else:
                setTheme(Theme.LIGHT)
            app = TrayApp()
            app.run()
    else:
        sys.exit()
