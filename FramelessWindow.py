# -*- coding: utf-8 -*-

import win32con
import win32gui
import darkdetect
from ctypes import cast
from ctypes.wintypes import LPRECT, MSG
from PyQt5.QtCore import Qt, QSize, QRect
from PyQt5.QtGui import QCloseEvent, QCursor
from PyQt5.QtWidgets import QApplication, QWidget, QDialog, QMainWindow
from qframelesswindow.utils import win32_utils as win_utils
from qframelesswindow.utils.win32_utils import Taskbar
from qframelesswindow.windows.c_structures import LPNCCALCSIZE_PARAMS
from qframelesswindow.windows.window_effect import WindowsWindowEffect
from qframelesswindow.titlebar import TitleBar
from qframelesswindow.windows import AcrylicWindow
from qframelesswindow.windows import WindowsFramelessWindow as FramelessWindow

AcrylicWindow = FramelessWindow


class FramelessDialog(QDialog, FramelessWindow):

    def __init__(self, parent=None):
        super().__init__(parent)
        self.titleBar.minBtn.hide()
        self.titleBar.maxBtn.hide()
        self.titleBar.setDoubleClickEnabled(False)
        self.windowEffect.disableMaximizeButton(self.winId())


class FramelessMainWindow(QMainWindow, FramelessWindow):

    def __init__(self, parent=None):
        super().__init__(parent)


class WindowsFramelessWindow(QWidget):

    BORDER_WIDTH = 5

    def __init__(self, parent=None):
        super().__init__(parent=parent)
        self.windowEffect = WindowsWindowEffect(self)
        self.titleBar = TitleBar(self)
        self._isSystemButtonVisible = False
        self._isResizeEnabled = True

        self.updateFrameless()

        if darkdetect.isDark():
            self.setStyleSheet("background-color: rgb(39, 39, 39);")
        else:
            self.setStyleSheet("background-color: rgb(249, 249, 249);")

        self.windowHandle().screenChanged.connect(self.__onScreenChanged)

        self.resize(500, 500)
        self.titleBar.raise_()

    def updateFrameless(self):
        if not win_utils.isWin7():
            self.setWindowFlags(self.windowFlags() | Qt.FramelessWindowHint)
        elif self.parent():
            self.setWindowFlags(self.parent().windowFlags() | Qt.FramelessWindowHint | Qt.WindowMinMaxButtonsHint)
        else:
            self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowMinMaxButtonsHint)

        self.windowEffect.addWindowAnimation(self.winId())
        if not isinstance(self, AcrylicWindow):
            self.windowEffect.addShadowEffect(self.winId())

    def setTitleBar(self, titleBar):
        self.titleBar.deleteLater()
        self.titleBar.hide()
        self.titleBar = titleBar
        self.titleBar.setParent(self)
        self.titleBar.raise_()

    def setResizeEnabled(self, isEnabled: bool):
        self._isResizeEnabled = isEnabled

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self.titleBar.resize(self.width(), self.titleBar.height())

    def isSystemButtonVisible(self):
        return self._isSystemButtonVisible

    def setSystemTitleBarButtonVisible(self, isVisible):
        pass

    def systemTitleBarRect(self, size: QSize) -> QRect:
        return QRect(0, 0, size.width(), size.height())

    def nativeEvent(self, eventType, message):
        msg = MSG.from_address(message.__int__())
        if not msg.hWnd:
            return super().nativeEvent(eventType, message)

        if msg.message == win32con.WM_NCHITTEST and self._isResizeEnabled:
            pos = QCursor.pos()
            xPos = pos.x() - self.x()
            yPos = pos.y() - self.y()
            w = self.frameGeometry().width()
            h = self.frameGeometry().height()

            bw = 0 if win_utils.isMaximized(msg.hWnd) or win_utils.isFullScreen(msg.hWnd) else self.BORDER_WIDTH
            lx = xPos < bw
            rx = xPos > w - bw
            ty = yPos < bw
            by = yPos > h - bw
            if lx and ty:
                return True, win32con.HTTOPLEFT
            elif rx and by:
                return True, win32con.HTBOTTOMRIGHT
            elif rx and ty:
                return True, win32con.HTTOPRIGHT
            elif lx and by:
                return True, win32con.HTBOTTOMLEFT
            elif ty:
                return True, win32con.HTTOP
            elif by:
                return True, win32con.HTBOTTOM
            elif lx:
                return True, win32con.HTLEFT
            elif rx:
                return True, win32con.HTRIGHT
        elif msg.message == win32con.WM_NCCALCSIZE:
            if msg.wParam:
                rect = cast(msg.lParam, LPNCCALCSIZE_PARAMS).contents.rgrc[0]
            else:
                rect = cast(msg.lParam, LPRECT).contents

            isMax = win_utils.isMaximized(msg.hWnd)
            isFull = win_utils.isFullScreen(msg.hWnd)

            if isMax and not isFull:
                ty = win_utils.getResizeBorderThickness(msg.hWnd, False)
                rect.top += ty
                rect.bottom -= ty

                tx = win_utils.getResizeBorderThickness(msg.hWnd, True)
                rect.left += tx
                rect.right -= tx

            if (isMax or isFull) and Taskbar.isAutoHide():
                position = Taskbar.getPosition(msg.hWnd)
                if position == Taskbar.LEFT:
                    rect.top += Taskbar.AUTO_HIDE_THICKNESS
                elif position == Taskbar.BOTTOM:
                    rect.bottom -= Taskbar.AUTO_HIDE_THICKNESS
                elif position == Taskbar.LEFT:
                    rect.left += Taskbar.AUTO_HIDE_THICKNESS
                elif position == Taskbar.RIGHT:
                    rect.right -= Taskbar.AUTO_HIDE_THICKNESS

            result = 0 if not msg.wParam else win32con.WVR_REDRAW
            return True, result

        return super().nativeEvent(eventType, message)

    def __onScreenChanged(self):
        hWnd = int(self.windowHandle().winId())
        win32gui.SetWindowPos(hWnd, None, 0, 0, 0, 0, win32con.SWP_NOMOVE |
                              win32con.SWP_NOSIZE | win32con.SWP_FRAMECHANGED)


class AcrylicWindow(WindowsFramelessWindow):

    def __init__(self, parent=None):
        super().__init__(parent=parent)
        self.__closedByKey = False
        self.setStyleSheet("AcrylicWindow{background:transparent}")

    def updateFrameless(self):
        super().updateFrameless()
        self.windowEffect.enableBlurBehindWindow(self.winId())

        if win_utils.isWin7() and self.parent():
            self.setWindowFlags(self.parent().windowFlags() | Qt.FramelessWindowHint | Qt.WindowMinMaxButtonsHint)
        else:
            self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowMinMaxButtonsHint)

        self.windowEffect.addWindowAnimation(self.winId())

        if win_utils.isWin7():
            self.windowEffect.addShadowEffect(self.winId())
            self.windowEffect.setAeroEffect(self.winId())
        else:
            self.windowEffect.setAcrylicEffect(self.winId())
            if win_utils.isGreaterEqualWin11():
                self.windowEffect.addShadowEffect(self.winId())

    def nativeEvent(self, eventType, message):
        msg = MSG.from_address(message.__int__())

        if msg.message == win32con.WM_SYSKEYDOWN:
            if msg.wParam == win32con.VK_F4:
                self.__closedByKey = True
                QApplication.sendEvent(self, QCloseEvent())
                return False, 0

        return super().nativeEvent(eventType, message)

    def closeEvent(self, e):
        if not self.__closedByKey or QApplication.quitOnLastWindowClosed():
            self.__closedByKey = False
            return super().closeEvent(e)

        self.__closedByKey = False
        self.hide()
