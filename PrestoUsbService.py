# -*- coding: utf-8 -*-

import os
import sys
import subprocess
import darkdetect
import PrestoResource
from webbrowser import open as WebOpen
from win32file import GetDiskFreeSpace
from win32api import GetVolumeInformation
from PyQt5.QtGui import QIcon, QColor, QPainterPath, QPainter
from PyQt5.QtCore import Qt, QPoint, QTimer, QDate, QRectF, QPropertyAnimation, QParallelAnimationGroup, \
    QEasingCurve, QEvent
from PyQt5.QtWidgets import QApplication, QHBoxLayout, QVBoxLayout, QLabel, QStackedWidget, QWidget, QGridLayout, \
    QFrame, QPushButton, QSpinBox, QLineEdit, QAction
from qfluentwidgets import setTheme, Theme, isDarkTheme, CheckBox, PrimaryPushButton, PushButton, \
    SubtitleLabel, Pivot, TransparentToolButton, RoundMenu, BodyLabel, CaptionLabel, Action, \
    TransparentPushButton, setThemeColor, PrimarySplitPushButton, ZhDatePicker, MaskDialogBase, \
    PrimaryDropDownPushButton, MenuAnimationType, setFont, ToolTipFilter, ToolTipPosition, ProgressRing
from qframelesswindow.titlebar import MinimizeButton, CloseButton, MaximizeButton
from qframelesswindow import TitleBarButton
from qfluentwidgets.common.style_sheet import FluentStyleSheet, themeColor
from qfluentwidgets.components.widgets.spin_box import SpinButton, SpinIcon
from qframelesswindow.utils import startSystemMove
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
        self.closeBtn.clicked.connect(self.Quit)

        self.window().installEventFilter(self)

    def Quit(self):
        self.window().close()
        sys.exit()

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
        self.minBtn.deleteLater()
        self.maxBtn.deleteLater()
        self.hBoxLayout.removeWidget(self.closeBtn)
        self.titleLabel = QLabel(self)
        self.titleLabel.setText("Presto")
        self.titleLabel.setObjectName('titleLabel')
        self.vBoxLayout = QVBoxLayout()
        self.buttonLayout = QHBoxLayout()
        self.buttonLayout.setSpacing(0)
        self.buttonLayout.setContentsMargins(0, 0, 0, 0)
        self.buttonLayout.setAlignment(Qt.AlignTop)
        self.buttonLayout.addWidget(self.closeBtn)
        self.titleLayout = QHBoxLayout()
        self.titleLayout.setContentsMargins(0, 8, 0, 0)
        self.titleLayout.addWidget(self.titleLabel)
        self.titleLayout.setAlignment(Qt.AlignTop)
        self.vBoxLayout.addLayout(self.buttonLayout)
        self.vBoxLayout.addStretch(1)
        self.hBoxLayout.addLayout(self.titleLayout, 0)
        self.hBoxLayout.addStretch(30)
        self.hBoxLayout.addLayout(self.vBoxLayout, 0)
        FluentStyleSheet.FLUENT_WINDOW.apply(self)


class MicaWindow(Window):

    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setTitleBar(FluentTitleBar(self))
        if sys.platform == 'win32' and sys.getwindowsversion().build >= 22000:
            self.windowEffect.setMicaEffect(self.winId(), isDarkTheme())


class MessageBoxBase(MaskDialogBase):
    """ Message box base """

    def __init__(self, parent=None):
        super().__init__(parent=parent)
        self.isDelete = None
        self.menu = RoundMenu(parent=self)
        self.passDeleteAction = Action(FIF.ACCEPT, '保留原有文件')
        self.deleteAction = Action(FIF.DELETE, '删除原有文件')
        self.menu.addAction(self.passDeleteAction)
        self.menu.addAction(self.deleteAction)

        self.buttonGroup = QFrame(self.widget)
        self.yesButton = PrimaryDropDownPushButton('确定', self.buttonGroup)
        self.yesButton.setMenu(self.menu)
        self.cancelButton = QPushButton('取消', self.buttonGroup)

        self.vBoxLayout = QVBoxLayout(self.widget)
        self.viewLayout = QVBoxLayout()
        self.buttonLayout = QHBoxLayout(self.buttonGroup)

        self.__initWidget()

    def __initWidget(self):
        self.__setQss()
        self.__initLayout()

        self.setShadowEffect(60, (0, 10), QColor(0, 0, 0, 50))
        self.setMaskColor(QColor(0, 0, 0, 76))

        self.yesButton.setAttribute(Qt.WA_LayoutUsesWidgetRect)
        self.cancelButton.setAttribute(Qt.WA_LayoutUsesWidgetRect)

        self.yesButton.setAttribute(Qt.WA_MacShowFocusRect, False)

        self.yesButton.setFocus()
        self.buttonGroup.setFixedHeight(81)

        self.passDeleteAction.triggered.connect(self.__onPassDeleteActionClicked)
        self.deleteAction.triggered.connect(self.__onDeleteActionClicked)
        self.cancelButton.clicked.connect(self.__onCancelButtonClicked)

    def __initLayout(self):
        self._hBoxLayout.removeWidget(self.widget)
        self._hBoxLayout.addWidget(self.widget, 1, Qt.AlignCenter)

        self.vBoxLayout.setSpacing(0)
        self.vBoxLayout.setContentsMargins(0, 0, 0, 0)
        self.vBoxLayout.addLayout(self.viewLayout, 1)
        self.vBoxLayout.addWidget(self.buttonGroup, 0, Qt.AlignBottom)

        self.viewLayout.setSpacing(12)
        self.viewLayout.setContentsMargins(24, 24, 24, 24)

        self.buttonLayout.setSpacing(12)
        self.buttonLayout.setContentsMargins(24, 24, 24, 24)
        self.buttonLayout.addWidget(self.yesButton, 1, Qt.AlignVCenter)
        self.buttonLayout.addWidget(self.cancelButton, 1, Qt.AlignVCenter)

    def validate(self) -> bool:
        return True

    def __onCancelButtonClicked(self):
        self.reject()

    def __onPassDeleteActionClicked(self):
        if self.validate():
            self.isDelete = False
            self.accept()

    def __onDeleteActionClicked(self):
        if self.validate():
            self.isDelete = True
            self.accept()

    def __setQss(self):
        self.buttonGroup.setObjectName('buttonGroup')
        self.cancelButton.setObjectName('cancelButton')
        FluentStyleSheet.DIALOG.apply(self)

    def hideYesButton(self):
        self.yesButton.hide()
        self.buttonLayout.insertStretch(0, 1)

    def hideCancelButton(self):
        self.cancelButton.hide()
        self.buttonLayout.insertStretch(0, 1)


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


class LatelyCopyMessageBox(MessageBoxBase):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.titleLabel = SubtitleLabel('复制 (最近文件)', self)
        self.textLabel = BodyLabel('范围 (单位: 天)', self)
        self.spinBox = SpinBox(self)
        self.spinBox.setFixedWidth(130)
        self.spinBox.setAccelerated(True)
        self.spinBox.setValue(7)
        self.spinBox.setMinimum(1)

        self.spinLayout = QHBoxLayout(self)
        self.spinLayout.addWidget(self.textLabel)
        self.spinLayout.addWidget(self.spinBox)

        self.viewLayout.addWidget(self.titleLabel)
        self.viewLayout.addLayout(self.spinLayout)


class DateCopyMessageBox(MessageBoxBase):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.titleLabel = SubtitleLabel('复制 (从时间戳)', self)
        self.fromDate = ZhDatePicker(self)
        self.toDate = ZhDatePicker(self)
        self.toDate.setDate(QDate.currentDate())
        self.fromLabel = BodyLabel('从: ', self)
        self.toLabel = BodyLabel('到: ', self)

        self.fromLayout = QHBoxLayout(self)
        self.toLayout = QHBoxLayout(self)

        self.fromLayout.addWidget(self.fromLabel)
        self.fromLayout.addWidget(self.fromDate)
        self.toLayout.addWidget(self.toLabel)
        self.toLayout.addWidget(self.toDate)
        self.viewLayout.addWidget(self.titleLabel)
        self.viewLayout.addLayout(self.fromLayout)
        self.viewLayout.addLayout(self.toLayout)


class OptionInterface(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent=parent)

        self.backBtn = TransparentPushButton(FIF.RETURN, '返回', self)
        self.backBtn.setFixedWidth(80)
        self.slectAll = CheckBox('全选', self)
        self.slectAll.setTristate(True)
        self.slectAll.setCheckState(Qt.Checked)
        self.slectAll.stateChanged.connect(self.onSlectAll)

        self.yuwen = CheckBox('语文', self)
        self.shuxue = CheckBox('数学', self)
        self.yingyu = CheckBox('英语', self)
        self.wuli = CheckBox('物理', self)
        self.huaxue = CheckBox('化学', self)
        self.shengwu = CheckBox('生物', self)
        self.zhengzhi = CheckBox('政治', self)
        self.lishi = CheckBox('历史', self)
        self.dili = CheckBox('地理', self)
        self.jishu = CheckBox('技术', self)
        self.ziliao = CheckBox('资料', self)
        self.yuwen.setChecked(True)
        self.shuxue.setChecked(True)
        self.yingyu.setChecked(True)
        self.wuli.setChecked(True)
        self.huaxue.setChecked(True)
        self.shengwu.setChecked(True)
        self.zhengzhi.setChecked(True)
        self.lishi.setChecked(True)
        self.dili.setChecked(True)
        self.jishu.setChecked(True)
        self.ziliao.setChecked(True)
        self.yuwen.stateChanged.connect(self.updateSlectAll)
        self.shuxue.stateChanged.connect(self.updateSlectAll)
        self.yingyu.stateChanged.connect(self.updateSlectAll)
        self.wuli.stateChanged.connect(self.updateSlectAll)
        self.huaxue.stateChanged.connect(self.updateSlectAll)
        self.shengwu.stateChanged.connect(self.updateSlectAll)
        self.zhengzhi.stateChanged.connect(self.updateSlectAll)
        self.lishi.stateChanged.connect(self.updateSlectAll)
        self.dili.stateChanged.connect(self.updateSlectAll)
        self.jishu.stateChanged.connect(self.updateSlectAll)
        self.ziliao.stateChanged.connect(self.updateSlectAll)

        self.syncAction = QAction(FIF.SYNC.icon(), '同步 (默认)')
        self.lowSyncAction = QAction(FIF.LEAF.icon(), '同步 (低占用)')
        self.latelyCopyAction = QAction(FIF.HISTORY.icon(), '复制 (最近文件)')
        self.dateCopyAction = QAction(FIF.DATE_TIME.icon(), '复制 (从时间戳)')
        self.exeMenu = RoundMenu(parent=self)
        self.exeMenu.addAction(self.syncAction)
        self.exeMenu.addAction(self.lowSyncAction)
        self.exeMenu.addAction(self.latelyCopyAction)
        self.exeMenu.addAction(self.dateCopyAction)

        self.exeBtn = PrimarySplitPushButton('    执行同步    ', self)
        self.exeBtn.setFlyout(self.exeMenu)
        self.exitBtn = PushButton('取消', self)
        self.exeBtn.setFixedWidth(145)
        self.exitBtn.setFixedWidth(145)

        self.exeBtn.clicked.connect(lambda: self.onSyncAction("/speed=full", False, '1'))
        self.syncAction.triggered.connect(lambda: self.onSyncAction("/speed=full", False, '1'))
        self.lowSyncAction.triggered.connect(lambda: self.onSyncAction("/low_io", False, '2'))
        self.latelyCopyAction.triggered.connect(self.onLatelyCopyAction)
        self.dateCopyAction.triggered.connect(self.onDateCopyAction)

        self.mainLayoout = QVBoxLayout(self)
        self.mainLayoout.setContentsMargins(0, 0, 0, 5)

        self.checkLayout = QGridLayout(self)
        self.checkLayout.setHorizontalSpacing(40)
        self.checkLayout.setVerticalSpacing(12)
        self.checkLayout.setContentsMargins(30, 0, 10, 20)
        self.checkLayout.addWidget(self.backBtn, 0, 0, 1, 1)
        self.checkLayout.addWidget(self.slectAll, 0, 2, 1, 1)
        self.checkLayout.addWidget(self.yuwen, 1, 0, 1, 1)
        self.checkLayout.addWidget(self.shuxue, 1, 1, 1, 1)
        self.checkLayout.addWidget(self.yingyu, 1, 2, 1, 1)
        self.checkLayout.addWidget(self.wuli, 2, 0, 1, 1)
        self.checkLayout.addWidget(self.huaxue, 2, 1, 1, 1)
        self.checkLayout.addWidget(self.shengwu, 2, 2, 1, 1)
        self.checkLayout.addWidget(self.zhengzhi, 3, 0, 1, 1)
        self.checkLayout.addWidget(self.lishi, 3, 1, 1, 1)
        self.checkLayout.addWidget(self.dili, 3, 2, 1, 1)
        self.checkLayout.addWidget(self.jishu, 4, 0, 1, 1)
        self.checkLayout.addWidget(self.ziliao, 4, 1, 1, 1)

        self.btnLayout = QHBoxLayout(self)
        self.btnLayout.setContentsMargins(20, 0, 20, 0)
        self.btnLayout.addWidget(self.exeBtn)
        self.btnLayout.addWidget(self.exitBtn)

        self.mainLayoout.addLayout(self.checkLayout)
        self.mainLayoout.addLayout(self.btnLayout)

        self.aniGroup = QParallelAnimationGroup(self)
        self.addAniWidget(self.backBtn, self.backBtn.geometry())
        self.addAniWidget(self.slectAll, self.slectAll.geometry())
        self.addAniWidget(self.yuwen, self.yuwen.geometry())
        self.addAniWidget(self.shuxue, self.shuxue.geometry())
        self.addAniWidget(self.yingyu, self.yingyu.geometry())
        self.addAniWidget(self.wuli, self.wuli.geometry())
        self.addAniWidget(self.huaxue, self.huaxue.geometry())
        self.addAniWidget(self.shengwu, self.shengwu.geometry())
        self.addAniWidget(self.zhengzhi, self.zhengzhi.geometry())
        self.addAniWidget(self.lishi, self.lishi.geometry())
        self.addAniWidget(self.dili, self.dili.geometry())
        self.addAniWidget(self.jishu, self.jishu.geometry())
        self.addAniWidget(self.ziliao, self.ziliao.geometry())
        self.addAniWidget(self.exeBtn, self.exeBtn.geometry())
        self.addAniWidget(self.exitBtn, self.exitBtn.geometry())

    def addAniWidget(self, w, position):
        ani = QPropertyAnimation(w, b'geometry')
        ani.setStartValue(position)
        ani.setDuration(200)
        ani.setEasingCurve(QEasingCurve.Linear)
        self.aniGroup.addAnimation(ani)

    def num(self):
        numTemp = 0
        if self.yuwen.isChecked():
            numTemp += 1
        if self.shuxue.isChecked():
            numTemp += 1
        if self.yingyu.isChecked():
            numTemp += 1
        if self.wuli.isChecked():
            numTemp += 1
        if self.huaxue.isChecked():
            numTemp += 1
        if self.shengwu.isChecked():
            numTemp += 1
        if self.zhengzhi.isChecked():
            numTemp += 1
        if self.lishi.isChecked():
            numTemp += 1
        if self.dili.isChecked():
            numTemp += 1
        if self.jishu.isChecked():
            numTemp += 1
        if self.ziliao.isChecked():
            numTemp += 1
        return numTemp

    def onSlectAll(self):
        if self.slectAll.checkState() == Qt.Checked:
            self.yuwen.setChecked(True)
            self.shuxue.setChecked(True)
            self.yingyu.setChecked(True)
            self.wuli.setChecked(True)
            self.huaxue.setChecked(True)
            self.shengwu.setChecked(True)
            self.zhengzhi.setChecked(True)
            self.lishi.setChecked(True)
            self.dili.setChecked(True)
            self.jishu.setChecked(True)
            self.ziliao.setChecked(True)
        elif self.slectAll.checkState() == Qt.Unchecked:
            self.yuwen.setChecked(False)
            self.shuxue.setChecked(False)
            self.yingyu.setChecked(False)
            self.wuli.setChecked(False)
            self.huaxue.setChecked(False)
            self.shengwu.setChecked(False)
            self.zhengzhi.setChecked(False)
            self.lishi.setChecked(False)
            self.dili.setChecked(False)
            self.jishu.setChecked(False)
            self.ziliao.setChecked(False)
        elif self.slectAll.checkState() == Qt.PartiallyChecked:
            if self.num() == 0:
                self.slectAll.setCheckState(Qt.Checked)
                self.yuwen.setChecked(True)
                self.shuxue.setChecked(True)
                self.yingyu.setChecked(True)
                self.wuli.setChecked(True)
                self.huaxue.setChecked(True)
                self.shengwu.setChecked(True)
                self.zhengzhi.setChecked(True)
                self.lishi.setChecked(True)
                self.dili.setChecked(True)
                self.jishu.setChecked(True)
                self.ziliao.setChecked(True)
            elif self.num() == 11:
                self.slectAll.setCheckState(Qt.Unchecked)
                self.yuwen.setChecked(False)
                self.shuxue.setChecked(False)
                self.yingyu.setChecked(False)
                self.wuli.setChecked(False)
                self.huaxue.setChecked(False)
                self.shengwu.setChecked(False)
                self.zhengzhi.setChecked(False)
                self.lishi.setChecked(False)
                self.dili.setChecked(False)
                self.jishu.setChecked(False)
                self.ziliao.setChecked(False)

    def updateSlectAll(self):
        if self.num() == 11:
            self.slectAll.setCheckState(Qt.Checked)
        elif self.num() == 0:
            self.slectAll.setCheckState(Qt.Unchecked)
        else:
            self.slectAll.setCheckState(Qt.PartiallyChecked)

    def onSyncAction(self, commandOption, isDelete, mode):
        self.isClicked = True

        arg = ["PrestoMain.exe"]
        arg.append(drive)

        if self.yuwen.isChecked():
            arg.append("1")
        else:
            arg.append("0")
        if self.shuxue.isChecked():
            arg.append("1")
        else:
            arg.append("0")
        if self.yingyu.isChecked():
            arg.append("1")
        else:
            arg.append("0")
        if self.wuli.isChecked():
            arg.append("1")
        else:
            arg.append("0")
        if self.huaxue.isChecked():
            arg.append("1")
        else:
            arg.append("0")
        if self.shengwu.isChecked():
            arg.append("1")
        else:
            arg.append("0")
        if self.zhengzhi.isChecked():
            arg.append("1")
        else:
            arg.append("0")
        if self.lishi.isChecked():
            arg.append("1")
        else:
            arg.append("0")
        if self.dili.isChecked():
            arg.append("1")
        else:
            arg.append("0")
        if self.jishu.isChecked():
            arg.append("1")
        else:
            arg.append("0")
        if self.ziliao.isChecked():
            arg.append("1")
        else:
            arg.append("0")

        arg.append(mode)
        arg.append(str(isDelete))
        arg.append(commandOption)
        subprocess.Popen(arg, shell=True)
        sys.exit()

    def onLatelyCopyAction(self):
        w = LatelyCopyMessageBox(self.window())
        if w.exec():
            self.onSyncAction(f"/from_date=-{w.spinBox.value()}D", w.isDelete, '3')
            sys.exit()

    def onDateCopyAction(self):
        w = DateCopyMessageBox(self.window())
        if w.exec():
            if w.fromDate.date.toString('yyyyMMdd') == '':
                w.fromDate.setDate(QDate.currentDate())
            if w.toDate.date > QDate.currentDate():
                w.toDate.setDate(QDate.currentDate())
            if w.fromDate.date > QDate.currentDate():
                w.toDate.setDate(QDate.currentDate())
            if w.fromDate.date > w.toDate.date:
                w.fromDate.setDate(w.toDate.date)

            self.onSyncAction(f"/from_date={w.fromDate.date.toString('yyyyMMdd')} /to_date={w.toDate.date.toString('yyyyMMdd')}", w.isDelete, '4')
            sys.exit()


class AskInterface(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent=parent)
        self.mainLayout = QVBoxLayout(self)
        self.mainLayout.setContentsMargins(0, 0, 0, 0)
        self.btnLayout = QHBoxLayout(self)
        self.infoLayout = QHBoxLayout(self)
        self.infoLabel = SubtitleLabel(self.getDriveName() + ' (' + drive + ')')
        self.infoBtn = TransparentToolButton(FIF.INFO, self)
        self.infoBtn.setToolTip('驱动器信息')
        self.infoBtn.installEventFilter(ToolTipFilter(self.infoBtn, 0, ToolTipPosition.BOTTOM))
        self.syncBtn = PrimaryPushButton(FIF.SYNC, '同步', self)
        self.openBtn = PushButton(FIF.FOLDER, '打开', self)
        self.btnLayout.addWidget(self.syncBtn)
        self.btnLayout.addWidget(self.openBtn)
        self.btnLayout.setContentsMargins(20, 10, 20, 10)
        self.infoLayout.setContentsMargins(20, 0, 20, 10)
        self.infoLayout.addWidget(self.infoLabel)
        self.infoLayout.addStretch(1)
        self.infoLayout.addWidget(self.infoBtn)
        self.mainLayout.addStretch()
        self.mainLayout.addLayout(self.infoLayout)
        self.mainLayout.addLayout(self.btnLayout)

    def getDriveName(self):
        try:
            if GetVolumeInformation(drive)[0] != '':
                return GetVolumeInformation(drive)[0]
            else:
                return "U盘"
        except:
            sys.exit()


class ProfileCard(QWidget):
    def __init__(self, ratio: int, name: str, size: str, parent=None):
        super().__init__(parent=parent)
        ratio = 1 if ratio == 0 else ratio

        self.ring = ProgressRing(self)
        self.nameLabel = BodyLabel(name, self)
        self.sizeLabel = CaptionLabel(size, self)

        color = QColor(206, 206, 206) if isDarkTheme() else QColor(96, 96, 96)
        self.sizeLabel.setStyleSheet('QLabel{color: ' + color.name() + '}')
        color = QColor(255, 255, 255) if isDarkTheme() else QColor(0, 0, 0)
        self.nameLabel.setStyleSheet('QLabel{color: ' + color.name() + '}')

        self.setFixedSize(280, 70)
        self.ring.setFixedSize(48, 48)
        self.ring.setValue(ratio)
        self.ring.move(2, 6)
        self.nameLabel.move(64, 13)
        self.sizeLabel.move(64, 32)


class MainWindow(MicaWindow):

    def __init__(self):
        super().__init__()
        self.isClicked = False
        self.opacity = 0.98
        setThemeColor(QColor(113, 89, 249))
        self.setWindowTitle("Presto")
        self.setWindowIcon(QIcon(':/icon.png'))
        self.setWindowOpacity(self.opacity)
        self.desktop = QApplication.screens()[0].size()
        self.setFixedSize(360, 145)
        self.move(self.desktop.width() - self.width() - 20, self.desktop.height() - self.height() - 60)

        QTimer.singleShot(10000, self.timeOut)

        self.pivot = Pivot(self)
        self.stackedWidget = QStackedWidget(self)
        self.askInterface = AskInterface(self)
        self.optionInterface = OptionInterface(self)
        self.addSubInterface(self.askInterface, 'askInterface', 'Ask')
        self.addSubInterface(self.optionInterface, 'optionInterface', 'Option')
        self.pivot.setVisible(False)

        self.askInterface.infoBtn.clicked.connect(self.infoBtnOn)
        self.askInterface.syncBtn.clicked.connect(self.onSyncBtn)
        self.askInterface.openBtn.clicked.connect(self.onOpenBtn)
        self.optionInterface.backBtn.clicked.connect(self.onBackBtn)
        self.optionInterface.exitBtn.clicked.connect(sys.exit)

        self.vBoxLayout = QVBoxLayout(self)
        self.vBoxLayout.addWidget(self.titleBar)
        self.vBoxLayout.addWidget(self.stackedWidget)
        self.vBoxLayout.setContentsMargins(0, 0, 0, 0)
        self.stackedWidget.setCurrentWidget(self.askInterface)
        self.pivot.setCurrentItem(self.askInterface.objectName())
        self.pivot.currentItemChanged.connect(lambda k: self.stackedWidget.setCurrentWidget(self.findChild(QWidget, k)))

        self.opacityAni = QPropertyAnimation(self, b'windowOpacity', self)
        self.opacityAni.setDuration(1000)
        self.opacityAni.setStartValue(1)
        self.opacityAni.setEndValue(0)
        self.opacityAni.finished.connect(self.Quit)

    def addSubInterface(self, widget: QLabel, objectName, text):
        widget.setObjectName(objectName)
        self.stackedWidget.addWidget(widget)
        self.pivot.addItem(routeKey=objectName, text=text)

    def timeOut(self):
        if not self.isClicked:
            self.opacityAni.start()

    def Quit(self):
        self.close()
        sys.exit()

    def getDriveName(self):
        try:
            if GetVolumeInformation(drive)[0] != '':
                return GetVolumeInformation(drive)[0]
            else:
                return "U盘"
        except:
            sys.exit()

    def getDriveSize(self):
        sectorsPerCluster, bytesPerSector, numFreeClusters, totalNumClusters = GetDiskFreeSpace(drive)
        freeSpace = format((numFreeClusters * sectorsPerCluster * bytesPerSector) / 1024 / 1024 / 1024, '.1f')
        totalSpace = format((sectorsPerCluster * bytesPerSector * totalNumClusters) / 1024 / 1024 / 1024, '.1f')
        return freeSpace + 'GB可用，共' + totalSpace + 'GB'

    def getDriveRatio(self):
        sectorsPerCluster, bytesPerSector, numFreeClusters, totalNumClusters = GetDiskFreeSpace(drive)
        freeSpace = numFreeClusters * sectorsPerCluster * bytesPerSector
        totalSpace = sectorsPerCluster * bytesPerSector * totalNumClusters
        return int((1 - freeSpace / totalSpace) * 100)

    def infoBtnOn(self):
        self.isClicked = True
        menu = RoundMenu(parent=self)
        card = ProfileCard(self.getDriveRatio(), self.getDriveName() + ' (' + drive + ')', self.getDriveSize(), menu)
        menu.addWidget(card, selectable=False)
        menu.addSeparator()
        settingAction = Action(FIF.SETTING, '设置')
        settingAction.triggered.connect(lambda: subprocess.Popen("PrestoSetting.exe", shell=True))
        menu.addAction(settingAction)
        helpAction = Action(FIF.HELP, '帮助')
        helpAction.triggered.connect(self.onHelpAction)
        menu.addAction(helpAction)
        menu.addSeparator()
        menu.addAction(Action(FIF.CLOSE, '关闭'))
        menu.exec(QPoint(self.x() + self.askInterface.infoBtn.x() - 315, self.y() + self.askInterface.infoBtn.y() - 55))

    def onHelpAction(self):
        if os.path.exists(os.path.abspath("./Doc/PrestoHelp.html")):
            os.startfile(os.path.abspath("./Doc/PrestoHelp.html"))
        else:
            WebOpen("https://sudo0015.github.io/post/Presto%20-bang-zhu.html")

    def onOpenBtn(self):
        os.startfile(drive)
        self.close()
        sys.exit()

    def onSyncBtn(self):
        self.isClicked = True
        self.stackedWidget.setCurrentWidget(self.optionInterface)
        self.pivot.setCurrentItem(self.optionInterface.objectName())
        self.pivot.currentItemChanged.connect(lambda k: self.stackedWidget.setCurrentWidget(self.findChild(QWidget, k)))
        self.setFixedSize(360, 300)
        self.move(self.desktop.width() - self.width() - 20, self.desktop.height() - self.height() - 60)
        self.optionInterface.aniGroup.start()

    def onBackBtn(self):
        self.isClicked = True
        self.stackedWidget.setCurrentWidget(self.askInterface)
        self.pivot.setCurrentItem(self.askInterface.objectName())
        self.pivot.currentItemChanged.connect(lambda k: self.stackedWidget.setCurrentWidget(self.findChild(QWidget, k)))
        self.setFixedSize(360, 145)
        self.move(self.desktop.width() - self.width() - 20, self.desktop.height() - self.height() - 60)


if __name__ == '__main__':
    try:
        if len(sys.argv) != 2:
            sys.exit()
        drive = sys.argv[1]
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
