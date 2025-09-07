# -*- coding: utf-8 -*-

import os
from enum import Enum
from qfluentwidgets import qconfig, QConfig, ConfigItem, OptionsConfigItem, BoolValidator, OptionsValidator, \
    FolderValidator, RangeConfigItem, RangeValidator, EnumSerializer


class BufSize(Enum):
    _32 = "32 MB"
    _64 = "64 MB"
    _128 = "128 MB"
    _256 = "256 MB"
    _512 = "512 MB"
    _1024 = "1 GB"


class Config(QConfig):
    AutoRun = ConfigItem("MainWindow", "AutoRun", True, BoolValidator())
    Notify = ConfigItem("MainWindow", "Notify", True, BoolValidator())
    IsSourceCloud = OptionsConfigItem("MainWindow", "IsSourceCloud", True, BoolValidator())

    sourceFolder = ConfigItem("Folders", "SourceFolder", "", FolderValidator())
    yuwenFolder = ConfigItem("Folders", "Yuwen", "", FolderValidator())
    shuxueFolder = ConfigItem("Folders", "Shuxue", "", FolderValidator())
    yingyuFolder = ConfigItem("Folders", "Yingyu", "", FolderValidator())
    wuliFolder = ConfigItem("Folders", "Wuli", "", FolderValidator())
    huaxueFolder = ConfigItem("Folders", "Huaxue", "", FolderValidator())
    shengwuFolder = ConfigItem("Folders", "Shengwu", "", FolderValidator())
    zhengzhiFolder = ConfigItem("Folders", "Zhengzhi", "", FolderValidator())
    lishiFolder = ConfigItem("Folders", "Lishi", "", FolderValidator())
    diliFolder = ConfigItem("Folders", "Dili", "", FolderValidator())
    jishuFolder = ConfigItem("Folders", "Jishu", "", FolderValidator())
    ziliaoFolder = ConfigItem("Folders", "Ziliao", "", FolderValidator())

    ScanCycle = RangeConfigItem("MainWindow", "ScanCycle", 10, RangeValidator(1, 50))
    ConcurrentProcess = ConfigItem("MainWindow", "ConcurrentProcess", 3, RangeValidator(1, 5))
    BufSize = OptionsConfigItem("MainWindow", "BufSize", BufSize._256, OptionsValidator(BufSize), EnumSerializer(BufSize))
    dpiScale = OptionsConfigItem("MainWindow", "DpiScale", "Auto", OptionsValidator([1, 1.25, 1.5, 1.75, 2, "Auto"]), restart=True)


YEAR = "2025"
VERSION = "7.1.0"
cfg = Config()
qconfig.load(os.path.join(os.path.expanduser('~'), '.Presto', 'config', 'config.json'), cfg)
