import sys, json
import pandas as pd
from os import path
from time import strftime, localtime
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import (QApplication, QMainWindow, QGridLayout, QHBoxLayout, QVBoxLayout, QFormLayout, QGroupBox,  QScrollArea, QTabWidget, QDialog, QFileDialog, 
                             QMessageBox, QWidget, QLabel, QPushButton, QDialogButtonBox, QSpinBox, QDoubleSpinBox, QLineEdit, QComboBox, QCheckBox)
from pyMEP import Quantity
from pyMEP.substance import *
from pyMEP.hvac.climatic import WeatherData, ReferenceDates
from pyMEP.hvac.coolingload import Setting, RTS
from pyMEP.hvac.lighting import lpd_df, LightingPowerDensities
from pyMEP.hvac.equipment import EquipmentPowerDensities
from pyMEP.hvac.people import human_vrp_df, human_hr_df, HumanHeatRate
from pyMEP.hvac.external_heat_gains import Wall
from pyMEP.charts.chart_2D import LineChart

from lib.CollapsibleBox import CollapsibleBox
from lib.BuildingGraphic import SideView, TopView
from lib.HourlyTable import numericHourlyTable, checkBoxHourlyTable
from lib._ThermalComfort import ComfortZone
from lib._resource import *
from lib.utils import *

__version__ = '0.1.0'

Q_ = Quantity
WATT_BTU_HR = 3.41214163

class Mainwindow(QMainWindow):
    def __init__(self):
        super(Mainwindow, self).__init__()
        home = path.dirname(__file__)
        self.setWindowIcon(QIcon(path.join(home, "res\\G17-2906204.ico")))
        title = "Radiant Time Series (RTS) method Cooling Load"
        title += f' (Build {strftime("%Y-%m-%d", localtime(path.getmtime(__file__)))})'
        self.setWindowTitle(title)

        pyqtMEPStylesheet = path.join(home, "res\\pyqtMEP.qss")
        loadStylesheet(pyqtMEPStylesheet)

        self.widget = QWidget(self)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(self.widget)
        rtsmGrid = QVBoxLayout(self.widget)
        rtsmGrid.setAlignment(Qt.AlignmentFlag.AlignTop)

        # region Climatic Desing Data
        climaticBox = CollapsibleBox("Climatic Design Data")
        rtsmGrid.addWidget(climaticBox)

        climaticGrid = QHBoxLayout()
        climaticGrid.setContentsMargins(0, 0, 0, 0)

        outsideGroup = QGroupBox('Outside Design Condition')
        insideGroup = QGroupBox('Inside Design Condition')
        climaticGrid.addWidget(outsideGroup, 1)
        climaticGrid.addWidget(insideGroup, 1)

        # region Outside
        outsidelayout = QGridLayout(outsideGroup)
        outsidelayout.setAlignment(Qt.AlignmentFlag.AlignTop)
        outsidelayout.setVerticalSpacing(0)
        outsidelayout.setColumnStretch(0, 3)
        outsidelayout.setColumnStretch(1, 2)
        outsidelayout.setColumnStretch(2, 1)
        outsidelayout.setColumnStretch(3, 2)

        outsidelayout.addWidget(QLabel('<b><div style="color: red">Location</div></b>'), 0, 0, Qt.AlignmentFlag.AlignRight)
        self.cmbRTSMLocation = QComboBox()
        self.cmbRTSMLocation.setFixedHeight(22)
        self.cmbRTSMLocation.addItems(locations.keys())
        outsidelayout.addWidget(self.cmbRTSMLocation, 0, 1, 1, 3)

        outsidelayout.addWidget(QLabel('Latitude'), 1, 0, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMLatitude = QDoubleSpinBox()
        self.tbRTSMLatitude.setDecimals(4)
        self.tbRTSMLatitude.setRange(-180, 180)
        outsidelayout.addWidget(self.tbRTSMLatitude, 1, 1)
        outsidelayout.addWidget(QLabel('(+N, -S)'), 1, 2, 1, 2)

        outsidelayout.addWidget(QLabel('Longtitude'), 2, 0, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMLongtitude = QDoubleSpinBox()
        self.tbRTSMLongtitude.setDecimals(4)
        self.tbRTSMLongtitude.setRange(-180, 180)
        outsidelayout.addWidget(self.tbRTSMLongtitude, 2, 1)
        outsidelayout.addWidget(QLabel('(+E, -W)'), 2, 2, 1, 2)

        outsidelayout.addWidget(QLabel('Altitude'), 3, 0, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMAltitude = QSpinBox()
        self.tbRTSMAltitude.setMaximum(1000)
        outsidelayout.addWidget(self.tbRTSMAltitude, 3, 1)
        outsidelayout.addWidget(QLabel('m (+Sea Level)'), 3, 2, 1, 2)

        outsidelayout.addWidget(QLabel('Outside DB'), 4, 0, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMOutsideDB = QDoubleSpinBox()
        outsidelayout.addWidget(self.tbRTSMOutsideDB, 4, 1)
        outsidelayout.addWidget(QLabel('°C'), 4, 2)

        outsidelayout.addWidget(QLabel('DB Range'), 5, 0, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMDBRange = QDoubleSpinBox()
        outsidelayout.addWidget(self.tbRTSMDBRange, 5, 1)
        outsidelayout.addWidget(QLabel('°C'), 5, 2)

        outsidelayout.addWidget(QLabel('Outside WB'), 6, 0, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMOutsideWB = QDoubleSpinBox()
        outsidelayout.addWidget(self.tbRTSMOutsideWB, 6, 1)
        outsidelayout.addWidget(QLabel('°C'), 6, 2)
        self.btnRTSM_DesignTempProfile = QPushButton('Profile')
        self.btnRTSM_DesignTempProfile.setFixedHeight(22)
        outsidelayout.addWidget(self.btnRTSM_DesignTempProfile, 6, 3)

        outsidelayout.addWidget(QLabel('taub'), 7, 0, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMtaub = QDoubleSpinBox()
        self.tbRTSMtaub.setDecimals(4)
        outsidelayout.addWidget(self.tbRTSMtaub, 7, 1)
        outsidelayout.addWidget(QLabel('taud'), 7, 2, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMtaud = QDoubleSpinBox()
        self.tbRTSMtaud.setDecimals(4)
        outsidelayout.addWidget(self.tbRTSMtaud, 7, 3)

        outsidelayout.addWidget(QLabel('UTC'), 8, 0, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMtz = QLineEdit()
        outsidelayout.addWidget(self.tbRTSMtz, 8, 1)
        ie = QLabel()
        ie.setAlignment(Qt.AlignmentFlag.AlignCenter)
        ie.setOpenExternalLinks(True)
        ie.setText("""<a href="https://ashrae-meteo.info/v2.0/places.php?continent=Asia"><img src="res/ASHRAE.png" width="40"></a>""")
        outsidelayout.addWidget(ie, 8, 2, 2, 2)

        outsidelayout.addWidget(QLabel('Design Month'), 9, 0, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMMonth = QLineEdit()
        self.tbRTSMMonth.setFixedHeight(22)
        outsidelayout.addWidget(self.tbRTSMMonth, 9, 1)
		# endregion

        # region Inside
        insidelayout = QGridLayout(insideGroup)
        insidelayout.setAlignment(Qt.AlignmentFlag.AlignTop)
        insidelayout.setVerticalSpacing(1)

        insidelayout.addWidget(QLabel('Name'), 0, 0, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMName = QLineEdit('Room1')
        self.tbRTSMName.setFixedHeight(24)
        insidelayout.addWidget(self.tbRTSMName, 0, 1, 1, 2)

        insidelayout.addWidget(QLabel('<div style="color: red">Inside DB'), 1, 0, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMInsideDB = QSpinBox()
        self.tbRTSMInsideDB.setValue(25)
        insidelayout.addWidget(self.tbRTSMInsideDB, 1, 1)
        insidelayout.addWidget(QLabel('°C'), 1, 2)

        insidelayout.addWidget(QLabel('<div style="color: red">Inside RH'), 2, 0, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMInsideRH = QSpinBox()
        self.tbRTSMInsideRH.setValue(55)
        insidelayout.addWidget(self.tbRTSMInsideRH, 2, 1)
        insidelayout.addWidget(QLabel('%'), 2, 2)

        insidelayout.addWidget(QLabel('<div style="color: red">Floor'), 3, 0, Qt.AlignmentFlag.AlignRight)
        self.cmbRTSMSpaceType = QComboBox()
        self.cmbRTSMSpaceType.addItems(space_types)
        insidelayout.addWidget(self.cmbRTSMSpaceType, 3, 1)

        self.spaceTypeGraphic = SideView(120, 115)
        self.spaceTypeGraphic.setContentsMargins(0,5,3,3)
        insidelayout.addWidget(self.spaceTypeGraphic, 4, 1, 1, 2, Qt.AlignmentFlag.AlignLeft)
		# endregion

        climaticBox.setContentLayout(climaticGrid)
		# endregion

        # region Architec Design Data
        architecBox = CollapsibleBox("Architec Design Data")
        rtsmGrid.addWidget(architecBox)
        architecGrid = QHBoxLayout()

		# region Letf Layout
        architecLeftlayout = QGridLayout()
        architecLeftlayout.setAlignment(Qt.AlignmentFlag.AlignTop)
        architecLeftlayout.setVerticalSpacing(0)
        architecLeftlayout.setColumnStretch(0, 3)
        architecLeftlayout.setColumnStretch(1, 3)
        architecLeftlayout.setColumnStretch(2, 1)
        architecLeftlayout.setColumnStretch(3, 3)
        architecLeftlayout.setColumnStretch(4, 3)

        architecLeftlayout.addWidget(QLabel('<div style="color: red">Orientation (from North)</div>'), 0, 0, 1, 3, Qt.AlignmentFlag.AlignRight)
        self.tbCompass = QSpinBox()
        self.tbCompass.setRange(-180, 180)
        architecLeftlayout.addWidget(self.tbCompass, 0, 3)
        architecLeftlayout.addWidget(QLabel('Degree'), 0, 4)

        architecLeftlayout.addWidget(QLabel('<div style="color: red">Room Height</div>'), 1, 1, 1, 2, Qt.AlignmentFlag.AlignRight)
        self.tbZoneHeight = QDoubleSpinBox()
        self.tbZoneHeight.setSingleStep(0.10)
        architecLeftlayout.addWidget(self.tbZoneHeight, 1, 3)
        architecLeftlayout.addWidget(QLabel('m'), 1, 4)

        architecLeftlayout.addWidget(QLabel('<div style="color: red">Length</div>'), 2, 1, 1, 2, Qt.AlignmentFlag.AlignRight)
        self.tbZoneLength = QDoubleSpinBox()
        self.tbZoneLength.setSingleStep(0.10)
        architecLeftlayout.addWidget(self.tbZoneLength, 2, 3)
        architecLeftlayout.addWidget(QLabel('m'), 2, 4)

        architecLeftlayout.addWidget(QLabel('<div style="color: red">Width</div>'), 3, 1, 1, 2, Qt.AlignmentFlag.AlignRight)
        self.tbZoneWidth = QDoubleSpinBox()
        self.tbZoneWidth.setSingleStep(0.10)
        architecLeftlayout.addWidget(self.tbZoneWidth, 3, 3)
        architecLeftlayout.addWidget(QLabel('m'), 3, 4)

		# region Roof & Floor
        architecLeftlayout.addWidget(QLabel('Roof'), 4, 0, Qt.AlignmentFlag.AlignRight)
        self.cmbRTSMRoofType = QComboBox()
        self.cmbRTSMRoofType.addItems(roofs.keys())
        self.setRTSMItemToolTip(self.cmbRTSMRoofType, roofs)
        architecLeftlayout.addWidget(self.cmbRTSMRoofType, 4, 1)
        architecLeftlayout.addWidget(QLabel('<b>U</>'), 4, 2, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMURoof = QDoubleSpinBox()
        self.tbRTSMURoof.setDecimals(3)
        architecLeftlayout.addWidget(self.tbRTSMURoof, 4, 3)
        architecLeftlayout.addWidget(QLabel('W/m² K'), 4, 4)

        architecLeftlayout.addWidget(QLabel('Floor'), 5, 0, Qt.AlignmentFlag.AlignRight)
        self.cmbRTSMFloorType = QComboBox()
        self.cmbRTSMFloorType.addItems(floors.keys())
        self.setRTSMItemToolTip(self.cmbRTSMFloorType, floors)
        architecLeftlayout.addWidget(self.cmbRTSMFloorType, 5, 1)
        architecLeftlayout.addWidget(QLabel('<b>U</>'), 5, 2, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMUFloor = QDoubleSpinBox()
        self.tbRTSMUFloor.setDecimals(3)
        architecLeftlayout.addWidget(self.tbRTSMUFloor, 5, 3)
        architecLeftlayout.addWidget(QLabel('W/m² K'), 5, 4)

        architecLeftlayout.addWidget(QLabel('Temperature Diff.'), 6, 1, 1, 2, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMOptionFloor = QDoubleSpinBox()
        self.tbRTSMOptionFloor.setValue(6.7)
        architecLeftlayout.addWidget(self.tbRTSMOptionFloor, 6, 3)
        architecLeftlayout.addWidget(QLabel('°C'), 6, 4)
		# endregion

		# region Wall A
        architecLeftlayout.addWidget(QLabel('Wall-A'), 7, 0, Qt.AlignmentFlag.AlignRight)
        self.cmbRTSMWall_A = QComboBox()
        self.cmbRTSMWall_A.addItems(walls.keys())
        self.setRTSMItemToolTip(self.cmbRTSMWall_A, walls)
        architecLeftlayout.addWidget(self.cmbRTSMWall_A, 7, 1)
        architecLeftlayout.addWidget(QLabel('<b>U</>'), 7, 2, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMUWall_A = QDoubleSpinBox()
        self.tbRTSMUWall_A.setDecimals(3)
        architecLeftlayout.addWidget(self.tbRTSMUWall_A, 7, 3)
        architecLeftlayout.addWidget(QLabel('W/m² K'), 7, 4)

        self.lbOptionWall_A = QLabel('Solar Absorptance')
        architecLeftlayout.addWidget(self.lbOptionWall_A, 8, 1, 1, 2, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMOptionWall_A = QDoubleSpinBox()
        self.tbRTSMOptionWall_A.setDecimals(3)
        self.tbRTSMOptionWall_A.setValue(0.052)
        architecLeftlayout.addWidget(self.tbRTSMOptionWall_A, 8, 3)
        self.lbOptionUnitWall_A = QLabel('°C')
        architecLeftlayout.addWidget(self.lbOptionUnitWall_A, 8, 4)
		# endregion

		# region Wall B
        architecLeftlayout.addWidget(QLabel('Wall-B'), 9, 0, Qt.AlignmentFlag.AlignRight)
        self.cmbRTSMWall_B = QComboBox()
        self.cmbRTSMWall_B.addItems(walls.keys())
        self.setRTSMItemToolTip(self.cmbRTSMWall_B, walls)
        architecLeftlayout.addWidget(self.cmbRTSMWall_B, 9, 1)
        architecLeftlayout.addWidget(QLabel('<b>U</>'), 9, 2, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMUWall_B = QDoubleSpinBox()
        self.tbRTSMUWall_B.setDecimals(3)
        architecLeftlayout.addWidget(self.tbRTSMUWall_B, 9, 3)
        architecLeftlayout.addWidget(QLabel('W/m² K'), 9, 4)

        self.lbOptionWall_B = QLabel('Solar Absorptance')
        architecLeftlayout.addWidget(self.lbOptionWall_B, 10, 1, 1, 2, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMOptionWall_B = QDoubleSpinBox()
        self.tbRTSMOptionWall_B.setDecimals(3)
        self.tbRTSMOptionWall_B.setValue(0.052)
        architecLeftlayout.addWidget(self.tbRTSMOptionWall_B, 10, 3)
        self.lbOptionUnitWall_B = QLabel('°C')
        architecLeftlayout.addWidget(self.lbOptionUnitWall_B, 10, 4)
		# endregion

		# region Wall C
        architecLeftlayout.addWidget(QLabel('Wall-C'), 11, 0, Qt.AlignmentFlag.AlignRight)
        self.cmbRTSMWall_C = QComboBox()
        self.cmbRTSMWall_C.addItems(walls.keys())
        self.setRTSMItemToolTip(self.cmbRTSMWall_C, walls)
        architecLeftlayout.addWidget(self.cmbRTSMWall_C, 11, 1)
        architecLeftlayout.addWidget(QLabel('<b>U</>'), 11, 2, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMUWall_C = QDoubleSpinBox()
        self.tbRTSMUWall_C.setDecimals(3)
        architecLeftlayout.addWidget(self.tbRTSMUWall_C, 11, 3)
        architecLeftlayout.addWidget(QLabel('W/m² K'), 11, 4)

        self.lbOptionWall_C = QLabel('Solar Absorptance')
        architecLeftlayout.addWidget(self.lbOptionWall_C, 12, 1, 1, 2, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMOptionWall_C = QDoubleSpinBox()
        self.tbRTSMOptionWall_C.setDecimals(3)
        self.tbRTSMOptionWall_C.setValue(0.052)
        architecLeftlayout.addWidget(self.tbRTSMOptionWall_C, 12, 3)
        self.lbOptionUnitWall_C = QLabel('°C')
        architecLeftlayout.addWidget(self.lbOptionUnitWall_C, 12, 4)
		# endregion

		# region Wall D
        architecLeftlayout.addWidget(QLabel('Wall-D'), 13, 0, Qt.AlignmentFlag.AlignRight)
        self.cmbRTSMWall_D = QComboBox()
        self.cmbRTSMWall_D.addItems(walls.keys())
        self.setRTSMItemToolTip(self.cmbRTSMWall_D, walls)
        architecLeftlayout.addWidget(self.cmbRTSMWall_D, 13, 1)
        architecLeftlayout.addWidget(QLabel('<b>U</>'), 13, 2, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMUWall_D = QDoubleSpinBox()
        self.tbRTSMUWall_D.setDecimals(3)
        architecLeftlayout.addWidget(self.tbRTSMUWall_D, 13, 3)
        architecLeftlayout.addWidget(QLabel('W/m² K'), 13, 4)

        self.lbOptionWall_D = QLabel('Solar Absorptance')
        architecLeftlayout.addWidget(self.lbOptionWall_D, 14, 1, 1, 2, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMOptionWall_D = QDoubleSpinBox()
        self.tbRTSMOptionWall_D.setDecimals(3)
        self.tbRTSMOptionWall_D.setValue(0.052)
        architecLeftlayout.addWidget(self.tbRTSMOptionWall_D, 14, 3)
        self.lbOptionUnitWall_D = QLabel('°C')
        architecLeftlayout.addWidget(self.lbOptionUnitWall_D, 14, 4)
		# endregion
		# endregion

		# region Right Layout
        architecRightlayout = QVBoxLayout()
        architecRightlayout.setAlignment(Qt.AlignmentFlag.AlignTop)

        self.architecGraphic = TopView(295, 150)
        architecRightlayout.addWidget(self.architecGraphic)

        self.windowTab = QTabWidget()
        architecRightlayout.addWidget(self.windowTab)
        self.windowAPage, self.windowBPage, self.windowCPage, self.windowDPage = QWidget(), QWidget(), QWidget(), QWidget()
        self.windowTab.addTab(self.windowAPage, "Win-A")
        self.windowTab.addTab(self.windowBPage, "Win-B")
        self.windowTab.addTab(self.windowCPage, "Win-C")
        self.windowTab.addTab(self.windowDPage, "Win-D")

		# region Window A
        windowAlayout = QGridLayout()
        windowAlayout.setVerticalSpacing(1)
        windowAlayout.setAlignment(Qt.AlignmentFlag.AlignTop)
        windowAlayout.setColumnStretch(1, 2)

        windowAlayout.addWidget(QLabel('Height'), 0, 0, Qt.AlignmentFlag.AlignRight)
        self.tbWinAHigh = QDoubleSpinBox()
        self.tbWinAHigh.setSingleStep(0.1)
        self.tbWinAHigh.setValue(1.5)
        windowAlayout.addWidget(self.tbWinAHigh, 0, 1)
        windowAlayout.addWidget(QLabel('m'), 0, 2)
        windowAlayout.addWidget(QLabel('Type'), 0, 3, Qt.AlignmentFlag.AlignRight)
        self.cmbRTSMWin_A = QComboBox()
        self.cmbRTSMWin_A.addItems(windows.keys())
        self.setRTSMItemToolTip(self.cmbRTSMWin_A, windows)
        windowAlayout.addWidget(self.cmbRTSMWin_A, 0, 4, 1, 2)

        windowAlayout.addWidget(QLabel('Width'), 1, 0, Qt.AlignmentFlag.AlignRight)
        self.tbWinAWidth = QDoubleSpinBox()
        self.tbWinAWidth.setSingleStep(0.1)
        self.tbWinAWidth.setValue(1.0)
        windowAlayout.addWidget(self.tbWinAWidth, 1, 1)
        windowAlayout.addWidget(QLabel('m'), 1, 2)
        windowAlayout.addWidget(QLabel('<b>U</b>'), 1, 3, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMUWin_A = QDoubleSpinBox()
        windowAlayout.addWidget(self.tbRTSMUWin_A, 1, 4)
        windowAlayout.addWidget(QLabel('W/m² K'), 1, 5)

        windowAlayout.addWidget(QLabel('SC'), 2, 3, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMSCWin_A = QDoubleSpinBox()
        windowAlayout.addWidget(self.tbRTSMSCWin_A, 2, 4)

        windowAlayout.addWidget(QLabel('Internal Shader'), 3, 0, 1, 4, Qt.AlignmentFlag.AlignRight)
        self.cmbRTSMInShaderWin_A = QComboBox()
        self.cmbRTSMInShaderWin_A.addItem('No Blinds')
        windowAlayout.addWidget(self.cmbRTSMInShaderWin_A, 3, 4, 1, 2)

        windowAlayout.addWidget(QLabel('External Shader'), 4, 0, 1, 4, Qt.AlignmentFlag.AlignRight)
        self.cmbRTSMExShaderWin_A = QComboBox()
        self.cmbRTSMExShaderWin_A.addItem('No Overhanges')
        windowAlayout.addWidget(self.cmbRTSMExShaderWin_A, 4, 4, 1, 2)

        self.windowAPage.setLayout(windowAlayout)
		# endregion

		# region Window B
        windowBlayout = QGridLayout()
        windowBlayout.setVerticalSpacing(1)
        windowBlayout.setAlignment(Qt.AlignmentFlag.AlignTop)
        windowBlayout.setColumnStretch(1, 2)

        windowBlayout.addWidget(QLabel('Height'), 0, 0, Qt.AlignmentFlag.AlignRight)
        self.tbWinBHigh = QDoubleSpinBox()
        self.tbWinBHigh.setSingleStep(0.1)
        windowBlayout.addWidget(self.tbWinBHigh, 0, 1)
        windowBlayout.addWidget(QLabel('m'), 0, 2)
        windowBlayout.addWidget(QLabel('Type'), 0, 3, Qt.AlignmentFlag.AlignRight)
        self.cmbRTSMWin_B = QComboBox()
        self.cmbRTSMWin_B.addItems(windows.keys())
        self.setRTSMItemToolTip(self.cmbRTSMWin_B, windows)
        windowBlayout.addWidget(self.cmbRTSMWin_B, 0, 4, 1, 2)

        windowBlayout.addWidget(QLabel('Width'), 1, 0, Qt.AlignmentFlag.AlignRight)
        self.tbWinBWidth = QDoubleSpinBox()
        self.tbWinBWidth.setSingleStep(0.1)
        windowBlayout.addWidget(self.tbWinBWidth, 1, 1)
        windowBlayout.addWidget(QLabel('m'), 1, 2)
        windowBlayout.addWidget(QLabel('<b>U</b>'), 1, 3, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMUWin_B = QDoubleSpinBox()
        windowBlayout.addWidget(self.tbRTSMUWin_B, 1, 4)
        windowBlayout.addWidget(QLabel('W/m² K'), 1, 5)

        windowBlayout.addWidget(QLabel('SC'), 2, 3, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMSCWin_B = QDoubleSpinBox()
        windowBlayout.addWidget(self.tbRTSMSCWin_B, 2, 4)

        windowBlayout.addWidget(QLabel('Internal Shader'), 3, 0, 1, 4, Qt.AlignmentFlag.AlignRight)
        self.cmbRTSMInShaderWin_B = QComboBox()
        self.cmbRTSMInShaderWin_B.addItem('No Blinds')
        windowBlayout.addWidget(self.cmbRTSMInShaderWin_B, 3, 4, 1, 2)

        windowBlayout.addWidget(QLabel('External Shader'), 4, 0, 1, 4, Qt.AlignmentFlag.AlignRight)
        self.cmbRTSMExShaderWin_B = QComboBox()
        self.cmbRTSMExShaderWin_B.addItem('No Overhanges')
        windowBlayout.addWidget(self.cmbRTSMExShaderWin_B, 4, 4, 1, 2)

        self.windowBPage.setLayout(windowBlayout)
		# endregion

		# region Window C
        windowClayout = QGridLayout()
        windowClayout.setVerticalSpacing(1)
        windowClayout.setAlignment(Qt.AlignmentFlag.AlignTop)
        windowClayout.setColumnStretch(1, 2)

        windowClayout.addWidget(QLabel('Height'), 0, 0, Qt.AlignmentFlag.AlignRight)
        self.tbWinCHigh = QDoubleSpinBox()
        self.tbWinCHigh.setSingleStep(0.1)
        windowClayout.addWidget(self.tbWinCHigh, 0, 1)
        windowClayout.addWidget(QLabel('m'), 0, 2)
        windowClayout.addWidget(QLabel('Type'), 0, 3, Qt.AlignmentFlag.AlignRight)
        self.cmbRTSMWin_C = QComboBox()
        self.cmbRTSMWin_C.addItems(windows.keys())
        self.setRTSMItemToolTip(self.cmbRTSMWin_C, windows)
        windowClayout.addWidget(self.cmbRTSMWin_C, 0, 4, 1, 2)

        windowClayout.addWidget(QLabel('Width'), 1, 0, Qt.AlignmentFlag.AlignRight)
        self.tbWinCWidth = QDoubleSpinBox()
        self.tbWinCWidth.setSingleStep(0.1)
        windowClayout.addWidget(self.tbWinCWidth, 1, 1)
        windowClayout.addWidget(QLabel('m'), 1, 2)
        windowClayout.addWidget(QLabel('<b>U</b>'), 1, 3, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMUWin_C = QDoubleSpinBox()
        windowClayout.addWidget(self.tbRTSMUWin_C, 1, 4)
        windowClayout.addWidget(QLabel('W/m² K'), 1, 5)

        windowClayout.addWidget(QLabel('SC'), 2, 3, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMSCWin_C = QDoubleSpinBox()
        windowClayout.addWidget(self.tbRTSMSCWin_C, 2, 4)

        windowClayout.addWidget(QLabel('Internal Shader'), 3, 0, 1, 4, Qt.AlignmentFlag.AlignRight)
        self.cmbRTSMInShaderWin_C = QComboBox()
        self.cmbRTSMInShaderWin_C.addItem('No Blinds')
        windowClayout.addWidget(self.cmbRTSMInShaderWin_C, 3, 4, 1, 2)

        windowClayout.addWidget(QLabel('External Shader'), 4, 0, 1, 4, Qt.AlignmentFlag.AlignRight)
        self.cmbRTSMExShaderWin_C = QComboBox()
        self.cmbRTSMExShaderWin_C.addItem('No Overhanges')
        windowClayout.addWidget(self.cmbRTSMExShaderWin_C, 4, 4, 1, 2)

        self.windowCPage.setLayout(windowClayout)
		# endregion

		# region Window D
        windowDlayout = QGridLayout()
        windowDlayout.setVerticalSpacing(1)
        windowDlayout.setAlignment(Qt.AlignmentFlag.AlignTop)
        windowDlayout.setColumnStretch(1, 2)

        windowDlayout.addWidget(QLabel('Height'), 0, 0, Qt.AlignmentFlag.AlignRight)
        self.tbWinDHigh = QDoubleSpinBox()
        self.tbWinDHigh.setSingleStep(0.1)
        windowDlayout.addWidget(self.tbWinDHigh, 0, 1)
        windowDlayout.addWidget(QLabel('m'), 0, 2)
        windowDlayout.addWidget(QLabel('Type'), 0, 3, Qt.AlignmentFlag.AlignRight)
        self.cmbRTSMWin_D = QComboBox()
        self.cmbRTSMWin_D.addItems(windows.keys())
        self.setRTSMItemToolTip(self.cmbRTSMWin_D, windows)
        windowDlayout.addWidget(self.cmbRTSMWin_D, 0, 4, 1, 2)

        windowDlayout.addWidget(QLabel('Width'), 1, 0, Qt.AlignmentFlag.AlignRight)
        self.tbWinDWidth = QDoubleSpinBox()
        self.tbWinDWidth.setSingleStep(0.1)
        windowDlayout.addWidget(self.tbWinDWidth, 1, 1)
        windowDlayout.addWidget(QLabel('m'), 1, 2)
        windowDlayout.addWidget(QLabel('<b>U</b>'), 1, 3, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMUWin_D = QDoubleSpinBox()
        windowDlayout.addWidget(self.tbRTSMUWin_D, 1, 4)
        windowDlayout.addWidget(QLabel('W/m² K'), 1, 5)

        windowDlayout.addWidget(QLabel('SC'), 2, 3, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMSCWin_D = QDoubleSpinBox()
        windowDlayout.addWidget(self.tbRTSMSCWin_D, 2, 4)

        windowDlayout.addWidget(QLabel('Internal Shader'), 3, 0, 1, 4, Qt.AlignmentFlag.AlignRight)
        self.cmbRTSMInShaderWin_D = QComboBox()
        self.cmbRTSMInShaderWin_D.addItem('No Blinds')
        windowDlayout.addWidget(self.cmbRTSMInShaderWin_D, 3, 4, 1, 2)

        windowDlayout.addWidget(QLabel('External Shader'), 4, 0, 1, 4, Qt.AlignmentFlag.AlignRight)
        self.cmbRTSMExShaderWin_D = QComboBox()
        self.cmbRTSMExShaderWin_D.addItem('No Overhanges')
        windowDlayout.addWidget(self.cmbRTSMExShaderWin_D, 4, 4, 1, 2)

        self.windowDPage.setLayout(windowDlayout)
		# endregion

		# endregion

        architecGrid.addLayout(architecLeftlayout)
        architecGrid.addLayout(architecRightlayout)
        architecBox.setContentLayout(architecGrid)
		# endregion

        # region Application Design Data
        applicationBox = CollapsibleBox("Application Design Data")
        rtsmGrid.addWidget(applicationBox)
        applicationGrid = QGridLayout()

        applicationGrid.addWidget(QLabel('Application'), 0, 0, Qt.AlignmentFlag.AlignRight)
        self.cmbRTSMRoomType = QComboBox()
        self.cmbRTSMRoomType.addItems(set(lpd_df['Space Types']))
        applicationGrid.addWidget(self.cmbRTSMRoomType, 0, 1)
        applicationGrid.addWidget(QLabel('Lighting PD (W/m²)'), 0, 2, 1, 2, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMLPD = QDoubleSpinBox()
        self.tbRTSMLPD.setToolTip('Lighting Power Densities\nASHRAE Fundamentals 2021 p18.5 Table 2')
        applicationGrid.addWidget(self.tbRTSMLPD, 0, 4)
        applicationGrid.addWidget(QLabel('F.Space'), 0, 5, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMLightF_SPACE = QDoubleSpinBox()
        self.tbRTSMLightF_SPACE.setToolTip('Space Fraction\nASHRAE Fundamentals 2021 p18.6')
        applicationGrid.addWidget(self.tbRTSMLightF_SPACE, 0, 6)
        applicationGrid.addWidget(QLabel('F.Rad'), 0, 7, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMLightF_RAD = QDoubleSpinBox()
        self.tbRTSMLightF_RAD.setToolTip('Radiative Fraction\nASHRAE Fundamentals 2021 p18.6')
        applicationGrid.addWidget(self.tbRTSMLightF_RAD, 0, 8)

        applicationGrid.addWidget(QLabel('Usage'), 1, 0, Qt.AlignmentFlag.AlignRight)
        self.LightingGrid = checkBoxHourlyTable(24, 20)
        self.LightingGrid.setFixedHeight(45)
        applicationGrid.addWidget(self.LightingGrid, 1, 1, 1, 8)

        applicationGrid.addWidget(QLabel('Activity'), 2, 0, Qt.AlignmentFlag.AlignRight)
        self.cmbRTSMActivity = QComboBox()
        self.cmbRTSMActivity.addItems(human_hr_df['Degree of Activity'].tolist())
        applicationGrid.addWidget(self.cmbRTSMActivity, 2, 1, 1, 2)
        applicationGrid.addWidget(QLabel('SH'), 2, 3, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMPeopleSH = QSpinBox()
        self.tbRTSMPeopleSH.setMaximum(500)
        self.tbRTSMPeopleSH.setToolTip('Sensible Heat (W/Person)\nASHRAE Fundamentals 2021 p18.4')
        applicationGrid.addWidget(self.tbRTSMPeopleSH, 2, 4)
        applicationGrid.addWidget(QLabel('LH'), 2, 5, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMPeopleLH = QSpinBox()
        self.tbRTSMPeopleLH.setMaximum(500)
        self.tbRTSMPeopleLH.setToolTip('Latent Heat (W/Person)\nASHRAE Fundamentals 2021 p18.4')
        applicationGrid.addWidget(self.tbRTSMPeopleLH, 2, 6)
        applicationGrid.addWidget(QLabel('F.Rad'), 2, 7, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMPeopleF_RAD = QSpinBox()
        self.tbRTSMPeopleF_RAD.setToolTip('Radiant Fraction\nASHRAE Fundamentals 2021 p18.4')
        applicationGrid.addWidget(self.tbRTSMPeopleF_RAD, 2, 8)

        applicationGrid.addWidget(QLabel('Persons'), 3, 0, Qt.AlignmentFlag.AlignRight)
        self.PeopleGrid = numericHourlyTable(24, 20)
        self.PeopleGrid.setFixedHeight(45)
        applicationGrid.addWidget(self.PeopleGrid, 3, 1, 1, 8)

        applicationGrid.addWidget(QLabel('Ventilation'), 4, 0, Qt.AlignmentFlag.AlignRight)
        self.cmbRTSMVentilation = QComboBox()
        self.cmbRTSMVentilation.addItems(ventilations.keys())
        applicationGrid.addWidget(self.cmbRTSMVentilation, 4, 1)
        self.tbRTSMVentilation = QDoubleSpinBox()
        self.tbRTSMVentilation.setMaximum(10000)
        self.tbRTSMVentilation.setFixedHeight(24)
        applicationGrid.addWidget(self.tbRTSMVentilation, 4, 2)
        applicationGrid.addWidget(QLabel('Rp'), 4, 3, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMRp = QDoubleSpinBox()
        self.tbRTSMRp.setToolTip('People Outdoor Air Rate (cfm/Person)\nASHRAE Standard 62.1-2022 p6.16')
        applicationGrid.addWidget(self.tbRTSMRp, 4, 4)
        applicationGrid.addWidget(QLabel('Ra'), 4, 5, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMRa = QDoubleSpinBox()
        self.tbRTSMRa.setToolTip('Area Outdoor Air Rate (cfm/Person)\nASHRAE Standard 62.1-2022 p6.16')
        applicationGrid.addWidget(self.tbRTSMRa, 4, 6)
        applicationGrid.addWidget(QLabel('Ez'), 4, 7, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMEz = QDoubleSpinBox()
        self.tbRTSMEz.setToolTip('Zone Air Distribution Effectiveness\nASHRAE Standard 62.1-2022 p6.21')
        self.tbRTSMEz.setValue(1)
        applicationGrid.addWidget(self.tbRTSMEz, 4, 8)

        applicationGrid.addWidget(QLabel('Equipment'), 5, 0, Qt.AlignmentFlag.AlignRight)
        self.cmbRTSMEquipment = QComboBox()
        self.cmbRTSMEquipment.addItems(equipments.keys())
        applicationGrid.addWidget(self.cmbRTSMEquipment, 5, 1, 1, 2)
        applicationGrid.addWidget(QLabel('SH'), 5, 3, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMEquipmentSH = QDoubleSpinBox()
        self.tbRTSMEquipmentSH.setMaximum(50000)
        self.tbRTSMEquipmentSH.setGroupSeparatorShown(True)
        self.tbRTSMEquipmentSH.setDecimals(0)
        applicationGrid.addWidget(self.tbRTSMEquipmentSH, 5, 4)
        applicationGrid.addWidget(QLabel('LH'), 5, 5, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMEquipmentLH = QDoubleSpinBox()
        self.tbRTSMEquipmentLH.setMaximum(50000)
        self.tbRTSMEquipmentLH.setGroupSeparatorShown(True)
        self.tbRTSMEquipmentLH.setDecimals(0)
        applicationGrid.addWidget(self.tbRTSMEquipmentLH, 5, 6)
        applicationGrid.addWidget(QLabel('No'), 5, 7, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMEquipmentNo = QSpinBox()
        self.tbRTSMEquipmentNo.setValue(1)
        applicationGrid.addWidget(self.tbRTSMEquipmentNo, 5, 8)

        applicationBox.setContentLayout(applicationGrid)
		# endregion

		# region Cooling Load Calcualtion
        coolingGroupBox = QGroupBox("Cooling Load")
        coolingGroupBox.setObjectName('Cooling_Load')
        rtsmGrid.addWidget(coolingGroupBox)
        coolingLayout = QHBoxLayout(coolingGroupBox)
		# region Calculate
        calculateLayout = QGridLayout()
        calculateLayout.setAlignment(Qt.AlignmentFlag.AlignTop)
        calculateLayout.setContentsMargins(0,24,0,0)
        calculateLayout.setVerticalSpacing(0)
        calculateLayout.setHorizontalSpacing(3)
        coolingLayout.addLayout(calculateLayout)

        calculateLayout.addWidget(QLabel('Safety (%)'), 0, 0, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMSafety = QSpinBox()
        self.tbRTSMSafety.setValue(5)
        self.tbRTSMSafety.setFixedHeight(24)
        calculateLayout.addWidget(self.tbRTSMSafety, 0, 1, 1, 2)

        self.btnRTSMCalculate = QPushButton('Calculate')
        self.btnRTSMCalculate.setStyleSheet('color: orange;')
        calculateLayout.addWidget(self.btnRTSMCalculate, 1, 0, Qt.AlignmentFlag.AlignRight)
        self.btnRTSMSolar = QPushButton('☼')
        self.btnRTSMSolar.setFixedSize(30, 24)
        calculateLayout.addWidget(self.btnRTSMSolar, 1, 1, Qt.AlignmentFlag.AlignLeft)
        self.btnRTSMConfig = QPushButton(QIcon('res/setting.png'),'')
        self.btnRTSMConfig.setFixedSize(30, 24)
        calculateLayout.addWidget(self.btnRTSMConfig, 1, 2, Qt.AlignmentFlag.AlignRight)

        self.tbRTSMOutTotalAll = QLineEdit('0')
        self.tbRTSMOutTotalAll.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.tbRTSMOutTotalAll.setStyleSheet('background: yellow; color: black;font-weight: bold;')
        self.tbRTSMOutTotalAll.setFixedHeight(22)
        self.tbRTSMOutTotalAll.setReadOnly(True)
        calculateLayout.addWidget(self.tbRTSMOutTotalAll, 2, 0)
        self.cmbRTSMOutUnit = QComboBox()
        self.cmbRTSMOutUnit.addItems(['Btu/hr','Watt'])
        calculateLayout.addWidget(self.cmbRTSMOutUnit, 2, 1, 1, 2)

        self.cbTSMExport = QCheckBox('TSM Export')
        self.cbTSMExport.setFixedHeight(26)
        calculateLayout.addWidget(self.cbTSMExport, 4, 0, 1, 2)

        self.btnRTSMOpen = QPushButton('Open')
        calculateLayout.addWidget(self.btnRTSMOpen, 5, 0)
        self.btnRTSMSave = QPushButton('Save')
        self.btnRTSMSave .setFixedWidth(62)
        calculateLayout.addWidget(self.btnRTSMSave, 5, 1, 1, 2)
        self.btnAbout = QPushButton('About')
        self.btnAbout.setFixedSize(145, 24)
        calculateLayout.addWidget(self.btnAbout, 6, 0, 1, 3)
		# endregion

		# region Component
        componentLayout = QGridLayout()
        componentLayout.setAlignment(Qt.AlignmentFlag.AlignTop)
        componentLayout.setSpacing(0)
        coolingLayout.addLayout(componentLayout)

        self.tbLoadComponemt = QLineEdit('Load Component')
        self.tbLoadComponemt.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.tbLoadComponemt.setReadOnly(True)
        self.tbLoadComponemt.setStyleSheet('border: 1px solid orange;')
        componentLayout.addWidget(self.tbLoadComponemt, 0, 0, 1, 6)
        loadComponemt = QLineEdit('Total Load')
        loadComponemt.setStyleSheet('font-weight: bold; border: 1px solid orange;')
        loadComponemt.setAlignment(Qt.AlignmentFlag.AlignCenter)
        loadComponemt.setReadOnly(True)
        loadComponemt.setFixedHeight(50)
        componentLayout.addWidget(loadComponemt, 0, 6, 2, 3)

        self.btnRTSMOutRoof = QPushButton('Roof')
        self.btnRTSMOutRoof.setFixedWidth(70)
        componentLayout.addWidget(self.btnRTSMOutRoof, 1, 0, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMOutRoof = QLineEdit()
        self.tbRTSMOutRoof.setReadOnly(True)
        componentLayout.addWidget(self.tbRTSMOutRoof, 1, 1)
        wallLabel = QLineEdit('Wall')
        wallLabel.setAlignment(Qt.AlignmentFlag.AlignCenter)
        wallLabel.setReadOnly(True)
        componentLayout.addWidget(wallLabel, 1, 2, 1, 2)
        windowLabel = QLineEdit('Window')
        windowLabel.setAlignment(Qt.AlignmentFlag.AlignCenter)
        windowLabel.setReadOnly(True)
        componentLayout.addWidget(windowLabel, 1, 4, 1, 2)

        self.btnRTSMOutCeiling = QPushButton('Ceiling')
        self.btnRTSMOutCeiling.setFixedWidth(70)
        componentLayout.addWidget(self.btnRTSMOutCeiling, 2, 0, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMOutCeiling = QLineEdit()
        self.tbRTSMOutCeiling.setReadOnly(True)
        componentLayout.addWidget(self.tbRTSMOutCeiling, 2, 1)
        self.btnRTSMOutWallA = QPushButton('A')
        self.btnRTSMOutWallA.setFixedWidth(20)
        componentLayout.addWidget(self.btnRTSMOutWallA, 2, 2)
        self.tbRTSMOutWallA = QLineEdit()
        self.tbRTSMOutWallA.setReadOnly(True)
        componentLayout.addWidget(self.tbRTSMOutWallA, 2, 3)
        self.btnRTSMOutWinA = QPushButton('A')
        self.btnRTSMOutWinA.setFixedWidth(20)
        componentLayout.addWidget(self.btnRTSMOutWinA, 2, 4)
        self.tbRTSMOutWinA = QLineEdit()
        self.tbRTSMOutWinA.setReadOnly(True)
        componentLayout.addWidget(self.tbRTSMOutWinA, 2, 5)
        self.btnRTSMOutExternal = QPushButton('External')
        self.btnRTSMOutExternal.setFixedWidth(60)
        componentLayout.addWidget(self.btnRTSMOutExternal, 2, 6)
        self.tbRTSMOutExternal = QLineEdit()
        self.tbRTSMOutExternal.setReadOnly(True)
        componentLayout.addWidget(self.tbRTSMOutExternal, 2, 7)

        self.btnRTSMOutFloor = QPushButton('Floor')
        self.btnRTSMOutFloor.setFixedWidth(70)
        componentLayout.addWidget(self.btnRTSMOutFloor, 3, 0, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMOutFloor = QLineEdit()
        self.tbRTSMOutFloor.setReadOnly(True)
        componentLayout.addWidget(self.tbRTSMOutFloor, 3, 1)
        self.btnRTSMOutWallB = QPushButton('B')
        self.btnRTSMOutWallB.setFixedWidth(20)
        componentLayout.addWidget(self.btnRTSMOutWallB, 3, 2)
        self.tbRTSMOutWallB = QLineEdit()
        self.tbRTSMOutWallB.setReadOnly(True)
        componentLayout.addWidget(self.tbRTSMOutWallB, 3, 3)
        self.btnRTSMOutWinB = QPushButton('B')
        self.btnRTSMOutWinB.setFixedWidth(20)
        componentLayout.addWidget(self.btnRTSMOutWinB, 3, 4)
        self.tbRTSMOutWinB = QLineEdit()
        self.tbRTSMOutWinB.setReadOnly(True)
        componentLayout.addWidget(self.tbRTSMOutWinB, 3, 5)
        self.btnRTSMOutInternal = QPushButton('Internal')
        self.btnRTSMOutInternal.setFixedWidth(60)
        componentLayout.addWidget(self.btnRTSMOutInternal, 3, 6)
        self.tbRTSMOutInternal = QLineEdit()
        self.tbRTSMOutInternal.setReadOnly(True)
        componentLayout.addWidget(self.tbRTSMOutInternal, 3, 7)

        self.btnRTSMOutLighting = QPushButton('Lighting')
        self.btnRTSMOutLighting.setFixedWidth(70)
        componentLayout.addWidget(self.btnRTSMOutLighting, 4, 0, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMOutLighting = QLineEdit()
        self.tbRTSMOutLighting.setReadOnly(True)
        componentLayout.addWidget(self.tbRTSMOutLighting, 4, 1)
        self.btnRTSMOutWallC = QPushButton('C')
        self.btnRTSMOutWallC.setFixedWidth(20)
        componentLayout.addWidget(self.btnRTSMOutWallC, 4, 2)
        self.tbRTSMOutWallC = QLineEdit()
        self.tbRTSMOutWallC.setReadOnly(True)
        componentLayout.addWidget(self.tbRTSMOutWallC, 4, 3)
        self.btnRTSMOutWinC = QPushButton('C')
        self.btnRTSMOutWinC.setFixedWidth(20)
        componentLayout.addWidget(self.btnRTSMOutWinC, 4, 4)
        self.tbRTSMOutWinC = QLineEdit()
        self.tbRTSMOutWinC.setReadOnly(True)
        componentLayout.addWidget(self.tbRTSMOutWinC, 4, 5)
        self.btnRTSMOutOA = QPushButton('Fresh Air')
        self.btnRTSMOutOA.setFixedWidth(60)
        componentLayout.addWidget(self.btnRTSMOutOA, 4, 6)
        self.tbRTSMOutOA = QLineEdit()
        self.tbRTSMOutOA.setReadOnly(True)
        componentLayout.addWidget(self.tbRTSMOutOA, 4, 7)

        self.btnRTSMOutEqp = QPushButton('Equipment')
        self.btnRTSMOutEqp.setFixedWidth(70)
        componentLayout.addWidget(self.btnRTSMOutEqp, 5, 0, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMOutEqp = QLineEdit()
        self.tbRTSMOutEqp.setReadOnly(True)
        componentLayout.addWidget(self.tbRTSMOutEqp, 5, 1)
        self.btnRTSMOutWallD = QPushButton('D')
        self.btnRTSMOutWallD.setFixedWidth(20)
        componentLayout.addWidget(self.btnRTSMOutWallD, 5, 2)
        self.tbRTSMOutWallD = QLineEdit()
        self.tbRTSMOutWallD.setReadOnly(True)
        componentLayout.addWidget(self.tbRTSMOutWallD, 5, 3)
        self.btnRTSMOutWinD = QPushButton('D')
        self.btnRTSMOutWinD.setFixedWidth(20)
        componentLayout.addWidget(self.btnRTSMOutWinD, 5, 4)
        self.tbRTSMOutWinD = QLineEdit()
        self.tbRTSMOutWinD.setReadOnly(True)
        componentLayout.addWidget(self.tbRTSMOutWinD, 5, 5)
        outSafety = QLineEdit('Safety')
        outSafety.setAlignment(Qt.AlignmentFlag.AlignCenter)
        outSafety.setReadOnly(True)
        outSafety.setFixedWidth(60)
        componentLayout.addWidget(outSafety, 5, 6)
        self.tbRTSMOutSafety = QLineEdit()
        self.tbRTSMOutSafety.setReadOnly(True)
        componentLayout.addWidget(self.tbRTSMOutSafety, 5, 7)

        self.btnRTSMOutPeople = QPushButton('People')
        self.btnRTSMOutPeople.setFixedWidth(70)
        componentLayout.addWidget(self.btnRTSMOutPeople, 6, 0, Qt.AlignmentFlag.AlignRight)
        self.tbRTSMOutPeople = QLineEdit()
        self.tbRTSMOutPeople.setReadOnly(True)
        componentLayout.addWidget(self.tbRTSMOutPeople, 6, 1)
        self.btnRTSMOutWallTotal = QPushButton('≡')
        self.btnRTSMOutWallTotal.setFixedWidth(20)
        componentLayout.addWidget(self.btnRTSMOutWallTotal, 6, 2)
        self.tbRTSMOutWallTotal = QLineEdit()
        self.tbRTSMOutWallTotal.setReadOnly(True)
        componentLayout.addWidget(self.tbRTSMOutWallTotal, 6, 3)
        self.btnRTSMOutWinTotal = QPushButton('≡')
        self.btnRTSMOutWinTotal.setFixedWidth(20)
        componentLayout.addWidget(self.btnRTSMOutWinTotal, 6, 4)
        self.tbRTSMOutWinTotal = QLineEdit()
        self.tbRTSMOutWinTotal.setReadOnly(True)
        componentLayout.addWidget(self.tbRTSMOutWinTotal, 6, 5)
        self.btnRTSMOutTotal = QPushButton('Total')
        self.btnRTSMOutTotal.setStyleSheet('font-weight: bold;')
        self.btnRTSMOutTotal.setFixedWidth(60)
        componentLayout.addWidget(self.btnRTSMOutTotal, 6, 6)
        self.tbRTSMOutTotal = QLineEdit()
        self.tbRTSMOutTotal.setReadOnly(True)
        componentLayout.addWidget(self.tbRTSMOutTotal, 6, 7)
		# endregion

		# endregion Cooling Load Calcualtion

        # region Signals
        self.cmbRTSMLocation.currentIndexChanged.connect(self.RTSMLocation_changed)
        self.btnRTSM_DesignTempProfile.clicked.connect(self.profileChart)
        self.cmbRTSMSpaceType.currentIndexChanged.connect(self.RTSMSpaceType_changed)
        self.tbCompass.valueChanged.connect(self.RTSMCompasse_changed)
        self.tbZoneLength.valueChanged.connect(self.RTSMLength_changed)
        self.tbZoneWidth.valueChanged.connect(self.RTSMWidth_changed)
        self.cmbRTSMRoofType.currentIndexChanged.connect(self.roofEnvelopment_changed)
        self.cmbRTSMFloorType.currentIndexChanged.connect(self.floorEnvelopment_changed)
        self.cmbRTSMWall_A.currentIndexChanged.connect(self.wallA_Envelopment_changed)
        self.cmbRTSMWall_B.currentIndexChanged.connect(self.wallB_Envelopment_changed)
        self.cmbRTSMWall_C.currentIndexChanged.connect(self.wallC_Envelopment_changed)
        self.cmbRTSMWall_D.currentIndexChanged.connect(self.wallD_Envelopment_changed)
        self.cmbRTSMWin_A.currentIndexChanged.connect(self.windowA_changed)
        self.cmbRTSMWin_B.currentIndexChanged.connect(self.windowB_changed)
        self.cmbRTSMWin_C.currentIndexChanged.connect(self.windowC_changed)
        self.cmbRTSMWin_D.currentIndexChanged.connect(self.windowD_changed)
        self.tbWinAHigh.valueChanged.connect(lambda checked, wall='Wall_A': self.ShowWindow(wall))
        self.tbWinAWidth.valueChanged.connect(lambda checked, wall='Wall_A': self.ShowWindow(wall))
        self.tbWinBHigh.valueChanged.connect(lambda checked, wall='Wall_B': self.ShowWindow(wall))
        self.tbWinBWidth.valueChanged.connect(lambda checked, wall='Wall_B': self.ShowWindow(wall))
        self.tbWinCHigh.valueChanged.connect(lambda checked, wall='Wall_C': self.ShowWindow(wall))
        self.tbWinCWidth.valueChanged.connect(lambda checked, wall='Wall_C': self.ShowWindow(wall))
        self.tbWinDHigh.valueChanged.connect(lambda checked, wall='Wall_D': self.ShowWindow(wall))
        self.tbWinDWidth.valueChanged.connect(lambda checked, wall='Wall_D': self.ShowWindow(wall))
        self.tbRTSMSafety.valueChanged.connect(self.RTSMCalculate)

        self.cmbRTSMRoomType.currentIndexChanged.connect(self.cmbRTSMRoomType_changed)
        self.cmbRTSMActivity.currentIndexChanged.connect(self.cmbRTSMActivity_changed)
        self.cmbRTSMVentilation.currentIndexChanged.connect(self.cmbRTSMVentilation_changed)
        self.cmbRTSMEquipment.currentIndexChanged.connect(self.cmbRTSMEquipment_changed)
        self.cmbRTSMOutUnit.currentIndexChanged.connect(self.RTSMCalculate)

        self.btnRTSMCalculate.clicked.connect(self.RTSMCalculate)
        self.btnRTSMSolar.clicked.connect(self.RTSMReport)
        self.btnRTSMConfig.clicked.connect(self.RTSMConfig)
        self.btnRTSMOutRoof.clicked.connect(self.RTSMReport)
        self.btnRTSMOutCeiling.clicked.connect(self.RTSMReport)
        self.btnRTSMOutFloor.clicked.connect(self.RTSMReport)
        self.btnRTSMOutLighting.clicked.connect(self.RTSMReport)
        self.btnRTSMOutEqp.clicked.connect(self.RTSMReport)
        self.btnRTSMOutPeople.clicked.connect(self.RTSMReport)
        self.btnRTSMOutWallA.clicked.connect(self.RTSMReport)
        self.btnRTSMOutWallB.clicked.connect(self.RTSMReport)
        self.btnRTSMOutWallC.clicked.connect(self.RTSMReport)
        self.btnRTSMOutWallD.clicked.connect(self.RTSMReport)
        self.btnRTSMOutWallTotal.clicked.connect(self.RTSMReport)
        self.btnRTSMOutWinA.clicked.connect(self.RTSMReport)
        self.btnRTSMOutWinB.clicked.connect(self.RTSMReport)
        self.btnRTSMOutWinC.clicked.connect(self.RTSMReport)
        self.btnRTSMOutWinD.clicked.connect(self.RTSMReport)
        self.btnRTSMOutWinTotal.clicked.connect(self.RTSMReport)
        self.btnRTSMOutExternal.clicked.connect(self.RTSMReport)
        self.btnRTSMOutInternal.clicked.connect(self.RTSMReport)
        self.btnRTSMOutOA.clicked.connect(self.RTSMReport)
        self.btnRTSMOutTotal.clicked.connect(self.RTSMReport)
        self.btnRTSMOpen.clicked.connect(self.RTSMOpen)
        self.btnRTSMSave.clicked.connect(self.RTSMSave)
        self.btnAbout.clicked.connect(self.About)
		# endregion

        self.zone_weather = WeatherData()
        self.zone = ComfortZone(self.tbRTSMName.text(), weather=self.zone_weather)
        self.cmbRTSMLocation.currentIndexChanged.emit(0)
        self.cmbRTSMSpaceType.currentIndexChanged.emit(0)
        self.cmbRTSMRoofType.currentIndexChanged.emit(0)
        self.cmbRTSMFloorType.currentIndexChanged.emit(0)
        self.tbCompass.setValue(120)
        self.tbZoneHeight.setValue(3.8)
        self.tbZoneLength.setValue(20)
        self.tbZoneWidth.setValue(10)
        self.cmbRTSMWall_D.setCurrentText('IW.1')
        self.cmbRTSMWall_C.currentIndexChanged.emit(0)
        self.cmbRTSMWall_B.currentIndexChanged.emit(0)
        self.cmbRTSMWall_A.currentIndexChanged.emit(0)
        self.cmbRTSMWin_A.setCurrentIndex(1)

        self.cmbRTSMRoomType.setCurrentText('Office')
        self.tbRTSMLightF_SPACE.setValue(Setting.Lighting_F_space)
        self.tbRTSMLightF_RAD.setValue(Setting.Lighting_F_rad)
        self.cmbRTSMActivity.currentIndexChanged.emit(0)
        self.cmbRTSMVentilation.currentIndexChanged.emit(0)
        self.cmbRTSMEquipment.setCurrentText(default_eqp)

        self.setCentralWidget(self.widget)
        self.setFixedWidth(600)

    def setRTSMItemToolTip(self, combo: QComboBox, toolTipDict):
        for i in range(combo.count()):
            combo.setItemData(i, toolTipDict.get(combo.itemText(i)).Description, Qt.ItemDataRole.ToolTipRole)

    def RTSMLocation_changed(self):
        location = locations.get(self.cmbRTSMLocation.currentText())
        self.zone_weather.ID = location.Location
        self.tbRTSMLatitude.setValue(location.Lat)
        self.tbRTSMLongtitude.setValue(location.Long)
        self.tbRTSMAltitude.setValue(location.Elev)
        self.tbRTSMOutsideDB.setValue(location.T_db_des)
        self.tbRTSMDBRange.setValue(location.T_db_rng)
        self.tbRTSMOutsideWB.setValue(location.T_wb_mc)
        self.tbRTSMtaub.setValue(location.taub)
        self.tbRTSMtaud.setValue(location.taud)
        self.tbRTSMtz.setText(('+' if location.tz>0 else '') + str(location.tz))
        self.tbRTSMMonth.setText(location.ReferenceDates)
        self.zone_weather.altitude = Q_(location.Elev, 'm')

    def profileChart(self):
        if self.tbRTSMOutTotalAll.text() == '0': self.RTSMCalculate()
        """Shows a diagram with the hourly values (in solar time) of the dry-bulb
            and the wet-bulb temperature on the date (the design day) indicated when
            instantiating the `WeatherData` object."""
        chart = LineChart(window_title = f'Design Temperature Profile {self.zone_weather.date}')
        if self.zone_weather.T_db_prof is not None:
            x_data = [hr for hr in range(1,25)]
            chart.add_xy_data(label='Dry-bulb (°C)',
                            x1_values=x_data,
                            y1_values=[T_db.to('degC').m for T_db in self.zone_weather.T_db_prof],
                            style_props={'marker': 's', 'linestyle': '--', 'color':'r'})
            chart.add_xy_data(label='Direct Solar (W/m²)',
                            x1_values=x_data,
                            y2_values=[e for e in self.zone_weather.position_df['Eb(W/m2)']],
                            style_props={'marker': 'x'})
            chart.add_xy_data(label='Diffuse Solar (W/m²)',
                            x1_values=x_data,
                            y2_values=[e for e in self.zone_weather.position_df['Ed(W/m2)']],
                            style_props={'marker': '|'})
        chart.x1.add_title('solar hour, Hr')
        chart.y1.add_title('temperature, °C')
        chart.add_y2_axis()
        chart.y2.add_title('Watt / m²')
        chart.add_legend(anchor='upper left', position = (0.01, 0.99))
        chart.show()

    def RTSMSpaceType_changed(self):
        self.spaceTypeGraphic.setFloor(self.cmbRTSMSpaceType.currentIndex())

    def RTSMCompasse_changed(self):
        self.architecGraphic.setRotate(self.tbCompass.value())

    def RTSMLength_changed(self):
        self.architecGraphic.setLength(self.tbZoneLength.value())

    def RTSMWidth_changed(self):
        self.architecGraphic.setWidth(self.tbZoneWidth.value())

    def roofEnvelopment_changed(self):
        roof = roofs.get(self.cmbRTSMRoofType.currentText())
        self.tbRTSMURoof.setValue(roof.U)
        self.cmbRTSMRoofType.setToolTip(roof.Description)

    def floorEnvelopment_changed(self):
        floor = floors.get(self.cmbRTSMFloorType.currentText())
        self.tbRTSMUFloor.setValue(floor.U)
        self.cmbRTSMFloorType.setToolTip(floor.Description)

    def wallA_Envelopment_changed(self):
        wall = walls.get(self.cmbRTSMWall_A.currentText())
        self.tbRTSMUWall_A.setValue(wall.U)
        self.windowTab.setTabEnabled(0, wall.isExternal)
        if wall.isExternal: self.windowTab.setCurrentIndex(0)
        self.architecGraphic.setWallType(wall='Wall_A', isExternal=wall.isExternal)
        self.ShowWindow(wall='Wall_A')
        if wall.isExternal:
            self.lbOptionWall_A.setText('Solar Absorptance')
            self.tbRTSMOptionWall_A.setValue(Setting.surface_absorptance)
            self.lbOptionUnitWall_A.setVisible(False)
        else:
            self.lbOptionWall_A.setText('Temp. Diff')
            self.tbRTSMOptionWall_A.setValue(Setting.delta_T.m)
            self.lbOptionUnitWall_A.setVisible(True)
        self.cmbRTSMWall_A.setToolTip(wall.Description)

    def wallB_Envelopment_changed(self):
        wall = walls.get(self.cmbRTSMWall_B.currentText())
        self.tbRTSMUWall_B.setValue(wall.U)
        self.windowTab.setTabEnabled(1, wall.isExternal)
        if wall.isExternal: self.windowTab.setCurrentIndex(1)
        self.architecGraphic.setWallType(wall='Wall_B', isExternal=wall.isExternal)
        self.ShowWindow(wall='Wall_B')
        if wall.isExternal:
            self.lbOptionWall_B.setText('Solar Absorptance')
            self.tbRTSMOptionWall_B.setValue(Setting.surface_absorptance)
            self.lbOptionUnitWall_B.setVisible(False)
        else:
            self.lbOptionWall_B.setText('Temp. Diff')
            self.tbRTSMOptionWall_B.setValue(Setting.delta_T.m)
            self.lbOptionUnitWall_B.setVisible(True)
        self.cmbRTSMWall_B.setToolTip(wall.Description)

    def wallC_Envelopment_changed(self):
        wall = walls.get(self.cmbRTSMWall_C.currentText())
        self.tbRTSMUWall_C.setValue(wall.U)
        self.windowTab.setTabEnabled(2, wall.isExternal)
        if wall.isExternal: self.windowTab.setCurrentIndex(2)
        self.architecGraphic.setWallType(wall='Wall_C', isExternal=wall.isExternal)
        self.ShowWindow(wall='Wall_C')
        if wall.isExternal:
            self.lbOptionWall_C.setText('Solar Absorptance')
            self.tbRTSMOptionWall_C.setValue(Setting.surface_absorptance)
            self.lbOptionUnitWall_C.setVisible(False)
        else:
            self.lbOptionWall_C.setText('Temp. Diff')
            self.tbRTSMOptionWall_C.setValue(Setting.delta_T.m)
            self.lbOptionUnitWall_C.setVisible(True)
        self.cmbRTSMWall_C.setToolTip(wall.Description)

    def wallD_Envelopment_changed(self):
        wall = walls.get(self.cmbRTSMWall_D.currentText())
        self.tbRTSMUWall_D.setValue(wall.U)
        self.windowTab.setTabEnabled(3, wall.isExternal)
        if wall.isExternal: self.windowTab.setCurrentIndex(3)
        self.architecGraphic.setWallType(wall='Wall_D', isExternal=wall.isExternal)
        self.ShowWindow(wall='Wall_D')
        if wall.isExternal:
            self.lbOptionWall_D.setText('Solar Absorptance')
            self.tbRTSMOptionWall_D.setValue(Setting.surface_absorptance)
            self.lbOptionUnitWall_D.setVisible(False)
        else:
            self.lbOptionWall_D.setText('Temp. Diff')
            self.tbRTSMOptionWall_D.setValue(Setting.delta_T.m)
            self.lbOptionUnitWall_D.setVisible(True)
        self.cmbRTSMWall_D.setToolTip(wall.Description)

    def windowA_changed(self):
        window = windows.get(self.cmbRTSMWin_A.currentText())
        self.tbRTSMUWin_A.setValue(window.U)
        self.tbRTSMSCWin_A.setValue(window.SC)
        if self.cmbRTSMWin_A.currentText() == 'None':
            self.tbWinAHigh.setValue(0)
            self.tbWinAHigh.setEnabled(False)
            self.tbWinAWidth.setValue(0)
            self.tbWinAWidth.setEnabled(False)
            self.ShowWindow(wall='Wall_A')
        else:
            self.tbWinAHigh.setEnabled(True)
            self.tbWinAWidth.setEnabled(True)
        self.cmbRTSMWin_A.setToolTip(window.Description)

    def windowB_changed(self):
        window = windows.get(self.cmbRTSMWin_B.currentText())
        self.tbRTSMUWin_B.setValue(window.U)
        self.tbRTSMSCWin_B.setValue(window.SC)
        if self.cmbRTSMWin_B.currentText() == 'None':
            self.tbWinBHigh.setValue(0)
            self.tbWinBHigh.setEnabled(False)
            self.tbWinBWidth.setValue(0)
            self.tbWinBWidth.setEnabled(False)
            self.ShowWindow(wall='Wall_B')
        else:
            self.tbWinBHigh.setEnabled(True)
            self.tbWinBWidth.setEnabled(True)
        self.cmbRTSMWin_B.setToolTip(window.Description)

    def windowC_changed(self):
        window = windows.get(self.cmbRTSMWin_C.currentText())
        self.tbRTSMUWin_C.setValue(window.U)
        self.tbRTSMSCWin_C.setValue(window.SC)
        if self.cmbRTSMWin_C.currentText() == 'None':
            self.tbWinCHigh.setValue(0)
            self.tbWinCHigh.setEnabled(False)
            self.tbWinCWidth.setValue(0)
            self.tbWinCWidth.setEnabled(False)
            self.ShowWindow(wall='Wall_C')
        else:
            self.tbWinCHigh.setEnabled(True)
            self.tbWinCWidth.setEnabled(True)
        self.cmbRTSMWin_C.setToolTip(window.Description)

    def windowD_changed(self):
        window = windows.get(self.cmbRTSMWin_D.currentText())
        self.tbRTSMUWin_D.setValue(window.U)
        self.tbRTSMSCWin_D.setValue(window.SC)
        if self.cmbRTSMWin_D.currentText() == 'None':
            self.tbWinDHigh.setValue(0)
            self.tbWinDHigh.setEnabled(False)
            self.tbWinDWidth.setValue(0)
            self.tbWinDWidth.setEnabled(False)
            self.ShowWindow(wall='Wall_D')
        else:
            self.tbWinDHigh.setEnabled(True)
            self.tbWinDWidth.setEnabled(True)
        self.cmbRTSMWin_D.setToolTip(window.Description)

    def ShowWindow(self, wall:str):
        wallExternal = True
        h, w = 0, 0
        match wall:
            case 'Wall_A':
                wallExternal = walls.get(self.cmbRTSMWall_A.currentText()).isExternal
                h = self.tbWinAHigh.value()
                w = self.tbWinAWidth.value()
            case 'Wall_B':
                wallExternal = walls.get(self.cmbRTSMWall_B.currentText()).isExternal
                h = self.tbWinBHigh.value()
                w = self.tbWinBWidth.value()
            case 'Wall_C':
                wallExternal = walls.get(self.cmbRTSMWall_C.currentText()).isExternal
                h = self.tbWinCHigh.value()
                w = self.tbWinCWidth.value()
            case 'Wall_D':
                wallExternal = walls.get(self.cmbRTSMWall_D.currentText()).isExternal
                h = self.tbWinDHigh.value()
                w = self.tbWinDWidth.value()
        isHasWindow = wallExternal and h>0 and w > 0
        self.architecGraphic.setWallWindow(wall= wall, isHasWindow= isHasWindow)

    def cmbRTSMRoomType_changed(self):
        self.tbRTSMLPD.setValue(LightingPowerDensities(space_type=self.cmbRTSMRoomType.currentText()))
        vrp = human_vrp_df[human_vrp_df['Space Types'].str.match(self.cmbRTSMRoomType.currentText())]
        self.tbRTSMRp.setValue(vrp['Rp'].values[0])
        self.tbRTSMRa.setValue(vrp['Ra'].values[0])

    def cmbRTSMActivity_changed(self):
        heatrate = HumanHeatRate(activity=self.cmbRTSMActivity.currentText())
        self.tbRTSMPeopleSH.setValue(heatrate[0])
        self.tbRTSMPeopleLH.setValue(heatrate[1])
        self.tbRTSMPeopleF_RAD.setValue(heatrate[2])

    def cmbRTSMVentilation_changed(self):
        self.tbRTSMVentilation.setValue(ventilations.get(self.cmbRTSMVentilation.currentText()))
        self.tbRTSMVentilation.setVisible(False if self.cmbRTSMVentilation.currentText()  == 'ASHRAE 62.1' else True)

    def vbzCalculate(self)->float:
        Rp = float(self.tbRTSMRp.value())
        Pz = max(self.PeopleGrid.value())
        Ra = float(self.tbRTSMRa.value())
        Az = self.zone.Area.to('foot ** 2').m
        Ez = float(self.tbRTSMEz.value())
        return (Rp * Pz + Ra * Az)/Ez

    def cmbRTSMEquipment_changed(self):
        sender = self.cmbRTSMEquipment.currentText()
        if sender == default_eqp:
            self.tbRTSMEquipmentSH.setValue(EquipmentPowerDensities(space_type=self.cmbRTSMRoomType.currentText()))
        else:
            self.tbRTSMEquipmentSH.setValue(equipments.get(sender).SH)
            self.tbRTSMEquipmentLH.setValue(equipments.get(sender).LH)

    def RTSMCalculate(self):
        self.zone.ID = self.tbRTSMName.text()
        self.zone_weather.fi = Q_(self.tbRTSMLatitude.value(), 'deg')
        self.zone_weather.L_loc = Q_(self.tbRTSMLongtitude.value(), 'deg')
        self.zone_weather.altitude = Q_(self.tbRTSMAltitude.value(), 'm')
        self.zone_weather.T_db_des = Q_(self.tbRTSMOutsideDB.value(), 'degC')
        self.zone_weather.T_db_rng = Q_(self.tbRTSMDBRange.value(), 'delta_degC')
        self.zone_weather.T_wb_mc = Q_(self.tbRTSMOutsideWB.value(), 'degC')
        self.zone_weather.taub = self.tbRTSMtaub.value()
        self.zone_weather.taud = self.tbRTSMtaud.value()
        self.zone_weather.tz = int(self.tbRTSMtz.text())
        self.zone_weather.date = ReferenceDates.get_date_for(self.tbRTSMMonth.text())
        self.zone_weather.synthetic_daily_db_profiles()

        Setting.Inside_DB = Q_(self.tbRTSMInsideDB.value(), 'degC')
        Setting.Inside_RH = self.tbRTSMInsideRH.value()/100
        self.zone.SpaceType = space_types.index(self.cmbRTSMSpaceType.currentText())
        self.zone.Oreintation = self.tbCompass.value()
        self.zone.Width = Q_(self.tbZoneWidth.value(), 'm')
        self.zone.Length = Q_(self.tbZoneLength.value(), 'm')
        self.zone.Height = Q_(self.tbZoneHeight.value() , 'm')

        # region Roof & Floor
        self.zone.Roof.weather_data = self.zone_weather
        self.zone.Roof.net_area = self.zone.Area
        self.zone.Roof.U = self.tbRTSMURoof.value()
        self.zone.Roof.CTS = roofs.get(self.cmbRTSMRoofType.currentText()).CTS

        self.zone.Floor.weather_data = self.zone_weather
        self.zone.Floor.net_area = self.zone.Area
        self.zone.Floor.U = self.tbRTSMURoof.value()
        self.zone.Floor.CTS = floors.get(self.cmbRTSMFloorType.currentText()).CTS
        self.zone.Floor.delta_T = Q_(self.tbRTSMOptionFloor.value(), 'delta_degC')
        # endregion

        # region Wall-A
        self.zone.Wall_A.weather_data = self.zone_weather
        self.zone.Wall_A.net_area = self.zone.Height * self.zone.Length
        self.zone.Wall_A.U = self.tbRTSMUWall_A.value()
        self.zone.Wall_A.CTS = walls.get(self.cmbRTSMWall_A.currentText()).CTS
        self.zone.Wall_A.wall_type = Wall.WallType.External if walls.get(self.cmbRTSMWall_A.currentText()).isExternal else Wall.WallType.Internal
        self.zone.Wall_A.clear_window()

        if walls.get(self.cmbRTSMWall_A.currentText()).isExternal:
            self.zone.Wall_A.surface_absorptance = self.tbRTSMOptionWall_A.value()
            win_w = Q_(self.tbWinAWidth.value(), 'm')
            win_h = Q_(self.tbWinAHigh.value(), 'm')
            if win_w>0 and win_h>0 and self.cmbRTSMWin_A.currentText() != 'None':
                self.zone.Wall_A.add_window(id='Win-A', width=win_w, height=win_h, U=self.tbRTSMUWin_A.value(), SC=self.tbRTSMSCWin_A.value())
                window = windows.get(self.cmbRTSMWin_A.currentText())
                win = self.zone.Wall_A.windows.get('Win-A')
                win.SHGCd = [window.SHGCd0, window.SHGCd4, window.SHGCd5, window.SHGCd6, window.SHGCd7, window.SHGCd8, 0]
                win.SHGCh = window.SHGCh
        else:
            self.zone.Wall_A.delta_T = Q_(self.tbRTSMOptionWall_A.value(), 'delta_degC')
        # endregion

        # region Wall-B
        self.zone.Wall_B.weather_data = self.zone_weather
        self.zone.Wall_B.net_area = self.zone.Height * self.zone.Width
        self.zone.Wall_B.U = self.tbRTSMUWall_B.value()
        self.zone.Wall_B.CTS = walls.get(self.cmbRTSMWall_B.currentText()).CTS
        self.zone.Wall_B.wall_type = Wall.WallType.External if walls.get(self.cmbRTSMWall_B.currentText()).isExternal else Wall.WallType.Internal
        self.zone.Wall_B.clear_window()

        if walls.get(self.cmbRTSMWall_B.currentText()).isExternal:
            self.zone.Wall_B.surface_absorptance = self.tbRTSMOptionWall_B.value()
            win_w = Q_(self.tbWinBWidth.value(), 'm')
            win_h = Q_(self.tbWinBHigh.value(), 'm')
            if win_w>0 and win_h>0 and self.cmbRTSMWin_B.currentText() != 'None':
                self.zone.Wall_B.add_window(id='Win-B', width=win_w, height=win_h, U=self.tbRTSMUWin_B.value(), SC=self.tbRTSMSCWin_B.value())
                window = windows.get(self.cmbRTSMWin_B.currentText())
                win = self.zone.Wall_B.windows.get('Win-B')
                win.SHGCd = [window.SHGCd0, window.SHGCd4, window.SHGCd5, window.SHGCd6, window.SHGCd7, window.SHGCd8, 0]
                win.SHGCh = window.SHGCh
        else:
            self.zone.Wall_B.delta_T = Q_(self.tbRTSMOptionWall_B.value(), 'delta_degC')
        # endregion

        # region Wall-C
        self.zone.Wall_C.weather_data = self.zone_weather
        self.zone.Wall_C.net_area = self.zone.Height * self.zone.Length
        self.zone.Wall_C.U = self.tbRTSMUWall_C.value()
        self.zone.Wall_C.CTS = walls.get(self.cmbRTSMWall_C.currentText()).CTS
        self.zone.Wall_C.wall_type = Wall.WallType.External if walls.get(self.cmbRTSMWall_C.currentText()).isExternal else Wall.WallType.Internal
        self.zone.Wall_C.clear_window()

        if walls.get(self.cmbRTSMWall_C.currentText()).isExternal:
            self.zone.Wall_C.surface_absorptance = self.tbRTSMOptionWall_C.value()
            win_w = Q_(self.tbWinCWidth.value(), 'm')
            win_h = Q_(self.tbWinCHigh.value(), 'm')
            if win_w>0 and win_h>0 and self.cmbRTSMWin_C.currentText() != 'None':
                self.zone.Wall_C.add_window(id='Win-C', width=win_w, height=win_h, U=self.tbRTSMUWin_C.value(), SC=self.tbRTSMSCWin_C.value())
                window = windows.get(self.cmbRTSMWin_C.currentText())
                win = self.zone.Wall_C.windows.get('Win-C')
                win.SHGCd = [window.SHGCd0, window.SHGCd4, window.SHGCd5, window.SHGCd6, window.SHGCd7, window.SHGCd8, 0]
                win.SHGCh = window.SHGCh
        else:
            self.zone.Wall_C.delta_T = Q_(self.tbRTSMOptionWall_C.value(), 'delta_degC')
        # endregion

        # region Wall-D
        self.zone.Wall_D.weather_data = self.zone_weather
        self.zone.Wall_D.net_area = self.zone.Height * self.zone.Width
        self.zone.Wall_D.U = self.tbRTSMUWall_D.value()
        self.zone.Wall_D.CTS = walls.get(self.cmbRTSMWall_D.currentText()).CTS
        self.zone.Wall_D.wall_type = Wall.WallType.External if walls.get(self.cmbRTSMWall_D.currentText()).isExternal else Wall.WallType.Internal
        self.zone.Wall_D.clear_window()

        if walls.get(self.cmbRTSMWall_D.currentText()).isExternal:
            self.zone.Wall_D.surface_absorptance = self.tbRTSMOptionWall_D.value()
            win_w = Q_(self.tbWinDWidth.value(), 'm')
            win_h = Q_(self.tbWinDHigh.value(), 'm')
            if win_w>0 and win_h>0 and self.cmbRTSMWin_D.currentText() != 'None':
                self.zone.Wall_D.add_window(id='Win-D', width=win_w, height=win_h, U=self.tbRTSMUWin_D.value(), SC=self.tbRTSMSCWin_D.value())
                window = windows.get(self.cmbRTSMWin_D.currentText())
                win = self.zone.Wall_D.windows.get('Win-D')
                win.SHGCd = [window.SHGCd0, window.SHGCd4, window.SHGCd5, window.SHGCd6, window.SHGCd7, window.SHGCd8, 0]
                win.SHGCh = window.SHGCh
        else:
            self.zone.Wall_D.delta_T = Q_(self.tbRTSMOptionWall_D.value(), 'delta_degC')
        # endregion

        zone_light = self.zone.Light_HeatGain.get_lighting('ls0')
        zone_light.power_density = Q_(self.tbRTSMLPD.value(), 'W / m ** 2')
        zone_light.A_floor = self.zone.Area
        zone_light.F_space = Q_(self.tbRTSMLightF_SPACE.value(),'').to('%')
        zone_light.F_rad = Q_(self.tbRTSMLightF_RAD.value(),'').to('%')
        self.zone.Light_HeatGain.UpdateUsageProfile(self.LightingGrid.value())

        self.zone.People_HeatGain.occupants.Q_dot_sen_person = Q_(self.tbRTSMPeopleSH.value(), 'W')
        self.zone.People_HeatGain.occupants.Q_dot_lat_person = Q_(self.tbRTSMPeopleLH.value(), 'W')
        self.zone.People_HeatGain.occupants.F_rad = Q_(self.tbRTSMPeopleF_RAD.value(), '%')
        self.zone.People_HeatGain.UpdateUsageProfile(self.PeopleGrid.value())

        self.tbRTSMVentilation.setVisible(True)
        match self.cmbRTSMVentilation.currentText():
            case 'ASHRAE 62.1':
                Vbz = self.vbzCalculate()
                self.tbRTSMVentilation.setValue(Vbz)
                self.zone.Ventilation = Q_(Vbz, 'cubic_foot/minute')
            case 'cfm/Person':
                self.zone.Ventilation = Q_(max(self.zone.People_HeatGain.usage_profile) * self.tbRTSMVentilation.value(), 'cubic_foot/minute')
            case 'cfm/ft²':
                self.zone.Ventilation = Q_(self.zone.Area.to('foot ** 2').m * self.tbRTSMVentilation.value(), 'cubic_foot/minute')
            case 'Air Change':
                self.zone.Ventilation = Q_((self.zone.Area * self.zone.Height).to('cubic_foot').m * self.tbRTSMVentilation.value()/60, 'cubic_foot/minute')
            case 'cfm':
                self.zone.Ventilation = Q_(self.tbRTSMVentilation.value(), 'cubic_foot/minute')

        zone_equipment = self.zone.Equipment_HeatGain.get_equipment('eqp0')
        eqp_sen = self.tbRTSMEquipmentSH.value()
        eqp_lat = self.tbRTSMEquipmentLH.value()
        if self.cmbRTSMEquipment.currentText() == default_eqp:
            eqp_sen = (self.zone.Area * self.tbRTSMEquipmentSH.value()).m
            eqp_lat = (self.zone.Area * self.tbRTSMEquipmentLH.value()).m
        zone_equipment.F_rad = Q_(Setting.Equipment_Generic_F_rad * 100, '%')
        zone_equipment.Q_dot_sen_pcs = Q_(eqp_sen, 'W') * self.tbRTSMEquipmentNo.value()
        zone_equipment.Q_dot_lat_pcs = Q_(eqp_lat, 'W') * self.tbRTSMEquipmentNo.value()
        self.zone.Equipment_HeatGain.UpdateUsageProfile(self.LightingGrid.value())
        self.zone.safety = Q_(self.tbRTSMSafety.value(),'%')
        Setting.tsm_export = self.cbTSMExport.isChecked()

        print('>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> CALCULATING <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<')
        self.zone.Calculate()

        i = self.zone.max_hr[0]
        _SI_UNIT = True if self.cmbRTSMOutUnit.currentText() == 'Watt' else False
        self.tbLoadComponemt.setText(f'Load Component ({self.cmbRTSMOutUnit.currentText()}) @Hour {i}')
        cl = self.zone.internal_load_df[i:i+1]['Lighting'].values[0]
        self.tbRTSMOutLighting.setText(f'{cl if _SI_UNIT else cl*WATT_BTU_HR:,.0f}')
        cl = self.zone.internal_load_df[i:i+1]['People'].values[0]
        self.tbRTSMOutPeople.setText(f'{cl if _SI_UNIT else cl*WATT_BTU_HR:,.0f}')
        cl = self.zone.internal_load_df[i:i+1]['Equipment'].values[0]
        self.tbRTSMOutEqp.setText(f'{cl if _SI_UNIT else cl*WATT_BTU_HR:,.0f}')

        cl = self.zone.external_load_df[i:i+1]['Roof'].values[0]
        self.tbRTSMOutRoof.setText(f'{cl if _SI_UNIT else cl*WATT_BTU_HR:,.0f}')
        cl = self.zone.external_load_df[i:i+1]['Floor'].values[0]
        self.tbRTSMOutFloor.setText(f'{cl if _SI_UNIT else cl*WATT_BTU_HR:,.0f}')
        cl = self.zone.external_load_df[i:i+1]['Ceiling'].values[0]
        self.tbRTSMOutCeiling.setText(f'{cl if _SI_UNIT else cl*WATT_BTU_HR:,.0f}')

        cl = self.zone.wall_load_df[i:i+1]['Wall-A'].values[0]
        self.tbRTSMOutWallA.setText(f'{cl if _SI_UNIT else cl*WATT_BTU_HR:,.0f}')
        cl = self.zone.wall_load_df[i:i+1]['Wall-B'].values[0]
        self.tbRTSMOutWallB.setText(f'{cl if _SI_UNIT else cl*WATT_BTU_HR:,.0f}')
        cl = self.zone.wall_load_df[i:i+1]['Wall-C'].values[0]
        self.tbRTSMOutWallC.setText(f'{cl if _SI_UNIT else cl*WATT_BTU_HR:,.0f}')
        cl = self.zone.wall_load_df[i:i+1]['Wall-D'].values[0]
        self.tbRTSMOutWallD.setText(f'{cl if _SI_UNIT else cl*WATT_BTU_HR:,.0f}')
        cl = self.zone.wall_load_df[i:i+1]['TOTAL_CL'].values[0]
        self.tbRTSMOutWallTotal.setText(f'{cl if _SI_UNIT else cl*WATT_BTU_HR:,.0f}')

        cl = self.zone.window_load_df[i:i+1]['Win-A'].values[0]
        self.tbRTSMOutWinA.setText(f'{cl if _SI_UNIT else cl*WATT_BTU_HR:,.0f}')
        cl = self.zone.window_load_df[i:i+1]['Win-B'].values[0]
        self.tbRTSMOutWinB.setText(f'{cl if _SI_UNIT else cl*WATT_BTU_HR:,.0f}')
        cl = self.zone.window_load_df[i:i+1]['Win-C'].values[0]
        self.tbRTSMOutWinC.setText(f'{cl if _SI_UNIT else cl*WATT_BTU_HR:,.0f}')
        cl = self.zone.window_load_df[i:i+1]['Win-D'].values[0]
        self.tbRTSMOutWinD.setText(f'{cl if _SI_UNIT else cl*WATT_BTU_HR:,.0f}')
        cl = self.zone.window_load_df[i:i+1]['TOTAL_CL'].values[0]
        self.tbRTSMOutWinTotal.setText(f'{cl if _SI_UNIT else cl*WATT_BTU_HR:,.0f}')

        cl = self.zone.external_load_df[i:i+1]['TOTAL_CL'].values[0]
        self.tbRTSMOutExternal.setText(f'{cl if _SI_UNIT else cl*WATT_BTU_HR:,.0f}')
        cl = self.zone.internal_load_df[i:i+1]['TOTAL_CL'].values[0]
        self.tbRTSMOutInternal.setText(f'{cl if _SI_UNIT else cl*WATT_BTU_HR:,.0f}')
        cl = self.zone.cooling_load_df[i:i+1]['ventilation'].values[0]
        self.tbRTSMOutOA.setText(f'{cl if _SI_UNIT else cl*WATT_BTU_HR:,.0f}')

        i_total_cl = self.zone.cooling_load_df[i:i+1]['TOTAL_CL'].values[0]
        i_safety = i_total_cl * self.zone.safety.to('').m
        i_total = i_total_cl + i_safety
        self.tbRTSMOutSafety.setText(f'{i_safety if _SI_UNIT else i_safety*WATT_BTU_HR:,.0f}')
        self.tbRTSMOutTotal.setText(f'{i_total if _SI_UNIT else i_total*WATT_BTU_HR:,.0f}')
        self.tbRTSMOutTotalAll.setText(self.tbRTSMOutTotal.text())

    def RTSMReport(self):
        sender = self.sender()
        if self.tbRTSMOutTotalAll.text() == '0': self.RTSMCalculate()
        pd.options.display.float_format = '{:.2f}'.format
        match sender:
            case self.btnRTSMSolar:
                print('\n§ SOLAR POSITION', self.zone_weather.ID, self.zone_weather.date, '§')
                print(f'Solar declination δ(selta) = {self.zone.weather_data.declination.to('deg'):~P.2f}')
                print(self.zone.weather_data.position_df)
            case self.btnRTSMOutRoof:
                print('\n§ ROOF COOLING LOAD §')
                print(f'Surface Azimuth Ψ(psi) = {self.zone.Roof.psi.to('deg'):~P.2f}')
                print(f'Surface Tilt Angle Σ(sigama) = {self.zone.Roof.sigma.to('deg'):~P.2f}')
                print(self.zone.Roof.solar_irradiance_df)
                print(f'Inside DB = {Setting.Inside_DB:~P.2f}')
                print(self.zone.Roof.cooling_load_df)
            case self.btnRTSMOutCeiling:
                print('\n§ CEILING COOLING LOAD §')
                print(self.zone.Ceiling.cooling_load_df)
            case self.btnRTSMOutFloor:
                print('\n§ FLOOR COOLING LOAD §')
                print(self.zone.Floor.cooling_load_df)

            case self.btnRTSMOutWallA:
                print('\n§ WALL-A COOLING LOAD §')
                print(f'Surface Azimuth Ψ(psi) = {self.zone.Wall_A.psi.to('deg'):~P.2f}')
                print(f'Surface Tilt Angle Σ(sigama) = {self.zone.Wall_A.sigma.to('deg'):~P.2f}')
                print(self.zone.Wall_A.solar_irradiance_df)
                print(f'Inside DB = {Setting.Inside_DB:~P.2f}')
                print(self.zone.Wall_A.cooling_load_df)
            case self.btnRTSMOutWallB:
                print('\n§ WALL-B COOLING LOAD §')
                print(f'Surface Azimuth Ψ(psi) = {self.zone.Wall_B.psi.to('deg'):~P.2f}')
                print(f'Surface Tilt Angle Σ(sigama) = {self.zone.Wall_B.sigma.to('deg'):~P.2f}')
                print(self.zone.Wall_B.solar_irradiance_df)
                print(f'Inside DB = {Setting.Inside_DB:~P.2f}')
                print(self.zone.Wall_B.cooling_load_df)
            case self.btnRTSMOutWallC:
                print('\n§ WALL-C COOLING LOAD §')
                print(f'Surface Azimuth Ψ(psi) = {self.zone.Wall_C.psi.to('deg'):~P.2f}')
                print(f'Surface Tilt Angle Σ(sigama) = {self.zone.Wall_C.sigma.to('deg'):~P.2f}')
                print(self.zone.Wall_C.solar_irradiance_df)
                print(f'Inside DB = {Setting.Inside_DB:~P.2f}')
                print(self.zone.Wall_C.cooling_load_df)
            case self.btnRTSMOutWallD:
                print('\n§ WALL-D COOLING LOAD §')
                print(f'Surface Azimuth Ψ(psi) = {self.zone.Wall_D.psi.to('deg'):~P.2f}')
                print(f'Surface Tilt Angle Σ(sigama) = {self.zone.Wall_D.sigma.to('deg'):~P.2f}')
                print(self.zone.Wall_D.solar_irradiance_df)
                print(f'Inside DB = {Setting.Inside_DB:~P.2f}')
                print(self.zone.Wall_D.cooling_load_df)
            case self.btnRTSMOutWallTotal:
                print('\n§ WALL COOLING LOAD §')
                print(self.zone.wall_load_df)
                self.RTSMChart(window_title='Wall Cooling Load')

            case self.btnRTSMOutWinA:
                print('\n§ WINDOW-A COOLING LOAD §')
                if self.zone.Wall_A.windows: win = self.zone.Wall_A.windows['Win-A']
                print(f'Surface Azimuth Ψ(psi) = {self.zone.Wall_A.psi.to('deg'):~P.2f}')
                print(f'Surface Tilt Angle Σ(sigama) = {self.zone.Wall_A.sigma.to('deg'):~P.2f}')
                print(self.zone.Wall_A.solar_irradiance_df)
                print(f'Hemispherical Diffuse = {win.SHGCh}')
                print(win.window_SHG_df)
                print(win.cooling_load_df)
            case self.btnRTSMOutWinB:
                print('\n§ WINDOW-B COOLING LOAD §')
                if self.zone.Wall_B.windows: win = self.zone.Wall_A.windows['Win-B']
                print(f'Surface Azimuth Ψ(psi) = {self.zone.Wall_B.psi.to('deg'):~P.2f}')
                print(f'Surface Tilt Angle Σ(sigama) = {self.zone.Wall_B.sigma.to('deg'):~P.2f}')
                print(self.zone.Wall_B.solar_irradiance_df)
                print(f'Hemispherical Diffuse = {win.SHGCh}')
                print(win.window_SHG_df)
                print(win.cooling_load_df)
            case self.btnRTSMOutWinC:
                print('\n§ WINDOW-C COOLING LOAD §')
                if self.zone.Wall_C.windows: win = self.zone.Wall_A.windows['Win-C']
                print(f'Surface Azimuth Ψ(psi) = {self.zone.Wall_C.psi.to('deg'):~P.2f}')
                print(f'Surface Tilt Angle Σ(sigama) = {self.zone.Wall_C.sigma.to('deg'):~P.2f}')
                print(self.zone.Wall_C.solar_irradiance_df)
                print(f'Hemispherical Diffuse = {win.SHGCh}')
                print(win.window_SHG_df)
                print(win.cooling_load_df)
            case self.btnRTSMOutWinD:
                print('\n§ WINDOW-D COOLING LOAD §')
                if self.zone.Wall_D.windows: win = self.zone.Wall_A.windows['Win-D']
                print(f'Surface Azimuth Ψ(psi) = {self.zone.Wall_D.psi.to('deg'):~P.2f}')
                print(f'Surface Tilt Angle Σ(sigama) = {self.zone.Wall_D.sigma.to('deg'):~P.2f}')
                print(self.zone.Wall_D.solar_irradiance_df)
                print(f'Hemispherical Diffuse = {win.SHGCh}')
                print(win.window_SHG_df)
                print(win.cooling_load_df)
            case self.btnRTSMOutWinTotal:
                print('\n§ WINDOW COOLING LOAD §')
                print(self.zone.window_load_df)
                self.RTSMChart(window_title='Window Cooling Load')

            case self.btnRTSMOutLighting:
                print('\n§ LIGHTING COOLING LOAD §')
                print(self.zone.Light_HeatGain.cooling_load_df)
            case self.btnRTSMOutPeople:
                print('\n§ PEOPLE COOLING LOAD §')
                print(self.zone.People_HeatGain.cooling_load_df)
            case self.btnRTSMOutEqp:
                print('\n§ EQUIPMENT COOLING LOAD §')
                print(self.zone.Equipment_HeatGain.cooling_load_df)

            case self.btnRTSMOutExternal:
                print('\n§ EXTERNAL COOLING LOAD §')
                print(self.zone.external_load_df)
                self.RTSMChart(window_title='External Cooling Load')
            case self.btnRTSMOutInternal:
                print('\n§ INTERNAL COOLING LOAD §')
                print(self.zone.internal_load_df)
                self.RTSMChart(window_title='Internal Cooling Load')
            case self.btnRTSMOutOA:
                print('\n§ VENTILATION COOLING LOAD §')
                print(self.zone.ventilation_load_df)
            case self.btnRTSMOutTotal:
                print('\n§', self.zone.ID, 'COOLING LOAD §')
                print(self.zone.cooling_load_df)
                self.RTSMChart(window_title='Total Cooling Load')

    def RTSMChart(self, window_title: str):
        sender = self.sender()
        chart = LineChart(window_title = window_title)
        x_data = [hr for hr in range(24)]
        try:
            match sender:
                case self.btnRTSMOutWallTotal:
                    chart.add_xy_data('Wall - A',x_data,[i for i in self.zone.wall_load_df['Wall-A']])
                    chart.add_xy_data('Wall - B',x_data,[i for i in self.zone.wall_load_df['Wall-B']])
                    chart.add_xy_data('Wall - C',x_data,[i for i in self.zone.wall_load_df['Wall-C']])
                    chart.add_xy_data('Wall - D',x_data,[i for i in self.zone.wall_load_df['Wall-D']])
                    chart.add_xy_data('Wall Cooling Load',x_data,[i for i in self.zone.wall_load_df['TOTAL_CL']])
                case self.btnRTSMOutWinTotal:
                    chart.add_xy_data('Window - A',x_data,[i for i in self.zone.window_load_df['Win-A']])
                    chart.add_xy_data('Window - B',x_data,[i for i in self.zone.window_load_df['Win-B']])
                    chart.add_xy_data('Window - C',x_data,[i for i in self.zone.window_load_df['Win-C']])
                    chart.add_xy_data('Window - D',x_data,[i for i in self.zone.window_load_df['Win-D']])
                    chart.add_xy_data('Window Cooling Load',x_data,[i for i in self.zone.window_load_df['TOTAL_CL']])
                case self.btnRTSMOutExternal:
                    chart.add_xy_data('Roof',x_data,[i for i in self.zone.external_load_df['Roof']])
                    chart.add_xy_data('Floor',x_data,[i for i in self.zone.external_load_df['Floor']])
                    chart.add_xy_data('Ceiling',x_data,[i for i in self.zone.external_load_df['Ceiling']])
                    chart.add_xy_data('Wall',x_data,[i for i in self.zone.external_load_df['Wall']])
                    chart.add_xy_data('Window',x_data,[i for i in self.zone.external_load_df['Window']])
                    chart.add_xy_data('External Cooling Load',x_data,[i for i in self.zone.external_load_df['TOTAL_CL']])
                case self.btnRTSMOutInternal:
                    chart.add_xy_data('Lighting',x_data,[i for i in self.zone.internal_load_df['Lighting']])
                    chart.add_xy_data('People',x_data,[i for i in self.zone.internal_load_df['People']])
                    chart.add_xy_data('Equipment',x_data,[i for i in self.zone.internal_load_df['Equipment']])
                    chart.add_xy_data('Internal Cooling Load',x_data,[i for i in self.zone.internal_load_df['TOTAL_CL']])
                case self.btnRTSMOutTotal:
                    chart.add_xy_data('External Cooling Load',x_data,[i for i in self.zone.cooling_load_df['external']])
                    chart.add_xy_data('Internal Cooling Load',x_data,[i for i in self.zone.cooling_load_df['internal']])
                    chart.add_xy_data('Ventilation Cooling Load',x_data,[i for i in self.zone.ventilation_load_df['TOTAL_CL']])
            chart.add_xy_data('Total Cooling Load',x_data,[i for i in self.zone.cooling_load_df['TOTAL_CL']],style_props={'marker': 's'})
        except KeyError:
            ...
        except:
            print("Something else went wrong")
        chart.x1.add_title('Local Standard Time (Hr)')
        chart.y1.add_title('Load, Watt')
        chart.y1.format_ticks(fmt='{x:,.0f}')
        chart.add_legend(anchor='upper left', position = (0.01, 0.99))
        chart.show()

    def RTSMConfig(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("ASHRAE Configuration")

        self.cmbConstruction = QComboBox()
        self.cmbConstruction.addItems(RTS.Room_construction)
        self.cmbConstruction.setCurrentText(Setting.NRTS_Room_construction)
        self.cmbCarpet = QComboBox()
        self.cmbCarpet.addItems(RTS.Carpet)
        self.cmbCarpet.setCurrentText(Setting.NRTS_Carpet)
        self.cmbGlass = QComboBox()
        self.cmbGlass.addItems(RTS.Glass)
        self.cmbGlass.setCurrentText(Setting.NRTS_Glass)

        formGroupBox = QGroupBox("Room Condition")
        layout = QFormLayout()
        layout.addRow("Construction:", self.cmbConstruction)
        layout.addRow("Carpet:", self.cmbCarpet)
        layout.addRow("Glass:", self.cmbGlass)
        formGroupBox.setLayout(layout)

        buttonBox = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttonBox.accepted.connect(dialog.accept)
        buttonBox.rejected.connect(dialog.reject)

        mainLayout = QVBoxLayout()
        mainLayout.addWidget(formGroupBox)
        mainLayout.addWidget(buttonBox)
        dialog.setLayout(mainLayout)

        if dialog.exec():
            Setting.NRTS_Room_construction = self.cmbConstruction.currentText()
            Setting.NRTS_Carpet = self.cmbCarpet.currentText()
            Setting.NRTS_Glass = self.cmbGlass.currentText()

    def RTSMOpen(self):
        fnames, filter = QFileDialog.getOpenFileNames(self, 'Open from file', '', 'RTSM (*.rtsm);;All files (*)')
        if fnames != [] and fnames[0]:
            with open(fnames[0], "r") as file:
                raw_data = file.read()
                try:
                    data = json.loads(raw_data)
                    ClimaticData = data['Climatic']
                    self.cmbRTSMLocation.setCurrentIndex(ClimaticData['Location'])
                    self.tbRTSMLatitude.setValue(ClimaticData['Latitude'])
                    self.tbRTSMLongtitude.setValue(ClimaticData['Longtitude'])
                    self.tbRTSMAltitude.setValue(ClimaticData['Altitude'])
                    self.tbRTSMOutsideDB.setValue(ClimaticData['OutsideDB'])
                    self.tbRTSMDBRange.setValue(ClimaticData['DBRange'])
                    self.tbRTSMOutsideWB.setValue(ClimaticData['OutsideWB'])
                    self.tbRTSMtaub.setValue(ClimaticData['taub'])
                    self.tbRTSMtaud.setValue(ClimaticData['taud'])
                    self.tbRTSMtz.setText(ClimaticData['tz'])
                    self.tbRTSMMonth.setText(ClimaticData['Month'])
                    self.tbRTSMName.setText(ClimaticData['Name'])
                    self.tbRTSMInsideDB.setValue(ClimaticData['InsideDB'])
                    self.tbRTSMInsideRH.setValue(ClimaticData['InsideRH'])
                    self.cmbRTSMSpaceType.setCurrentIndex(ClimaticData['SpaceType'])

                    ArchitecData = data['Architec']
                    self.tbCompass.setValue(ArchitecData['Compass'])
                    self.tbZoneHeight.setValue(ArchitecData['ZoneHeight'])
                    self.tbZoneLength.setValue(ArchitecData['ZoneLength'])
                    self.tbZoneWidth.setValue(ArchitecData['ZoneWidth'])
                    self.cmbRTSMRoofType.setCurrentIndex(ArchitecData['RoofType'])
                    self.tbRTSMURoof.setValue(ArchitecData['URoof'])
                    self.cmbRTSMFloorType.setCurrentIndex(ArchitecData['FloorType'])
                    self.tbRTSMUFloor.setValue(ArchitecData['UFloor'])
                    self.tbRTSMOptionFloor.setValue(ArchitecData['OptionFloor'])
                    self.cmbRTSMWall_A.setCurrentIndex(ArchitecData['Wall_A'])
                    self.tbRTSMUWall_A.setValue(ArchitecData['UWall_A'])
                    self.tbRTSMOptionWall_A.setValue(ArchitecData['OptionWall_A'])
                    self.cmbRTSMWall_B.setCurrentIndex(ArchitecData['Wall_B'])
                    self.tbRTSMUWall_B.setValue(ArchitecData['UWall_B'])
                    self.tbRTSMOptionWall_B.setValue(ArchitecData['OptionWall_B'])
                    self.cmbRTSMWall_C.setCurrentIndex(ArchitecData['Wall_C'])
                    self.tbRTSMUWall_C.setValue(ArchitecData['UWall_C'])
                    self.tbRTSMOptionWall_C.setValue(ArchitecData['OptionWall_C'])
                    self.cmbRTSMWall_D.setCurrentIndex(ArchitecData['Wall_D'])
                    self.tbRTSMUWall_D.setValue(ArchitecData['UWall_D'])
                    self.tbRTSMOptionWall_D.setValue(ArchitecData['OptionWall_D'])
                    self.tbWinAHigh.setValue(ArchitecData['WinAHigh'])
                    self.cmbRTSMWin_A.setCurrentIndex(ArchitecData['Win_A'])
                    self.tbWinAWidth.setValue(ArchitecData['WinAWidth'])
                    self.tbRTSMUWin_A.setValue(ArchitecData['UWin_A'])
                    self.tbRTSMSCWin_A.setValue(ArchitecData['SCWin_A'])
                    self.cmbRTSMInShaderWin_A.setCurrentIndex(ArchitecData['InShaderWin_A'])
                    self.cmbRTSMExShaderWin_A.setCurrentIndex(ArchitecData['ExShaderWin_A'])
                    self.tbWinBHigh.setValue(ArchitecData['WinBHigh'])
                    self.cmbRTSMWin_B.setCurrentIndex(ArchitecData['Win_B'])
                    self.tbWinBWidth.setValue(ArchitecData['WinBWidth'])
                    self.tbRTSMUWin_B.setValue(ArchitecData['UWin_B'])
                    self.tbRTSMSCWin_B.setValue(ArchitecData['SCWin_B'])
                    self.cmbRTSMInShaderWin_B.setCurrentIndex(ArchitecData['InShaderWin_B'])
                    self.cmbRTSMExShaderWin_B.setCurrentIndex(ArchitecData['ExShaderWin_B'])
                    self.tbWinCHigh.setValue(ArchitecData['WinCHigh'])
                    self.cmbRTSMWin_C.setCurrentIndex(ArchitecData['Win_C'])
                    self.tbWinCWidth.setValue(ArchitecData['WinCWidth'])
                    self.tbRTSMUWin_C.setValue(ArchitecData['UWin_C'])
                    self.tbRTSMSCWin_C.setValue(ArchitecData['SCWin_C'])
                    self.cmbRTSMInShaderWin_C.setCurrentIndex(ArchitecData['InShaderWin_C'])
                    self.cmbRTSMExShaderWin_C.setCurrentIndex(ArchitecData['ExShaderWin_C'])
                    self.tbWinDHigh.setValue(ArchitecData['WinDHigh'])
                    self.cmbRTSMWin_D.setCurrentIndex(ArchitecData['Win_D'])
                    self.tbWinDWidth.setValue(ArchitecData['WinDWidth'])
                    self.tbRTSMUWin_D.setValue(ArchitecData['UWin_D'])
                    self.tbRTSMSCWin_D.setValue(ArchitecData['SCWin_D'])
                    self.cmbRTSMInShaderWin_D.setCurrentIndex(ArchitecData['InShaderWin_D'])
                    self.cmbRTSMExShaderWin_D.setCurrentIndex(ArchitecData['ExShaderWin_D'])

                    ApplicationData = data['Application']
                    self.cmbRTSMRoomType.setCurrentIndex(ApplicationData['RoomType'])
                    self.tbRTSMLPD.setValue(ApplicationData['LPD'])
                    self.tbRTSMLightF_SPACE.setValue(ApplicationData['LightF_SPACE'])
                    self.tbRTSMLightF_RAD.setValue(ApplicationData['LightF_RAD'])
                    self.LightingGrid.setValue(ApplicationData['LightingGrid'])
                    self.cmbRTSMActivity.setCurrentIndex(ApplicationData['Activity'])
                    self.tbRTSMPeopleSH.setValue(ApplicationData['PeopleSH'])
                    self.tbRTSMPeopleLH.setValue(ApplicationData['PeopleLH'])
                    self.tbRTSMPeopleF_RAD.setValue(ApplicationData['PeopleF_RAD'])
                    self.PeopleGrid.setValue(ApplicationData['PeopleGrid'])
                    self.cmbRTSMVentilation.setCurrentIndex(ApplicationData['Ventilation'])
                    self.tbRTSMVentilation.setValue(ApplicationData['tbVentilation'])
                    self.tbRTSMRp.setValue(ApplicationData['Rp'])
                    self.tbRTSMRa.setValue(ApplicationData['Ra'])
                    self.tbRTSMEz.setValue(ApplicationData['Ez'])
                    self.cmbRTSMEquipment.setCurrentIndex(ApplicationData['Equipment'])
                    self.tbRTSMEquipmentSH.setValue(ApplicationData['EquipmentSH'])
                    self.tbRTSMEquipmentLH.setValue(ApplicationData['EquipmentLH'])
                    self.tbRTSMEquipmentNo.setValue(ApplicationData['EquipmentNo'])

                    self.tbRTSMSafety.setValue(data['Safety'])
                    self.cmbRTSMOutUnit.setCurrentIndex(data['OutUnit'])
                except json.JSONDecodeError:
                    print("%s is not a valid JSON file" % os.path.basename(fnames[0]))
                except Exception as e:
                    print(e)

    def RTSMSave(self):
        new_name = self.tbRTSMName.text()
        filename, filter = QFileDialog.getSaveFileName(self, 'Save to file', new_name, 'RTSM (*.rtsm);;All files (*)')
        if filename == '': return False
        with open(filename, "w") as file:
            ClimaticData = dict([
                ('Location', self.cmbRTSMLocation.currentIndex()),
                ('Latitude', self.tbRTSMLatitude.value()),
                ('Longtitude', self.tbRTSMLongtitude.value()),
                ('Altitude', self.tbRTSMAltitude.value()),
                ('OutsideDB', self.tbRTSMOutsideDB.value()),
                ('DBRange', self.tbRTSMDBRange.value()),
                ('OutsideWB', self.tbRTSMOutsideWB.value()),
                ('taub', self.tbRTSMtaub.value()),
                ('taud', self.tbRTSMtaud.value()),
                ('tz', self.tbRTSMtz.text()),
                ('Month', self.tbRTSMMonth.text()),
                ('Name', self.tbRTSMName.text()),
                ('InsideDB', self.tbRTSMInsideDB.value()),
                ('InsideRH', self.tbRTSMInsideRH.value()),
                ('SpaceType', self.cmbRTSMSpaceType.currentIndex()),])
            ArchitecData = dict([
                ('Compass', self.tbCompass.value()),
                ('ZoneHeight', self.tbZoneHeight.value()),
                ('ZoneLength', self.tbZoneLength.value()),
                ('ZoneWidth', self.tbZoneWidth.value()),

                ('RoofType', self.cmbRTSMRoofType.currentIndex()),
                ('URoof', self.tbRTSMURoof.value()),
                ('FloorType', self.cmbRTSMFloorType.currentIndex()),
                ('UFloor', self.tbRTSMUFloor.value()),
                ('OptionFloor', self.tbRTSMOptionFloor.value()),

                ('Wall_A', self.cmbRTSMWall_A.currentIndex()),
                ('UWall_A', self.tbRTSMUWall_A.value()),
                ('OptionWall_A', self.tbRTSMOptionWall_A.value()),
                ('Wall_B', self.cmbRTSMWall_B.currentIndex()),
                ('UWall_B', self.tbRTSMUWall_B.value()),
                ('OptionWall_B', self.tbRTSMOptionWall_B.value()),
                ('Wall_C', self.cmbRTSMWall_C.currentIndex()),
                ('UWall_C', self.tbRTSMUWall_C.value()),
                ('OptionWall_C', self.tbRTSMOptionWall_C.value()),
                ('Wall_D', self.cmbRTSMWall_D.currentIndex()),
                ('UWall_D', self.tbRTSMUWall_D.value()),
                ('OptionWall_D', self.tbRTSMOptionWall_D.value()),

                ('WinAHigh', self.tbWinAHigh.value()),
                ('Win_A', self.cmbRTSMWin_A.currentIndex()),
                ('WinAWidth', self.tbWinAWidth.value()),
                ('UWin_A', self.tbRTSMUWin_A.value()),
                ('SCWin_A', self.tbRTSMSCWin_A.value()),
                ('InShaderWin_A', self.cmbRTSMInShaderWin_A.currentIndex()),
                ('ExShaderWin_A', self.cmbRTSMExShaderWin_A.currentIndex()),
                ('WinBHigh', self.tbWinBHigh.value()),
                ('Win_B', self.cmbRTSMWin_B.currentIndex()),
                ('WinBWidth', self.tbWinBWidth.value()),
                ('UWin_B', self.tbRTSMUWin_B.value()),
                ('SCWin_B', self.tbRTSMSCWin_B.value()),
                ('InShaderWin_B', self.cmbRTSMInShaderWin_B.currentIndex()),
                ('ExShaderWin_B', self.cmbRTSMExShaderWin_B.currentIndex()),
                ('WinCHigh', self.tbWinCHigh.value()),
                ('Win_C', self.cmbRTSMWin_C.currentIndex()),
                ('WinCWidth', self.tbWinCWidth.value()),
                ('UWin_C', self.tbRTSMUWin_C.value()),
                ('SCWin_C', self.tbRTSMSCWin_C.value()),
                ('InShaderWin_C', self.cmbRTSMInShaderWin_C.currentIndex()),
                ('ExShaderWin_C', self.cmbRTSMExShaderWin_C.currentIndex()),
                ('WinDHigh', self.tbWinDHigh.value()),
                ('Win_D', self.cmbRTSMWin_D.currentIndex()),
                ('WinDWidth', self.tbWinDWidth.value()),
                ('UWin_D', self.tbRTSMUWin_D.value()),
                ('SCWin_D', self.tbRTSMSCWin_D.value()),
                ('InShaderWin_D', self.cmbRTSMInShaderWin_D.currentIndex()),
                ('ExShaderWin_D', self.cmbRTSMExShaderWin_D.currentIndex()),])
            ApplicationData = dict([
                ('RoomType', self.cmbRTSMRoomType.currentIndex()),
                ('LPD', self.tbRTSMLPD.value()),
                ('LightF_SPACE', self.tbRTSMLightF_SPACE.value()),
                ('LightF_RAD', self.tbRTSMLightF_RAD.value()),
                ('LightingGrid', self.LightingGrid.value()),
                ('Activity', self.cmbRTSMActivity.currentIndex()),
                ('PeopleSH', self.tbRTSMPeopleSH.value()),
                ('PeopleLH', self.tbRTSMPeopleLH.value()),
                ('PeopleF_RAD', self.tbRTSMPeopleF_RAD.value()),
                ('PeopleGrid', self.PeopleGrid.value()),
                ('Ventilation', self.cmbRTSMVentilation.currentIndex()),
                ('tbVentilation', self.tbRTSMVentilation.value()),
                ('Rp', self.tbRTSMRp.value()),
                ('Ra', self.tbRTSMRa.value()),
                ('Ez', self.tbRTSMEz.value()),
                ('Equipment', self.cmbRTSMEquipment.currentIndex()),
                ('EquipmentSH', self.tbRTSMEquipmentSH.value()),
                ('EquipmentLH', self.tbRTSMEquipmentLH.value()),
                ('EquipmentNo', self.tbRTSMEquipmentNo.value()),])
            data = dict([
                ('Climatic', ClimaticData),
                ('Architec', ArchitecData),
                ('Application', ApplicationData),
                ('Safety', self.tbRTSMSafety.value()),
                ('OutUnit', self.cmbRTSMOutUnit.currentIndex()),
                ])
            file.write( json.dumps( data, indent=4 ) )
 
    def About(self):
        QMessageBox.about(self, "Radiant Time Series (RTS) method Cooling Load",
                "Design cooling loads are based on the assumption of <b>steady-periodic conditions</b><br>"
                "Thus, the heat gain for a particular component at a particular hour is the same as<br>"
                "24 h prior, which is the same as 48 h prior, etc.<br>"
                "This assumption is the basis for the RTS derivation from the HB method.<br><br>"
                "For more information: <b>ASHRAE Handbook - Fundamentals, SI Edition</b><br>"
                "Copyright © 2011-2025 (\xa9) <a href='mailto:euttanal@betagro.com/'>L.Euttana</a>"
                ", <a href='https://github.com/hs0wkc/pyRTSM'>GitHub</a>"
                ", <a href='https://www.youtube.com/channel/UCh8nxggUtig4fX6qfsIIdvA'>Youtube</a><br><br>"
                )

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    wnd = Mainwindow()
    wnd.show()
    sys.exit(app.exec())
